/**
 * ============================================================================
 * AI_GATEWAY.GS — PAYMENT ALL
 * ============================================================================
 * Generic, permission-controlled bridge between AI and the Payment All workbook.
 *
 * READ tools:
 *   - AI may inspect any business sheet, schema, rows, search, and aggregate.
 *   - Sensitive credential columns are always blocked.
 *
 * SAFE actions:
 *   - syncFlow
 *   - syncCustomers
 *   - sync both
 *   - invoice preview
 *   - generate/regenerate invoice PDF
 *   - create a DRAFT manual payout (an intention; only a human can set it PAID)
 *
 * PROTECTED actions (PIN required; PIN is NEVER sent to OpenAI):
 *   - create invoice
 *   - change customer payout fee
 *
 * IMPORTANT:
 *   Existing accounting logic remains in Flow.gs / CustomerV2.3.gs / InvoicePdf.gs.
 *   This file does not reimplement accounting formulas.
 * ============================================================================
 */

const PAYMENT_AI_CFG = Object.freeze({
  HEADER_ROW: 2,
  DATA_START_ROW: 3,
  BUSINESS_TIMEZONE: 'America/New_York',
  MAX_READ_ROWS: 200,
  MAX_SEARCH_HITS: 120,
  MAX_TOOL_OUTPUT_CHARS: 85000,
  PENDING_ACTION_TTL_MS: 5 * 60 * 1000,
  MAX_PIN_ATTEMPTS: 5,

  PROP_SPREADSHEET_ID: 'PAYMENT_SPREADSHEET_ID',
  PROP_ACTION_PIN_HASH: 'PAYMENT_ACTION_PIN_HASH',
  PENDING_PREFIX: 'PAYMENT_AI_PENDING_ACTION_',

  SENSITIVE_HEADERS: [
    'Client ID',
    'Client Secret',
    'API Key',
    'API Token',
    'Access Token',
    'Refresh Token',
    'Password',
    'PIN',
    'Bot Token',
    'OpenAI API Key'
  ]
});


/************************* TOOL DEFINITIONS *************************/

function payAiToolDefinitions_() {
  const filterItem = {
    type: 'object',
    properties: {
      column: { type: 'string', description: 'Exact sheet header name.' },
      op: {
        type: 'string',
        enum: ['eq','neq','contains','starts_with','gt','gte','lt','lte','between','in','blank','not_blank','date_between'],
        description: 'Filter operator.'
      },
      value: { type: 'string', description: 'Primary comparison value. Use YYYY-MM-DD for dates. For in, comma-separate values. Use empty string when unused.' },
      value2: { type: 'string', description: 'Second value for between/date_between; otherwise empty string.' }
    },
    required: ['column','op','value','value2'],
    additionalProperties: false
  };

  const metricItem = {
    type: 'object',
    properties: {
      op: { type: 'string', enum: ['count','sum','avg','min','max'] },
      column: { type: 'string', description: 'Exact numeric header. For count of rows, use empty string.' },
      alias: { type: 'string', description: 'Short output label.' }
    },
    required: ['op','column','alias'],
    additionalProperties: false
  };

  return [
    {
      type: 'function',
      name: 'get_system_guide',
      description: 'Read the Payment All architecture/business-rule guide. Use when you need to understand which sheet owns a concept or how customer vs treasury logic works.',
      strict: true,
      parameters: { type: 'object', properties: {}, required: [], additionalProperties: false }
    },
    {
      type: 'function',
      name: 'list_sheets',
      description: 'List workbook sheets, row counts, and visible headers. Credential columns are omitted.',
      strict: true,
      parameters: { type: 'object', properties: {}, required: [], additionalProperties: false }
    },
    {
      type: 'function',
      name: 'get_sheet_schema',
      description: 'Get exact visible headers and row count for one sheet. Use before querying an unfamiliar sheet.',
      strict: true,
      parameters: {
        type: 'object',
        properties: { sheet: { type: 'string' } },
        required: ['sheet'],
        additionalProperties: false
      }
    },
    {
      type: 'function',
      name: 'read_sheet_rows',
      description: 'Read filtered rows from one sheet. Can paginate and sort. Prefer filters and only needed columns instead of pulling an entire large sheet at once.',
      strict: true,
      parameters: {
        type: 'object',
        properties: {
          sheet: { type: 'string' },
          columns: { type: 'array', items: { type: 'string' }, description: 'Exact headers to return. Empty array means all non-sensitive columns.' },
          filters: { type: 'array', items: filterItem },
          limit: { type: 'integer', minimum: 1, maximum: 200 },
          offset: { type: 'integer', minimum: 0 },
          sort_by: { type: 'string', description: 'Exact header to sort by, or empty string.' },
          sort_direction: { type: 'string', enum: ['asc','desc','none'] }
        },
        required: ['sheet','columns','filters','limit','offset','sort_by','sort_direction'],
        additionalProperties: false
      }
    },
    {
      type: 'function',
      name: 'aggregate_sheet',
      description: 'Calculate count/sum/avg/min/max across ALL matching rows in Apps Script, so the AI does not have to manually add money. Supports one optional group_by field.',
      strict: true,
      parameters: {
        type: 'object',
        properties: {
          sheet: { type: 'string' },
          filters: { type: 'array', items: filterItem },
          metrics: { type: 'array', minItems: 1, items: metricItem },
          group_by: { type: 'string', description: 'Exact header to group by, or empty string.' }
        },
        required: ['sheet','filters','metrics','group_by'],
        additionalProperties: false
      }
    },
    {
      type: 'function',
      name: 'search_workbook',
      description: 'Search visible business data across selected sheets (or all sheets) for a phrase such as transaction ID, invoice ID, customer name, prefix, reference, or key.',
      strict: true,
      parameters: {
        type: 'object',
        properties: {
          query: { type: 'string' },
          sheets: { type: 'array', items: { type: 'string' }, description: 'Empty array means all sheets.' },
          limit: { type: 'integer', minimum: 1, maximum: 120 }
        },
        required: ['query','sheets','limit'],
        additionalProperties: false
      }
    },
    {
      type: 'function',
      name: 'run_safe_action',
      description: 'Run an approved non-destructive operational action. Invoice preview is safe. PDF generation changes PDF metadata only. create_manual_payout writes a DRAFT payout row for money the owner is about to spend from a treasury account: it is an intention, not a payment — no balance changes and no transaction exists until the owner sets Status to PAID by hand. Use empty strings and 0 for parameters not used by the selected action.',
      strict: true,
      parameters: {
        type: 'object',
        properties: {
          action: { type: 'string', enum: ['sync_flow','sync_customers','sync_both','preview_invoice','generate_invoice_pdf','validate_customer_system','create_manual_payout'] },
          customer_key: { type: 'string' },
          from_date: { type: 'string', description: 'YYYY-MM-DD or DD/MM/YYYY; empty when unused.' },
          to_date: { type: 'string', description: 'YYYY-MM-DD or DD/MM/YYYY; empty when unused.' },
          as_of_date: { type: 'string', description: 'YYYY-MM-DD or DD/MM/YYYY; empty means today for preview.' },
          settlement_key: { type: 'string' },
          force: { type: 'boolean' },
          advance_amount: { type: 'number', description: 'preview_invoice only: money to advance to the customer on this invoice, on top of what they earned. It increases Amount Due now and is recovered from later invoices. Use 0 for a normal invoice.' },
          advance_note: { type: 'string', description: 'preview_invoice only: why the advance is being given. Empty otherwise.' },
          provider: { type: 'string', description: 'create_manual_payout only: USDT or PINGPONG. Empty otherwise.' },
          account_key: { type: 'string', description: 'create_manual_payout only: which treasury account the money will leave, e.g. USDT1. Must already exist and not be disabled. Empty otherwise.' },
          amount: { type: 'number', description: 'create_manual_payout only: amount to pay out, greater than 0. Use 0 otherwise.' },
          fee: { type: 'number', description: 'create_manual_payout only: known fee, 0 if none.' },
          payee: { type: 'string', description: 'create_manual_payout only: who or what is being paid. Empty otherwise.' },
          category: { type: 'string', description: 'create_manual_payout only: OWNER_DRAW (the owner taking their own service-fee earnings), EXPENSE, REFUND, TRANSFER, ADJUSTMENT or OTHER. Never CUSTOMER_PAYOUT — invoices own that. Empty otherwise.' },
          note: { type: 'string', description: 'create_manual_payout only: what this payment is for, in the owner\'s own words. Empty otherwise.' }
        },
        required: ['action','customer_key','from_date','to_date','as_of_date','settlement_key','force','advance_amount','advance_note','provider','account_key','amount','fee','payee','category','note'],
        additionalProperties: false
      }
    },
    {
      type: 'function',
      name: 'request_protected_action',
      description: 'Prepare a financial/data-changing action that requires the owner PIN. This NEVER executes the action immediately. It validates/previews, stores a 5-minute pending action, and asks the user to send /pin <PIN>. The PIN is handled locally and never sent to the AI.',
      strict: true,
      parameters: {
        type: 'object',
        properties: {
          action: { type: 'string', enum: ['create_invoice','set_customer_payout_fee'] },
          customer_key: { type: 'string' },
          from_date: { type: 'string' },
          to_date: { type: 'string' },
          as_of_date: { type: 'string' },
          advance_amount: { type: 'number', description: 'create_invoice only: money to advance to the customer on this invoice, recovered from later invoices. Use 0 unless the owner explicitly asked for an advance.' },
          advance_note: { type: 'string', description: 'create_invoice only: why the advance is being given. Empty otherwise.' },
          payout_fee_percent: { type: 'number', minimum: 0, maximum: 100 }
        },
        required: ['action','customer_key','from_date','to_date','as_of_date','advance_amount','advance_note','payout_fee_percent'],
        additionalProperties: false
      }
    }
  ];
}


/************************* TOOL DISPATCH *************************/

function payAiDispatchTool_(name, args, ctx) {
  try {
    switch (name) {
      case 'get_system_guide':
        return payAiToolResult_(true, { guide: payAiRuntimeGuide_() });
      case 'list_sheets':
        return payAiToolResult_(true, payAiListSheets_());
      case 'get_sheet_schema':
        return payAiToolResult_(true, payAiGetSheetSchema_(args.sheet));
      case 'read_sheet_rows':
        return payAiToolResult_(true, payAiReadSheetRows_(args));
      case 'aggregate_sheet':
        return payAiToolResult_(true, payAiAggregateSheet_(args));
      case 'search_workbook':
        return payAiToolResult_(true, payAiSearchWorkbook_(args));
      case 'run_safe_action':
        return payAiToolResult_(true, payAiRunSafeAction_(args));
      case 'request_protected_action':
        return payAiToolResult_(true, payAiRequestProtectedAction_(ctx, args));
      default:
        return payAiToolResult_(false, null, 'Unknown tool: ' + name);
    }
  } catch (err) {
    return payAiToolResult_(false, null, err && err.message ? err.message : String(err));
  }
}

function payAiToolResult_(ok, data, error) {
  const out = { ok: !!ok };
  if (ok) out.data = data;
  else out.error = error || 'Unknown error';
  return out;
}


/************************* SPREADSHEET CONTEXT *************************/

function payAiSpreadsheet_() {
  const props = PropertiesService.getScriptProperties();
  let id = String(props.getProperty(PAYMENT_AI_CFG.PROP_SPREADSHEET_ID) || '').trim();

  if (!id) {
    const active = SpreadsheetApp.getActive();
    if (!active) throw new Error('PAYMENT_SPREADSHEET_ID is not configured. Run setupPaymentAi() from the Payment All spreadsheet first.');
    id = active.getId();
    props.setProperty(PAYMENT_AI_CFG.PROP_SPREADSHEET_ID, id);
  }

  return SpreadsheetApp.openById(id);
}

function payAiActivateSpreadsheet_() {
  const ss = payAiSpreadsheet_();
  SpreadsheetApp.setActiveSpreadsheet(ss);
  return ss;
}

function payAiSheet_(name) {
  const ss = payAiSpreadsheet_();
  const sh = ss.getSheetByName(String(name || '').trim());
  if (!sh) throw new Error('Sheet not found: ' + name);
  return sh;
}

function payAiHeaderMap_(sh) {
  if (!sh || sh.getLastColumn() < 1) return {};
  const headers = sh.getRange(PAYMENT_AI_CFG.HEADER_ROW, 1, 1, sh.getLastColumn()).getDisplayValues()[0];
  const map = {};
  headers.forEach(function(h, i) {
    h = String(h || '').trim();
    if (h) map[h] = i + 1;
  });
  return map;
}

function payAiVisibleHeaders_(sh) {
  const map = payAiHeaderMap_(sh);
  return Object.keys(map).filter(function(h) { return !payAiIsSensitiveHeader_(h); });
}

function payAiIsSensitiveHeader_(header) {
  const h = String(header || '').trim().toLowerCase();
  if (!h) return false;
  if (PAYMENT_AI_CFG.SENSITIVE_HEADERS.some(function(x) { return h === x.toLowerCase(); })) return true;
  return /(client\s*secret|api\s*key|api\s*token|access\s*token|refresh\s*token|password|passcode|\bpin\b|bot\s*token|secret\s*key)/i.test(h);
}

function payAiAssertVisibleHeader_(map, header) {
  const h = String(header || '').trim();
  if (!h || !map[h]) throw new Error('Header not found: ' + h);
  if (payAiIsSensitiveHeader_(h)) throw new Error('Sensitive header is not available to AI: ' + h);
  return h;
}


/************************* READ / SEARCH *************************/

function payAiListSheets_() {
  const ss = payAiSpreadsheet_();
  return ss.getSheets().map(function(sh) {
    return {
      sheet: sh.getName(),
      data_rows: Math.max(0, sh.getLastRow() - PAYMENT_AI_CFG.DATA_START_ROW + 1),
      headers: payAiVisibleHeaders_(sh)
    };
  });
}

function payAiGetSheetSchema_(sheetName) {
  const sh = payAiSheet_(sheetName);
  return {
    sheet: sh.getName(),
    header_row: PAYMENT_AI_CFG.HEADER_ROW,
    data_start_row: PAYMENT_AI_CFG.DATA_START_ROW,
    data_rows: Math.max(0, sh.getLastRow() - PAYMENT_AI_CFG.DATA_START_ROW + 1),
    headers: payAiVisibleHeaders_(sh)
  };
}

function payAiReadSheetRows_(args) {
  const sh = payAiSheet_(args.sheet);
  const map = payAiHeaderMap_(sh);
  const visible = payAiVisibleHeaders_(sh);

  let columns = Array.isArray(args.columns) ? args.columns.map(String) : [];
  if (!columns.length) columns = visible.slice();
  columns = columns.map(function(c) { return payAiAssertVisibleHeader_(map, c); });

  const filters = Array.isArray(args.filters) ? args.filters : [];
  filters.forEach(function(f) { payAiAssertVisibleHeader_(map, f.column); });

  const sortBy = String(args.sort_by || '').trim();
  if (sortBy) payAiAssertVisibleHeader_(map, sortBy);

  const rows = payAiReadObjects_(sh, map).filter(function(r) {
    return payAiRowMatchesFilters_(r, filters);
  });

  if (sortBy && args.sort_direction !== 'none') {
    const dir = args.sort_direction === 'desc' ? -1 : 1;
    rows.sort(function(a, b) { return dir * payAiCompareValues_(a[sortBy], b[sortBy]); });
  }

  const offset = Math.max(0, Number(args.offset) || 0);
  const limit = Math.min(PAYMENT_AI_CFG.MAX_READ_ROWS, Math.max(1, Number(args.limit) || 50));
  const page = rows.slice(offset, offset + limit);

  const data = page.map(function(r) {
    const out = { _row: r._row };
    columns.forEach(function(c) { out[c] = payAiJsonValue_(r[c]); });
    return out;
  });

  return {
    sheet: sh.getName(),
    matched_rows: rows.length,
    offset: offset,
    returned_rows: data.length,
    has_more: offset + data.length < rows.length,
    columns: columns,
    rows: data
  };
}

function payAiAggregateSheet_(args) {
  const sh = payAiSheet_(args.sheet);
  const map = payAiHeaderMap_(sh);
  const filters = Array.isArray(args.filters) ? args.filters : [];
  filters.forEach(function(f) { payAiAssertVisibleHeader_(map, f.column); });

  const metrics = Array.isArray(args.metrics) ? args.metrics : [];
  if (!metrics.length) throw new Error('At least one metric is required.');
  metrics.forEach(function(m) {
    if (m.op !== 'count' || String(m.column || '').trim()) payAiAssertVisibleHeader_(map, m.column);
  });

  const groupBy = String(args.group_by || '').trim();
  if (groupBy) payAiAssertVisibleHeader_(map, groupBy);

  const rows = payAiReadObjects_(sh, map).filter(function(r) {
    return payAiRowMatchesFilters_(r, filters);
  });

  const groups = {};
  rows.forEach(function(r) {
    const key = groupBy ? String(payAiJsonValue_(r[groupBy])) : '__ALL__';
    if (!groups[key]) groups[key] = [];
    groups[key].push(r);
  });

  if (!rows.length && !groupBy) groups.__ALL__ = [];

  const result = Object.keys(groups).map(function(groupKey) {
    const rs = groups[groupKey];
    const out = {};
    if (groupBy) out[groupBy] = groupKey;

    metrics.forEach(function(m) {
      const alias = String(m.alias || '').trim() || (m.op + (m.column ? '_' + m.column : ''));
      const values = m.column ? rs.map(function(r) { return payAiToNumber_(r[m.column]); }).filter(function(n) { return n !== null; }) : [];

      if (m.op === 'count') out[alias] = m.column ? values.length : rs.length;
      else if (m.op === 'sum') out[alias] = payAiRound_(values.reduce(function(a, b) { return a + b; }, 0), 6);
      else if (m.op === 'avg') out[alias] = values.length ? payAiRound_(values.reduce(function(a, b) { return a + b; }, 0) / values.length, 6) : 0;
      else if (m.op === 'min') out[alias] = values.length ? Math.min.apply(null, values) : 0;
      else if (m.op === 'max') out[alias] = values.length ? Math.max.apply(null, values) : 0;
    });

    return out;
  });

  return {
    sheet: sh.getName(),
    matched_rows: rows.length,
    group_by: groupBy,
    results: result
  };
}

function payAiSearchWorkbook_(args) {
  const query = String(args.query || '').trim();
  if (!query) throw new Error('Search query is required.');
  const q = query.toLowerCase();
  const ss = payAiSpreadsheet_();
  const requested = Array.isArray(args.sheets) ? args.sheets.map(String).filter(Boolean) : [];
  const limit = Math.min(PAYMENT_AI_CFG.MAX_SEARCH_HITS, Math.max(1, Number(args.limit) || 50));
  const hits = [];

  const sheets = requested.length
    ? requested.map(function(n) { const s = ss.getSheetByName(n); if (!s) throw new Error('Sheet not found: ' + n); return s; })
    : ss.getSheets();

  for (let s = 0; s < sheets.length && hits.length < limit; s++) {
    const sh = sheets[s];
    const map = payAiHeaderMap_(sh);
    const headers = payAiVisibleHeaders_(sh);
    const rows = payAiReadObjects_(sh, map);

    for (let i = 0; i < rows.length && hits.length < limit; i++) {
      const row = rows[i];
      const matchedCols = [];
      headers.forEach(function(h) {
        const value = payAiJsonValue_(row[h]);
        if (String(value === null || value === undefined ? '' : value).toLowerCase().indexOf(q) !== -1) matchedCols.push(h);
      });
      if (!matchedCols.length) continue;

      const preview = { _row: row._row };
      matchedCols.slice(0, 5).forEach(function(h) { preview[h] = payAiJsonValue_(row[h]); });
      ['Customer Key','Transaction Key','Transaction ID','PayPal Transaction ID','Invoice ID','Invoice Number','Settlement Key','Payout Key','Account Key','Label'].forEach(function(h) {
        if (map[h] && !payAiIsSensitiveHeader_(h) && row[h] !== '' && row[h] !== null && row[h] !== undefined) preview[h] = payAiJsonValue_(row[h]);
      });

      hits.push({ sheet: sh.getName(), matched_columns: matchedCols.slice(0, 10), row: preview });
    }
  }

  return { query: query, returned_hits: hits.length, truncated: hits.length >= limit, hits: hits };
}

function payAiReadObjects_(sh, map) {
  const lastRow = sh.getLastRow();
  if (lastRow < PAYMENT_AI_CFG.DATA_START_ROW) return [];
  const lastCol = sh.getLastColumn();
  const values = sh.getRange(PAYMENT_AI_CFG.DATA_START_ROW, 1, lastRow - PAYMENT_AI_CFG.DATA_START_ROW + 1, lastCol).getValues();
  const byIndex = {};
  Object.keys(map).forEach(function(h) { byIndex[map[h] - 1] = h; });

  return values.map(function(row, i) {
    const o = { _row: PAYMENT_AI_CFG.DATA_START_ROW + i };
    row.forEach(function(v, j) {
      const h = byIndex[j];
      if (h && !payAiIsSensitiveHeader_(h)) o[h] = v;
    });
    return o;
  });
}


/************************* FILTERING *************************/

function payAiRowMatchesFilters_(row, filters) {
  return (filters || []).every(function(f) {
    const actual = row[String(f.column || '').trim()];
    const op = String(f.op || 'eq');
    const v1 = String(f.value === null || f.value === undefined ? '' : f.value).trim();
    const v2 = String(f.value2 === null || f.value2 === undefined ? '' : f.value2).trim();

    if (op === 'blank') return actual === '' || actual === null || actual === undefined;
    if (op === 'not_blank') return !(actual === '' || actual === null || actual === undefined);

    if (op === 'contains') return String(payAiJsonValue_(actual)).toLowerCase().indexOf(v1.toLowerCase()) !== -1;
    if (op === 'starts_with') return String(payAiJsonValue_(actual)).toLowerCase().indexOf(v1.toLowerCase()) === 0;
    if (op === 'in') {
      const set = v1.split(',').map(function(x) { return x.trim().toLowerCase(); }).filter(Boolean);
      return set.indexOf(String(payAiJsonValue_(actual)).trim().toLowerCase()) !== -1;
    }
    if (op === 'date_between') {
      const a = payAiDateKey_(actual), b = payAiDateKey_(v1), c = payAiDateKey_(v2);
      return !!a && !!b && !!c && a >= b && a <= c;
    }

    const actualNum = payAiToNumber_(actual);
    const v1Num = payAiToNumber_(v1);
    const v2Num = payAiToNumber_(v2);
    const numeric = actualNum !== null && v1Num !== null;

    if (op === 'between') return numeric && v2Num !== null && actualNum >= v1Num && actualNum <= v2Num;
    if (op === 'gt') return numeric ? actualNum > v1Num : payAiCompareValues_(actual, v1) > 0;
    if (op === 'gte') return numeric ? actualNum >= v1Num : payAiCompareValues_(actual, v1) >= 0;
    if (op === 'lt') return numeric ? actualNum < v1Num : payAiCompareValues_(actual, v1) < 0;
    if (op === 'lte') return numeric ? actualNum <= v1Num : payAiCompareValues_(actual, v1) <= 0;

    const aText = String(payAiJsonValue_(actual)).trim().toLowerCase();
    const bText = v1.toLowerCase();
    if (op === 'neq') return aText !== bText;
    return aText === bText;
  });
}

function payAiCompareValues_(a, b) {
  if (a instanceof Date || b instanceof Date) {
    const ad = payAiDateMs_(a), bd = payAiDateMs_(b);
    if (ad !== null && bd !== null) return ad === bd ? 0 : (ad < bd ? -1 : 1);
  }
  const an = payAiToNumber_(a), bn = payAiToNumber_(b);
  if (an !== null && bn !== null) return an === bn ? 0 : (an < bn ? -1 : 1);
  const as = String(payAiJsonValue_(a)).toLowerCase(), bs = String(payAiJsonValue_(b)).toLowerCase();
  return as === bs ? 0 : (as < bs ? -1 : 1);
}


/************************* SAFE ACTIONS *************************/

function payAiRunSafeAction_(args) {
  payAiActivateSpreadsheet_();
  const action = String(args.action || '');

  if (action === 'sync_flow') {
    payAiRequireFunction_('syncFlow');
    const result = syncFlow();
    return { action: action, result: payAiJsonSafe_(result), message: 'Flow sync completed.' };
  }

  if (action === 'sync_customers') {
    payAiRequireFunction_('syncCustomers');
    const result = syncCustomers();
    return { action: action, result: payAiJsonSafe_(result), message: 'Customer sync completed.' };
  }

  if (action === 'sync_both') {
    payAiRequireFunction_('syncFlow');
    payAiRequireFunction_('syncCustomers');
    const flow = syncFlow();
    const customers = syncCustomers();
    return { action: action, flow: payAiJsonSafe_(flow), customers: payAiJsonSafe_(customers), message: 'Flow + customer sync completed.' };
  }

  if (action === 'preview_invoice') {
    payAiRequireFunction_('previewCustomerInvoice');
    const preview = previewCustomerInvoice(payAiInvoiceParamsFromArgs_(args));
    return payAiSummarizePreview_(preview);
  }

  if (action === 'generate_invoice_pdf') {
    if (!String(args.settlement_key || '').trim()) throw new Error('settlement_key is required.');
    if (typeof invPdfGenerate_ !== 'function') throw new Error('InvoicePdf function invPdfGenerate_() is not available.');
    const result = invPdfGenerate_(String(args.settlement_key).trim(), !!args.force);
    return { action: action, result: payAiJsonSafe_(result) };
  }

  if (action === 'validate_customer_system') {
    payAiRequireFunction_('validateCustomerSystem');
    return { action: action, result: payAiJsonSafe_(validateCustomerSystem()) };
  }

  if (action === 'create_manual_payout') {
    return payAiCreateManualPayout_(args);
  }

  throw new Error('Unsupported safe action: ' + action);
}


/************************* MANUAL PAYOUT *************************/

/*
 * Chỉ những loại này. CUSTOMER_PAYOUT cố tình vắng mặt: đó là lãnh địa của
 * engine hoá đơn, nơi mỗi đồng phải truy được về allocation đã tiêu. Một dòng
 * tay đội lốt payout khách sẽ trốn được cổng settlement của Flow.
 *
 * Danh sách gốc nằm ở flowPayoutManualCategories_ trong Flow.gs, dùng chung
 * với dropdown trên sheet. Đọc qua HÀM chứ không giữ một bản sao: hai danh
 * sách rời nhau thì sớm muộn cũng lệch, và lệch ở đây nghĩa là AI ghi được
 * giá trị mà dropdown không cho phép.
 */
function payAiManualCategories_() {
  return flowPayoutManualCategories_();
}


/*
 * Lệnh chi viết tay, luôn dừng ở DRAFT.
 *
 * Vì sao không cần PIN: DRAFT không phải giao dịch, nó là ý định. Flow loại
 * DRAFT khỏi mọi số dư (flowTransactionCounts_), và không có transaction nào
 * được sinh ra cho tới khi có người đổi Status thành PAID trong sheet. Nên thứ
 * hàm này tạo ra là một tờ giấy chờ ký, không phải tiền rời đi.
 *
 * Điều đó chỉ đúng chừng nào các bất biến dưới đây còn nguyên. Chúng không
 * phải kiểm tra phòng hờ — chúng là toàn bộ lý do action này được xếp loại an
 * toàn. Ai đó thêm đường cho AI ghi Status khác DRAFT là phá vỡ lập luận này.
 */
function payAiCreateManualPayout_(args) {
  const ss = payAiSpreadsheet_();
  payAiActivateSpreadsheet_();

  const provider = String(args.provider || '').trim().toUpperCase();
  if (provider !== 'USDT' && provider !== 'PINGPONG') {
    throw new Error('provider must be USDT or PINGPONG (Flow only settles those). Got: ' + (provider || '(empty)'));
  }

  const category = String(args.category || 'EXPENSE').trim().toUpperCase();
  if (category === 'CUSTOMER_PAYOUT') {
    throw new Error('CUSTOMER_PAYOUT belongs to the invoice engine. Create an invoice instead; a manual row cannot carry a settlement.');
  }
  if (payAiManualCategories_().indexOf(category) === -1) {
    throw new Error('category must be one of: ' + payAiManualCategories_().join(', ') + '. Got: ' + category);
  }

  const amount = payAiToNumber_(args.amount);
  if (amount === null || !(amount > 0)) throw new Error('amount must be greater than 0.');

  const feeRaw = payAiToNumber_(args.fee);
  const fee = feeRaw === null ? 0 : feeRaw;
  if (fee < 0) throw new Error('fee cannot be negative.');

  const payee = String(args.payee || '').trim();
  if (!payee) throw new Error('payee is required: say who or what is being paid.');

  const note = String(args.note || '').trim();
  if (!note) throw new Error('note is required: say what this payment is for.');

  /*
   * Tài khoản phải có thật và đang bật. Kiểm ở đây để hỏng thì hỏng ngay lúc
   * hỏi, thay vì đẻ ra một dòng mà lần sync sau mới dán nhãn ACCOUNT NOT FOUND
   * rồi nằm đó chờ ai đó tình cờ nhìn thấy.
   */
  const accountKey = String(args.account_key || '').trim().toUpperCase();
  if (!accountKey) throw new Error('account_key is required: which treasury account does the money leave from?');
  const known = payAiTreasuryAccountKeys_(ss, provider);
  if (known.indexOf(accountKey) === -1) {
    throw new Error('Account not found or disabled on the ' + provider + ' sheet: ' + accountKey +
      '. Enabled accounts: ' + (known.length ? known.join(', ') : '(none)'));
  }

  const sh = ss.getSheetByName('Payouts');
  if (!sh) throw new Error('Missing Payouts sheet.');
  if (typeof custHeaderMap_ !== 'function' || typeof custWritePayoutDraftRow_ !== 'function' ||
      typeof custNextKey_ !== 'function') {
    throw new Error('Customer module is not loaded; manual payout needs its Payouts helpers.');
  }
  const h = custHeaderMap_(sh);
  ['Payout Key','Provider','Account Key','Payee','Category','Amount','Fee','Currency','Status','Note']
    .forEach(function(k) { if (!h[k]) throw new Error('Payouts sheet missing header: ' + k); });

  const now = new Date();
  const payoutKey = custNextKey_(sh, h, 'Payout Key', 'PO-', 6);
  const currency = provider === 'USDT' ? 'USDT' : 'USD';

  const obj = {
    'Payout Key': payoutKey,
    'Payout Date': '',
    'Provider': provider,
    'Account Key': accountKey,
    'Payee': payee,
    'Category': category,
    'Amount': amount,
    'Fee': fee,
    'Total Deducted': payAiRound_(amount + fee, 2),
    'Currency': currency,
    'Reference': '',
    /* Bất biến 1: không bao giờ là gì khác DRAFT. */
    'Status': 'DRAFT',
    'Validation': 'DRAFT',
    'Transaction Key': '',
    'Created At': now,
    'Paid At': '',
    'Note': note + ' | MANUAL VIA AI ' + Utilities.formatDate(now, 'Asia/Ho_Chi_Minh', 'yyyy-MM-dd HH:mm'),
    /* Bất biến 2: không đội lốt được payout khách. */
    'Customer Key': '',
    'Settlement Key': ''
  };

  const row = custWritePayoutDraftRow_(ss, sh, h, obj);
  SpreadsheetApp.flush();

  /*
   * Cố tình KHÔNG tạo dòng Transaction. Flow sở hữu sheet đó và sẽ dựng
   * POTX|<payoutKey> đúng cách vào lúc Status thành PAID. Viết trước một dòng
   * DRAFT ở đây là đoán trước ngữ nghĩa của Flow cho một khoản chi mà nó chưa
   * từng thấy.
   */
  return {
    action: 'create_manual_payout',
    payout_key: payoutKey,
    row: row,
    provider: provider,
    account_key: accountKey,
    payee: payee,
    category: category,
    amount: amount,
    fee: fee,
    total_deducted: obj['Total Deducted'],
    currency: currency,
    status: 'DRAFT',
    message: 'DRAFT payout ' + payoutKey + ' created on row ' + row + '. Nothing has moved and no balance changed: ' +
      'it stays an intention until the owner sets Status to PAID on the Payouts sheet. ' +
      'Report the payout key and say plainly that it is still DRAFT.'
  };
}


/*
 * Khoá tài khoản đang bật của một provider chi tiền. Sheet Pingpong viết hoa
 * kiểu khác USDT, nên tra tên sheet chứ đừng suy ra từ tên provider.
 */
function payAiTreasuryAccountKeys_(ss, provider) {
  const sheetName = provider === 'USDT' ? 'USDT' : (provider === 'PINGPONG' ? 'Pingpong' : '');
  const sh = sheetName ? ss.getSheetByName(sheetName) : null;
  if (!sh) return [];

  const map = payAiHeaderMap_(sh);
  if (!map['Account Key']) return [];

  const out = [];
  payAiReadObjects_(sh, map).forEach(function(r) {
    const key = String(r['Account Key'] || '').trim().toUpperCase();
    const status = map['Status'] ? String(r['Status'] || '').trim().toUpperCase() : '';
    if (key && status !== 'DISABLED') out.push(key);
  });
  return out;
}

function payAiInvoiceParamsFromArgs_(args) {
  const out = {
    customer_key: String(args.customer_key || '').trim().toUpperCase(),
    from_date: String(args.from_date || '').trim(),
    to_date: String(args.to_date || '').trim(),
    as_of_date: String(args.as_of_date || '').trim()
  };

  /*
   * Chỉ gắn khi thật sự có số. Để advance_amount = 0 lọt vào cũng vô hại, nhưng
   * bỏ hẳn thì các đường gọi cũ giữ nguyên hình dạng params như trước.
   *
   * Số này đi thẳng vào fingerprint bên Customer.gs, nên preview và create bắt
   * buộc thấy cùng một con số — không có cách nào duyệt 0 rồi tạo 1000.
   */
  const advance = payAiToNumber_(args.advance_amount);
  if (advance !== null && advance > 0) {
    out.advance_amount = advance;
    out.advance_note = String(args.advance_note || '').trim();
  }

  return out;
}

function payAiSummarizePreview_(p) {
  if (!p) return { action: 'preview_invoice', message: 'Preview completed.', preview: null };
  return {
    action: 'preview_invoice',
    /*
     * Dấu vân tay của chính bộ số dưới đây. Nó đi kèm ra tới Telegram và được
     * cất vào pending action, để lúc gõ PIN còn có thứ mà đối chiếu: cái được
     * tạo phải đúng là cái người ta đã nhìn.
     */
    fingerprint: p.fingerprint,
    customer_key: p.customer && p.customer.customerKey,
    customer_name: p.customer && p.customer.customerName,
    period_from: p.params && p.params.fromKey,
    period_to: p.params && p.params.toKey,
    as_of: p.params && p.params.asOfKey,
    eligible_transactions: p.items ? p.items.length : (p.eligibleTransactions || 0),
    gross_sales: p.grossSales,
    refunds: p.refunds,
    paypal_fees: p.paypalFees,
    receiving_fees: p.receivingFees,
    available_amount: p.availableAmount,
    conversion_rate: p.conversionRate,
    conversion_cost: p.conversionCost,
    transfer_rate: p.transferRate,
    transfer_minimum: p.transferMinimum,
    transfer_cost: p.transferCost,
    downstream_cost: p.downstreamCost,
    net_before_payout_fee: p.netBeforePayoutFee,
    payout_fee_rate: p.payoutFeeRate,
    payout_fee_amount: p.payoutFeeAmount,
    amount_due: p.amountDue,
    source_currency: p.sourceCurrency,
    payout_currency: p.payoutCurrency,
    carry_forward_amount: p.carryForwardAmount,
    late_refund_applied: p.lateRefundApplied,
    late_refund_carry_forward: p.lateRefundCarryForward,
    advance_given: p.advanceGiven,
    advance_recovered: p.advanceRecovered,
    advance_outstanding_before: p.advanceOutstandingBefore,
    advance_outstanding_after: p.advanceOutstandingAfter,
    message: 'Invoice preview completed. No invoice was created.'
  };
}


/************************* PROTECTED ACTIONS *************************/

function payAiRequestProtectedAction_(ctx, args) {
  if (!ctx || !ctx.chatId) throw new Error('Protected actions require a Telegram chat context.');
  const action = String(args.action || '');
  let previewInfo = null;

  payAiActivateSpreadsheet_();

  if (action === 'create_invoice') {
    payAiRequireFunction_('previewCustomerInvoice');
    previewInfo = payAiSummarizePreview_(previewCustomerInvoice(payAiInvoiceParamsFromArgs_(args)));
  } else if (action === 'set_customer_payout_fee') {
    const current = payAiGetCustomerRow_(args.customer_key);
    previewInfo = {
      customer_key: current['Customer Key'],
      customer_name: current['Customer Name'],
      current_payout_fee_percent: payAiRound_((payAiToNumber_(current['Payout Fee %']) || 0) * 100, 4),
      new_payout_fee_percent: Number(args.payout_fee_percent)
    };
  } else {
    throw new Error('Unsupported protected action: ' + action);
  }

  const pending = {
    action: action,
    args: args,
    preview: previewInfo,
    /*
     * PIN xác nhận Ý ĐỊNH, không xác nhận SỐ TIỀN: giữa lúc xem preview và lúc
     * gõ PIN có thể trôi qua vài phút, mà createCustomerInvoice() gọi
     * syncCustomers_() trước khi dựng lại preview — và trigger giờ đồng bộ hai
     * lần mỗi ngày. Dữ liệu PayPal mới rơi vào khoảng đó là hoá đơn được tạo
     * khác với hoá đơn đã được duyệt, im lặng.
     *
     * Cất vân tay lúc duyệt để lúc thực thi có thể từ chối thay vì tạo bừa.
     */
    fingerprint: previewInfo && previewInfo.fingerprint ? String(previewInfo.fingerprint) : '',
    createdAt: Date.now(),
    expiresAt: Date.now() + PAYMENT_AI_CFG.PENDING_ACTION_TTL_MS,
    attempts: 0
  };

  PropertiesService.getScriptProperties().setProperty(
    PAYMENT_AI_CFG.PENDING_PREFIX + String(ctx.chatId),
    JSON.stringify(pending)
  );

  return {
    status: 'PIN_REQUIRED',
    action: action,
    expires_in_minutes: Math.round(PAYMENT_AI_CFG.PENDING_ACTION_TTL_MS / 60000),
    preview: previewInfo,
    instruction: 'Ask the owner to send /pin <ACTION_PIN>. Do not ask for or expose any Google, PayPal, Telegram, or OpenAI password/token.'
  };
}

function payAiExecutePendingAction_(chatId, rawPin) {
  const props = PropertiesService.getScriptProperties();
  const key = PAYMENT_AI_CFG.PENDING_PREFIX + String(chatId);
  const raw = props.getProperty(key);
  if (!raw) return { ok: false, error: 'No pending protected action. Ask AI to prepare the action first.' };

  let pending;
  try { pending = JSON.parse(raw); } catch (_) { props.deleteProperty(key); return { ok: false, error: 'Pending action is corrupted. Prepare it again.' }; }

  if (!pending.expiresAt || Date.now() > Number(pending.expiresAt)) {
    props.deleteProperty(key);
    return { ok: false, error: 'Protected action expired. Prepare it again.' };
  }

  const savedHash = String(props.getProperty(PAYMENT_AI_CFG.PROP_ACTION_PIN_HASH) || '');
  if (!savedHash) return { ok: false, error: 'Action PIN is not configured. Run setPaymentAiActionPin() from the spreadsheet.' };

  const enteredHash = payAiSha256Hex_(String(rawPin || ''));
  if (enteredHash !== savedHash) {
    pending.attempts = Number(pending.attempts || 0) + 1;
    if (pending.attempts >= PAYMENT_AI_CFG.MAX_PIN_ATTEMPTS) {
      props.deleteProperty(key);
      return { ok: false, error: 'Wrong PIN too many times. Pending action cancelled.' };
    }
    props.setProperty(key, JSON.stringify(pending));
    return { ok: false, error: 'Wrong action PIN. Attempts remaining: ' + (PAYMENT_AI_CFG.MAX_PIN_ATTEMPTS - pending.attempts) };
  }

  const execute = function() {
    let result;
    try {
      result = payAiPerformProtectedAction_(pending.action, pending.args, pending.fingerprint);
    } catch (err) {
      /*
       * Hỏng thì bỏ luôn pending, kể cả lỗi tạm như hết giờ chờ lock. Sau một
       * lần tạo hoá đơn thất bại, không ai biết chắc nó đã đi tới đâu, nên bắt
       * xem lại số rồi duyệt lại là đúng — dùng lại chữ ký duyệt cũ cho lần thử
       * sau là biến PIN thành tấm séc khống.
       */
      props.deleteProperty(key);
      throw err;
    }
    props.deleteProperty(key);
    return { ok: true, action: pending.action, result: payAiJsonSafe_(result) };
  };

  if (typeof paymentAllWithScriptLock_ === 'function') {
    return paymentAllWithScriptLock_('payAiExecutePendingAction', execute);
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try { return execute(); }
  finally { lock.releaseLock(); }
}

function payAiCancelPendingAction_(chatId) {
  PropertiesService.getScriptProperties().deleteProperty(PAYMENT_AI_CFG.PENDING_PREFIX + String(chatId));
  return { ok: true, message: 'Pending protected action cancelled.' };
}

function payAiPerformProtectedAction_(action, args, approvedFingerprint) {
  payAiActivateSpreadsheet_();

  if (action === 'create_invoice') {
    payAiRequireFunction_('previewCustomerInvoice');
    payAiRequireFunction_('createCustomerInvoice');
    const params = payAiInvoiceParamsFromArgs_(args);
    const preview = previewCustomerInvoice(params);
    const previewSummary = payAiSummarizePreview_(preview);

    /*
     * Chữ ký của bộ số đã được duyệt bằng PIN. createCustomerInvoice() dựng
     * preview của riêng nó sau khi sync rồi so với chữ ký này; lệch thì nó ném
     * lỗi thay vì tạo một hoá đơn khác với hoá đơn đã duyệt.
     *
     * Chỉ gắn khi thật sự có chữ ký. Đường gọi cũ không có thì hành vi giữ
     * nguyên như trước — nhưng đường Telegram thì luôn có.
     */
    const expected = String(approvedFingerprint || '').trim();
    if (expected) params.expected_fingerprint = expected;

    const created = createCustomerInvoice(params);
    return { preview: previewSummary, created: payAiJsonSafe_(created) };
  }

  if (action === 'set_customer_payout_fee') {
    const percent = Number(args.payout_fee_percent);
    if (!isFinite(percent) || percent < 0 || percent > 100) throw new Error('payout_fee_percent must be between 0 and 100.');

    const ss = payAiSpreadsheet_();
    const sh = ss.getSheetByName('Customers');
    if (!sh) throw new Error('Missing Customers sheet.');
    const map = payAiHeaderMap_(sh);
    ['Customer Key','Payout Fee %'].forEach(function(h) { if (!map[h]) throw new Error('Customers missing header: ' + h); });
    const rows = payAiReadObjects_(sh, map);
    const key = String(args.customer_key || '').trim().toUpperCase();
    const row = rows.find(function(r) { return String(r['Customer Key'] || '').trim().toUpperCase() === key; });
    if (!row) throw new Error('Customer not found: ' + key);

    const oldRate = payAiToNumber_(row['Payout Fee %']) || 0;
    const newRate = percent / 100;
    sh.getRange(row._row, map['Payout Fee %']).setValue(newRate).setNumberFormat('0.00%');
    SpreadsheetApp.flush();
    if (typeof syncCustomers === 'function') syncCustomers();

    return { customer_key: key, old_percent: oldRate * 100, new_percent: percent };
  }

  throw new Error('Unsupported protected action: ' + action);
}

function payAiGetCustomerRow_(customerKey) {
  const sh = payAiSheet_('Customers');
  const map = payAiHeaderMap_(sh);
  const key = String(customerKey || '').trim().toUpperCase();
  const row = payAiReadObjects_(sh, map).find(function(r) {
    return String(r['Customer Key'] || '').trim().toUpperCase() === key;
  });
  if (!row) throw new Error('Customer not found: ' + key);
  return row;
}


/************************* RUNTIME GUIDE *************************/

function payAiRuntimeGuide_() {
  if (typeof getPaymentAiReadme === 'function') {
    try { return String(getPaymentAiReadme() || ''); } catch (_) {}
  }
  if (typeof PAYMENT_AI_SYSTEM_GUIDE !== 'undefined') return String(PAYMENT_AI_SYSTEM_GUIDE || '');
  return 'Payment All: Flow.gs owns treasury/reconciliation. CustomerV2.3 owns customer ledger/allocation/settlement. Invoice PDF is presentation only. Customer money becomes invoiceable only when Customer Allocation is AVAILABLE after PayPal withdrawal is matched to the receiving provider.';
}


/************************* HELPERS *************************/

function payAiRequireFunction_(name) {
  const available = {
    syncFlow: typeof syncFlow === 'function',
    syncCustomers: typeof syncCustomers === 'function',
    previewCustomerInvoice: typeof previewCustomerInvoice === 'function',
    createCustomerInvoice: typeof createCustomerInvoice === 'function',
    previewCustomerInvoiceFromControl: typeof previewCustomerInvoiceFromControl === 'function',
    validateCustomerSystem: typeof validateCustomerSystem === 'function',
    createCustomerInvoiceFromControl: typeof createCustomerInvoiceFromControl === 'function'
  };
  if (!available[name]) throw new Error('Required function is not available in this Apps Script project: ' + name + '()');
}

function payAiJsonValue_(v) {
  if (v instanceof Date) {
    if (isNaN(v.getTime())) return '';
    return Utilities.formatDate(v, PAYMENT_AI_CFG.BUSINESS_TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
  }
  if (typeof v === 'number' && !isFinite(v)) return 0;
  return v;
}

function payAiJsonSafe_(v) {
  if (v === undefined) return null;
  try { return JSON.parse(JSON.stringify(v, function(k, x) { return payAiJsonValue_(x); })); }
  catch (_) { return String(v); }
}

/*
 * Không dùng instanceof: nó so theo constructor của từng realm, nên hỏng khi
 * chạy trong sandbox test — và một phép kiểm mà test không xác minh được thì
 * chẳng bảo vệ được gì. toString đúng ở mọi realm.
 */
function payAiIsDate_(v) {
  return Object.prototype.toString.call(v) === '[object Date]';
}


function payAiToNumber_(v) {
  if (typeof v === 'number' && isFinite(v)) return v;
  if (v === null || v === undefined || v === '') return null;
  /* Ngày tháng không phải số — xem chú thích ở custNum_ bên Customer.gs. */
  if (payAiIsDate_(v)) return null;
  let s = String(v).trim().replace(/[^\d,.-]/g, '');
  if (!s) return null;
  if (s.indexOf(',') !== -1 && s.indexOf('.') !== -1) {
    if (s.lastIndexOf(',') > s.lastIndexOf('.')) s = s.replace(/\./g, '').replace(',', '.');
    else s = s.replace(/,/g, '');
  } else if (s.indexOf(',') !== -1) {
    const parts = s.split(',');
    if (parts.length === 2 && parts[1].length <= 4) s = parts[0] + '.' + parts[1];
    else s = s.replace(/,/g, '');
  }
  const n = Number(s);
  return isFinite(n) ? n : null;
}

function payAiRound_(n, digits) {
  const p = Math.pow(10, Number(digits || 0));
  return Math.round((Number(n) + Number.EPSILON) * p) / p;
}

function payAiParseDateInput_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return value;
  const s = String(value || '').trim();
  if (!s) return null;
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), 12, 0, 0, 0);
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]), 12, 0, 0, 0);
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

function payAiDateKey_(v) {
  if (v instanceof Date && !isNaN(v.getTime())) return Utilities.formatDate(v, PAYMENT_AI_CFG.BUSINESS_TIMEZONE, 'yyyy-MM-dd');
  const d = payAiParseDateInput_(v);
  if (!d) return '';
  return Utilities.formatDate(d, PAYMENT_AI_CFG.BUSINESS_TIMEZONE, 'yyyy-MM-dd');
}

function payAiDateMs_(v) {
  const d = v instanceof Date ? v : payAiParseDateInput_(v);
  return d && !isNaN(d.getTime()) ? d.getTime() : null;
}

function payAiSha256Hex_(text) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(text || ''), Utilities.Charset.UTF_8);
  return bytes.map(function(b) {
    const n = b < 0 ? b + 256 : b;
    return ('0' + n.toString(16)).slice(-2);
  }).join('');
}
