/************************************************************
 * PAYPAL RAW COLLECTOR
 *
 * NHIỆM VỤ DUY NHẤT:
 * PayPal API -> raw data
 *
 * KHÔNG:
 * - classify transaction
 * - calculate Balance Impact
 * - normalize status
 * - reconcile
 * - calculate Grey / USDT
 ************************************************************/

const PP_CFG = Object.freeze({
  SPREADSHEET_ID: '', // để '' nếu script bound với Payment All

  ACCOUNT_SHEET: 'Account',
  PAYPAL_SHEET: 'Paypal',

  HEADER_ROW: 2,
  DATA_START_ROW: 3,

  API_BASE: 'https://api-m.paypal.com',

  PAGE_SIZE: 500,
  CHUNK_DAYS: 30,
  OVERLAP_HOURS: 6
});


/************************************************************
 * ENTRY
 ************************************************************/

function syncPayPalAllAccounts() {
  const ss = ppGetSpreadsheet_();

  const accountSheet = ss.getSheetByName(PP_CFG.ACCOUNT_SHEET);
  const paypalSheet = ss.getSheetByName(PP_CFG.PAYPAL_SHEET);

  if (!accountSheet) throw new Error('Không tìm thấy Account.');
  if (!paypalSheet) throw new Error('Không tìm thấy Paypal.');

  ppEnsureRawStatusHeader_(paypalSheet);

  const accountHeaders = ppHeaderMap_(accountSheet);
  const paypalHeaders = ppHeaderMap_(paypalSheet);

  ppValidateAccount_(accountHeaders);
  ppValidatePaypal_(paypalHeaders);

  const accounts = ppReadAccounts_(accountSheet, accountHeaders);

  accounts
    .filter(a => a.provider === 'PAYPAL' && a.enabled)
    .forEach(account => {
      try {
        ppSyncAccount_(
          account,
          accountSheet,
          accountHeaders,
          paypalSheet,
          paypalHeaders
        );
      } catch (e) {
        ppSetAccountField_(
          accountSheet,
          account.rowNumber,
          accountHeaders,
          'API Status',
          'ERROR: ' + ppTruncate_(e.message || String(e), 300)
        );

        Logger.log(
          'PayPal ERROR ' +
          account.accountKey +
          ': ' +
          (e.stack || e)
        );
      }
    });
}


/************************************************************
 * TEST / ONE ACCOUNT
 *
 * Ví dụ:
 * syncPayPalOneAccount('PP1')
 ************************************************************/

function syncPayPalOneAccount(accountKey) {
  const ss = ppGetSpreadsheet_();

  const accountSheet = ss.getSheetByName(PP_CFG.ACCOUNT_SHEET);
  const paypalSheet = ss.getSheetByName(PP_CFG.PAYPAL_SHEET);

  if (!accountSheet || !paypalSheet) {
    throw new Error('Thiếu Account hoặc Paypal.');
  }

  ppEnsureRawStatusHeader_(paypalSheet);

  const accountHeaders = ppHeaderMap_(accountSheet);
  const paypalHeaders = ppHeaderMap_(paypalSheet);

  ppValidateAccount_(accountHeaders);
  ppValidatePaypal_(paypalHeaders);

  const accounts = ppReadAccounts_(accountSheet, accountHeaders);

  const account = accounts.find(
    a => a.accountKey === String(accountKey).trim()
  );

  if (!account) {
    throw new Error('Không tìm thấy account ' + accountKey);
  }

  ppSyncAccount_(
    account,
    accountSheet,
    accountHeaders,
    paypalSheet,
    paypalHeaders
  );
}


/************************************************************
 * SYNC ONE
 ************************************************************/

function ppSyncAccount_(
  account,
  accountSheet,
  accountHeaders,
  paypalSheet,
  paypalHeaders
) {
  ppSetAccountField_(
    accountSheet,
    account.rowNumber,
    accountHeaders,
    'API Status',
    'SYNCING'
  );

  const token = ppGetToken_(
    account.clientId,
    account.clientSecret
  );

  const state = ppLoadPaypalState_(
    paypalSheet,
    paypalHeaders
  );

  const now = new Date();

  let from;

  if (account.lastSync) {
    from = new Date(
      account.lastSync.getTime() -
      PP_CFG.OVERLAP_HOURS * 60 * 60 * 1000
    );
  } else if (account.initialFrom) {
    from = new Date(account.initialFrom);
  } else {
    from = new Date(
      now.getTime() -
      30 * 24 * 60 * 60 * 1000
    );
  }

  let cursor = new Date(from);
  let fetched = 0;

  while (cursor < now) {
    let end = new Date(
      cursor.getTime() +
      PP_CFG.CHUNK_DAYS * 24 * 60 * 60 * 1000 -
      1000
    );

    if (end > now) end = new Date(now);

    const rows = ppFetchTransactions_(
      token,
      cursor,
      end
    );

    rows.forEach(detail => {
      const record = ppBuildRawRecord_(
        detail,
        account
      );

      if (!record) return;

      ppUpsertPaypalRaw_(
        state,
        paypalHeaders,
        record
      );

      fetched++;
    });

    cursor = new Date(end.getTime() + 1000);
  }

  ppWritePaypalState_(
    paypalSheet,
    paypalHeaders,
    state
  );

  /**********************************************************
   * BALANCE = PayPal API trả trực tiếp
   * Đây KHÔNG phải calculation.
   **********************************************************/

  try {
    const balance = ppFetchBalance_(token);

    if (balance) {
      ppSetAccountField_(
        accountSheet,
        account.rowNumber,
        accountHeaders,
        'PayPal Account ID',
        balance.accountId
      );

      ppSetAccountField_(
        accountSheet,
        account.rowNumber,
        accountHeaders,
        'Total Balance',
        balance.total
      );

      ppSetAccountField_(
        accountSheet,
        account.rowNumber,
        accountHeaders,
        'Available Balance',
        balance.available
      );

      ppSetAccountField_(
        accountSheet,
        account.rowNumber,
        accountHeaders,
        'Withheld Balance',
        balance.withheld
      );

      ppSetAccountField_(
        accountSheet,
        account.rowNumber,
        accountHeaders,
        'Currency',
        balance.currency
      );

      ppSetAccountField_(
        accountSheet,
        account.rowNumber,
        accountHeaders,
        'Last Refresh',
        new Date()
      );
    }
  } catch (e) {
    Logger.log(
      'Balance API warning ' +
      account.accountKey +
      ': ' +
      e.message
    );
  }

  ppSetAccountField_(
    accountSheet,
    account.rowNumber,
    accountHeaders,
    'Last Sync',
    now
  );

  ppSetAccountField_(
    accountSheet,
    account.rowNumber,
    accountHeaders,
    'API Status',
    'OK · ' + fetched + ' records'
  );

  SpreadsheetApp.flush();

  Logger.log(
    account.accountKey +
    ' OK · ' +
    fetched +
    ' raw records'
  );
}


/************************************************************
 * OAUTH
 ************************************************************/

function ppGetToken_(clientId, clientSecret) {
  const auth = Utilities.base64Encode(
    clientId + ':' + clientSecret
  );

  const response = UrlFetchApp.fetch(
    PP_CFG.API_BASE + '/v1/oauth2/token',
    {
      method: 'post',
      contentType: 'application/x-www-form-urlencoded',
      payload: 'grant_type=client_credentials',
      headers: {
        Authorization: 'Basic ' + auth
      },
      muteHttpExceptions: true
    }
  );

  const code = response.getResponseCode();
  const text = response.getContentText();

  if (code !== 200) {
    throw new Error(
      'OAuth HTTP ' + code + ': ' + text
    );
  }

  const json = JSON.parse(text);

  if (!json.access_token) {
    throw new Error('PayPal không trả access_token.');
  }

  return json.access_token;
}


/************************************************************
 * TRANSACTION SEARCH RAW
 ************************************************************/

function ppFetchTransactions_(token, startDate, endDate) {
  const all = [];

  let page = 1;

  while (true) {
    const url =
      PP_CFG.API_BASE +
      '/v1/reporting/transactions' +
      '?start_date=' +
      encodeURIComponent(startDate.toISOString()) +
      '&end_date=' +
      encodeURIComponent(endDate.toISOString()) +
      '&fields=all' +
      '&page_size=' +
      PP_CFG.PAGE_SIZE +
      '&page=' +
      page;

    const response = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: {
        Authorization: 'Bearer ' + token
      },
      muteHttpExceptions: true
    });

    const code = response.getResponseCode();
    const text = response.getContentText();

    if (code !== 200) {
      throw new Error(
        'Transaction Search HTTP ' +
        code +
        ': ' +
        text
      );
    }

    const json = JSON.parse(text);

    const details =
      Array.isArray(json.transaction_details)
        ? json.transaction_details
        : [];

    all.push.apply(all, details);

    const totalPages =
      Number(json.total_pages || 0);

    if (
      details.length < PP_CFG.PAGE_SIZE ||
      (totalPages && page >= totalPages)
    ) {
      break;
    }

    page++;
  }

  return all;
}


/************************************************************
 * BALANCE RAW
 ************************************************************/

function ppFetchBalance_(token) {
  const url =
    PP_CFG.API_BASE +
    '/v1/reporting/balances?currency_code=USD';

  const response = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: {
      Authorization: 'Bearer ' + token
    },
    muteHttpExceptions: true
  });

  const code = response.getResponseCode();
  const text = response.getContentText();

  if (code !== 200) {
    throw new Error(
      'Balance HTTP ' + code + ': ' + text
    );
  }

  const json = JSON.parse(text);

  const balances =
    Array.isArray(json.balances)
      ? json.balances
      : [];

  if (!balances.length) return null;

  const usd =
    balances.find(
      b =>
        String(
          b.currency ||
          b.total_balance?.currency_code ||
          ''
        ).toUpperCase() === 'USD'
    ) || balances[0];

  return {
    accountId: json.account_id || '',
    total: ppMoneyValue_(usd.total_balance),
    available: ppMoneyValue_(usd.available_balance),
    withheld: ppMoneyValue_(usd.withheld_balance),

    currency:
      usd.currency ||
      usd.total_balance?.currency_code ||
      'USD'
  };
}


/************************************************************
 * API OBJECT -> RAW ROW
 *
 * Derived columns are intentionally blank/preserved:
 * Type
 * Status
 * Invoice Prefix
 * Balance Impact
 ************************************************************/

function ppBuildRawRecord_(detail, account) {
  const info =
    detail.transaction_info || {};

  const payer =
    detail.payer_info || {};

  const cart =
    detail.cart_info || {};

  const transactionId =
    ppString_(info.transaction_id);

  if (!transactionId) return null;

  const eventCode =
    ppString_(info.transaction_event_code);

  const rawStatus =
    ppString_(info.transaction_status);

  const amount =
    ppNumber_(info.transaction_amount?.value);

  const fee =
    ppNumber_(info.fee_amount?.value);

  const currency =
    ppString_(
      info.transaction_amount?.currency_code ||
      info.fee_amount?.currency_code
    ).toUpperCase();

  let buyerName = '';

  if (payer.payer_name) {
    buyerName =
      ppString_(
        payer.payer_name.alternate_full_name
      );

    if (!buyerName) {
      buyerName = ppString_(
        [
          payer.payer_name.given_name,
          payer.payer_name.surname
        ]
          .filter(Boolean)
          .join(' ')
      );
    }
  }

  const items =
    Array.isArray(cart.item_details)
      ? cart.item_details
      : [];

  const itemText =
    items
      .map(i => ppString_(i.item_name))
      .filter(Boolean)
      .slice(0, 5)
      .join(' | ');

  const uniqueKey =
    account.accountKey +
    '|' +
    transactionId +
    '|' +
    eventCode;

  return {
    'Account Key': account.accountKey,
    'Label': account.label,

    'Initiation Date':
      ppDate_(info.transaction_initiation_date),

    'Updated Date':
      ppDate_(info.transaction_updated_date),

    'Transaction ID': transactionId,

    'Reference ID':
      ppString_(info.paypal_reference_id),

    'Event Code': eventCode,

    // FLOW owns these:
    'Type': '',
    'Status': '',

    'Raw Status': rawStatus,

    'Invoice ID':
      ppString_(info.invoice_id),

    // FLOW owns:
    'Invoice Prefix': '',

    'Buyer Name': buyerName,

    'Buyer Email':
      ppString_(payer.email_address),

    'Item': itemText,

    'Transaction Amount': amount,
    'Fee Amount': fee,

    // FLOW owns:
    'Balance Impact': '',

    'Currency': currency,

    'Ending Balance':
      ppMoneyValue_(info.ending_balance),

    'Available Balance':
      ppMoneyValue_(info.available_balance),

    'Instrument':
      ppString_(
        info.instrument_type ||
        detail.instrument_type
      ),

    'Subject':
      ppString_(info.transaction_subject),

    'Sync Time': new Date(),

    'Unique Key': uniqueKey
  };
}


/************************************************************
 * ACCOUNT
 ************************************************************/

function ppReadAccounts_(sheet, h) {
  const lastRow = sheet.getLastRow();

  if (lastRow < PP_CFG.DATA_START_ROW) {
    return [];
  }

  const rows = sheet.getRange(
    PP_CFG.DATA_START_ROW,
    1,
    lastRow - PP_CFG.DATA_START_ROW + 1,
    sheet.getLastColumn()
  ).getValues();

  return rows
    .map((row, i) => ({
      rowNumber: PP_CFG.DATA_START_ROW + i,

      accountKey:
        ppString_(row[h['Account Key']]),

      label:
        ppString_(row[h['Label']]),

      provider:
        ppString_(row[h['Provider']]).toUpperCase(),

      clientId:
        ppString_(row[h['Client ID']]),

      clientSecret:
        ppString_(row[h['Client Secret']]),

      enabled:
        ppBool_(row[h['Enabled']]),

      initialFrom:
        ppValidDate_(row[h['Initial From']]),

      lastSync:
        ppValidDate_(row[h['Last Sync']])
    }))
    .filter(a => a.accountKey);
}


/************************************************************
 * STATE
 ************************************************************/

function ppLoadPaypalState_(sheet, h) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  let rows = [];

  if (lastRow >= PP_CFG.DATA_START_ROW) {
    rows = sheet.getRange(
      PP_CFG.DATA_START_ROW,
      1,
      lastRow - PP_CFG.DATA_START_ROW + 1,
      lastCol
    ).getValues();

    rows = rows.filter(
      row => row.some(v => v !== '' && v !== null)
    );
  }

  const keyMap = {};

  rows.forEach((row, i) => {
    const key =
      ppString_(row[h['Unique Key']]);

    if (key) keyMap[key] = i;
  });

  return {
    rows,
    keyMap,
    columnCount: lastCol,
    oldRowCount:
      Math.max(
        0,
        lastRow - PP_CFG.DATA_START_ROW + 1
      )
  };
}


/************************************************************
 * UPSERT RAW
 *
 * Không overwrite các field do Flow quản lý.
 ************************************************************/

function ppUpsertPaypalRaw_(state, h, record) {
  const key = record['Unique Key'];

  let row;
  let exists = false;

  if (
    Object.prototype.hasOwnProperty.call(
      state.keyMap,
      key
    )
  ) {
    row =
      state.rows[state.keyMap[key]];

    exists = true;
  } else {
    row =
      new Array(state.columnCount).fill('');

    state.rows.push(row);
    state.keyMap[key] =
      state.rows.length - 1;
  }

  const flowOwned = {
    'Type': true,
    'Status': true,
    'Invoice Prefix': true,
    'Balance Impact': true
  };

  Object.keys(record).forEach(header => {
    if (
      !Object.prototype.hasOwnProperty.call(
        h,
        header
      )
    ) {
      return;
    }

    if (flowOwned[header]) {
      if (!exists) {
        row[h[header]] = '';
      }

      return;
    }

    row[h[header]] = record[header];
  });
}


/************************************************************
 * WRITE
 ************************************************************/

function ppWritePaypalState_(sheet, h, state) {
  const dateIndex = h['Initiation Date'];

  state.rows.sort(
    (a, b) =>
      ppMillis_(b[dateIndex]) -
      ppMillis_(a[dateIndex])
  );

  const clearRows = Math.max(
    state.oldRowCount,
    state.rows.length
  );

  if (clearRows > 0) {
    sheet.getRange(
      PP_CFG.DATA_START_ROW,
      1,
      clearRows,
      state.columnCount
    ).clearContent();
  }

  if (state.rows.length) {
    sheet.getRange(
      PP_CFG.DATA_START_ROW,
      1,
      state.rows.length,
      state.columnCount
    ).setValues(state.rows);
  }
}


/************************************************************
 * RAW STATUS HEADER
 ************************************************************/

function ppEnsureRawStatusHeader_(sheet) {
  const lastCol =
    Math.max(1, sheet.getLastColumn());

  const headers = sheet.getRange(
    PP_CFG.HEADER_ROW,
    1,
    1,
    lastCol
  ).getValues()[0];

  const exists =
    headers.some(
      h => ppString_(h) === 'Raw Status'
    );

  if (!exists) {
    sheet.getRange(
      PP_CFG.HEADER_ROW,
      lastCol + 1
    ).setValue('Raw Status');
  }
}


/************************************************************
 * VALIDATE
 ************************************************************/

function ppValidateAccount_(h) {
  ppRequire_(h, [
    'Account Key',
    'Label',
    'Provider',
    'Client ID',
    'Client Secret',
    'Enabled',
    'Initial From',
    'Last Sync',
    'API Status',
    'PayPal Account ID',
    'Total Balance',
    'Available Balance',
    'Withheld Balance',
    'Currency',
    'Last Refresh'
  ], 'Account');
}

function ppValidatePaypal_(h) {
  ppRequire_(h, [
    'Account Key',
    'Label',
    'Initiation Date',
    'Updated Date',
    'Transaction ID',
    'Reference ID',
    'Event Code',
    'Type',
    'Status',
    'Raw Status',
    'Invoice ID',
    'Invoice Prefix',
    'Buyer Name',
    'Buyer Email',
    'Item',
    'Transaction Amount',
    'Fee Amount',
    'Balance Impact',
    'Currency',
    'Ending Balance',
    'Available Balance',
    'Instrument',
    'Subject',
    'Sync Time',
    'Unique Key'
  ], 'Paypal');
}


/************************************************************
 * GENERIC
 ************************************************************/

function ppGetSpreadsheet_() {
  if (PP_CFG.SPREADSHEET_ID) {
    return SpreadsheetApp.openById(
      PP_CFG.SPREADSHEET_ID
    );
  }

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  if (!ss) {
    throw new Error(
      'Không có active spreadsheet. ' +
      'Hãy điền PP_CFG.SPREADSHEET_ID.'
    );
  }

  return ss;
}

function ppHeaderMap_(sheet) {
  const row = sheet.getRange(
    PP_CFG.HEADER_ROW,
    1,
    1,
    sheet.getLastColumn()
  ).getValues()[0];

  const map = {};

  row.forEach((v, i) => {
    const name = ppString_(v);

    if (!name) return;

    if (
      Object.prototype.hasOwnProperty.call(
        map,
        name
      )
    ) {
      throw new Error(
        'Duplicate header "' +
        name +
        '" trong ' +
        sheet.getName()
      );
    }

    map[name] = i;
  });

  return map;
}

function ppRequire_(map, required, sheetName) {
  const missing = required.filter(
    h =>
      !Object.prototype.hasOwnProperty.call(
        map,
        h
      )
  );

  if (missing.length) {
    throw new Error(
      sheetName +
      ' thiếu header: ' +
      missing.join(', ')
    );
  }
}

function ppSetAccountField_(
  sheet,
  row,
  h,
  name,
  value
) {
  sheet.getRange(
    row,
    h[name] + 1
  ).setValue(value);
}

function ppMoneyValue_(obj) {
  if (
    obj === null ||
    obj === undefined ||
    obj === ''
  ) {
    return '';
  }

  if (typeof obj === 'number') return obj;

  if (typeof obj === 'string') {
    return ppNumber_(obj);
  }

  if (
    typeof obj === 'object' &&
    obj.value !== undefined
  ) {
    return ppNumber_(obj.value);
  }

  return '';
}

function ppNumber_(v) {
  if (typeof v === 'number') return v;

  const n = Number(
    String(v || '')
      .replace(/,/g, '')
      .replace(/[^0-9.\-]/g, '')
  );

  return isNaN(n) ? 0 : n;
}

function ppString_(v) {
  return String(
    v === null || v === undefined
      ? ''
      : v
  ).trim();
}

function ppBool_(v) {
  if (v === true) return true;

  const s =
    ppString_(v).toUpperCase();

  return (
    s === 'TRUE' ||
    s === 'YES' ||
    s === '1' ||
    s === 'ACTIVE'
  );
}

function ppDate_(v) {
  if (!v) return '';

  const d = new Date(v);

  return isNaN(d.getTime())
    ? ''
    : d;
}

function ppValidDate_(v) {
  if (!v) return null;

  if (v instanceof Date) {
    return isNaN(v.getTime())
      ? null
      : v;
  }

  const d = new Date(v);

  return isNaN(d.getTime())
    ? null
    : d;
}

function ppMillis_(v) {
  const d = ppValidDate_(v);

  return d ? d.getTime() : 0;
}

function ppTruncate_(v, max) {
  const s = String(v || '');

  return s.length > max
    ? s.substring(0, max)
    : s;
}