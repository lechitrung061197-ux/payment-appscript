/*************************************************
 * CUSTOMER V2.3.GS — CONVERSION + TRANSFER PRICING + LATE REFUND
 *
 * TARGET FLOW:
 * PayPal sale
 *   -> Customer Ledger
 *   -> PayPal withdrawal
 *   -> Receiving Provider MATCHED
 *   -> Customer Allocation AVAILABLE
 *   -> Rule-based conversion + transfer pricing
 *   -> Preview
 *   -> Settlement / Invoice
 *   -> Payout DRAFT
 *
 * CUSTOMER SOURCE OF TRUTH:
 * PayPal Total
 * -> Matched Amount
 * -> Available Amount
 * -> Invoiced Amount
 * -> Outstanding Amount
 *
 * IMPORTANT:
 * - Customer does NOT need to be traced through Grey -> Grey -> USDT.
 * - Transaction + Flow continue to manage real treasury movement.
 * - PayPal Fee / Refund / Receiving Fee are ACTUAL.
 * - Downstream pricing comes from Fee Rules:
 *     Conversion Cost = Available Amount * Conversion Rate
 *     Transfer Cost   = MAX(Available Amount * Transfer Rate, Transfer Minimum)
 *     Downstream Cost = Conversion Cost + Transfer Cost
 * - Customer payout fee still comes from Customers.
 *
 * Customer Transfer Allocation is NOT used by V2.
 *
 * LATE REFUND RULE:
 * - Refund known before the first invoice is handled normally in that invoice.
 * - Refund appearing after an invoice is ISSUED never rewrites the old invoice.
 * - The delta becomes LATE_REFUND_ADJUSTMENT on the next invoice.
 * - If the adjustment is larger than the next payable base, Amount Due = 0 and
 *   the unused refund balance carries forward to a later invoice.
 * - Settlement Items are the audit / consumption record, so syncCustomers()
 *   remains idempotent and cannot consume the same late refund twice.
 *************************************************/

const CUST_CFG = Object.freeze({
  CUSTOMERS_SHEET: 'Customers',
  PAYPAL_SHEET: 'Paypal',
  TRANSACTION_SHEET: 'Transaction',
  LEDGER_SHEET: 'Customer Ledger',
  ALLOCATION_SHEET: 'Customer Allocation',
  FEE_RULES_SHEET: 'Fee Rules',
  SETTLEMENTS_SHEET: 'Settlements',
  SETTLEMENT_ITEMS_SHEET: 'Settlement Items',
  PAYOUTS_SHEET: 'Payouts',
  CONTROL_SHEET: 'Invoice Control',
  HEADER_ROW: 2,
  DATA_START_ROW: 3,
  BUSINESS_TIMEZONE: 'America/New_York',
  EPSILON: 0.005,
  PREVIEW_PROP: 'CUSTOMER_INVOICE_PREVIEW_V5_CONVERSION_TRANSFER_LATE_REFUND'
});

/*********************** PUBLIC ************************/
function syncCustomers() { return syncCustomers_(); }
function runCustomerControl() {
  const ss = SpreadsheetApp.getActive();
  const sh = custRequireSheet_(ss, CUST_CFG.CONTROL_SHEET);
  const mode = custUpper_(sh.getRange('B3').getValue()) || 'CUSTOMER STATEMENT';
  if (mode === 'INVOICE PREVIEW') return previewCustomerInvoiceFromControl();
  return previewCustomerStatementFromControl();
}
function clearCustomerControl() { custClearControlOutput_(SpreadsheetApp.getActive()); }

function setupCustomerModule() {
  const ss = SpreadsheetApp.getActive();
  const customers = custRequireSheet_(ss, CUST_CFG.CUSTOMERS_SHEET);
  const ledger = custRequireSheet_(ss, CUST_CFG.LEDGER_SHEET);
  const alloc = custRequireSheet_(ss, CUST_CFG.ALLOCATION_SHEET);
  const feeRules = custRequireSheet_(ss, CUST_CFG.FEE_RULES_SHEET);
  const settlements = custRequireSheet_(ss, CUST_CFG.SETTLEMENTS_SHEET);
  const items = custRequireSheet_(ss, CUST_CFG.SETTLEMENT_ITEMS_SHEET);
  const payouts = custRequireSheet_(ss, CUST_CFG.PAYOUTS_SHEET);
  const control = custRequireSheet_(ss, CUST_CFG.CONTROL_SHEET);

  const hc = custHeaderMap_(customers);
  const hl = custHeaderMap_(ledger);
  const ha = custHeaderMap_(alloc);
  const hf = custHeaderMap_(feeRules);
  const hs = custHeaderMap_(settlements);
  const hi = custHeaderMap_(items);

  custRequireHeaders_(CUST_CFG.CUSTOMERS_SHEET, hc, [
    'Customer Key','Customer Name','Status','Email','Prefixes','Payout Fee %',
    'Default Payout Provider','Default Payout Account Key','PayPal Total','Matched Amount',
    'Available Amount','Invoiced Amount','Outstanding Amount','Pending in PayPal',
    'Last Settlement','Last Payout','Last Sync','Sync Status','Note'
  ]);
  custRequireHeaders_(CUST_CFG.LEDGER_SHEET, hl, [
    'Ledger Key','Transaction Date','Customer Key','Customer Name','PayPal Account Key',
    'PayPal Transaction ID','Reference ID','Invoice ID','Invoice Prefix','Event Type',
    'Gross Amount','PayPal Fee','Refund Amount','Net Amount','Currency','Allocated Amount',
    'Invoiced Amount','Remaining Amount','Source Unique Key','Last Sync','Status','Note'
  ]);
  custRequireHeaders_(CUST_CFG.ALLOCATION_SHEET, ha, [
    'Allocation Key','Allocation Date','Customer Key','Customer Ledger Key','PayPal Account Key',
    'PayPal Source Transaction ID','PayPal Withdrawal Transaction Key','PayPal Withdrawal Transaction ID',
    'PayPal Withdrawal Date','PayPal Withdrawal Status','Allocated Amount','Currency','Receiving Provider',
    'Receiving Account Key','Receiving Transaction Key','Receiving Reference','Receiving Fee Allocated',
    'Available Amount','Available Date','Allocation Status','Invoiced Amount','Remaining Available',
    'Settlement Key','Source Unique Key','Last Sync','Note'
  ]);
  custRequireHeaders_(CUST_CFG.FEE_RULES_SHEET, hf, [
    'Rule Key','Enabled','Receiving Provider','Payout Provider','Currency','Conversion Rate',
    'Transfer Rate','Transfer Minimum','Effective From','Effective To','Priority','Note'
  ]);
  custRequireHeaders_(CUST_CFG.SETTLEMENTS_SHEET, hs, [
    'Settlement Key','Invoice Number','Customer Key','Customer Name','Period From','Period To',
    'Gross Sales','Refunds','PayPal Fees','Receiving Fees','Available Amount','Conversion Rate',
    'Conversion Cost','Transfer Rate','Transfer Minimum','Transfer Cost','Downstream Cost',
    'Net Before Payout Fee','Payout Fee %','Payout Fee Amount','Amount Due','Amount Paid',
    'Outstanding Payable','Currency','PayPal Total Snapshot','Matched Snapshot','Available Snapshot',
    'Invoiced Snapshot','Outstanding Snapshot','Fee Rule Key','Payout Provider','Payout Account Key',
    'Payout Currency','Status','Created At','Issued At','Paid At','Last Sync','PDF File ID','PDF URL',
    'PDF Generated At','Note'
  ]);
  custRequireHeaders_(CUST_CFG.SETTLEMENT_ITEMS_SHEET, hi, [
    'Settlement Item Key','Settlement Key','Invoice Number','Customer Key','Line Type','Transaction Date',
    'PayPal Account Key','PayPal Transaction ID','Reference ID','Invoice ID','Invoice Prefix','Item','Buyer Name',
    'Gross Amount','PayPal Fee','Refund Amount','PayPal Net Amount','Receiving Fee Allocated','Available Amount',
    'Amount Included','Currency','Customer Ledger Key','Customer Allocation Key','Source Unique Key','Created At','Note'
  ]);

  custEnsureHeaders_(payouts, ['Customer Key','Settlement Key']);

  const n = Math.max(customers.getMaxRows() - CUST_CFG.DATA_START_ROW + 1, 1);
  customers.getRange(CUST_CFG.DATA_START_ROW, hc['Status'], n, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['ACTIVE','DISABLED'], true).setAllowInvalid(false).build()
  );
  customers.getRange(CUST_CFG.DATA_START_ROW, hc['Default Payout Provider'], n, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['USDT','PINGPONG'], true).setAllowInvalid(false).build()
  );
  customers.getRange(CUST_CFG.DATA_START_ROW, hc['Payout Fee %'], n, 1).setNumberFormat('0.00%');

  const fn = Math.max(feeRules.getMaxRows() - CUST_CFG.DATA_START_ROW + 1, 1);
  feeRules.getRange(CUST_CFG.DATA_START_ROW, hf['Conversion Rate'], fn, 1).setNumberFormat('0.000');
  feeRules.getRange(CUST_CFG.DATA_START_ROW, hf['Transfer Rate'], fn, 1).setNumberFormat('0.000');
  feeRules.getRange(CUST_CFG.DATA_START_ROW, hf['Transfer Minimum'], fn, 1).setNumberFormat('0.00');

  const sn = Math.max(settlements.getMaxRows() - CUST_CFG.DATA_START_ROW + 1, 1);
  ['Conversion Rate','Transfer Rate'].forEach(k=>{
    settlements.getRange(CUST_CFG.DATA_START_ROW, hs[k], sn, 1).setNumberFormat('0.000');
  });
  ['Gross Sales','Refunds','PayPal Fees','Receiving Fees','Available Amount','Conversion Cost',
   'Transfer Minimum','Transfer Cost','Downstream Cost','Net Before Payout Fee','Payout Fee Amount',
   'Amount Due','Amount Paid','Outstanding Payable','PayPal Total Snapshot','Matched Snapshot',
   'Available Snapshot','Invoiced Snapshot','Outstanding Snapshot'].forEach(k=>{
    settlements.getRange(CUST_CFG.DATA_START_ROW, hs[k], sn, 1).setNumberFormat('0.00');
  });
  settlements.getRange(CUST_CFG.DATA_START_ROW, hs['Payout Fee %'], sn, 1).setNumberFormat('0.00%');

  control.getRange('B3').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['CUSTOMER STATEMENT','INVOICE PREVIEW'], true).setAllowInvalid(false).build()
  );
  if (!custStr_(control.getRange('B3').getValue())) control.getRange('B3').setValue('CUSTOMER STATEMENT');
  control.getRange('B5:B7').setNumberFormat('dd/MM/yyyy');
  if (!control.getRange('B7').getValue()) control.getRange('B7').setValue(new Date());

  custRefreshInvoiceControlCustomerDropdown_();
  SpreadsheetApp.getActive().toast('Customer V2.3 setup completed.', 'Customer V2.3', 5);
}

/*********************** MAIN SYNC ************************/
function syncCustomers_() {
  const ss = SpreadsheetApp.getActive();
  const now = new Date();
  const shCustomers = custRequireSheet_(ss, CUST_CFG.CUSTOMERS_SHEET);
  const shPaypal = custRequireSheet_(ss, CUST_CFG.PAYPAL_SHEET);
  const shTx = custRequireSheet_(ss, CUST_CFG.TRANSACTION_SHEET);
  const shLedger = custRequireSheet_(ss, CUST_CFG.LEDGER_SHEET);
  const shAlloc = custRequireSheet_(ss, CUST_CFG.ALLOCATION_SHEET);
  const shPayouts = custRequireSheet_(ss, CUST_CFG.PAYOUTS_SHEET);

  custEnsureHeaders_(shPayouts, ['Customer Key','Settlement Key']);

  const hCustomers = custHeaderMap_(shCustomers);
  const hPaypal = custHeaderMap_(shPaypal);
  const hTx = custHeaderMap_(shTx);
  const hLedger = custHeaderMap_(shLedger);
  const hAlloc = custHeaderMap_(shAlloc);

  custRequireHeaders_(CUST_CFG.CUSTOMERS_SHEET, hCustomers, [
    'Customer Key','Customer Name','Status','Email','Prefixes','Payout Fee %',
    'Default Payout Provider','Default Payout Account Key','PayPal Total','Matched Amount',
    'Available Amount','Invoiced Amount','Outstanding Amount','Pending in PayPal',
    'Last Settlement','Last Payout','Last Sync','Sync Status','Note'
  ]);
  custRequireHeaders_(CUST_CFG.PAYPAL_SHEET, hPaypal, [
    'Account Key','Initiation Date','Transaction ID','Reference ID','Event Code','Status','Invoice ID','Invoice Prefix',
    'Buyer Name','Item','Transaction Amount','Fee Amount','Balance Impact','Currency','Unique Key'
  ]);
  custRequireHeaders_(CUST_CFG.TRANSACTION_SHEET, hTx, [
    'Transaction Key','Transaction Date','Provider','Account Key','Type','Direction','Gross Amount','Net Amount','Fee Amount',
    'Currency','Reference','Counterparty Provider','Counterparty Account Key','Reconcile Status','Matched Transaction Key',
    'Source','Source ID','Sync Time','Unique Key'
  ]);
  custRequireHeaders_(CUST_CFG.LEDGER_SHEET, hLedger, [
    'Ledger Key','Transaction Date','Customer Key','Customer Name','PayPal Account Key','PayPal Transaction ID','Reference ID',
    'Invoice ID','Invoice Prefix','Event Type','Gross Amount','PayPal Fee','Refund Amount','Net Amount','Currency','Allocated Amount',
    'Invoiced Amount','Remaining Amount','Source Unique Key','Last Sync','Status','Note'
  ]);
  custRequireHeaders_(CUST_CFG.ALLOCATION_SHEET, hAlloc, [
    'Allocation Key','Allocation Date','Customer Key','Customer Ledger Key','PayPal Account Key','PayPal Source Transaction ID',
    'PayPal Withdrawal Transaction Key','PayPal Withdrawal Transaction ID','PayPal Withdrawal Date','PayPal Withdrawal Status',
    'Allocated Amount','Currency','Receiving Provider','Receiving Account Key','Receiving Transaction Key','Receiving Reference',
    'Receiving Fee Allocated','Available Amount','Available Date','Allocation Status','Invoiced Amount','Remaining Available',
    'Settlement Key','Source Unique Key','Last Sync','Note'
  ]);

  let customerState;
  try { customerState = custReadCustomers_(shCustomers, hCustomers); }
  catch (err) { custSetGlobalCustomerError_(shCustomers, hCustomers, now, err.message); throw err; }

  const paypalRows = custReadPaypalRows_(shPaypal, hPaypal, customerState);
  const ledgerBuild = custBuildCustomerLedger_(paypalRows, customerState, {}, now);
  const ledgerRecords = ledgerBuild.records;
  custAssertSingleCurrencyPerCustomer_(ledgerRecords);

  /*
   * Settlement Items are read BEFORE rebuilding allocations so an already
   * invoiced sale can keep its historical allocation even if a later PayPal
   * refund reduces the current sale Net Amount.
   */
  const lateRefundState = custBuildLateRefundState_(ss, ledgerRecords);

  const txRows = custReadTransactionRows_(shTx, hTx);
  const oldAlloc = custReadExistingAllocationState_(shAlloc, hAlloc);
  const allocations = custBuildCustomerAllocations_(
    ledgerBuild.allSourceRecords,
    ledgerRecords,
    txRows,
    oldAlloc,
    lateRefundState,
    now
  );

  custFinalizeFullLedger_(ledgerRecords, allocations, []);
  custSyncSettlementPayments_(ss, now);
  custWriteCustomerLedger_(shLedger, hLedger, ledgerRecords);
  custWriteCustomerAllocations_(shAlloc, hAlloc, allocations);
  custUpdateCustomerSummary_(ss, customerState, ledgerRecords, allocations, lateRefundState, now);
  custRefreshInvoiceControlCustomerDropdown_();

  SpreadsheetApp.flush();
  const refundCarry = custRound2_(Object.keys(lateRefundState).reduce((sum,k)=>sum+custNum_(lateRefundState[k].outstanding),0));
  SpreadsheetApp.getActive().toast(
    'CUSTOMER V2.2 OK · ' + ledgerRecords.length + ' ledger · ' + allocations.length + ' alloc' +
    (refundCarry > CUST_CFG.EPSILON ? ' · refund carry ' + refundCarry.toFixed(2) : ''),
    'Customer V2.2', 6
  );
  return {customers: customerState.customers.length, ledgerRows: ledgerRecords.length, allocations: allocations.length, lateRefundOutstanding: refundCarry};
}

/*********************** CUSTOMER MASTER ************************/
function custReadCustomers_(sheet, headers) {
  const lastRow = sheet.getLastRow();
  const state = {customers:[], customersByKey:{}, prefixMap:{}, knownPrefixes:[]};
  if (lastRow < CUST_CFG.DATA_START_ROW) return state;
  const values = sheet.getRange(CUST_CFG.DATA_START_ROW,1,lastRow-CUST_CFG.DATA_START_ROW+1,sheet.getLastColumn()).getValues();
  const owners = {};
  values.forEach((row,i) => {
    const rowNumber = CUST_CFG.DATA_START_ROW + i;
    const key = custCanonicalKey_(row[headers['Customer Key']-1]);
    const name = custStr_(row[headers['Customer Name']-1]);
    const status = custUpper_(row[headers['Status']-1]) || 'ACTIVE';
    const rawPrefixes = custStr_(row[headers['Prefixes']-1]);
    if (!key && !name && !rawPrefixes) return;
    if (!key) throw new Error('Customers row ' + rowNumber + ': Customer Key is blank.');
    if (state.customersByKey[key]) throw new Error('DUPLICATE CUSTOMER KEY: ' + key);
    const prefixes = custParsePrefixes_(rawPrefixes);
    const local = {};
    prefixes.forEach(p => { if (local[p]) throw new Error('DUPLICATE PREFIX: ' + p + ' appears twice in ' + key); local[p]=true; });
    const customer = {
      rowNumber, customerKey:key, customerName:name, status, prefixes,
      payoutFeeRate:custPercent_(row[headers['Payout Fee %']-1]),
      defaultPayoutProvider:custUpper_(row[headers['Default Payout Provider']-1]),
      defaultPayoutAccountKey:custCanonicalKey_(row[headers['Default Payout Account Key']-1])
    };
    state.customers.push(customer); state.customersByKey[key]=customer;
    prefixes.forEach(p => {
      if (owners[p]) throw new Error('DUPLICATE PREFIX: ' + p + ' belongs to ' + owners[p].customerKey + ' and ' + key);
      owners[p]=customer; state.prefixMap[p]=customer;
    });
  });
  state.knownPrefixes = Object.keys(state.prefixMap).sort((a,b)=>b.length-a.length);
  return state;
}

/*********************** PAYPAL READ + LEDGER ************************/
function custReadPaypalRows_(sheet, headers, customerState) {
  const lastRow = sheet.getLastRow();
  if (lastRow < CUST_CFG.DATA_START_ROW) return [];
  const values = sheet.getRange(CUST_CFG.DATA_START_ROW,1,lastRow-CUST_CFG.DATA_START_ROW+1,sheet.getLastColumn()).getValues();
  const out=[];
  values.forEach((row,i)=>{
    const txId=custStr_(row[headers['Transaction ID']-1]);
    const accountKey=custCanonicalKey_(row[headers['Account Key']-1]);
    if (!txId || !accountKey) return;
    const invoiceId=custStr_(row[headers['Invoice ID']-1]);
    let prefix=custCanonicalPrefix_(row[headers['Invoice Prefix']-1]);
    if ((!prefix || !customerState.prefixMap[prefix]) && invoiceId) {
      const p=custMatchInvoiceIdToKnownPrefix_(invoiceId, customerState.knownPrefixes); if (p) prefix=p;
    }
    out.push({
      sourceRow:CUST_CFG.DATA_START_ROW+i, accountKey, transactionDate:custDate_(row[headers['Initiation Date']-1]),
      transactionId:txId, referenceId:custStr_(row[headers['Reference ID']-1]), eventCode:custUpper_(row[headers['Event Code']-1]),
      status:custUpper_(row[headers['Status']-1]), invoiceId, invoicePrefix:prefix,
      buyerName:custStr_(row[headers['Buyer Name']-1]), item:custStr_(row[headers['Item']-1]),
      transactionAmount:custNum_(row[headers['Transaction Amount']-1]), feeAmount:custNum_(row[headers['Fee Amount']-1]),
      balanceImpact:custNum_(row[headers['Balance Impact']-1]), currency:custUpper_(row[headers['Currency']-1]),
      uniqueKey:custStr_(row[headers['Unique Key']-1])
    });
  });
  return out;
}

function custBuildCustomerLedger_(paypalRows, customerState, oldLedgerState, now) {
  const byTx={};
  paypalRows.forEach(r=>{
    const k=custPaypalTxIndexKey_(r.accountKey,r.transactionId);
    if (!byTx[k] || (!byTx[k].invoicePrefix && r.invoicePrefix) || (!custIsIncomingPayment_(byTx[k]) && custIsIncomingPayment_(r))) byTx[k]=r;
  });
  const rootSales={};
  paypalRows.forEach(r=>{ if (custIsIncomingPayment_(r)) rootSales[custPaypalTxIndexKey_(r.accountKey,r.transactionId)]=r; });
  const eventsByRoot={};
  paypalRows.forEach(r=>{
    const root=custResolveRootSale_(r,byTx,rootSales); if(!root) return;
    const k=custPaypalTxIndexKey_(root.accountKey,root.transactionId); (eventsByRoot[k]||(eventsByRoot[k]=[])).push(r);
  });

  const records=[]; const allSourceRecords=[];
  Object.keys(rootSales).forEach(rootKey=>{
    const root=rootSales[rootKey]; const events=eventsByRoot[rootKey]||[root];
    const customer=custResolveCustomerForPaypalRow_(root,byTx,customerState);
    let fee=0, refund=0, hold=0;
    events.forEach(ev=>{
      fee += -ev.feeAmount;
      if (Math.abs(ev.feeAmount)<CUST_CFG.EPSILON && custIsStandaloneFeeCharge_(ev.eventCode) && ev.transactionAmount<0) fee += Math.abs(ev.transactionAmount);
      if (Math.abs(ev.feeAmount)<CUST_CFG.EPSILON && custIsFeeRecovery_(ev.eventCode) && ev.transactionAmount>0) fee -= Math.abs(ev.transactionAmount);
      if (custIsCustomerLossEvent_(ev.eventCode) && ev.transactionAmount<0) refund += Math.abs(ev.transactionAmount);
      if (custIsCustomerRecoveryEvent_(ev.eventCode) && ev.transactionAmount>0) refund -= Math.abs(ev.transactionAmount);
      if (custIsHoldEvent_(ev.eventCode)) hold += custEventMagnitude_(ev);
      if (custIsHoldReleaseEvent_(ev.eventCode)) hold -= custEventMagnitude_(ev);
    });
    fee=Math.max(0,custRound2_(fee)); refund=Math.max(0,custRound2_(refund));
    const gross=Math.max(0,custRound2_(root.transactionAmount));
    const net=Math.max(0,custRound2_(gross-fee-refund));
    hold=Math.min(Math.max(0,custRound2_(hold)),net);
    const sourceAvailable=Math.max(0,custRound2_(net-hold));
    const common={
      rootKey, accountKey:root.accountKey, transactionDate:root.transactionDate, transactionId:root.transactionId,
      currency:root.currency, grossAmount:gross, paypalFee:fee, refundAmount:refund, netAmount:net,
      holdAmount:hold, sourceAvailable,
      customerKey:customer?customer.customerKey:'', customerName:customer?customer.customerName:'',
      invoiceId:root.invoiceId, invoicePrefix:custResolvePrefixForPaypalRow_(root,byTx,customerState)||'',
      referenceId:root.referenceId, buyerName:root.buyerName, item:root.item, uniqueKey:root.uniqueKey
    };
    allSourceRecords.push(Object.assign({},common));
    if (!customer) return;
    const ledgerKey='CL|'+root.accountKey+'|'+root.transactionId;
    records.push({
      ledgerKey, transactionDate:root.transactionDate, customerKey:customer.customerKey, customerName:customer.customerName,
      paypalAccountKey:root.accountKey, paypalTransactionId:root.transactionId, referenceId:root.referenceId,
      invoiceId:root.invoiceId, invoicePrefix:common.invoicePrefix, item:root.item, buyerName:root.buyerName,
      eventType:'SALE', grossAmount:gross, paypalFee:fee, refundAmount:refund, netAmount:net, currency:root.currency,
      sourceAvailable, allocatedAmount:0, invoicedAmount:0, remainingAmount:sourceAvailable,
      sourceUniqueKey:root.uniqueKey, lastSync:now, status:'PENDING', note:''
    });
  });
  records.sort(custSortDateDesc_('transactionDate','ledgerKey'));
  allSourceRecords.sort((a,b)=>custTime_(a.transactionDate)-custTime_(b.transactionDate));
  return {records, allSourceRecords};
}

/*********************** PAYPAL RELATIONS ************************/
function custResolveRootSale_(row, byTx, rootSales) {
  let cur=row; const seen={};
  for(let i=0;i<10;i++){
    if(!cur) return null; const k=custPaypalTxIndexKey_(cur.accountKey,cur.transactionId); if(seen[k]) return null; seen[k]=true;
    if(rootSales[k]) return rootSales[k]; if(!cur.referenceId) return null;
    cur=byTx[custPaypalTxIndexKey_(cur.accountKey,cur.referenceId)];
  }
  return null;
}
function custResolveCustomerForPaypalRow_(row, byTx, customerState) {
  let cur=row; const seen={};
  for(let i=0;i<10;i++){
    if(!cur) return null; const k=custPaypalTxIndexKey_(cur.accountKey,cur.transactionId); if(seen[k]) return null; seen[k]=true;
    const p=custResolvePrefixDirect_(cur,customerState); if(p&&customerState.prefixMap[p]) return customerState.prefixMap[p];
    if(!cur.referenceId) return null; cur=byTx[custPaypalTxIndexKey_(cur.accountKey,cur.referenceId)];
  }
  return null;
}
function custResolvePrefixForPaypalRow_(row, byTx, customerState) {
  let cur=row; const seen={};
  for(let i=0;i<10;i++){
    if(!cur) return ''; const k=custPaypalTxIndexKey_(cur.accountKey,cur.transactionId); if(seen[k]) return ''; seen[k]=true;
    const p=custResolvePrefixDirect_(cur,customerState); if(p) return p; if(!cur.referenceId) return '';
    cur=byTx[custPaypalTxIndexKey_(cur.accountKey,cur.referenceId)];
  }
  return '';
}
function custResolvePrefixDirect_(row,state){ const p=custCanonicalPrefix_(row.invoicePrefix); if(p&&state.prefixMap[p]) return p; return custMatchInvoiceIdToKnownPrefix_(row.invoiceId,state.knownPrefixes); }
function custMatchInvoiceIdToKnownPrefix_(invoiceId,prefixes){ const id=custUpper_(invoiceId); if(!id)return ''; for(let i=0;i<prefixes.length;i++){const p=prefixes[i]; if(id===p||id.indexOf(p+'-')===0||id.indexOf(p+'_')===0)return p;} return ''; }

/*********************** EVENT CLASSIFICATION ************************/
function custIsIncomingPayment_(r){ return /^T00\d{2}$/.test(r.eventCode||'') && r.transactionAmount>CUST_CFG.EPSILON && ['DENIED','REVERSED','V'].indexOf(r.status)===-1; }
function custIsCustomerLossEvent_(c){ return ['T1100','T1106','T1107','T1116','T1118','T1120','T1201'].indexOf(c)!==-1; }
function custIsCustomerRecoveryEvent_(c){ return ['T1202','T1205'].indexOf(c)!==-1; }
function custIsStandaloneFeeCharge_(c){ return ['T0106','T0114'].indexOf(c)!==-1; }
function custIsFeeRecovery_(c){ return ['T1108','T1109'].indexOf(c)!==-1; }
function custIsHoldEvent_(c){ return ['T1110','T1500','T1501','T1502','T1503','T2101','T2103','T2105','T2107','T2111','T2114','T2120'].indexOf(c)!==-1; }
function custIsHoldReleaseEvent_(c){ return ['T1105','T1111','T2102','T2104','T2106','T2108','T2112','T2115','T2121'].indexOf(c)!==-1; }
function custEventMagnitude_(e){ return Math.abs(e.transactionAmount)>=CUST_CFG.EPSILON?Math.abs(e.transactionAmount):Math.abs(e.balanceImpact); }

/*********************** CENTRAL TX READ ************************/
function custReadTransactionRows_(sheet, h){
  const last=sheet.getLastRow(); if(last<CUST_CFG.DATA_START_ROW)return [];
  return sheet.getRange(CUST_CFG.DATA_START_ROW,1,last-CUST_CFG.DATA_START_ROW+1,sheet.getLastColumn()).getValues().map((r,i)=>({
    sourceRow:CUST_CFG.DATA_START_ROW+i, transactionKey:custStr_(r[h['Transaction Key']-1]), transactionDate:custDate_(r[h['Transaction Date']-1]),
    provider:custUpper_(r[h['Provider']-1]), accountKey:custCanonicalKey_(r[h['Account Key']-1]), type:custUpper_(r[h['Type']-1]),
    direction:custUpper_(r[h['Direction']-1]), grossAmount:custNum_(r[h['Gross Amount']-1]), netAmount:custNum_(r[h['Net Amount']-1]),
    feeAmount:custNum_(r[h['Fee Amount']-1]), currency:custUpper_(r[h['Currency']-1]), reference:custStr_(r[h['Reference']-1]),
    counterpartyProvider:custUpper_(r[h['Counterparty Provider']-1]), counterpartyAccountKey:custCanonicalKey_(r[h['Counterparty Account Key']-1]),
    reconcileStatus:custUpper_(r[h['Reconcile Status']-1]), matchedTransactionKey:custStr_(r[h['Matched Transaction Key']-1]),
    source:custUpper_(r[h['Source']-1]), sourceId:custStr_(r[h['Source ID']-1]), syncTime:custDate_(r[h['Sync Time']-1]), uniqueKey:custStr_(r[h['Unique Key']-1])
  })).filter(x=>x.transactionKey);
}
/*********************** EXISTING STATES ************************/
function custReadExistingAllocationState_(sheet,h){
  const out={}; const last=sheet.getLastRow(); if(last<CUST_CFG.DATA_START_ROW)return out;
  sheet.getRange(CUST_CFG.DATA_START_ROW,1,last-CUST_CFG.DATA_START_ROW+1,sheet.getLastColumn()).getValues().forEach(r=>{
    const k=custStr_(r[h['Allocation Key']-1]); if(!k)return;
    out[k]={
      customerLedgerKey:custStr_(r[h['Customer Ledger Key']-1]),
      allocatedAmount:custNum_(r[h['Allocated Amount']-1]),
      receivingProvider:custStr_(r[h['Receiving Provider']-1]),
      receivingAccountKey:custStr_(r[h['Receiving Account Key']-1]),
      receivingTransactionKey:custStr_(r[h['Receiving Transaction Key']-1]),
      receivingReference:custStr_(r[h['Receiving Reference']-1]),
      receivingFeeAllocated:custNum_(r[h['Receiving Fee Allocated']-1]),
      availableAmount:custNum_(r[h['Available Amount']-1]),
      availableDate:custDate_(r[h['Available Date']-1]),
      allocationStatus:custStr_(r[h['Allocation Status']-1]),
      invoicedAmount:custNum_(r[h['Invoiced Amount']-1]),
      settlementKey:custStr_(r[h['Settlement Key']-1])
    };
  });
  return out;
}

/*********************** LATE REFUND STATE ************************/
/*
 * No new sheet is needed.
 *
 * First SALE Settlement Item for a ledger = historical snapshot at the time
 * the sale first entered an issued invoice.
 *
 * Current Refund Amount - first invoice Refund Amount = refund delta that
 * appeared after invoicing. LATE_REFUND_ADJUSTMENT Settlement Items are the
 * immutable consumption history for that delta.
 */
function custBuildLateRefundState_(ss, ledgerRows){
  const shSet=custRequireSheet_(ss,CUST_CFG.SETTLEMENTS_SHEET), hs=custHeaderMap_(shSet);
  const shItems=custRequireSheet_(ss,CUST_CFG.SETTLEMENT_ITEMS_SHEET), hi=custHeaderMap_(shItems);
  const sets=custReadRowsObjects_(shSet,hs), items=custReadRowsObjects_(shItems,hi);
  const active={};
  sets.forEach(s=>{
    const key=custStr_(s['Settlement Key']); if(!key||custUpper_(s['Status'])==='CANCELLED')return;
    active[key]={issuedAt:custDate_(s['Issued At']||s['Created At'])};
  });

  const byLedger={};
  items.forEach(it=>{
    const sk=custStr_(it['Settlement Key']); if(!active[sk])return;
    const lk=custStr_(it['Customer Ledger Key']); if(!lk)return;
    (byLedger[lk]||(byLedger[lk]=[])).push(it);
  });

  const out={};
  ledgerRows.forEach(l=>{
    const ledgerKey=custLedgerField_(l,'ledgerKey','Ledger Key'); if(!ledgerKey)return;
    const history=byLedger[ledgerKey]||[];
    const sales=history.filter(it=>custUpper_(it['Line Type']||'SALE')==='SALE');
    if(!sales.length)return;

    sales.sort((a,b)=>{
      const ad=custDate_(a['Created At'])||active[custStr_(a['Settlement Key'])].issuedAt;
      const bd=custDate_(b['Created At'])||active[custStr_(b['Settlement Key'])].issuedAt;
      return custTime_(ad)-custTime_(bd);
    });

    const first=sales[0];
    const firstSettlement=active[custStr_(first['Settlement Key'])]||{};
    const firstInvoiceAt=custDate_(first['Created At'])||firstSettlement.issuedAt||null;
    const baselineRefund=Math.max(0,custNum_(first['Refund Amount']));
    const currentRefund=Math.max(0,custNum_(custLedgerField_(l,'refundAmount','Refund Amount')));
    const totalLate=custRound2_(Math.max(0,currentRefund-baselineRefund));

    let consumed=0;
    history.filter(it=>custUpper_(it['Line Type'])==='LATE_REFUND_ADJUSTMENT').forEach(it=>{
      const explicit=Math.max(0,custNum_(it['Refund Amount']));
      const fallback=Math.max(0,-custNum_(it['Amount Included']));
      consumed+=explicit>CUST_CFG.EPSILON?explicit:fallback;
    });
    consumed=custRound2_(consumed);

    out[ledgerKey]={
      ledgerKey:ledgerKey,
      customerKey:custLedgerField_(l,'customerKey','Customer Key'),
      baselineRefund:baselineRefund,
      currentRefund:currentRefund,
      totalLate:totalLate,
      consumed:consumed,
      outstanding:custRound2_(Math.max(0,totalLate-consumed)),
      overConsumed:custRound2_(Math.max(0,consumed-totalLate)),
      firstInvoiceAt:firstInvoiceAt,
      currency:custUpper_(custLedgerField_(l,'currency','Currency')),
      paypalTransactionId:custLedgerField_(l,'paypalTransactionId','PayPal Transaction ID')
    };
  });
  return out;
}

function custLedgerField_(row,prop,header){
  if(row&&Object.prototype.hasOwnProperty.call(row,prop))return row[prop];
  return row?row[header]:'';
}

/*
 * Refund event map is only used to respect the invoice Through Date.
 * The accounting amount still comes from the immutable invoice snapshot vs
 * the current ledger refund delta above.
 */
function custBuildRefundEventsByLedger_(ss){
  const shCustomers=custRequireSheet_(ss,CUST_CFG.CUSTOMERS_SHEET), hc=custHeaderMap_(shCustomers);
  const customerState=custReadCustomers_(shCustomers,hc);
  const shPaypal=custRequireSheet_(ss,CUST_CFG.PAYPAL_SHEET), hp=custHeaderMap_(shPaypal);
  const paypalRows=custReadPaypalRows_(shPaypal,hp,customerState);
  const byTx={}, rootSales={};
  paypalRows.forEach(r=>{
    const k=custPaypalTxIndexKey_(r.accountKey,r.transactionId);
    if(!byTx[k]||(!byTx[k].invoicePrefix&&r.invoicePrefix)||(!custIsIncomingPayment_(byTx[k])&&custIsIncomingPayment_(r)))byTx[k]=r;
    if(custIsIncomingPayment_(r))rootSales[k]=r;
  });
  const out={};
  paypalRows.forEach(ev=>{
    let delta=0;
    if(custIsCustomerLossEvent_(ev.eventCode)&&ev.transactionAmount<0)delta=Math.abs(ev.transactionAmount);
    else if(custIsCustomerRecoveryEvent_(ev.eventCode)&&ev.transactionAmount>0)delta=-Math.abs(ev.transactionAmount);
    else return;
    const root=custResolveRootSale_(ev,byTx,rootSales); if(!root)return;
    const customer=custResolveCustomerForPaypalRow_(root,byTx,customerState); if(!customer)return;
    const lk='CL|'+root.accountKey+'|'+root.transactionId;
    (out[lk]||(out[lk]=[])).push({date:ev.transactionDate,delta:custRound2_(delta),transactionId:ev.transactionId,uniqueKey:ev.uniqueKey});
  });
  Object.keys(out).forEach(k=>out[k].sort((a,b)=>custTime_(a.date)-custTime_(b.date)));
  return out;
}

function custBuildEligibleLateRefundAdjustments_(ss,customer,ledgerByKey,lateState,throughKey){
  const eventMap=custBuildRefundEventsByLedger_(ss);
  const entries=[];
  Object.keys(lateState).forEach(lk=>{
    const st=lateState[lk];
    if(custCanonicalKey_(st.customerKey)!==customer.customerKey||st.outstanding<=CUST_CFG.EPSILON)return;
    const l=ledgerByKey[lk]; if(!l)return;
    const saleKey=custBusinessDateKey_(custDate_(l['Transaction Date'])); if(!saleKey||saleKey>throughKey)return;

    const events=eventMap[lk]||[];
    let lateThrough=0, latestDate=null, sawDated=false;
    events.forEach(ev=>{
      const d=custDate_(ev.date); if(!d)return;
      if(st.firstInvoiceAt&&d.getTime()<=st.firstInvoiceAt.getTime())return;
      const k=custBusinessDateKey_(d); if(!k||k>throughKey)return;
      sawDated=true; lateThrough+=custNum_(ev.delta);
      if(!latestDate||d>latestDate)latestDate=d;
    });
    lateThrough=custRound2_(Math.max(0,lateThrough));

    /*
     * Normal forward invoices use event-dated amount. Fallback to the current
     * outstanding only when PayPal supplied no usable refund event date.
     */
    let eligible=sawDated?custRound2_(Math.max(0,lateThrough-st.consumed)):st.outstanding;
    eligible=custRound2_(Math.min(st.outstanding,eligible));
    if(eligible<=CUST_CFG.EPSILON)return;

    entries.push({
      ledgerKey:lk,
      outstanding:eligible,
      eventDate:latestDate||custDate_(l['Transaction Date']),
      currency:custUpper_(l['Currency']),
      ledger:l,
      baselineRefund:st.baselineRefund,
      currentRefund:st.currentRefund
    });
  });
  entries.sort((a,b)=>custTime_(a.eventDate)-custTime_(b.eventDate)||a.ledgerKey.localeCompare(b.ledgerKey));
  return entries;
}

/*********************** PHASE 2: PAYPAL -> GREY ************************/
function custBuildCustomerAllocations_(allSourceRecords, ledgerRecords, txRows, oldState, lateRefundState, now){
  const ledgerByTx={}; ledgerRecords.forEach(r=>ledgerByTx[custPaypalTxIndexKey_(r.paypalAccountKey,r.paypalTransactionId)]=r);
  const txByKey={}; txRows.forEach(t=>txByKey[t.transactionKey]=t);

  /* Historical invoiced allocations are a floor and can never shrink. */
  const lockedByLedger={};
  Object.keys(oldState).forEach(k=>{
    const o=oldState[k]; if(!o.settlementKey||!o.customerLedgerKey)return;
    lockedByLedger[o.customerLedgerKey]=custRound2_((lockedByLedger[o.customerLedgerKey]||0)+custNum_(o.allocatedAmount));
  });

  const pools={};
  allSourceRecords.forEach(src=>{
    const ledgerKey='CL|'+src.accountKey+'|'+src.transactionId;
    const late=lateRefundState[ledgerKey]||null;
    let sourceForAllocation=custNum_(src.sourceAvailable);

    /*
     * If this sale was already invoiced, a later refund is NOT allowed to
     * rewrite the historical PayPal withdrawal allocation. Restore only the
     * late-refund delta to the FIFO source pool. The same delta is charged as
     * LATE_REFUND_ADJUSTMENT on a later invoice instead.
     */
    if(late&&late.totalLate>CUST_CFG.EPSILON){
      const preLateNet=Math.max(0,custRound2_(custNum_(src.grossAmount)-custNum_(src.paypalFee)-custNum_(late.baselineRefund)));
      sourceForAllocation=Math.min(preLateNet,custRound2_(sourceForAllocation+late.totalLate));
    }
    sourceForAllocation=Math.max(sourceForAllocation,custNum_(lockedByLedger[ledgerKey]));
    sourceForAllocation=custRound2_(sourceForAllocation);
    if(sourceForAllocation<=CUST_CFG.EPSILON)return;

    const k=custMoneyPoolKey_(src.accountKey,src.currency);
    (pools[k]||(pools[k]=[])).push({src,ledgerKey,remaining:sourceForAllocation});
  });
  Object.keys(pools).forEach(k=>{pools[k].sort((a,b)=>custTime_(a.src.transactionDate)-custTime_(b.src.transactionDate));pools[k]._cursor=0;});

  const withdrawals=txRows.filter(t=>
    t.provider==='PAYPAL'&&t.type==='PAYPAL_WITHDRAWAL'&&t.direction==='OUT'&&
    Math.abs(t.grossAmount)>CUST_CFG.EPSILON&&t.reconcileStatus!=='RETURNED'
  ).sort((a,b)=>custTime_(a.transactionDate)-custTime_(b.transactionDate));

  const out=[]; const generated={};
  withdrawals.forEach(w=>{
    const pool=pools[custMoneyPoolKey_(w.accountKey,w.currency)]; if(!pool||!w.transactionDate)return;
    const gross=Math.abs(custRound2_(w.grossAmount)); let rem=gross;
    const receivingTx=w.matchedTransactionKey?txByKey[w.matchedTransactionKey]:null;
    const match=custDescribeReceivingMatch_(w,receivingTx);
    const receivingFeeTotal=match.available?custGetReceivingTransactionFee_(receivingTx):0;

    while(rem>CUST_CFG.EPSILON){
      const idx=pool._cursor||0; if(idx>=pool.length)break;
      const item=pool[idx], src=item.src;
      if(!src.transactionDate||src.transactionDate.getTime()>w.transactionDate.getTime())break;
      if(item.remaining<=CUST_CFG.EPSILON){pool._cursor=idx+1;continue;}
      const take=custRound2_(Math.min(item.remaining,rem)); if(take<=CUST_CFG.EPSILON)break;
      const ledger=ledgerByTx[custPaypalTxIndexKey_(src.accountKey,src.transactionId)]||null;
      if(ledger){
        const key='CA|'+w.transactionKey+'|'+ledger.ledgerKey; const old=oldState[key]||null;
        const locked=!!(old&&old.settlementKey);
        if(locked&&Math.abs(old.allocatedAmount-take)>0.01)throw new Error('LOCKED ALLOCATION CHANGED: '+key);

        const computedFee=match.available&&gross>CUST_CFG.EPSILON?custRound2_(receivingFeeTotal*take/gross):0;
        const computedAvailable=match.available?Math.max(0,custRound2_(take-computedFee)):0;
        const receiveFee=locked?custNum_(old.receivingFeeAllocated):computedFee;
        const available=locked?custNum_(old.availableAmount):computedAvailable;
        const invoiced=old?Math.min(available,Math.max(0,custNum_(old.invoicedAmount))):0;

        out.push({
          allocationKey:key, allocationDate:w.transactionDate, customerKey:ledger.customerKey, customerLedgerKey:ledger.ledgerKey,
          paypalAccountKey:w.accountKey, paypalSourceTransactionId:ledger.paypalTransactionId,
          paypalWithdrawalTransactionKey:w.transactionKey,
          paypalWithdrawalTransactionId:w.sourceId||w.reference||w.transactionKey,
          paypalWithdrawalDate:w.transactionDate, paypalWithdrawalStatus:w.reconcileStatus,
          allocatedAmount:locked?custNum_(old.allocatedAmount):take, currency:w.currency,
          receivingProvider:locked?(old.receivingProvider||match.provider):match.provider,
          receivingAccountKey:locked?(old.receivingAccountKey||match.accountKey):match.accountKey,
          receivingTransactionKey:locked?(old.receivingTransactionKey||match.transactionKey):match.transactionKey,
          receivingReference:locked?(old.receivingReference||match.reference):match.reference,
          receivingFeeAllocated:custRound2_(receiveFee), availableAmount:custRound2_(available),
          availableDate:locked?(old.availableDate||match.date):(match.available?match.date:null),
          allocationStatus:locked?(custUpper_(old.allocationStatus)||'AVAILABLE'):match.status,
          invoicedAmount:custRound2_(invoiced), remainingAvailable:custRound2_(Math.max(0,available-invoiced)),
          settlementKey:old?old.settlementKey:'', sourceUniqueKey:w.uniqueKey, lastSync:now,
          note:locked&&lateRefundState[ledger.ledgerKey]&&lateRefundState[ledger.ledgerKey].totalLate>CUST_CFG.EPSILON?'LOCKED · LATE REFUND HANDLED BY ADJUSTMENT':''
        });
        generated[key]=true;
      }
      item.remaining=custRound2_(item.remaining-take); rem=custRound2_(rem-take);
      if(item.remaining<=CUST_CFG.EPSILON)pool._cursor=idx+1;
    }
  });
  Object.keys(oldState).forEach(k=>{if(oldState[k].settlementKey&&!generated[k])throw new Error('LOCKED ALLOCATION DISAPPEARED: '+k);});
  out.sort(custSortDateDesc_('paypalWithdrawalDate','allocationKey'));
  return out;
}
function custDescribeReceivingMatch_(w,tx){
  const fallbackProvider=w.counterpartyProvider||'';
  const fallbackAccount=w.counterpartyAccountKey||'';

  /*
   * IMPORTANT V2.1:
   * Flow.gs is the source of truth for reconciliation.
   * If the PayPal withdrawal is ARRIVED and Matched Transaction Key resolves
   * to a real non-PayPal receiving transaction, customer funds are AVAILABLE.
   *
   * Do NOT require Direction=IN here. The proven legacy customer code did not
   * require that extra condition, and some normalized receiving rows may not
   * carry the direction exactly as expected even though Flow has already
   * reconciled them correctly.
   *
   * Current proven Grey route remains explicitly accepted:
   *   PAYPAL_WITHDRAWAL ARRIVED -> GREY / GREY_FUNDING
   */
  const isProvenGreyMatch = !!(
    w.reconcileStatus==='ARRIVED' &&
    tx &&
    tx.provider==='GREY' &&
    tx.type==='GREY_FUNDING'
  );

  const isGenericReceivingMatch = !!(
    w.reconcileStatus==='ARRIVED' &&
    tx &&
    tx.provider &&
    tx.provider!=='PAYPAL'
  );

  if(isProvenGreyMatch || isGenericReceivingMatch){
    return {
      available:true,
      status:'AVAILABLE',
      provider:tx.provider,
      accountKey:tx.accountKey,
      transactionKey:tx.transactionKey,
      reference:tx.reference,
      date:tx.transactionDate
    };
  }

  let status='IN_TRANSIT';
  if(w.reconcileStatus==='OVERDUE')status='OVERDUE';
  else if(w.reconcileStatus==='UNMATCHED')status='UNMATCHED';
  else if(w.reconcileStatus==='ARRIVED'&&!tx)status='ERROR_NO_RECEIVING_MATCH';
  else if(w.reconcileStatus==='ARRIVED'&&tx)status='ERROR_RECEIVING_MATCH_REJECTED';

  return {
    available:false,
    status,
    provider:fallbackProvider,
    accountKey:fallbackAccount,
    transactionKey:tx?tx.transactionKey:'',
    reference:tx?tx.reference:'',
    date:tx?tx.transactionDate:null
  };
}
function custGetReceivingTransactionFee_(tx){
  if(!tx)return 0;
  const f=Math.abs(custNum_(tx.feeAmount));
  if(f>CUST_CFG.EPSILON)return custRound2_(f);
  return custRound2_(Math.max(0,Math.abs(custNum_(tx.grossAmount))-Math.abs(custNum_(tx.netAmount))));
}
/*********************** FINAL LEDGER V2 ************************/
function custFinalizeFullLedger_(ledger, allocations, transfers){
  const byLedger={}; allocations.forEach(a=>(byLedger[a.customerLedgerKey]||(byLedger[a.customerLedgerKey]=[])).push(a));
  ledger.forEach(r=>{
    const aa=byLedger[r.ledgerKey]||[];
    let allocated=0,invoiced=0,ready=0,inTransit=0;
    aa.forEach(a=>{
      allocated+=custNum_(a.allocatedAmount); invoiced+=custNum_(a.invoicedAmount);
      if(a.allocationStatus==='AVAILABLE')ready+=custNum_(a.remainingAvailable); else inTransit+=custNum_(a.allocatedAmount);
    });
    r.allocatedAmount=custRound2_(allocated);
    r.invoicedAmount=custRound2_(invoiced);
    r.remainingAmount=custRound2_(Math.max(0,r.sourceAvailable-r.allocatedAmount));
    if(r.invoicedAmount>CUST_CFG.EPSILON&&ready<=CUST_CFG.EPSILON&&r.remainingAmount<=CUST_CFG.EPSILON&&inTransit<=CUST_CFG.EPSILON)r.status='INVOICED';
    else if(ready>CUST_CFG.EPSILON&&(r.remainingAmount>CUST_CFG.EPSILON||inTransit>CUST_CFG.EPSILON))r.status='MIXED';
    else if(ready>CUST_CFG.EPSILON)r.status='AVAILABLE';
    else if(inTransit>CUST_CFG.EPSILON)r.status='IN_TRANSIT';
    else r.status='PENDING';
  });
}

/*********************** WRITE DERIVED TABLES ************************/
function custWriteCustomerLedger_(sheet,h,records){
  custClearData_(sheet); if(!records.length)return; const lc=sheet.getLastColumn();
  const rows=records.map(r=>{const x=new Array(lc).fill('');const p=(k,v)=>custPut_(x,h,k,v);
    p('Ledger Key',r.ledgerKey);p('Transaction Date',r.transactionDate);p('Customer Key',r.customerKey);p('Customer Name',r.customerName);
    p('PayPal Account Key',r.paypalAccountKey);p('PayPal Transaction ID',r.paypalTransactionId);p('Reference ID',r.referenceId);
    p('Invoice ID',r.invoiceId);p('Invoice Prefix',r.invoicePrefix);p('Event Type',r.eventType);p('Gross Amount',r.grossAmount);
    p('PayPal Fee',r.paypalFee);p('Refund Amount',r.refundAmount);p('Net Amount',r.netAmount);p('Currency',r.currency);
    p('Allocated Amount',r.allocatedAmount);p('Invoiced Amount',r.invoicedAmount);p('Remaining Amount',r.remainingAmount);
    p('Source Unique Key',r.sourceUniqueKey);p('Last Sync',r.lastSync);p('Status',r.status);p('Note',r.note||'');return x;});
  sheet.getRange(CUST_CFG.DATA_START_ROW,1,rows.length,lc).setValues(rows);
  custFormatCols_(sheet,h,rows.length,['Transaction Date','Last Sync'],['Gross Amount','PayPal Fee','Refund Amount','Net Amount','Allocated Amount','Invoiced Amount','Remaining Amount']);
}
function custWriteCustomerAllocations_(sheet,h,records){
  custClearData_(sheet); if(!records.length)return; const lc=sheet.getLastColumn();
  const rows=records.map(r=>{const x=new Array(lc).fill('');const p=(k,v)=>custPut_(x,h,k,v);
    p('Allocation Key',r.allocationKey);p('Allocation Date',r.allocationDate);p('Customer Key',r.customerKey);p('Customer Ledger Key',r.customerLedgerKey);
    p('PayPal Account Key',r.paypalAccountKey);p('PayPal Source Transaction ID',r.paypalSourceTransactionId);
    p('PayPal Withdrawal Transaction Key',r.paypalWithdrawalTransactionKey);p('PayPal Withdrawal Transaction ID',r.paypalWithdrawalTransactionId);
    p('PayPal Withdrawal Date',r.paypalWithdrawalDate);p('PayPal Withdrawal Status',r.paypalWithdrawalStatus);
    p('Allocated Amount',r.allocatedAmount);p('Currency',r.currency);p('Receiving Provider',r.receivingProvider);p('Receiving Account Key',r.receivingAccountKey);
    p('Receiving Transaction Key',r.receivingTransactionKey);p('Receiving Reference',r.receivingReference);p('Receiving Fee Allocated',r.receivingFeeAllocated);
    p('Available Amount',r.availableAmount);p('Available Date',r.availableDate);p('Allocation Status',r.allocationStatus);
    p('Invoiced Amount',r.invoicedAmount);p('Remaining Available',r.remainingAvailable);p('Settlement Key',r.settlementKey);
    p('Source Unique Key',r.sourceUniqueKey);p('Last Sync',r.lastSync);p('Note',r.note||'');return x;});
  sheet.getRange(CUST_CFG.DATA_START_ROW,1,rows.length,lc).setValues(rows);
  custFormatCols_(sheet,h,rows.length,['Allocation Date','PayPal Withdrawal Date','Available Date','Last Sync'],['Allocated Amount','Receiving Fee Allocated','Available Amount','Invoiced Amount','Remaining Available']);
}
/*********************** SETTLEMENT PAYMENT SYNC ************************/
function custSyncSettlementPayments_(ss, now){
  const shSet=custRequireSheet_(ss,CUST_CFG.SETTLEMENTS_SHEET), shPay=custRequireSheet_(ss,CUST_CFG.PAYOUTS_SHEET);
  custEnsureHeaders_(shPay,['Customer Key','Settlement Key']);
  const hs=custHeaderMap_(shSet), hp=custHeaderMap_(shPay);
  const settlements=custReadRowsObjects_(shSet,hs), payouts=custReadRowsObjects_(shPay,hp);
  const paidBySet={};
  payouts.forEach(p=>{if(custUpper_(p['Status'])!=='PAID'||!custStr_(p['Settlement Key']))return; const k=custStr_(p['Settlement Key']); paidBySet[k]=(paidBySet[k]||0)+custNum_(p['Amount']);});
  settlements.forEach(s=>{
    const key=custStr_(s['Settlement Key']); if(!key||custUpper_(s['Status'])==='CANCELLED')return;
    const due=custNum_(s['Amount Due']); const paid=custRound2_(Math.min(due,paidBySet[key]||0)); const bal=custRound2_(Math.max(0,due-paid));
    let status=custUpper_(s['Status'])||'ISSUED'; if(paid>CUST_CFG.EPSILON)status=bal<=CUST_CFG.EPSILON?'PAID':'PARTIALLY_PAID';
    custSetCellByHeader_(shSet,hs,s._row,'Amount Paid',paid);custSetCellByHeader_(shSet,hs,s._row,'Outstanding Payable',bal);
    custSetCellByHeader_(shSet,hs,s._row,'Status',status);custSetCellByHeader_(shSet,hs,s._row,'Last Sync',now);
    if(status==='PAID'&&!s['Paid At'])custSetCellByHeader_(shSet,hs,s._row,'Paid At',now);
  });
}

/*********************** CUSTOMER SUMMARY ************************/
function custUpdateCustomerSummary_(ss, state, ledger, allocations, lateRefundState, now){
  const sh=custRequireSheet_(ss,CUST_CFG.CUSTOMERS_SHEET), h=custHeaderMap_(sh);
  const shSet=custRequireSheet_(ss,CUST_CFG.SETTLEMENTS_SHEET), hs=custHeaderMap_(shSet), sets=custReadRowsObjects_(shSet,hs);
  const shPay=custRequireSheet_(ss,CUST_CFG.PAYOUTS_SHEET), hp=custHeaderMap_(shPay), payouts=custReadRowsObjects_(shPay,hp);
  const byCustLedger={}, byCustAlloc={}, refundCarryByCustomer={};
  ledger.forEach(r=>(byCustLedger[r.customerKey]||(byCustLedger[r.customerKey]=[])).push(r));
  allocations.forEach(a=>(byCustAlloc[a.customerKey]||(byCustAlloc[a.customerKey]=[])).push(a));
  Object.keys(lateRefundState||{}).forEach(k=>{
    const st=lateRefundState[k], ck=custCanonicalKey_(st.customerKey); if(!ck)return;
    refundCarryByCustomer[ck]=custRound2_((refundCarryByCustomer[ck]||0)+custNum_(st.outstanding));
  });

  state.customers.forEach(c=>{
    const ll=byCustLedger[c.customerKey]||[], aa=byCustAlloc[c.customerKey]||[];
    let paypalTotal=0,matched=0,available=0,invoiced=0;
    ll.forEach(r=>paypalTotal+=custNum_(r.netAmount));
    aa.forEach(a=>{
      if(a.allocationStatus==='AVAILABLE'){
        matched+=custNum_(a.allocatedAmount); available+=custNum_(a.availableAmount);
      }
      invoiced+=custNum_(a.invoicedAmount);
    });
    paypalTotal=custRound2_(paypalTotal); matched=custRound2_(matched); available=custRound2_(available);
    invoiced=custRound2_(Math.min(available,invoiced));
    const refundCarry=custRound2_(refundCarryByCustomer[c.customerKey]||0);
    const outstanding=custRound2_(Math.max(0,available-invoiced-refundCarry));
    const pending=custRound2_(Math.max(0,paypalTotal-matched));
    let lastSettlement=null,lastPayout=null;
    sets.filter(s=>custCanonicalKey_(s['Customer Key'])===c.customerKey&&custUpper_(s['Status'])!=='CANCELLED').forEach(s=>{const d=custDate_(s['Issued At']||s['Created At']);if(d&&(!lastSettlement||d>lastSettlement))lastSettlement=d;});
    payouts.filter(p=>custCanonicalKey_(p['Customer Key'])===c.customerKey&&custUpper_(p['Status'])==='PAID').forEach(p=>{const d=custDate_(p['Paid At']||p['Payout Date']);if(d&&(!lastPayout||d>lastPayout))lastPayout=d;});
    [['PayPal Total',paypalTotal],['Matched Amount',matched],['Available Amount',available],['Invoiced Amount',invoiced],['Outstanding Amount',outstanding],['Pending in PayPal',pending],['Last Settlement',lastSettlement||''],['Last Payout',lastPayout||''],['Last Sync',now],['Sync Status','OK · '+ll.length+' ledger · '+aa.length+' alloc'+(refundCarry>CUST_CFG.EPSILON?' · refund carry '+refundCarry.toFixed(2):'')]]
      .forEach(([k,v])=>custSetCellByHeader_(sh,h,c.rowNumber,k,v));
  });
}
/*********************** FEE RULES ************************/
function custFindFeeRule_(ss, receivingProvider, payoutProvider, currency, asOfKey){
  const sh=custRequireSheet_(ss,CUST_CFG.FEE_RULES_SHEET), h=custHeaderMap_(sh);
  custRequireHeaders_(CUST_CFG.FEE_RULES_SHEET,h,[
    'Rule Key','Enabled','Receiving Provider','Payout Provider','Currency','Conversion Rate',
    'Transfer Rate','Transfer Minimum','Effective From','Effective To','Priority','Note'
  ]);
  const rows=custReadRowsObjects_(sh,h);
  const rp=custUpper_(receivingProvider), pp=custUpper_(payoutProvider), cc=custUpper_(currency);
  const candidates=rows.filter(r=>{
    if(!custBool_(r['Enabled']))return false;
    const rr=custUpper_(r['Receiving Provider']), pr=custUpper_(r['Payout Provider']), cr=custUpper_(r['Currency']);
    if(rr!==rp&&rr!=='*')return false;
    if(pr!==pp&&pr!=='*')return false;
    if(cr!==cc&&cr!=='*')return false;
    const fk=custConfigDateKey_(ss,r['Effective From']), tk=custConfigDateKey_(ss,r['Effective To']);
    if(fk&&asOfKey<fk)return false;
    if(tk&&asOfKey>tk)return false;
    return true;
  });
  if(!candidates.length)throw new Error('No active Fee Rule for '+rp+' -> '+pp+' / '+cc+' on '+asOfKey);
  candidates.sort((a,b)=>{
    const pd=custNum_(b['Priority'])-custNum_(a['Priority']);
    if(pd)return pd;
    return (custConfigDateKey_(ss,b['Effective From'])||'').localeCompare(custConfigDateKey_(ss,a['Effective From'])||'');
  });
  const r=candidates[0];
  return {
    ruleKey:custStr_(r['Rule Key']),
    receivingProvider:rp,
    payoutProvider:pp,
    currency:cc,
    conversionRate:Math.max(0,custNum_(r['Conversion Rate'])),
    transferRate:Math.max(0,custNum_(r['Transfer Rate'])),
    transferMinimum:Math.max(0,custNum_(r['Transfer Minimum'])),
    priority:custNum_(r['Priority'])
  };
}

/*
 * Pricing rule:
 *   Conversion Cost = base * Conversion Rate
 *   Transfer Cost   = MAX(base * Transfer Rate, Transfer Minimum)
 *   Downstream Cost = Conversion Cost + Transfer Cost
 *
 * Total downstream cost is capped at the available base so an extremely
 * small invoice can never create a negative payable base.
 */
function custCalculateDownstreamPricing_(baseAmount,rule){
  const base=Math.max(0,custNum_(baseAmount));
  if(base<=CUST_CFG.EPSILON){
    return {conversionCost:0,transferCost:0,downstreamCost:0};
  }

  const conversionCost=custRound2_(Math.min(base,base*custNum_(rule.conversionRate)));
  const transferRaw=Math.max(base*custNum_(rule.transferRate),custNum_(rule.transferMinimum));
  const transferCost=custRound2_(Math.min(Math.max(0,base-conversionCost),transferRaw));
  const downstreamCost=custRound2_(Math.min(base,conversionCost+transferCost));

  return {conversionCost,transferCost,downstreamCost};
}

/*********************** CONTROL / STATEMENT ************************/
function previewCustomerStatementFromControl(){
  const ss=SpreadsheetApp.getActive(); syncCustomers_();
  const params=custReadControlParams_(ss), customer=custGetCustomerByKey_(ss,params.customerKey);
  custFillControlCustomerConfig_(ss,customer); custClearControlPreview_(ss);
  const shL=custRequireSheet_(ss,CUST_CFG.LEDGER_SHEET), hL=custHeaderMap_(shL);
  const rows=custReadRowsObjects_(shL,hL).filter(r=>custCanonicalKey_(r['Customer Key'])===params.customerKey&&custBusinessDateKeyInRange_(custDate_(r['Transaction Date']),params.fromKey,params.toKey));
  const selected={}; rows.forEach(r=>selected[custStr_(r['Ledger Key'])]=true);
  const shA=custRequireSheet_(ss,CUST_CFG.ALLOCATION_SHEET), hA=custHeaderMap_(shA);
  const allocs=custReadRowsObjects_(shA,hA).filter(a=>custCanonicalKey_(a['Customer Key'])===params.customerKey&&selected[custStr_(a['Customer Ledger Key'])]);
  let paypalTotal=0,matched=0,available=0,invoiced=0;
  rows.forEach(r=>paypalTotal+=custNum_(r['Net Amount']));
  allocs.forEach(a=>{if(custUpper_(a['Allocation Status'])==='AVAILABLE'){matched+=custNum_(a['Allocated Amount']);available+=custNum_(a['Available Amount']);}invoiced+=custNum_(a['Invoiced Amount']);});
  paypalTotal=custRound2_(paypalTotal);matched=custRound2_(matched);available=custRound2_(available);invoiced=custRound2_(invoiced);
  const outstanding=custRound2_(Math.max(0,available-invoiced)),pending=custRound2_(Math.max(0,paypalTotal-matched));
  const control=custRequireSheet_(ss,CUST_CFG.CONTROL_SHEET);
  [['B14',paypalTotal],['B15',matched],['B16',available],['B17',invoiced],['B18',outstanding],['B19',pending]].forEach(([a,v])=>control.getRange(a).setValue(v));
  custWriteControlDetails_(control,rows.map(r=>({date:custBusinessDateTimeDisplay_(custDate_(r['Transaction Date'])),tx:r['PayPal Transaction ID'],invoice:r['Invoice ID'],item:custLedgerItemLookup_(ss,r['PayPal Transaction ID'],r['PayPal Account Key']),gross:r['Gross Amount'],paypalFee:r['PayPal Fee'],refund:r['Refund Amount'],paypalNet:r['Net Amount'],receivingFee:'',available:'',included:'',status:r['Status']})));
  control.getRange('E21').setValue('OK · US TIME '+CUST_CFG.BUSINESS_TIMEZONE);control.getRange('E22').setValue('STATEMENT READY');
  return {rows:rows.length,paypalTotal,matched,available,invoiced,outstanding,pending};
}

/*********************** INVOICE PREVIEW ************************/
function previewCustomerInvoiceFromControl(){
  const ss=SpreadsheetApp.getActive(); syncCustomers_(); const preview=custBuildInvoicePreview_(ss,true); custRenderInvoicePreview_(ss,preview);
  PropertiesService.getDocumentProperties().setProperty(CUST_CFG.PREVIEW_PROP, JSON.stringify({fingerprint:preview.fingerprint,customerKey:preview.customer.customerKey,from:preview.params.fromKey,to:preview.params.toKey,asOf:preview.params.asOfKey}));
  return preview;
}
function custBuildInvoicePreview_(ss, requireEligible){
  const params=custReadControlParams_(ss), customer=custGetCustomerByKey_(ss,params.customerKey);
  if(customer.status==='DISABLED')throw new Error('Customer is DISABLED.');
  if(!customer.defaultPayoutProvider||!customer.defaultPayoutAccountKey)throw new Error('Customer Default Payout Provider / Default Payout Account Key is required.');

  const shA=custRequireSheet_(ss,CUST_CFG.ALLOCATION_SHEET), hA=custHeaderMap_(shA), allocs=custReadRowsObjects_(shA,hA);
  const shL=custRequireSheet_(ss,CUST_CFG.LEDGER_SHEET), hL=custHeaderMap_(shL), ledgers=custReadRowsObjects_(shL,hL), ledgerByKey={};
  ledgers.forEach(l=>ledgerByKey[custStr_(l['Ledger Key'])]=l);

  const selected=allocs.filter(a=>{
    if(custCanonicalKey_(a['Customer Key'])!==customer.customerKey)return false;
    if(custUpper_(a['Allocation Status'])!=='AVAILABLE')return false;
    if(custNum_(a['Remaining Available'])<=CUST_CFG.EPSILON)return false;
    const l=ledgerByKey[custStr_(a['Customer Ledger Key'])]; if(!l)return false;
    const k=custBusinessDateKey_(custDate_(l['Transaction Date'])); return !!k&&k<=params.toKey;
  });
  if(requireEligible&&!selected.length)throw new Error('No eligible uninvoiced AVAILABLE funds through '+params.toDisplay+'.');

  const providers=[...new Set(selected.map(a=>custUpper_(a['Receiving Provider'])).filter(Boolean))];
  if(providers.length>1)throw new Error('Mixed receiving providers in one invoice are not enabled yet: '+providers.join(', '));
  const receivingProvider=providers[0]||'';
  const currencies=[...new Set(selected.map(a=>custUpper_(a['Currency'])).filter(Boolean))];
  if(currencies.length>1)throw new Error('Invoice preview mixes currencies: '+currencies.join(', '));
  const sourceCurrency=currencies[0]||'USD';
  const feeRule=custFindFeeRule_(ss,receivingProvider,customer.defaultPayoutProvider,sourceCurrency,params.asOfKey);

  const lateState=custBuildLateRefundState_(ss,ledgers);
  const lateEntries=custBuildEligibleLateRefundAdjustments_(ss,customer,ledgerByKey,lateState,params.toKey);
  lateEntries.forEach(x=>{if(x.currency&&x.currency!==sourceCurrency)throw new Error('Late refund currency '+x.currency+' does not match invoice currency '+sourceCurrency+'.');});

  const grouped={}; let gross=0,saleRefund=0,paypalFee=0,receivingFee=0,availableTotal=0,currentAmount=0,carryAmount=0;
  selected.forEach(a=>{
    const l=ledgerByKey[custStr_(a['Customer Ledger Key'])]; if(!l)throw new Error('Missing Customer Ledger: '+a['Customer Ledger Key']);
    const remaining=custNum_(a['Remaining Available']), allocAvail=custNum_(a['Available Amount']), allocSource=custNum_(a['Allocated Amount']);
    const includeRatio=allocAvail>CUST_CFG.EPSILON?Math.min(1,remaining/allocAvail):0;
    const sourceIncluded=custRound2_(allocSource*includeRatio), receiveIncluded=custRound2_(custNum_(a['Receiving Fee Allocated'])*includeRatio);

    /*
     * Once a sale has invoice history, late refund delta is represented only
     * by LATE_REFUND_ADJUSTMENT. Do not let the new refund distort the SALE
     * portion again.
     */
    const lk=custStr_(a['Customer Ledger Key']), lateForSale=lateState[lk]||null;
    const saleRefundBasis=lateForSale?custNum_(lateForSale.baselineRefund):custNum_(l['Refund Amount']);
    const saleNetBasis=Math.max(0,custRound2_(custNum_(l['Gross Amount'])-custNum_(l['PayPal Fee'])-saleRefundBasis));
    const sourceRatio=saleNetBasis>CUST_CFG.EPSILON?sourceIncluded/saleNetBasis:0;
    gross+=custNum_(l['Gross Amount'])*sourceRatio; saleRefund+=saleRefundBasis*sourceRatio; paypalFee+=custNum_(l['PayPal Fee'])*sourceRatio; receivingFee+=receiveIncluded; availableTotal+=remaining;
    const sourceKey=custBusinessDateKey_(custDate_(l['Transaction Date'])), periodClass=sourceKey<params.fromKey?'CARRY_FORWARD':'CURRENT_PERIOD';
    if(periodClass==='CARRY_FORWARD')carryAmount+=remaining;else currentAmount+=remaining;
    if(!grouped[lk])grouped[lk]={ledger:l,allocationKeys:[],receivingFeeAllocated:0,availableAmount:0,amountIncluded:0,sourceIncluded:0,hasCarry:false,hasCurrent:false};
    grouped[lk].allocationKeys.push(custStr_(a['Allocation Key']));
    grouped[lk].receivingFeeAllocated=custRound2_(grouped[lk].receivingFeeAllocated+receiveIncluded);
    grouped[lk].availableAmount=custRound2_(grouped[lk].availableAmount+allocAvail);
    grouped[lk].amountIncluded=custRound2_(grouped[lk].amountIncluded+remaining);
    grouped[lk].sourceIncluded=custRound2_(grouped[lk].sourceIncluded+sourceIncluded);
    if(periodClass==='CARRY_FORWARD')grouped[lk].hasCarry=true;else grouped[lk].hasCurrent=true;
  });

  gross=custRound2_(gross); saleRefund=custRound2_(saleRefund); paypalFee=custRound2_(paypalFee); receivingFee=custRound2_(receivingFee);
  availableTotal=custRound2_(availableTotal); currentAmount=custRound2_(currentAmount); carryAmount=custRound2_(carryAmount);

  const downstreamPricing=custCalculateDownstreamPricing_(availableTotal,feeRule);
  const conversionCost=downstreamPricing.conversionCost;
  const transferCost=downstreamPricing.transferCost;
  const downstreamCost=downstreamPricing.downstreamCost;
  const beforeLateRefund=custRound2_(Math.max(0,availableTotal-downstreamCost));
  const lateRefundOutstanding=custRound2_(lateEntries.reduce((sum,x)=>sum+custNum_(x.outstanding),0));
  const lateRefundApplied=custRound2_(Math.min(beforeLateRefund,lateRefundOutstanding));
  const lateRefundCarryForward=custRound2_(Math.max(0,lateRefundOutstanding-lateRefundApplied));
  const netBeforePayoutFee=custRound2_(Math.max(0,beforeLateRefund-lateRefundApplied));
  const payoutFeeRate=customer.payoutFeeRate, payoutFeeAmount=custRound2_(netBeforePayoutFee*payoutFeeRate), amountDue=custRound2_(Math.max(0,netBeforePayoutFee-payoutFeeAmount));
  const totalRefunds=custRound2_(saleRefund+lateRefundApplied);
  const cRow=custReadCustomerSheetRow_(ss,customer.customerKey);

  const saleItems=Object.keys(grouped).map(k=>{
    const x=grouped[k]; let pc='CURRENT_PERIOD'; if(x.hasCarry&&!x.hasCurrent)pc='CARRY_FORWARD'; else if(x.hasCarry&&x.hasCurrent)pc='MIXED_PERIOD';
    return {lineType:'SALE',ledgerKey:k,allocationKeys:x.allocationKeys,transactionDate:x.ledger['Transaction Date'],transactionDateDisplay:custBusinessDateTimeDisplay_(custDate_(x.ledger['Transaction Date'])),periodClass:pc,
      paypalAccountKey:x.ledger['PayPal Account Key'],paypalTransactionId:x.ledger['PayPal Transaction ID'],referenceId:x.ledger['Reference ID'],invoiceId:x.ledger['Invoice ID'],invoicePrefix:x.ledger['Invoice Prefix'],
      item:custLedgerItemLookup_(ss,x.ledger['PayPal Transaction ID'],x.ledger['PayPal Account Key']),buyerName:'',grossAmount:custNum_(x.ledger['Gross Amount']),paypalFee:custNum_(x.ledger['PayPal Fee']),
      refundAmount:lateState[k]?custNum_(lateState[k].baselineRefund):custNum_(x.ledger['Refund Amount']),
      paypalNetAmount:lateState[k]?Math.max(0,custRound2_(custNum_(x.ledger['Gross Amount'])-custNum_(x.ledger['PayPal Fee'])-custNum_(lateState[k].baselineRefund))):custNum_(x.ledger['Net Amount']),
      receivingFeeAllocated:x.receivingFeeAllocated,availableAmount:x.availableAmount,amountIncluded:x.amountIncluded,currency:custUpper_(x.ledger['Currency']),sourceUniqueKey:custStr_(x.ledger['Source Unique Key']),note:pc==='CARRY_FORWARD'?'CARRY_FORWARD':(pc==='MIXED_PERIOD'?'MIXED_PERIOD':'')};
  });

  let refundRemaining=lateRefundApplied;
  const adjustmentItems=[];
  lateEntries.forEach(x=>{
    if(refundRemaining<=CUST_CFG.EPSILON)return;
    const take=custRound2_(Math.min(refundRemaining,x.outstanding)); if(take<=CUST_CFG.EPSILON)return;
    const l=x.ledger;
    adjustmentItems.push({
      lineType:'LATE_REFUND_ADJUSTMENT',ledgerKey:x.ledgerKey,allocationKeys:[],transactionDate:x.eventDate,transactionDateDisplay:custBusinessDateTimeDisplay_(x.eventDate),periodClass:'LATE_REFUND',
      paypalAccountKey:l['PayPal Account Key'],paypalTransactionId:l['PayPal Transaction ID'],referenceId:l['Reference ID'],invoiceId:l['Invoice ID'],invoicePrefix:l['Invoice Prefix'],
      item:'Late refund adjustment · '+(custLedgerItemLookup_(ss,l['PayPal Transaction ID'],l['PayPal Account Key'])||l['Invoice ID']||l['PayPal Transaction ID']),buyerName:'',
      grossAmount:0,paypalFee:0,refundAmount:take,paypalNetAmount:-take,receivingFeeAllocated:0,availableAmount:0,amountIncluded:-take,currency:custUpper_(l['Currency']),
      sourceUniqueKey:'LATE_REFUND|'+x.ledgerKey,note:'Refund occurred after prior invoice · applied '+take.toFixed(2)
    });
    refundRemaining=custRound2_(refundRemaining-take);
  });

  const items=saleItems.concat(adjustmentItems).sort((a,b)=>custTime_(a.transactionDate)-custTime_(b.transactionDate)||a.lineType.localeCompare(b.lineType));
  const payoutCurrency=custDefaultPayoutCurrency_(customer.defaultPayoutProvider,sourceCurrency);
  const fingerprint=custFingerprint_({customer:customer.customerKey,from:params.fromKey,through:params.toKey,asOf:params.asOfKey,timezone:CUST_CFG.BUSINESS_TIMEZONE,
    payoutFeeRate,feeRule:[feeRule.ruleKey,feeRule.conversionRate,feeRule.transferRate,feeRule.transferMinimum],
    allocations:selected.map(a=>[a['Allocation Key'],custNum_(a['Remaining Available']),custNum_(a['Receiving Fee Allocated']),custStr_(a['Settlement Key'])]).sort(),
    lateRefunds:lateEntries.map(x=>[x.ledgerKey,custNum_(x.outstanding),custTime_(x.eventDate)]).sort()});

  return {params,customer,allocations:selected,items,saleItems,adjustmentItems,receivingProvider,sourceCurrency,payoutCurrency,
    grossSales:gross,refunds:totalRefunds,saleRefunds:saleRefund,paypalFees:paypalFee,receivingFees:receivingFee,availableAmount:availableTotal,
    currentPeriodAmount:currentAmount,carryForwardAmount:carryAmount,lateRefundOutstanding,lateRefundApplied,lateRefundCarryForward,
    feeRule,conversionRate:feeRule.conversionRate,conversionCost,transferRate:feeRule.transferRate,transferMinimum:feeRule.transferMinimum,transferCost,downstreamCost,netBeforePayoutFee,payoutFeeRate,payoutFeeAmount,amountDue,
    paypalTotalSnapshot:custNum_(cRow['PayPal Total']),matchedSnapshot:custNum_(cRow['Matched Amount']),availableSnapshot:custNum_(cRow['Available Amount']),invoicedSnapshot:custNum_(cRow['Invoiced Amount']),outstandingSnapshot:custNum_(cRow['Outstanding Amount']),fingerprint};
}
function custRenderInvoicePreview_(ss,p){
  const sh=custRequireSheet_(ss,CUST_CFG.CONTROL_SHEET);
  custFillControlCustomerConfig_(ss,p.customer);
  custClearControlPreview_(ss);

  /* Fill left-side customer summary during invoice preview. */
  const pendingSnapshot=custRound2_(Math.max(0,p.paypalTotalSnapshot-p.matchedSnapshot));
  [['B14',p.paypalTotalSnapshot],['B15',p.matchedSnapshot],['B16',p.availableSnapshot],['B17',p.invoicedSnapshot],['B18',p.outstandingSnapshot],['B19',pendingSnapshot]]
    .forEach(([a,v])=>sh.getRange(a).setValue(v));

  const validation='OK · '+p.feeRule.ruleKey+' · Carry-forward '+p.carryForwardAmount.toFixed(2)+' '+p.sourceCurrency+
    (p.lateRefundApplied>CUST_CFG.EPSILON?' · Late refund -'+p.lateRefundApplied.toFixed(2):'')+
    (p.lateRefundCarryForward>CUST_CFG.EPSILON?' · Refund carry '+p.lateRefundCarryForward.toFixed(2):'');

  [
    ['E4',p.saleItems.length],
    ['E5',p.grossSales],
    ['E6',p.refunds],
    ['E7',p.paypalFees],
    ['E8',p.receivingFees],
    ['E9',p.availableAmount],
    ['E10',p.conversionRate],
    ['E11',p.conversionCost],
    ['E12',p.transferRate],
    ['E13',p.transferMinimum],
    ['E14',p.transferCost],
    ['E15',p.downstreamCost],
    ['E16',p.netBeforePayoutFee],
    ['E17',p.payoutFeeRate],
    ['E18',p.payoutFeeAmount],
    ['E19',p.amountDue],
    ['E21',validation],
    ['E22','READY · '+p.amountDue.toFixed(2)+' '+p.payoutCurrency]
  ].forEach(([a,v])=>sh.getRange(a).setValue(v));

  sh.getRange('E10').setNumberFormat('0.000');
  sh.getRange('E11').setNumberFormat('0.00');
  sh.getRange('E12').setNumberFormat('0.000');
  sh.getRange('E13:E16').setNumberFormat('0.00');
  sh.getRange('E17').setNumberFormat('0.00%');
  sh.getRange('E18:E19').setNumberFormat('0.00');

  custWriteControlDetails_(sh,p.items.map(i=>({
    date:i.transactionDateDisplay,
    tx:i.paypalTransactionId,
    invoice:i.invoiceId,
    item:i.item,
    gross:i.grossAmount,
    paypalFee:i.paypalFee,
    refund:i.refundAmount,
    paypalNet:i.paypalNetAmount,
    receivingFee:i.receivingFeeAllocated,
    available:i.availableAmount,
    included:i.amountIncluded,
    status:i.lineType==='LATE_REFUND_ADJUSTMENT'?'LATE_REFUND':(i.periodClass==='CARRY_FORWARD'?'CARRY_FORWARD':(i.periodClass==='MIXED_PERIOD'?'MIXED':'CURRENT'))
  })));
}

/*********************** CREATE INVOICE ************************/
function createCustomerInvoiceFromControl(){
  const ss=SpreadsheetApp.getActive(); syncCustomers_(); const p=custBuildInvoicePreview_(ss,true);
  const savedRaw=PropertiesService.getDocumentProperties().getProperty(CUST_CFG.PREVIEW_PROP); if(!savedRaw)throw new Error('Run INVOICE PREVIEW first.');
  const saved=JSON.parse(savedRaw); if(saved.fingerprint!==p.fingerprint)throw new Error('Source data changed. Run INVOICE PREVIEW again.');
  const now=new Date();
  const shSet=custRequireSheet_(ss,CUST_CFG.SETTLEMENTS_SHEET), shItems=custRequireSheet_(ss,CUST_CFG.SETTLEMENT_ITEMS_SHEET), shAlloc=custRequireSheet_(ss,CUST_CFG.ALLOCATION_SHEET), shPay=custRequireSheet_(ss,CUST_CFG.PAYOUTS_SHEET), shTx=custRequireSheet_(ss,CUST_CFG.TRANSACTION_SHEET);
  custEnsureHeaders_(shPay,['Customer Key','Settlement Key']);
  const hs=custHeaderMap_(shSet), hi=custHeaderMap_(shItems), ha=custHeaderMap_(shAlloc), hp=custHeaderMap_(shPay), htx=custHeaderMap_(shTx);
  const settlementKey=custNextKey_(shSet,hs,'Settlement Key','SET-',6), invoiceNumber=custNextKey_(shSet,hs,'Invoice Number','INV-',6);

  const notes=[];
  if(p.carryForwardAmount>CUST_CFG.EPSILON)notes.push('Includes prior eligible carry-forward: '+p.carryForwardAmount.toFixed(2)+' '+p.sourceCurrency);
  if(p.lateRefundApplied>CUST_CFG.EPSILON)notes.push('Late refund adjustment applied: '+p.lateRefundApplied.toFixed(2)+' '+p.sourceCurrency);
  if(p.lateRefundCarryForward>CUST_CFG.EPSILON)notes.push('Late refund balance carried forward: '+p.lateRefundCarryForward.toFixed(2)+' '+p.sourceCurrency);

  custAppendObjectRow_(shSet,hs,{
    'Settlement Key':settlementKey,'Invoice Number':invoiceNumber,'Customer Key':p.customer.customerKey,'Customer Name':p.customer.customerName,
    'Period From':p.params.fromDateValue,'Period To':p.params.toDateValue,'Gross Sales':p.grossSales,'Refunds':p.refunds,'PayPal Fees':p.paypalFees,'Receiving Fees':p.receivingFees,
    'Available Amount':p.availableAmount,
    'Conversion Rate':p.conversionRate,'Conversion Cost':p.conversionCost,
    'Transfer Rate':p.transferRate,'Transfer Minimum':p.transferMinimum,'Transfer Cost':p.transferCost,
    'Downstream Cost':p.downstreamCost,'Net Before Payout Fee':p.netBeforePayoutFee,
    'Payout Fee %':p.payoutFeeRate,'Payout Fee Amount':p.payoutFeeAmount,'Amount Due':p.amountDue,'Amount Paid':0,'Outstanding Payable':p.amountDue,'Currency':p.sourceCurrency,
    'PayPal Total Snapshot':p.paypalTotalSnapshot,'Matched Snapshot':p.matchedSnapshot,'Available Snapshot':p.availableSnapshot,'Invoiced Snapshot':p.invoicedSnapshot,'Outstanding Snapshot':p.outstandingSnapshot,
    'Fee Rule Key':p.feeRule.ruleKey,'Payout Provider':p.customer.defaultPayoutProvider,'Payout Account Key':p.customer.defaultPayoutAccountKey,'Payout Currency':p.payoutCurrency,
    'Status':'ISSUED','Created At':now,'Issued At':now,'Paid At':'','Last Sync':now,'PDF File ID':'','PDF URL':'','PDF Generated At':'','Note':notes.join(' | ')
  });

  p.items.forEach((it,idx)=>custAppendObjectRow_(shItems,hi,{
    'Settlement Item Key':settlementKey+'|SI|'+String(idx+1).padStart(4,'0'),'Settlement Key':settlementKey,'Invoice Number':invoiceNumber,'Customer Key':p.customer.customerKey,'Line Type':it.lineType||'SALE',
    'Transaction Date':it.transactionDate,'PayPal Account Key':it.paypalAccountKey,'PayPal Transaction ID':it.paypalTransactionId,'Reference ID':it.referenceId,'Invoice ID':it.invoiceId,'Invoice Prefix':it.invoicePrefix,'Item':it.item,'Buyer Name':it.buyerName,
    'Gross Amount':it.grossAmount,'PayPal Fee':it.paypalFee,'Refund Amount':it.refundAmount,'PayPal Net Amount':it.paypalNetAmount,'Receiving Fee Allocated':it.receivingFeeAllocated,'Available Amount':it.availableAmount,'Amount Included':it.amountIncluded,'Currency':it.currency,
    'Customer Ledger Key':it.ledgerKey,'Customer Allocation Key':(it.allocationKeys||[]).join(', '),'Source Unique Key':it.sourceUniqueKey,'Created At':now,'Note':it.note||''
  }));

  /* Positive AVAILABLE allocations are consumed even when refund makes Amount Due = 0. */
  const allocRows=custReadRowsObjects_(shAlloc,ha), rowByKey={}; allocRows.forEach(r=>rowByKey[custStr_(r['Allocation Key'])]=r);
  p.allocations.forEach(a=>{
    const key=custStr_(a['Allocation Key']), r=rowByKey[key]; if(!r)throw new Error('Allocation row disappeared: '+key);
    const available=custNum_(r['Available Amount']), already=custNum_(r['Invoiced Amount']), remaining=custNum_(r['Remaining Available']);
    const newInvoiced=custRound2_(Math.min(available,already+remaining));
    custSetCellByHeader_(shAlloc,ha,r._row,'Invoiced Amount',newInvoiced);custSetCellByHeader_(shAlloc,ha,r._row,'Remaining Available',0);custSetCellByHeader_(shAlloc,ha,r._row,'Settlement Key',settlementKey);
  });

  let payoutKey='', txKey='';
  if(p.amountDue>CUST_CFG.EPSILON){
    payoutKey=custNextKey_(shPay,hp,'Payout Key','PO-',6); txKey='POTX|'+payoutKey;
    custWritePayoutDraftRow_(ss,shPay,hp,{'Payout Key':payoutKey,'Payout Date':'','Provider':p.customer.defaultPayoutProvider,'Account Key':p.customer.defaultPayoutAccountKey,'Payee':p.customer.customerName,'Category':'CUSTOMER_PAYOUT','Amount':p.amountDue,'Fee':0,'Total Deducted':p.amountDue,'Currency':p.payoutCurrency,'Reference':invoiceNumber,'Status':'DRAFT','Validation':'DRAFT','Transaction Key':txKey,'Created At':now,'Paid At':'','Note':'Auto-created from '+invoiceNumber,'Customer Key':p.customer.customerKey,'Settlement Key':settlementKey});
    custUpsertDraftPayoutTransaction_(shTx,htx,{transactionKey:txKey,provider:p.customer.defaultPayoutProvider,accountKey:p.customer.defaultPayoutAccountKey,amount:p.amountDue,currency:p.payoutCurrency,reference:invoiceNumber,payoutKey,now});
  }

  PropertiesService.getDocumentProperties().deleteProperty(CUST_CFG.PREVIEW_PROP); syncCustomers_();
  const statusText=payoutKey?('CREATED · '+invoiceNumber+' · '+payoutKey):('CREATED · '+invoiceNumber+' · NO PAYOUT · REFUND CARRY '+p.lateRefundCarryForward.toFixed(2));
  custRequireSheet_(ss,CUST_CFG.CONTROL_SHEET).getRange('E22').setValue(statusText);
  SpreadsheetApp.getActive().toast(payoutKey?(invoiceNumber+' created · '+payoutKey+' DRAFT'):(invoiceNumber+' created · Amount Due 0 · no payout'),'Customer Invoice V2.3',7);
  return {settlementKey,invoiceNumber,payoutKey,transactionKey:txKey,lateRefundApplied:p.lateRefundApplied,lateRefundCarryForward:p.lateRefundCarryForward};
}
function custUpsertDraftPayoutTransaction_(sheet,h,x){
  const rows=custReadRowsObjects_(sheet,h); let row=rows.find(r=>custStr_(r['Transaction Key'])===x.transactionKey||custStr_(r['Unique Key'])==='PAYOUT|'+x.payoutKey);
  const obj={'Transaction Key':x.transactionKey,'Transaction Date':x.now,'Provider':x.provider,'Account Key':x.accountKey,'Label':'','Type':x.provider==='PINGPONG'?'PINGPONG_PAYOUT':'USDT_PAYOUT','Direction':'OUT','Gross Amount':x.amount,'Net Amount':x.amount,'Fee Amount':0,'Currency':x.currency,'Reference':x.reference,'Counterparty Provider':'CUSTOMER','Counterparty Account Key':'','Reconcile Status':'DRAFT','Matched Transaction Key':'','Source':'PAYOUTS','Source ID':x.payoutKey,'Sync Time':x.now,'Unique Key':'PAYOUT|'+x.payoutKey};
  if(row){Object.keys(obj).forEach(k=>custSetCellByHeader_(sheet,h,row._row,k,obj[k]));}else custAppendObjectRow_(sheet,h,obj);
}

/*********************** PAYOUT TEMPLATE ROW ************************/
/*
 * Customer invoice payouts should reuse the first EMPTY pre-formatted
 * Payouts row before appending a new row.
 *
 * Provider/Currency alone do NOT make a row occupied because the user may
 * pre-fill those columns on blank template rows (as in rows 3,4,5...).
 *
 * This preserves dropdowns/data validation already prepared in the sheet.
 */
function custWritePayoutDraftRow_(ss,sheet,h,obj){
  let row=custFindReusablePayoutRow_(sheet,h);

  if(!row){
    row=custCreatePayoutTemplateRow_(sheet,h);
  }

  Object.keys(obj).forEach(k=>custSetCellByHeader_(sheet,h,row,k,obj[k]));

  // Make the generated row usable even if it was appended rather than reused.
  custEnsurePayoutRowValidations_(ss,sheet,h,row,custUpper_(obj['Provider']));

  return row;
}

function custFindReusablePayoutRow_(sheet,h){
  const start=CUST_CFG.DATA_START_ROW;
  const last=Math.max(sheet.getLastRow(),start);
  const lc=sheet.getLastColumn();
  if(lc<1)return null;

  const values=sheet.getRange(start,1,last-start+1,lc).getValues();

  // Intentionally ignore Provider and Currency because they can be defaults
  // on otherwise-empty template rows.
  const meaningful=[
    'Payout Key','Payout Date','Account Key','Payee','Category',
    'Amount','Fee','Total Deducted','Reference','Status','Validation',
    'Transaction Key','Created At','Paid At','Note','Customer Key','Settlement Key'
  ].filter(k=>h[k]);

  for(let i=0;i<values.length;i++){
    const row=values[i];
    const hasInput=meaningful.some(k=>{
      const v=row[h[k]-1];
      return v!=='' && v!==null && typeof v!=='undefined';
    });
    if(!hasInput)return start+i;
  }

  return null;
}

function custCreatePayoutTemplateRow_(sheet,h){
  let row=sheet.getLastRow()+1;
  if(row< CUST_CFG.DATA_START_ROW)row=CUST_CFG.DATA_START_ROW;

  if(row>sheet.getMaxRows()){
    sheet.insertRowsAfter(sheet.getMaxRows(),row-sheet.getMaxRows());
  }

  // Copy ONLY data validation from the first prepared Payouts data row.
  const templateRow=CUST_CFG.DATA_START_ROW;
  if(templateRow<=sheet.getMaxRows()){
    const src=sheet.getRange(templateRow,1,1,sheet.getLastColumn());
    const dst=sheet.getRange(row,1,1,sheet.getLastColumn());
    const validations=src.getDataValidations();
    if(validations && validations[0] && validations[0].some(v=>v!==null)){
      dst.setDataValidations(validations);
    }
  }

  return row;
}

function custEnsurePayoutRowValidations_(ss,sheet,h,row,provider){
  if(h['Provider']){
    sheet.getRange(row,h['Provider']).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(['USDT','PINGPONG'],true)
        .setAllowInvalid(false)
        .build()
    );
  }

  if(h['Status']){
    sheet.getRange(row,h['Status']).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(['DRAFT','PAID','CANCELLED'],true)
        .setAllowInvalid(false)
        .build()
    );
  }

  custSetPayoutAccountValidation_(ss,sheet,h,row,provider);
}

function custSetPayoutAccountValidation_(ss,payoutSheet,h,row,provider){
  if(!h['Account Key'])return;

  const target=payoutSheet.getRange(row,h['Account Key']);
  const providerName=custUpper_(provider);
  const sourceName=providerName==='USDT'?'USDT':(providerName==='PINGPONG'?'Pingpong':'');

  if(!sourceName){
    target.clearDataValidations();
    return;
  }

  const source=ss.getSheetByName(sourceName);
  if(!source || source.getLastColumn()<1){
    target.clearDataValidations();
    return;
  }

  const hs=custHeaderMap_(source);
  if(!hs['Account Key']){
    target.clearDataValidations();
    return;
  }

  const last=source.getLastRow();
  if(last<CUST_CFG.DATA_START_ROW){
    target.clearDataValidations();
    return;
  }

  const vals=source.getRange(
    CUST_CFG.DATA_START_ROW,1,last-CUST_CFG.DATA_START_ROW+1,source.getLastColumn()
  ).getValues();

  const keys=[];
  vals.forEach(r=>{
    const key=custCanonicalKey_(r[hs['Account Key']-1]);
    const status=hs['Status']?custUpper_(r[hs['Status']-1]):'';
    if(key && status!=='DISABLED')keys.push(key);
  });

  const unique=[...new Set(keys)];
  if(!unique.length){
    target.clearDataValidations();
    return;
  }

  target.setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(unique,true)
      .setAllowInvalid(false)
      .build()
  );
}

function repairCustomerPayoutDropdowns(){
  const ss=SpreadsheetApp.getActive();
  const sheet=custRequireSheet_(ss,CUST_CFG.PAYOUTS_SHEET);
  const h=custHeaderMap_(sheet);
  const last=sheet.getLastRow();
  let fixed=0;

  if(last>=CUST_CFG.DATA_START_ROW){
    const vals=sheet.getRange(
      CUST_CFG.DATA_START_ROW,1,last-CUST_CFG.DATA_START_ROW+1,sheet.getLastColumn()
    ).getValues();

    vals.forEach((r,i)=>{
      const payoutKey=h['Payout Key']?custStr_(r[h['Payout Key']-1]):'';
      const category=h['Category']?custUpper_(r[h['Category']-1]):'';
      if(!payoutKey && category!=='CUSTOMER_PAYOUT')return;
      const row=CUST_CFG.DATA_START_ROW+i;
      const provider=h['Provider']?custUpper_(r[h['Provider']-1]):'';
      custEnsurePayoutRowValidations_(ss,sheet,h,row,provider);
      fixed++;
    });
  }

  SpreadsheetApp.getActive().toast('Payout dropdowns repaired · '+fixed+' row(s)','Customer',5);
  return fixed;
}

/*********************** CONTROL HELPERS ************************/
function handleCustomerControlEdit_(e){ if(!e||!e.range||e.range.getSheet().getName()!==CUST_CFG.CONTROL_SHEET)return; const a=e.range.getA1Notation(); if(['B3','B4','B5','B6','B7'].indexOf(a)!==-1){PropertiesService.getDocumentProperties().deleteProperty(CUST_CFG.PREVIEW_PROP);custClearControlPreview_(SpreadsheetApp.getActive());if(a==='B4'){try{custFillControlCustomerConfig_(SpreadsheetApp.getActive(),custGetCustomerByKey_(SpreadsheetApp.getActive(),custCanonicalKey_(e.range.getValue())));}catch(_){}}} }
function custReadControlParams_(ss){
  const sh=custRequireSheet_(ss,CUST_CFG.CONTROL_SHEET), key=custCanonicalKey_(sh.getRange('B4').getValue());
  if(!key)throw new Error('Select Customer in B4.');
  const fromInfo=custReadControlDateCell_(sh.getRange('B5')), toInfo=custReadControlDateCell_(sh.getRange('B6'));
  let asOfInfo=custReadControlDateCell_(sh.getRange('B7'));
  if(!fromInfo||!toInfo)throw new Error('From Date and To Date are required.');
  if(fromInfo.key>toInfo.key)throw new Error('From Date cannot be after To Date.');
  if(!asOfInfo){sh.getRange('B7').setValue(new Date());asOfInfo=custReadControlDateCell_(sh.getRange('B7'));}
  return {customerKey:key,fromKey:fromInfo.key,toKey:toInfo.key,asOfKey:asOfInfo.key,fromDisplay:fromInfo.display,toDisplay:toInfo.display,asOfDisplay:asOfInfo.display,fromDateValue:fromInfo.value,toDateValue:toInfo.value,asOfDateValue:asOfInfo.value};
}
function custGetCustomerByKey_(ss,key){ const sh=custRequireSheet_(ss,CUST_CFG.CUSTOMERS_SHEET),h=custHeaderMap_(sh),state=custReadCustomers_(sh,h),c=state.customersByKey[custCanonicalKey_(key)]; if(!c)throw new Error('Customer not found: '+key); return c; }
function custFillControlCustomerConfig_(ss,c){const sh=custRequireSheet_(ss,CUST_CFG.CONTROL_SHEET);sh.getRange('B9').setValue(c.customerName);sh.getRange('B10').setValue(c.payoutFeeRate).setNumberFormat('0.00%');sh.getRange('B11').setValue(c.defaultPayoutProvider);sh.getRange('B12').setValue(c.defaultPayoutAccountKey);}
function custClearControlOutput_(ss){const sh=custRequireSheet_(ss,CUST_CFG.CONTROL_SHEET);sh.getRange('B9:B19').clearContent();custClearControlPreview_(ss);}
function custClearControlPreview_(ss){const sh=custRequireSheet_(ss,CUST_CFG.CONTROL_SHEET);sh.getRange('E4:E22').clearContent();const max=Math.max(0,sh.getLastRow()-28);if(max>0)sh.getRange(29,1,max,12).clearContent();}
function custWriteControlDetails_(sh,rows){
  const max=Math.max(0,sh.getLastRow()-28);if(max>0)sh.getRange(29,1,max,12).clearContent();if(!rows.length)return;
  const vals=rows.map(r=>[r.date,r.tx,r.invoice,r.item,r.gross,r.paypalFee,r.refund,r.paypalNet,r.receivingFee,r.available,r.included,r.status]);
  sh.getRange(29,1,vals.length,12).setValues(vals);sh.getRange(29,1,vals.length,1).setNumberFormat('@');sh.getRange(29,5,vals.length,7).setNumberFormat('0.00');
}
function custReadCustomerSheetRow_(ss,key){ const sh=custRequireSheet_(ss,CUST_CFG.CUSTOMERS_SHEET),h=custHeaderMap_(sh); return custReadRowsObjects_(sh,h).find(r=>custCanonicalKey_(r['Customer Key'])===key)||{}; }
function custLedgerItemLookup_(ss,txId,accountKey){ const sh=custRequireSheet_(ss,CUST_CFG.PAYPAL_SHEET),h=custHeaderMap_(sh); const rows=custReadRowsObjects_(sh,h); const r=rows.find(x=>custStr_(x['Transaction ID'])===custStr_(txId)&&custCanonicalKey_(x['Account Key'])===custCanonicalKey_(accountKey)); return r?custStr_(r['Item']):''; }

/*********************** ON EDIT CUSTOMER ************************/
function handleCustomerEdit_(e){
  if(!e||!e.range)return; const sh=e.range.getSheet(); if(sh.getName()!==CUST_CFG.CUSTOMERS_SHEET||e.range.getRow()<CUST_CFG.DATA_START_ROW)return;
  const h=custHeaderMap_(sh); if(!h['Customer Key']||!h['Prefixes'])return;
  const sc=e.range.getColumn(),ec=sc+e.range.getNumColumns()-1,single=e.range.getNumRows()===1&&e.range.getNumColumns()===1;
  const keyCol=h['Customer Key'],prefixCol=h['Prefixes'],statusCol=h['Status'],providerCol=h['Default Payout Provider'];
  if(single){const c=e.range.getColumn();if(c===keyCol)e.range.setValue(custCanonicalKey_(e.range.getValue()));if(c===prefixCol)e.range.setValue(custParsePrefixes_(e.range.getValue()).join(', '));if(c===statusCol||c===providerCol)e.range.setValue(custUpper_(e.range.getValue()));}
  if((keyCol>=sc&&keyCol<=ec)||(prefixCol>=sc&&prefixCol<=ec)){try{custReadCustomers_(sh,custHeaderMap_(sh));if(single)e.range.setNote('');}catch(err){if(single)e.range.setValue(typeof e.oldValue==='undefined'?'':e.oldValue).setNote(err.message);SpreadsheetApp.getActive().toast(err.message,'CUSTOMER ERROR',8);return;}}
  custRefreshInvoiceControlCustomerDropdown_();
}
function custRefreshInvoiceControlCustomerDropdown_(){ const ss=SpreadsheetApp.getActive(),sh=ss.getSheetByName(CUST_CFG.CUSTOMERS_SHEET),ctl=ss.getSheetByName(CUST_CFG.CONTROL_SHEET); if(!sh||!ctl)return; const h=custHeaderMap_(sh),keys=custReadRowsObjects_(sh,h).filter(r=>custUpper_(r['Status'])!=='DISABLED'&&custStr_(r['Customer Key'])).map(r=>custCanonicalKey_(r['Customer Key'])); const cell=ctl.getRange('B4'); if(!keys.length){cell.clearDataValidations();return;} cell.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(keys,true).setAllowInvalid(false).build()); }

/*********************** VALIDATION ************************/
function validateCustomerSystem(){
  const ss=SpreadsheetApp.getActive(); syncCustomers_();
  const shTx=custRequireSheet_(ss,CUST_CFG.TRANSACTION_SHEET), hTx=custHeaderMap_(shTx), txs=custReadTransactionRows_(shTx,hTx);
  const shA=custRequireSheet_(ss,CUST_CFG.ALLOCATION_SHEET), hA=custHeaderMap_(shA), allocs=custReadRowsObjects_(shA,hA);
  const shL=custRequireSheet_(ss,CUST_CFG.LEDGER_SHEET), hL=custHeaderMap_(shL), ledgers=custReadRowsObjects_(shL,hL);
  const late=custBuildLateRefundState_(ss,ledgers);
  const sums={},issues=[];
  allocs.forEach(a=>{
    const k=custStr_(a['PayPal Withdrawal Transaction Key']);sums[k]=(sums[k]||0)+custNum_(a['Allocated Amount']);
    const allocated=custNum_(a['Allocated Amount']),fee=custNum_(a['Receiving Fee Allocated']),available=custNum_(a['Available Amount']),invoiced=custNum_(a['Invoiced Amount']),remaining=custNum_(a['Remaining Available']);
    if(fee-allocated>0.02)issues.push('RECEIVING FEE > ALLOCATED '+a['Allocation Key']);
    if(available-allocated>0.02)issues.push('AVAILABLE > ALLOCATED '+a['Allocation Key']);
    if(invoiced-available>0.02)issues.push('INVOICED > AVAILABLE '+a['Allocation Key']);
    if(Math.abs(custRound2_(available-invoiced-remaining))>0.02)issues.push('REMAINING MISMATCH '+a['Allocation Key']);
  });
  txs.filter(t=>t.provider==='PAYPAL'&&t.type==='PAYPAL_WITHDRAWAL'&&t.direction==='OUT'&&t.reconcileStatus!=='RETURNED').forEach(w=>{
    const allocated=custRound2_(sums[w.transactionKey]||0),gross=Math.abs(custRound2_(w.grossAmount));if(allocated-gross>0.02)issues.push('OVER ALLOCATED '+w.transactionKey+' '+allocated+' > '+gross);
  });
  Object.keys(late).forEach(k=>{if(late[k].overConsumed>0.02)issues.push('LATE REFUND OVER-CONSUMED '+k+' '+late[k].overConsumed);});
  const carry=custRound2_(Object.keys(late).reduce((sum,k)=>sum+custNum_(late[k].outstanding),0));
  const msg=issues.length?('ISSUES '+issues.length+': '+issues.slice(0,4).join(' | ')):('OK · V2.2 allocation valid'+(carry>CUST_CFG.EPSILON?' · refund carry '+carry.toFixed(2):''));
  SpreadsheetApp.getActive().toast(msg,'Customer Validation',8); return {ok:!issues.length,issues,lateRefundOutstanding:carry};
}

/*********************** GENERIC ************************/
function custRequireSheet_(ss,name){const s=ss.getSheetByName(name);if(!s)throw new Error('Missing required sheet: '+name);return s;}
function custHeaderMap_(sheet){if(!sheet||sheet.getLastColumn()<1)return{};const vals=sheet.getRange(CUST_CFG.HEADER_ROW,1,1,sheet.getLastColumn()).getDisplayValues()[0],m={};vals.forEach((v,i)=>{const k=String(v||'').trim();if(k)m[k]=i+1;});return m;}
function custRequireHeaders_(name,h,required){const miss=required.filter(k=>!h[k]);if(miss.length)throw new Error(name+' missing headers: '+miss.join(', '));}
function custEnsureHeaders_(sheet,names){let h=custHeaderMap_(sheet),last=sheet.getLastColumn();names.forEach(name=>{if(!h[name]){last++;sheet.getRange(CUST_CFG.HEADER_ROW,last).setValue(name);h[name]=last;}});return h;}
function custClearData_(sheet){const last=sheet.getLastRow();if(last>=CUST_CFG.DATA_START_ROW)sheet.getRange(CUST_CFG.DATA_START_ROW,1,last-CUST_CFG.DATA_START_ROW+1,sheet.getLastColumn()).clearContent();}
function custPut_(row,h,k,v){if(h[k])row[h[k]-1]=(v===undefined||v===null)?'':v;}
function custSetCellByHeader_(sheet,h,row,k,v){if(h[k])sheet.getRange(row,h[k]).setValue(v===undefined||v===null?'':v);}
function custAppendObjectRow_(sheet,h,obj){const lc=sheet.getLastColumn(),row=new Array(lc).fill('');Object.keys(obj).forEach(k=>custPut_(row,h,k,obj[k]));sheet.getRange(sheet.getLastRow()+1,1,1,lc).setValues([row]);}
function custReadRowsObjects_(sheet,h){const last=sheet.getLastRow();if(last<CUST_CFG.DATA_START_ROW)return[];const inv={};Object.keys(h).forEach(k=>inv[h[k]-1]=k);return sheet.getRange(CUST_CFG.DATA_START_ROW,1,last-CUST_CFG.DATA_START_ROW+1,sheet.getLastColumn()).getValues().map((r,i)=>{const o={_row:CUST_CFG.DATA_START_ROW+i};r.forEach((v,j)=>{if(inv[j])o[inv[j]]=v;});return o;});}
function custFormatCols_(sheet,h,n,dateCols,numCols){dateCols.forEach(k=>{if(h[k])sheet.getRange(CUST_CFG.DATA_START_ROW,h[k],n,1).setNumberFormat('dd/MM/yyyy HH:mm:ss');});numCols.forEach(k=>{if(h[k])sheet.getRange(CUST_CFG.DATA_START_ROW,h[k],n,1).setNumberFormat('0.00');});}
function custStr_(v){return(v===null||v===undefined)?'':String(v).trim();}
function custUpper_(v){return custStr_(v).toUpperCase();}
function custCanonicalKey_(v){return custUpper_(v);}
function custCanonicalPrefix_(v){return custUpper_(v).replace(/\s+/g,'');}
function custParsePrefixes_(v){const s=custStr_(v);return s?s.split(',').map(custCanonicalPrefix_).filter(Boolean):[];}
function custPaypalTxIndexKey_(a,t){return custCanonicalKey_(a)+'|'+custStr_(t);}
function custMoneyPoolKey_(a,c){return custCanonicalKey_(a)+'|'+custUpper_(c);}
function custNum_(v){if(typeof v==='number'&&isFinite(v))return v;let s=custStr_(v);if(!s)return 0;s=s.replace(/[^\d,.\-]/g,'');if(s.indexOf(',')!==-1&&s.indexOf('.')!==-1){if(s.lastIndexOf(',')>s.lastIndexOf('.'))s=s.replace(/\./g,'').replace(',','.');else s=s.replace(/,/g,'');}else if(s.indexOf(',')!==-1){const p=s.split(',');s=(p.length===2&&p[1].length<=3)?p[0]+'.'+p[1]:s.replace(/,/g,'');}const n=Number(s);return isFinite(n)?n:0;}
function custPercent_(v){if(typeof v==='number'&&isFinite(v))return v>=1?v/100:v;const s=custStr_(v);if(!s)return 0;const n=custNum_(s);return s.indexOf('%')!==-1?n/100:(n>=1?n/100:n);}
function custBool_(v){if(v===true)return true;if(v===false)return false;const s=custUpper_(v);return ['TRUE','YES','Y','1','ENABLED','ACTIVE'].indexOf(s)!==-1;}
function custConfigDateKey_(ss,v){if(!v)return '';const d=custDate_(v);if(!d)return '';const tz=ss.getSpreadsheetTimeZone()||Session.getScriptTimeZone()||'GMT';return Utilities.formatDate(d,tz,'yyyy-MM-dd');}
function custDefaultPayoutCurrency_(provider,sourceCurrency){return custUpper_(provider)==='USDT'?'USDT':(custUpper_(sourceCurrency)||'USD');}
function custRound2_(v){return Math.round((custNum_(v)+Number.EPSILON)*100)/100;}
function custDate_(v){if(v instanceof Date)return isNaN(v.getTime())?null:v;if(!v)return null;const d=new Date(v);return isNaN(d.getTime())?null:d;}
function custDateOnly_(v){const d=custDate_(v);if(!d)return null;return new Date(d.getFullYear(),d.getMonth(),d.getDate());}
function custEndOfDay_(d){return new Date(d.getFullYear(),d.getMonth(),d.getDate(),23,59,59,999);}

// Business-date helpers.
// All customer statement/invoice period logic uses America/New_York,
// never the Apps Script or spreadsheet timezone.
function custBusinessDateKey_(d){
  const x=custDate_(d);
  return x?Utilities.formatDate(x,CUST_CFG.BUSINESS_TIMEZONE,'yyyy-MM-dd'):'';
}
function custBusinessDateTimeDisplay_(d){
  const x=custDate_(d);
  return x?Utilities.formatDate(x,CUST_CFG.BUSINESS_TIMEZONE,'dd/MM/yyyy HH:mm:ss'):'';
}
function custBusinessDateKeyInRange_(d,fromKey,toKey){
  const k=custBusinessDateKey_(d);
  return !!k&&k>=fromKey&&k<=toKey;
}
function custReadControlDateCell_(range){
  const raw=custStr_(range.getDisplayValue());
  const value=range.getValue();
  if(!raw&&!value)return null;

  // B5:B7 are formatted dd/MM/yyyy by setupCustomerModule.
  const m=raw.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if(m){
    const dd=String(parseInt(m[1],10)).padStart(2,'0');
    const mm=String(parseInt(m[2],10)).padStart(2,'0');
    const yyyy=m[3];
    return {key:yyyy+'-'+mm+'-'+dd,display:dd+'/'+mm+'/'+yyyy,value:value};
  }

  const d=custDate_(value);
  if(!d)return null;
  // Fallback only; the display parser above is preferred.
  const tz=SpreadsheetApp.getActive().getSpreadsheetTimeZone()||Session.getScriptTimeZone()||'GMT';
  return {
    key:Utilities.formatDate(d,tz,'yyyy-MM-dd'),
    display:Utilities.formatDate(d,tz,'dd/MM/yyyy'),
    value:value
  };
}
function custDateKey_(d){return custBusinessDateKey_(d);}
function custDateInRange_(d,from,to){return !!d&&d.getTime()>=from.getTime()&&d.getTime()<=to.getTime();}
function custTime_(d){return d instanceof Date&&!isNaN(d.getTime())?d.getTime():0;}
function custSortDateDesc_(dateProp,keyProp){return(a,b)=>{const d=custTime_(b[dateProp])-custTime_(a[dateProp]);return d||String(a[keyProp]||'').localeCompare(String(b[keyProp]||''));};}
function custAssertSingleCurrencyPerCustomer_(records){const m={};records.forEach(r=>{if(!r.customerKey||!r.currency||Math.abs(r.netAmount)<=CUST_CFG.EPSILON)return;(m[r.customerKey]||(m[r.customerKey]={}))[r.currency]=true;});Object.keys(m).forEach(k=>{const a=Object.keys(m[k]);if(a.length>1)throw new Error('MULTI CURRENCY CUSTOMER: '+k+' = '+a.join(', '));});}
function custSetGlobalCustomerError_(sheet,h,now,msg){if(!h['Last Sync']||!h['Sync Status'])return;custReadRowsObjects_(sheet,h).forEach(r=>{if(!custStr_(r['Customer Key']))return;custSetCellByHeader_(sheet,h,r._row,'Last Sync',now);custSetCellByHeader_(sheet,h,r._row,'Sync Status','ERROR · '+msg);});}
function custNextKey_(sheet,h,header,prefix,width){let max=0;custReadRowsObjects_(sheet,h).forEach(r=>{const s=custStr_(r[header]);if(s.indexOf(prefix)===0){const n=parseInt(s.slice(prefix.length),10);if(isFinite(n)&&n>max)max=n;}});return prefix+String(max+1).padStart(width,'0');}
function custFingerprint_(obj){const raw=JSON.stringify(obj);const bytes=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,raw);return bytes.map(b=>('0'+((b<0?b+256:b).toString(16))).slice(-2)).join('');}