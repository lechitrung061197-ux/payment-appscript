/**
 * ============================================================================
 * INVOICEUI.GS — PAYMENT ALL
 * ============================================================================
 * Modal dialog for previewing and creating a customer invoice.
 *
 * Presentation only. Every number comes from previewCustomerInvoice() /
 * createCustomerInvoice(), the same params API the AI gateway uses, so this
 * file contains no accounting logic and writes nothing to Invoice Control.
 * The sheet flow keeps working untouched and stays the fallback.
 *
 * Safety note: the dialog always sends back the fingerprint it previewed, so
 * createCustomerInvoice() refuses if the underlying data moved in between.
 * ============================================================================
 */

const INV_UI_CFG = Object.freeze({
  DIALOG_FILE: 'InvoiceDialog',
  DIALOG_TITLE: 'Tạo Invoice',
  WIDTH: 920,
  HEIGHT: 680,
  MAX_ITEMS: 400
});


/*********************** MENU ************************/

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Payment All')
    .addItem('Tạo Invoice…', 'showInvoiceDialog')
    .addToUi();
}


function showInvoiceDialog() {
  const html = HtmlService
    .createHtmlOutputFromFile(INV_UI_CFG.DIALOG_FILE)
    .setWidth(INV_UI_CFG.WIDTH)
    .setHeight(INV_UI_CFG.HEIGHT);

  SpreadsheetApp.getUi().showModalDialog(html, INV_UI_CFG.DIALOG_TITLE);
}


/*********************** BRIDGE ************************/

/*
 * custReadCustomers_ throws on duplicate Customer Key / Prefix, so a broken
 * Customers sheet surfaces the moment the dialog opens rather than at create.
 */
function uiListCustomers() {
  const ss = SpreadsheetApp.getActive();
  const sh = custRequireSheet_(ss, CUST_CFG.CUSTOMERS_SHEET);
  const state = custReadCustomers_(sh, custHeaderMap_(sh));

  return state.customers
    .filter(function (c) { return custUpper_(c.status) !== 'DISABLED'; })
    .map(function (c) {
      return {
        customerKey: c.customerKey,
        customerName: c.customerName,
        payoutFeeRate: custNum_(c.payoutFeeRate),
        payoutProvider: c.defaultPayoutProvider || '',
        payoutAccountKey: c.defaultPayoutAccountKey || ''
      };
    });
}


function uiPreviewInvoice(params) {
  return uiShapePreview_(previewCustomerInvoice(params));
}


function uiCreateInvoice(params) {
  const r = createCustomerInvoice(params) || {};
  return {
    settlementKey: custStr_(r.settlementKey),
    invoiceNumber: custStr_(r.invoiceNumber),
    payoutKey: custStr_(r.payoutKey),
    existing: !!r.existing,
    lateRefundCarryForward: custNum_(r.lateRefundCarryForward)
  };
}


function uiGenerateInvoicePdf(settlementKey) {
  const key = custStr_(settlementKey);
  if (!key) throw new Error('Missing settlement key.');
  if (typeof invPdfGenerate_ !== 'function') {
    throw new Error('InvoicePdf module is not available.');
  }

  const r = invPdfGenerate_(key, false) || {};
  return {
    invoiceNumber: custStr_(r.invoiceNumber),
    url: custStr_(r.url),
    existing: !!r.existing
  };
}


/*
 * Flattens the preview for the client.
 *
 * Dates go across as pre-formatted strings on purpose: invoice period logic
 * runs in America/New_York, and handing the browser a Date would let it
 * re-render in the viewer's own timezone and show a different day.
 */
function uiShapePreview_(p) {
  if (!p) throw new Error('Preview returned nothing.');

  const items = (p.items || []).slice(0, INV_UI_CFG.MAX_ITEMS).map(function (i) {
    return {
      date: custStr_(i.transactionDateDisplay),
      transactionId: custStr_(i.paypalTransactionId),
      invoiceId: custStr_(i.invoiceId),
      item: custStr_(i.item),
      gross: custNum_(i.grossAmount),
      paypalFee: custNum_(i.paypalFee),
      refund: custNum_(i.refundAmount),
      paypalNet: custNum_(i.paypalNetAmount),
      receivingFee: custNum_(i.receivingFeeAllocated),
      included: custNum_(i.amountIncluded),
      kind: i.lineType === 'LATE_REFUND_ADJUSTMENT'
        ? 'LATE_REFUND'
        : (i.periodClass === 'CARRY_FORWARD'
            ? 'CARRY_FORWARD'
            : (i.periodClass === 'MIXED_PERIOD' ? 'MIXED' : 'CURRENT'))
    };
  });

  return {
    fingerprint: custStr_(p.fingerprint),

    customerKey: p.customer && p.customer.customerKey,
    customerName: p.customer && p.customer.customerName,
    periodFrom: p.params && p.params.fromDisplay,
    periodTo: p.params && p.params.toDisplay,
    asOf: p.params && p.params.asOfDisplay,

    sourceCurrency: custStr_(p.sourceCurrency),
    payoutCurrency: custStr_(p.payoutCurrency),
    receivingProvider: custStr_(p.receivingProvider),
    feeRuleKey: p.feeRule && custStr_(p.feeRule.ruleKey),

    grossSales: custNum_(p.grossSales),
    refunds: custNum_(p.refunds),
    paypalFees: custNum_(p.paypalFees),
    receivingFees: custNum_(p.receivingFees),
    availableAmount: custNum_(p.availableAmount),

    conversionRate: custNum_(p.conversionRate),
    conversionCost: custNum_(p.conversionCost),
    transferRate: custNum_(p.transferRate),
    transferCost: custNum_(p.transferCost),
    downstreamCost: custNum_(p.downstreamCost),

    netBeforePayoutFee: custNum_(p.netBeforePayoutFee),
    payoutFeeRate: custNum_(p.payoutFeeRate),
    payoutFeeAmount: custNum_(p.payoutFeeAmount),
    amountDue: custNum_(p.amountDue),

    carryForwardAmount: custNum_(p.carryForwardAmount),
    lateRefundApplied: custNum_(p.lateRefundApplied),
    lateRefundCarryForward: custNum_(p.lateRefundCarryForward),

    itemCount: (p.items || []).length,
    itemsTruncated: (p.items || []).length > INV_UI_CFG.MAX_ITEMS,
    items: items
  };
}
