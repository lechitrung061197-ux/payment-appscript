/*************************************************
 * INVOICEPDF.GS
 *
 * Customer-facing PDF generator.
 *
 * Reads ONLY:
 * - Settlements
 * - Settlement Items
 *
 * Writes ONLY PDF metadata back to Settlements:
 * - PDF File ID
 * - PDF URL
 * - PDF Generated At
 *
 * Does NOT change:
 * - Customer Allocation
 * - Payouts
 * - Transaction
 * - settlement money/status
 *
 * V2.3 customer-facing rules:
 * - Internal receiving/conversion/transfer costs are combined on the PDF
 *   into one simple line: Transfer Fees.
 * - Payout Fee % is read from each Settlement (never hard-coded).
 * - Customer-facing wording stays simple; internal accounting terminology
 *   remains in Sheets only.
 *************************************************/

const INV_PDF_CFG = Object.freeze({
  SETTLEMENTS_SHEET: 'Settlements',
  ITEMS_SHEET: 'Settlement Items',

  HEADER_ROW: 2,
  DATA_START_ROW: 3,

  BUSINESS_TIMEZONE: 'America/New_York',

  ISSUER_NAME: 'VNS Centralize',
  ISSUER_CONTACT: '',
  LOGO_URL: 'https://shields.vnsmerchant.com/images/logo/vns-logo.png',
  LOGO_FILE_NAME: 'VNS-Centralize-Logo.png',
  SCRIPT_PROPERTY_LOGO_ID: 'PAYMENT_ALL_INVOICE_LOGO_FILE_ID',

  PDF_TITLE: 'INVOICE',
  DRIVE_FOLDER_NAME: 'Payment All - Customer Invoices',
  SCRIPT_PROPERTY_FOLDER_ID: 'PAYMENT_ALL_INVOICE_PDF_FOLDER_ID',

  // Portrait US Letter in points: 8.5 x 11 in.
  PAGE_WIDTH: 612,
  PAGE_HEIGHT: 792,
  PAGE_MARGIN: 34,

  ACCENT_COLOR: '#B01835',
  DARK_COLOR: '#1F2937',
  MUTED_COLOR: '#6B7280',
  LIGHT_BG: '#F5F6F8',
  BORDER_COLOR: '#D9DDE3',
  WHITE: '#FFFFFF',

  FONT_FAMILY: 'Arial',
  DETAIL_FONT_SIZE: 7,
  NORMAL_FONT_SIZE: 9,
  SMALL_FONT_SIZE: 7.5,

  // Dynamic transaction paging. Rows are packed based on estimated rendered
  // height instead of a fixed number of transactions per page.
  // Dense dynamic pagination calibrated for the current portrait layout.
  // The previous values were intentionally conservative and left too much
  // white space. These values let short/medium rows pack much more tightly
  // while long Item descriptions still consume more budget.
  DETAIL_PAGE_MAX_UNITS: 28.0,
  DETAIL_PAGE_MAX_ROWS: 18,
  DETAIL_ITEM_CHARS_PER_LINE: 42
});


/*************************************************
 * PUBLIC
 *************************************************/

/**
 * Run once.
 * Adds PDF metadata columns to Settlements if missing
 * and creates/reuses the Drive folder.
 */
function setupInvoicePdfModule() {
  const ss = SpreadsheetApp.getActive();
  const settlements = invPdfRequireSheet_(ss, INV_PDF_CFG.SETTLEMENTS_SHEET);

  invPdfEnsureColumn_(settlements, 'PDF File ID');
  invPdfEnsureColumn_(settlements, 'PDF URL');
  invPdfEnsureColumn_(settlements, 'PDF Generated At');

  const folder = invPdfGetOrCreateFolder_();

  // Cache the VNS logo in Drive once. If the remote image is temporarily
  // unavailable, PDF generation still works with a text-only fallback.
  let logoReady = false;
  try {
    logoReady = !!invPdfGetLogoBlob_();
  } catch (err) {
    logoReady = false;
  }

  SpreadsheetApp.getActive().toast(
    'PDF module ready · ' + folder.getName() + (logoReady ? ' · VNS logo ready' : ' · logo fallback'),
    'Invoice PDF',
    6
  );
}


/**
 * Select a settlement row in the Settlements tab, then run this.
 * If that settlement already has a PDF, it returns the existing PDF
 * instead of creating a duplicate.
 */
function generateInvoicePdfFromSelectedSettlement() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getActiveSheet();
  const range = ss.getActiveRange();

  if (!sheet || sheet.getName() !== INV_PDF_CFG.SETTLEMENTS_SHEET) {
    throw new Error('Select a row in the Settlements sheet first.');
  }

  if (!range || range.getRow() < INV_PDF_CFG.DATA_START_ROW) {
    throw new Error('Select a settlement data row (row 3 or below).');
  }

  const headers = invPdfHeaderMap_(sheet);
  invPdfRequireHeaders_(INV_PDF_CFG.SETTLEMENTS_SHEET, headers, [
    'Settlement Key'
  ]);

  const settlementKey = invPdfStr_(
    sheet.getRange(range.getRow(), headers['Settlement Key']).getValue()
  );

  if (!settlementKey) {
    throw new Error('Selected row has no Settlement Key.');
  }

  return invPdfGenerate_(settlementKey, false);
}


/**
 * Generates the newest settlement's PDF.
 * Useful immediately after CREATE INVOICE.
 */
function generateLatestInvoicePdf() {
  const ss = SpreadsheetApp.getActive();
  const sheet = invPdfRequireSheet_(ss, INV_PDF_CFG.SETTLEMENTS_SHEET);
  const headers = invPdfHeaderMap_(sheet);

  invPdfRequireHeaders_(INV_PDF_CFG.SETTLEMENTS_SHEET, headers, [
    'Settlement Key',
    'Created At'
  ]);

  const lastRow = sheet.getLastRow();
  if (lastRow < INV_PDF_CFG.DATA_START_ROW) {
    throw new Error('No settlements found.');
  }

  const values = sheet.getRange(
    INV_PDF_CFG.DATA_START_ROW,
    1,
    lastRow - INV_PDF_CFG.DATA_START_ROW + 1,
    sheet.getLastColumn()
  ).getValues();

  let newestKey = '';
  let newestTime = -1;

  values.forEach(function(row) {
    const key = invPdfStr_(row[headers['Settlement Key'] - 1]);
    if (!key) return;

    // An unfinished or failed create is not an invoice; never pick it as "latest".
    if (headers['Status'] &&
        !invPdfIsEffectiveSettlementStatus_(row[headers['Status'] - 1])) {
      return;
    }

    const created = invPdfDate_(row[headers['Created At'] - 1]);
    const time = created ? created.getTime() : 0;

    if (time > newestTime) {
      newestTime = time;
      newestKey = key;
    }
  });

  if (!newestKey) {
    throw new Error('Could not find an issued settlement to generate.');
  }

  return invPdfGenerate_(newestKey, false);
}


/**
 * Regenerates the PDF for the selected settlement.
 * Does NOT delete the old file; the new PDF gets a timestamp suffix
 * and the Settlements row is updated to point to the newest version.
 */
function regenerateInvoicePdfFromSelectedSettlement() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getActiveSheet();
  const range = ss.getActiveRange();

  if (!sheet || sheet.getName() !== INV_PDF_CFG.SETTLEMENTS_SHEET) {
    throw new Error('Select a row in the Settlements sheet first.');
  }

  if (!range || range.getRow() < INV_PDF_CFG.DATA_START_ROW) {
    throw new Error('Select a settlement data row (row 3 or below).');
  }

  const headers = invPdfHeaderMap_(sheet);
  const settlementKey = invPdfStr_(
    sheet.getRange(range.getRow(), headers['Settlement Key']).getValue()
  );

  if (!settlementKey) {
    throw new Error('Selected row has no Settlement Key.');
  }

  return invPdfGenerate_(settlementKey, true);
}


/**
 * Batch helper: generate PDFs for settlements that do not have one yet.
 */
function generateMissingInvoicePdfs() {
  const ss = SpreadsheetApp.getActive();
  const sheet = invPdfRequireSheet_(ss, INV_PDF_CFG.SETTLEMENTS_SHEET);

  invPdfEnsureColumn_(sheet, 'PDF File ID');
  invPdfEnsureColumn_(sheet, 'PDF URL');
  invPdfEnsureColumn_(sheet, 'PDF Generated At');

  const headers = invPdfHeaderMap_(sheet);
  const lastRow = sheet.getLastRow();

  if (lastRow < INV_PDF_CFG.DATA_START_ROW) {
    return 0;
  }

  const values = sheet.getRange(
    INV_PDF_CFG.DATA_START_ROW,
    1,
    lastRow - INV_PDF_CFG.DATA_START_ROW + 1,
    sheet.getLastColumn()
  ).getValues();

  const keys = [];

  values.forEach(function(row) {
    const key = invPdfStr_(row[headers['Settlement Key'] - 1]);
    const pdfId = invPdfStr_(row[headers['PDF File ID'] - 1]);

    // Skip non-issued settlements, otherwise one FAILED row aborts the batch.
    if (headers['Status'] &&
        !invPdfIsEffectiveSettlementStatus_(row[headers['Status'] - 1])) {
      return;
    }

    if (key && !pdfId) {
      keys.push(key);
    }
  });

  keys.forEach(function(key) {
    invPdfGenerate_(key, false);
  });

  SpreadsheetApp.getActive().toast(
    'Generated ' + keys.length + ' PDF(s).',
    'Invoice PDF',
    5
  );

  return keys.length;
}


/*************************************************
 * MAIN GENERATOR
 *************************************************/

function invPdfGenerate_(settlementKey, forceRegenerate) {
  if (typeof paymentAllWithScriptLock_ === 'function') {
    return paymentAllWithScriptLock_(
      'invPdfGenerate',
      function() {
        return invPdfGenerateCore_(settlementKey, forceRegenerate);
      }
    );
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    return invPdfGenerateCore_(settlementKey, forceRegenerate);
  } finally {
    lock.releaseLock();
  }
}


function invPdfGenerateCore_(settlementKey, forceRegenerate) {
  const ss = SpreadsheetApp.getActive();
  const settlementsSheet = invPdfRequireSheet_(ss, INV_PDF_CFG.SETTLEMENTS_SHEET);
  const itemsSheet = invPdfRequireSheet_(ss, INV_PDF_CFG.ITEMS_SHEET);

  // Safe schema extension only.
  invPdfEnsureColumn_(settlementsSheet, 'PDF File ID');
  invPdfEnsureColumn_(settlementsSheet, 'PDF URL');
  invPdfEnsureColumn_(settlementsSheet, 'PDF Generated At');

  const settlementHeaders = invPdfHeaderMap_(settlementsSheet);
  const itemHeaders = invPdfHeaderMap_(itemsSheet);

  invPdfRequireHeaders_(INV_PDF_CFG.SETTLEMENTS_SHEET, settlementHeaders, [
    'Settlement Key',
    'Invoice Number',
    'Customer Key',
    'Customer Name',
    'Period From',
    'Period To',
    'Gross Sales',
    'Refunds',
    'PayPal Fees',
    'Receiving Fees',
    'Available Amount',
    'Conversion Rate',
    'Conversion Cost',
    'Transfer Rate',
    'Transfer Minimum',
    'Transfer Cost',
    'Downstream Cost',
    'Net Before Payout Fee',
    'Payout Fee %',
    'Payout Fee Amount',
    'Amount Due',
    'Amount Paid',
    'Outstanding Payable',
    'Currency',
    'PayPal Total Snapshot',
    'Matched Snapshot',
    'Available Snapshot',
    'Invoiced Snapshot',
    'Outstanding Snapshot',
    'Payout Provider',
    'Payout Account Key',
    'Payout Currency',
    'Status',
    'Created At',
    'Issued At',
    'PDF File ID',
    'PDF URL',
    'PDF Generated At'
  ]);

  invPdfRequireHeaders_(INV_PDF_CFG.ITEMS_SHEET, itemHeaders, [
    'Settlement Item Key',
    'Settlement Key',
    'Invoice Number',
    'Customer Key',
    'Line Type',
    'Transaction Date',
    'PayPal Account Key',
    'PayPal Transaction ID',
    'Reference ID',
    'Invoice ID',
    'Invoice Prefix',
    'Item',
    'Buyer Name',
    'Gross Amount',
    'PayPal Fee',
    'Refund Amount',
    'PayPal Net Amount',
    'Receiving Fee Allocated',
    'Available Amount',
    'Amount Included',
    'Currency',
    'Customer Ledger Key',
    'Customer Allocation Key',
    'Source Unique Key',
    'Created At',
    'Note'
  ]);

  const settlement = invPdfFindSettlement_(
    settlementsSheet,
    settlementHeaders,
    settlementKey
  );

  if (!settlement) {
    throw new Error('Settlement not found: ' + settlementKey);
  }

  /*
   * Only a real invoice may become a customer-facing document. A settlement at
   * CREATING is an unfinished create and FAILED/CANCELLED never happened, so
   * rendering either would put a document in a customer's hands for money that
   * was never invoiced. PDF generation is a SAFE action for the AI gateway, so
   * this is the only gate on that path.
   */
  invPdfRequireEffectiveSettlement_(settlement, settlementHeaders, settlementKey);

  const existingPdfId = invPdfStr_(settlement.values[settlementHeaders['PDF File ID'] - 1]);
  const existingPdfUrl = invPdfStr_(settlement.values[settlementHeaders['PDF URL'] - 1]);

  if (!forceRegenerate && existingPdfId) {
    SpreadsheetApp.getActive().toast(
      'PDF already exists for ' + settlementKey,
      'Invoice PDF',
      5
    );

    return {
      settlementKey: settlementKey,
      fileId: existingPdfId,
      url: existingPdfUrl,
      existing: true
    };
  }

  const items = invPdfReadSettlementItems_(
    itemsSheet,
    itemHeaders,
    settlementKey
  );

  if (!items.length) {
    throw new Error('No Settlement Items found for ' + settlementKey);
  }

  const invoiceNumber = invPdfStr_(
    settlement.values[settlementHeaders['Invoice Number'] - 1]
  ) || settlementKey;

  const customerName = invPdfStr_(
    settlement.values[settlementHeaders['Customer Name'] - 1]
  );

  const periodFromDisplay = settlement.display[settlementHeaders['Period From'] - 1];
  const periodToDisplay = settlement.display[settlementHeaders['Period To'] - 1];

  const fileName = invPdfBuildFilename_(
    invoiceNumber,
    customerName,
    periodFromDisplay,
    periodToDisplay,
    forceRegenerate
  );

  const folder = invPdfGetOrCreateFolder_();

  const tempDoc = DocumentApp.create('TMP_' + fileName.replace(/\.pdf$/i, ''));
  const body = tempDoc.getBody();

  invPdfConfigureDocument_(body);
  invPdfRenderHeader_(body, settlement, settlementHeaders);
  invPdfRenderSummary_(body, settlement, settlementHeaders);
  invPdfRenderFundsNotIncluded_(body, settlement, settlementHeaders);
  invPdfRenderNotes_(body);
  invPdfRenderTransactionPages_(body, items, itemHeaders, settlement, settlementHeaders);

  tempDoc.saveAndClose();

  const tempFile = DriveApp.getFileById(tempDoc.getId());
  const pdfBlob = tempFile.getAs(MimeType.PDF).setName(fileName);
  const pdfFile = folder.createFile(pdfBlob);

  // Temp Google Doc is only an intermediate renderer.
  tempFile.setTrashed(true);

  const now = new Date();

  settlementsSheet
    .getRange(settlement.rowNumber, settlementHeaders['PDF File ID'])
    .setValue(pdfFile.getId());

  settlementsSheet
    .getRange(settlement.rowNumber, settlementHeaders['PDF URL'])
    .setValue(pdfFile.getUrl());

  settlementsSheet
    .getRange(settlement.rowNumber, settlementHeaders['PDF Generated At'])
    .setValue(now)
    .setNumberFormat('dd/MM/yyyy HH:mm:ss');

  SpreadsheetApp.flush();

  SpreadsheetApp.getActive().toast(
    invoiceNumber + ' PDF created.',
    'Invoice PDF',
    6
  );

  return {
    settlementKey: settlementKey,
    invoiceNumber: invoiceNumber,
    fileId: pdfFile.getId(),
    url: pdfFile.getUrl(),
    existing: false
  };
}


/*************************************************
 * DOCUMENT RENDERING
 *************************************************/

function invPdfConfigureDocument_(body) {
  body.setPageWidth(INV_PDF_CFG.PAGE_WIDTH);
  body.setPageHeight(INV_PDF_CFG.PAGE_HEIGHT);

  body.setMarginTop(INV_PDF_CFG.PAGE_MARGIN);
  body.setMarginBottom(INV_PDF_CFG.PAGE_MARGIN);
  body.setMarginLeft(INV_PDF_CFG.PAGE_MARGIN);
  body.setMarginRight(INV_PDF_CFG.PAGE_MARGIN);
}


function invPdfRenderHeader_(body, settlement, h) {
  const v = settlement.values;
  const d = settlement.display;

  const invoiceNumber = invPdfStr_(v[h['Invoice Number'] - 1]);
  const customerName = invPdfStr_(v[h['Customer Name'] - 1]);
  const customerKey = invPdfStr_(v[h['Customer Key'] - 1]);
  const status = invPdfStr_(v[h['Status'] - 1]) || 'ISSUED';
  const issuedAt = invPdfDate_(v[h['Issued At'] - 1]) || invPdfDate_(v[h['Created At'] - 1]);
  const payoutProvider = invPdfStr_(v[h['Payout Provider'] - 1]);
  const payoutAccount = invPdfStr_(v[h['Payout Account Key'] - 1]);
  const payoutCurrency = h['Payout Currency']
    ? (invPdfStr_(v[h['Payout Currency'] - 1]) || invPdfStr_(v[h['Currency'] - 1]) || 'USD')
    : (invPdfStr_(v[h['Currency'] - 1]) || 'USD');

  /******** BRAND + INVOICE TITLE — SAME FIRST PAGE ********/
  const top = body.appendTable([['', '']]);
  top.setBorderWidth(0);

  const brandCell = top.getRow(0).getCell(0);
  const invoiceCell = top.getRow(0).getCell(1);
  invPdfPadCell_(brandCell, 0, 0, 4, 0);
  invPdfPadCell_(invoiceCell, 0, 0, 4, 0);

  const brandP = brandCell.getChild(0).asParagraph();
  brandP.setAlignment(DocumentApp.HorizontalAlignment.LEFT);

  // IMPORTANT: read ORIGINAL dimensions before resizing.
  // The previous version resized width first, then calculated ratio from the
  // original height, which could make the logo hundreds of points tall and
  // force the invoice across several pages.
  const logoBlob = invPdfGetLogoBlobSafe_();
  if (logoBlob) {
    try {
      const image = brandP.appendInlineImage(logoBlob);
      const originalW = Math.max(image.getWidth(), 1);
      const originalH = Math.max(image.getHeight(), 1);
      const maxW = 68;
      const maxH = 48;
      const scale = Math.min(maxW / originalW, maxH / originalH, 1);
      image.setWidth(Math.max(1, Math.round(originalW * scale)));
      image.setHeight(Math.max(1, Math.round(originalH * scale)));
    } catch (err) {
      // Text fallback below is sufficient.
    }
  }

  const issuer = brandCell.appendParagraph(INV_PDF_CFG.ISSUER_NAME);
  issuer.editAsText()
    .setFontFamily(INV_PDF_CFG.FONT_FAMILY)
    .setFontSize(10.5)
    .setBold(true)
    .setForegroundColor(INV_PDF_CFG.DARK_COLOR);

  if (INV_PDF_CFG.ISSUER_CONTACT) {
    const contact = brandCell.appendParagraph(INV_PDF_CFG.ISSUER_CONTACT);
    contact.editAsText()
      .setFontFamily(INV_PDF_CFG.FONT_FAMILY)
      .setFontSize(INV_PDF_CFG.SMALL_FONT_SIZE)
      .setForegroundColor(INV_PDF_CFG.MUTED_COLOR);
  }

  const invoiceTitle = invoiceCell.getChild(0).asParagraph();
  invoiceTitle.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  invoiceTitle.setText('INVOICE');
  invoiceTitle.editAsText()
    .setFontFamily(INV_PDF_CFG.FONT_FAMILY)
    .setFontSize(22)
    .setBold(true)
    .setForegroundColor(INV_PDF_CFG.DARK_COLOR);

  const invNo = invoiceCell.appendParagraph(invoiceNumber);
  invNo.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  invNo.editAsText()
    .setFontFamily(INV_PDF_CFG.FONT_FAMILY)
    .setFontSize(11)
    .setBold(true)
    .setForegroundColor(INV_PDF_CFG.ACCENT_COLOR);

  const invStatus = invoiceCell.appendParagraph(status.toUpperCase());
  invStatus.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  invStatus.editAsText()
    .setFontFamily(INV_PDF_CFG.FONT_FAMILY)
    .setFontSize(7.5)
    .setBold(true)
    .setForegroundColor(INV_PDF_CFG.MUTED_COLOR);

  invPdfAccentRule_(body);

  /******** CUSTOMER + INVOICE DETAILS ********/
  const info = body.appendTable([['', '']]);
  info.setBorderWidth(0);

  const customerCell = info.getRow(0).getCell(0);
  const metaCell = info.getRow(0).getCell(1);
  invPdfPadCell_(customerCell, 6, 4, 6, 0);
  invPdfPadCell_(metaCell, 6, 4, 6, 0);

  const billLabel = customerCell.getChild(0).asParagraph();
  billLabel.setText('BILLED TO');
  invPdfStyleLabelParagraph_(billLabel);

  const customerP = customerCell.appendParagraph(customerName || customerKey);
  customerP.editAsText()
    .setFontFamily(INV_PDF_CFG.FONT_FAMILY)
    .setFontSize(14)
    .setBold(true)
    .setForegroundColor(INV_PDF_CFG.DARK_COLOR);

  const customerIdP = customerCell.appendParagraph('Customer ID: ' + customerKey);
  customerIdP.editAsText()
    .setFontFamily(INV_PDF_CFG.FONT_FAMILY)
    .setFontSize(INV_PDF_CFG.SMALL_FONT_SIZE)
    .setForegroundColor(INV_PDF_CFG.MUTED_COLOR);

  const metaLabel = metaCell.getChild(0).asParagraph();
  metaLabel.setText('INVOICE DETAILS');
  metaLabel.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  invPdfStyleLabelParagraph_(metaLabel);

  invPdfAppendMetaLine_(metaCell, 'Period',
    invPdfFormatPeriodDisplay_(d[h['Period From'] - 1]) + ' - ' +
    invPdfFormatPeriodDisplay_(d[h['Period To'] - 1]));

  invPdfAppendMetaLine_(metaCell, 'Issued',
    issuedAt ? Utilities.formatDate(issuedAt, INV_PDF_CFG.BUSINESS_TIMEZONE, 'MMM d, yyyy') : '');

  invPdfAppendMetaLine_(metaCell, 'Payout',
    [payoutProvider, payoutAccount].filter(function(x) { return !!x; }).join(' / '));

  invPdfAppendMetaLine_(metaCell, 'Currency', payoutCurrency);

  body.appendParagraph('').editAsText().setFontSize(2);
}

function invPdfRenderSummary_(body, settlement, h) {
  const v = settlement.values;

  const sourceCurrency = invPdfStr_(v[h['Currency'] - 1]) || 'USD';
  const payoutCurrency = invPdfStr_(v[h['Payout Currency'] - 1]) || sourceCurrency;

  const payoutFeeRate = invPdfNum_(v[h['Payout Fee %'] - 1]);
  const amountDue = invPdfNum_(v[h['Amount Due'] - 1]);

  /*
   * CUSTOMER-FACING FEE WORDING
   * ---------------------------
   * The Sheet keeps these costs separate for audit:
   *   Receiving Fees + Conversion Cost + Transfer Cost
   *
   * The PDF intentionally combines them into ONE simple line:
   *   Transfer Fees
   *
   * Downstream Cost is NOT shown separately because it already equals
   * Conversion Cost + Transfer Cost and would make the invoice look like
   * the same fees were charged twice.
   */
  const transferFees = invPdfNum_(v[h['Receiving Fees'] - 1])
    + invPdfNum_(v[h['Conversion Cost'] - 1])
    + invPdfNum_(v[h['Transfer Cost'] - 1]);

  const heading = body.appendParagraph('INVOICE SUMMARY');
  heading.editAsText()
    .setFontFamily(INV_PDF_CFG.FONT_FAMILY)
    .setBold(true)
    .setFontSize(10.5)
    .setForegroundColor(INV_PDF_CFG.DARK_COLOR);

  /*
   * TẠM ỨNG
   * -------
   * Hai dòng này nằm SAU dòng phí, phản chiếu đúng thứ tự tính bên
   * Customer.gs: tiền ứng là vốn của chủ cho mượn, không phải doanh thu của
   * khách, nên phí không chạm vào nó ở cả hai chiều.
   *
   * Cột chỉ mọc ra khi bản Customer.gs mới ghi vào, nên hoá đơn cũ không có
   * cột này vẫn render như trước — đọc thiếu cột phải là im lặng, không phải nổ.
   */
  const advanceGiven = h['Advance Amount'] ? invPdfNum_(v[h['Advance Amount'] - 1]) : 0;
  const advanceRepaid = h['Advance Recovered'] ? invPdfNum_(v[h['Advance Recovered'] - 1]) : 0;

  const rows = [
    ['Sales', invPdfMoney_(v[h['Gross Sales'] - 1], sourceCurrency)],
    ['Refunds', invPdfMoneySignedFee_(v[h['Refunds'] - 1], sourceCurrency)],
    ['PayPal Fees', invPdfMoneySignedFee_(v[h['PayPal Fees'] - 1], sourceCurrency)],
    ['Transfer Fees', invPdfMoneySignedFee_(transferFees, sourceCurrency)],
    ['Net Before Payout Fee', invPdfMoney_(v[h['Net Before Payout Fee'] - 1], payoutCurrency)],
    ['Payout Fee (' + invPdfPercent_(payoutFeeRate) + ')', invPdfMoneySignedFee_(v[h['Payout Fee Amount'] - 1], payoutCurrency)]
  ];

  if (advanceRepaid > 0.0001) {
    rows.push(['Advance Repayment', invPdfMoneySignedFee_(advanceRepaid, payoutCurrency)]);
  }
  if (advanceGiven > 0.0001) {
    rows.push(['Advance', '+' + advanceGiven.toFixed(2) + ' ' + payoutCurrency]);
  }

  rows.push(['AMOUNT DUE', invPdfMoney_(amountDue, payoutCurrency)]);

  const table = body.appendTable(rows);
  table.setBorderWidth(0.35);
  try { table.setBorderColor(INV_PDF_CFG.BORDER_COLOR); } catch (err) {}

  for (let r = 0; r < rows.length; r++) {
    const label = table.getRow(r).getCell(0);
    const amount = table.getRow(r).getCell(1);

    invPdfPadCell_(label, 4, 4, 7, 7);
    invPdfPadCell_(amount, 4, 4, 7, 7);

    if (r === rows.length - 1) {
      label.setBackgroundColor(INV_PDF_CFG.ACCENT_COLOR);
      amount.setBackgroundColor(INV_PDF_CFG.ACCENT_COLOR);
      label.editAsText()
        .setFontFamily(INV_PDF_CFG.FONT_FAMILY)
        .setFontSize(10)
        .setBold(true)
        .setForegroundColor(INV_PDF_CFG.WHITE);
      amount.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
      amount.editAsText()
        .setFontFamily(INV_PDF_CFG.FONT_FAMILY)
        .setFontSize(12.5)
        .setBold(true)
        .setForegroundColor(INV_PDF_CFG.WHITE);
      continue;
    }

    if (r % 2 === 1) {
      label.setBackgroundColor(INV_PDF_CFG.LIGHT_BG);
      amount.setBackgroundColor(INV_PDF_CFG.LIGHT_BG);
    }

    label.editAsText()
      .setFontFamily(INV_PDF_CFG.FONT_FAMILY)
      .setFontSize(8.25)
      .setForegroundColor(INV_PDF_CFG.DARK_COLOR);

    amount.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
    amount.editAsText()
      .setFontFamily(INV_PDF_CFG.FONT_FAMILY)
      .setFontSize(8.25)
      .setForegroundColor(INV_PDF_CFG.DARK_COLOR);
  }

  /*
   * Khách phải thấy phần ứng còn nợ. Không hiện thì kỳ sau bị trừ mà không
   * hiểu vì sao, và người phải giải thích là bạn.
   */
  const advanceLeft = h['Advance Outstanding After'] ? invPdfNum_(v[h['Advance Outstanding After'] - 1]) : 0;
  if (advanceLeft > 0.0001) {
    const advNote = body.appendParagraph(
      'Advance balance carried forward: ' + advanceLeft.toFixed(2) + ' ' + payoutCurrency +
      ' — recovered from future invoices.'
    );
    advNote.editAsText()
      .setFontFamily(INV_PDF_CFG.FONT_FAMILY)
      .setFontSize(7.5)
      .setForegroundColor(INV_PDF_CFG.MUTED_COLOR);
  }

  const paid = invPdfNum_(v[h['Amount Paid'] - 1]);
  const balance = invPdfNum_(v[h['Outstanding Payable'] - 1]);

  if (paid > 0.0001 || balance < amountDue - 0.0001) {
    body.appendParagraph('').editAsText().setFontSize(2);

    const paymentRows = [
      ['Amount Paid', invPdfMoney_(paid, payoutCurrency)],
      ['Balance Due', invPdfMoney_(balance, payoutCurrency)]
    ];

    const paymentTable = body.appendTable(paymentRows);
    paymentTable.setBorderWidth(0.35);
    try { paymentTable.setBorderColor(INV_PDF_CFG.BORDER_COLOR); } catch (err) {}

    for (let r = 0; r < paymentRows.length; r++) {
      const label = paymentTable.getRow(r).getCell(0);
      const amount = paymentTable.getRow(r).getCell(1);
      invPdfPadCell_(label, 4, 4, 7, 7);
      invPdfPadCell_(amount, 4, 4, 7, 7);
      label.editAsText().setFontFamily(INV_PDF_CFG.FONT_FAMILY).setBold(true).setFontSize(8.25);
      amount.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
      amount.editAsText().setFontFamily(INV_PDF_CFG.FONT_FAMILY).setFontSize(8.25);
    }
  }

  body.appendParagraph('').editAsText().setFontSize(3);
}

function invPdfRenderFundsNotIncluded_(body, settlement, h) {
  const v = settlement.values;
  const sourceCurrency = invPdfStr_(v[h['Currency'] - 1]) || 'USD';

  /*
   * V2 no longer exposes customer-facing Hold / In-Transit buckets in the
   * Settlement. Customer ownership becomes available when the PayPal
   * withdrawal reaches and matches the receiving provider.
   *
   * Keep the old, easy-to-understand PDF wording by showing only the amount
   * that has not reached the receiving side yet:
   *   Pending in PayPal = PayPal Total Snapshot - Matched Snapshot
   */
  const paypalTotal = invPdfNum_(v[h['PayPal Total Snapshot'] - 1]);
  const matched = invPdfNum_(v[h['Matched Snapshot'] - 1]);
  const pending = Math.max(0, paypalTotal - matched);

  const heading = body.appendParagraph('FUNDS NOT INCLUDED IN THIS INVOICE');
  heading.editAsText()
    .setFontFamily(INV_PDF_CFG.FONT_FAMILY)
    .setBold(true)
    .setFontSize(9.5)
    .setForegroundColor(INV_PDF_CFG.DARK_COLOR);

  const rows = [
    ['Pending in PayPal', invPdfMoney_(pending, sourceCurrency)]
  ];

  const table = body.appendTable(rows);
  table.setBorderWidth(0.35);
  try { table.setBorderColor(INV_PDF_CFG.BORDER_COLOR); } catch (err) {}

  for (let r = 0; r < rows.length; r++) {
    const label = table.getRow(r).getCell(0);
    const amount = table.getRow(r).getCell(1);
    invPdfPadCell_(label, 3.5, 3.5, 7, 7);
    invPdfPadCell_(amount, 3.5, 3.5, 7, 7);
    label.editAsText().setFontFamily(INV_PDF_CFG.FONT_FAMILY).setFontSize(8);
    amount.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
    amount.editAsText().setFontFamily(INV_PDF_CFG.FONT_FAMILY).setFontSize(8);
  }

  const p = body.appendParagraph(
    'This amount is not included in the amount due above and may be included in a future invoice once it becomes available.'
  );
  p.editAsText()
    .setFontFamily(INV_PDF_CFG.FONT_FAMILY)
    .setFontSize(6.75)
    .setForegroundColor(INV_PDF_CFG.MUTED_COLOR)
    .setItalic(true);

  body.appendParagraph('').editAsText().setFontSize(2);
}

function invPdfRenderNotes_(body) {
  const p = body.appendParagraph(
    'Note: Amount Included is the amount included in this invoice. It may be lower than PayPal Net when part of a transaction settles in a later invoice.'
  );

  p.editAsText()
    .setFontFamily(INV_PDF_CFG.FONT_FAMILY)
    .setFontSize(6.75)
    .setForegroundColor(INV_PDF_CFG.MUTED_COLOR)
    .setItalic(true);
}

function invPdfRenderTransactionPages_(body, items, h, settlement, sh) {
  if (!items || !items.length) return;

  const invoiceNumber = invPdfStr_(settlement.values[sh['Invoice Number'] - 1]);
  const sourceCurrency = invPdfStr_(settlement.values[sh['Currency'] - 1]) || 'USD';

  // Build pages dynamically from estimated row height. Short item names allow
  // more rows on a page; long item names consume more of the page budget.
  const chunks = invPdfBuildDynamicTransactionChunks_(items, h);
  const totalChunks = chunks.length;

  // Page 1 contains all principal invoice information. Details start on page 2.
  body.appendPageBreak();

  let renderedCount = 0;

  for (let chunkIndex = 0; chunkIndex < totalChunks; chunkIndex++) {
    const chunk = chunks[chunkIndex];
    const from = renderedCount;

    invPdfRenderTransactionMiniHeader_(body, invoiceNumber, chunkIndex + 1, totalChunks);

    const heading = body.appendParagraph('TRANSACTION DETAILS - ' + sourceCurrency);
    heading.editAsText()
      .setFontFamily(INV_PDF_CFG.FONT_FAMILY)
      .setBold(true)
      .setFontSize(10.5)
      .setForegroundColor(INV_PDF_CFG.DARK_COLOR);

    const rows = [[
      'Date',
      'Transaction / Invoice',
      'Item',
      'Gross',
      'PayPal Fee',
      'Refund',
      'PayPal Net',
      'Included'
    ]];

    chunk.forEach(function(item) {
      const v = item.values;
      const txDate = invPdfDate_(v[h['Transaction Date'] - 1]);
      const txId = invPdfStr_(v[h['PayPal Transaction ID'] - 1]);
      const invoiceId = invPdfStr_(v[h['Invoice ID'] - 1]);

      const lineType = invPdfStr_(v[h['Line Type'] - 1]).toUpperCase();
      const isLateRefund = lineType === 'LATE_REFUND_ADJUSTMENT';
      const itemText = invPdfStr_(v[h['Item'] - 1]);

      rows.push([
        txDate
          ? Utilities.formatDate(txDate, INV_PDF_CFG.BUSINESS_TIMEZONE, 'MM/dd/yy')
          : '',
        txId + (invoiceId ? '\n' + invoiceId : ''),
        isLateRefund && itemText.toLowerCase().indexOf('late refund') !== 0
          ? 'Late refund · ' + itemText
          : itemText,
        invPdfNumberPlain_(v[h['Gross Amount'] - 1]),
        invPdfNumberNegative_(v[h['PayPal Fee'] - 1]),
        invPdfNumberNegative_(v[h['Refund Amount'] - 1]),
        invPdfNumberSignedPlain_(v[h['PayPal Net Amount'] - 1]),
        invPdfNumberSignedPlain_(v[h['Amount Included'] - 1])
      ]);
    });

    const table = body.appendTable(rows);
    table.setBorderWidth(0.25);
    try { table.setBorderColor(INV_PDF_CFG.BORDER_COLOR); } catch (err) {}

    // Portrait-friendly column widths. This gives Transaction ID and Item more
    // room and prevents ugly character-by-character wrapping.
    invPdfSetTableColumnWidths_(table, [42, 92, 145, 48, 54, 48, 54, 54]);

    const headerRow = table.getRow(0);
    for (let c = 0; c < headerRow.getNumCells(); c++) {
      const cell = headerRow.getCell(c);
      cell.setBackgroundColor(INV_PDF_CFG.ACCENT_COLOR);
      invPdfPadCell_(cell, 4, 4, 3, 3);
      cell.editAsText()
        .setFontFamily(INV_PDF_CFG.FONT_FAMILY)
        .setBold(true)
        .setFontSize(6.6)
        .setForegroundColor(INV_PDF_CFG.WHITE);
    }

    for (let r = 1; r < table.getNumRows(); r++) {
      const row = table.getRow(r);
      const alt = r % 2 === 0;

      for (let c = 0; c < row.getNumCells(); c++) {
        const cell = row.getCell(c);
        // Compact vertical padding so detail pages use the printable area.
        invPdfPadCell_(cell, 2, 2, 2, 2);
        if (alt) cell.setBackgroundColor(INV_PDF_CFG.LIGHT_BG);

        // Transaction/Invoice gets a slightly smaller font so IDs stay on one
        // line whenever possible; Item gets the normal detail size.
        const fontSize = (c === 1) ? 6.2 : INV_PDF_CFG.DETAIL_FONT_SIZE;
        cell.editAsText()
          .setFontFamily(INV_PDF_CFG.FONT_FAMILY)
          .setFontSize(fontSize)
          .setForegroundColor(INV_PDF_CFG.DARK_COLOR);

        if (c >= 3) {
          cell.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
        }
      }
    }

    const countP = body.appendParagraph(
      'Transactions ' + (from + 1) + '-' + (from + chunk.length) + ' of ' + items.length
    );
    countP.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
    countP.editAsText()
      .setFontFamily(INV_PDF_CFG.FONT_FAMILY)
      .setFontSize(6.5)
      .setForegroundColor(INV_PDF_CFG.MUTED_COLOR);

    renderedCount += chunk.length;

    if (chunkIndex < totalChunks - 1) {
      body.appendPageBreak();
    }
  }
}

/**
 * Dynamic pagination for transaction detail. DocumentApp does not expose the
 * final rendered row height, so we use a content-based estimate.
 * This dense version is calibrated to fill portrait pages much more fully
 * while still charging extra budget for long Item descriptions.
 */
function invPdfBuildDynamicTransactionChunks_(items, h) {
  const maxUnits = Math.max(8, Number(INV_PDF_CFG.DETAIL_PAGE_MAX_UNITS) || 15.2);
  const maxRows = Math.max(6, Number(INV_PDF_CFG.DETAIL_PAGE_MAX_ROWS) || 13);

  const chunks = [];
  let chunk = [];
  let usedUnits = 0;

  items.forEach(function(item) {
    const units = invPdfEstimateTransactionRowUnits_(item, h);

    // Start a new page only after the current page already contains enough
    // content. This prevents one unusually long row from leaving huge gaps.
    const wouldOverflow = chunk.length > 0 && (
      usedUnits + units > maxUnits ||
      chunk.length >= maxRows
    );

    if (wouldOverflow) {
      chunks.push(chunk);
      chunk = [];
      usedUnits = 0;
    }

    chunk.push(item);
    usedUnits += units;
  });

  if (chunk.length) chunks.push(chunk);
  return chunks;
}

function invPdfEstimateTransactionRowUnits_(item, h) {
  const v = item.values;
  const itemText = invPdfStr_(v[h['Item'] - 1]);
  const txId = invPdfStr_(v[h['PayPal Transaction ID'] - 1]);
  const invoiceId = invPdfStr_(v[h['Invoice ID'] - 1]);

  const charsPerLine = Math.max(24, Number(INV_PDF_CFG.DETAIL_ITEM_CHARS_PER_LINE) || 38);

  // Item column is the main source of variable row height. Respect explicit
  // line breaks as well as approximate wrapping at the configured width.
  const itemSegments = itemText ? itemText.split(/\r?\n/) : [''];
  let itemLines = 0;
  itemSegments.forEach(function(segment) {
    itemLines += Math.max(1, Math.ceil(segment.length / charsPerLine));
  });

  // Transaction / Invoice is normally two lines. Very long IDs get a small
  // extra allowance so they do not force unexpected pagination.
  const txLines = (txId ? Math.max(1, Math.ceil(txId.length / 22)) : 1) + (invoiceId ? 1 : 0);
  const visibleLines = Math.max(2, itemLines, txLines);

  // Base row + incremental height for wrapped lines. A small extra charge is
  // added for very long product descriptions.
  // Aggressive-but-safe estimate for the 7pt detail table. The old
  // estimator overcharged wrapped text and caused large blank areas.
  let units = 0.82 + Math.max(0, visibleLines - 1) * 0.24;
  if (itemText.length > 170) units += 0.15;
  if (itemText.length > 280) units += 0.20;

  return Math.min(units, 3.2);
}

function invPdfSetTableColumnWidths_(table, widths) {
  if (!table || !widths || !widths.length) return;
  for (let r = 0; r < table.getNumRows(); r++) {
    const row = table.getRow(r);
    const count = Math.min(row.getNumCells(), widths.length);
    for (let c = 0; c < count; c++) {
      try { row.getCell(c).setWidth(widths[c]); } catch (err) {}
    }
  }
}


function invPdfRenderTransactionMiniHeader_(body, invoiceNumber, pageNo, totalPages) {
  const table = body.appendTable([['VNS Centralize', invoiceNumber + ' · Detail ' + pageNo + '/' + totalPages]]);
  table.setBorderWidth(0);

  const left = table.getRow(0).getCell(0);
  const right = table.getRow(0).getCell(1);

  left.editAsText()
    .setFontFamily(INV_PDF_CFG.FONT_FAMILY)
    .setFontSize(8)
    .setBold(true)
    .setForegroundColor(INV_PDF_CFG.ACCENT_COLOR);

  right.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  right.editAsText()
    .setFontFamily(INV_PDF_CFG.FONT_FAMILY)
    .setFontSize(7.5)
    .setForegroundColor(INV_PDF_CFG.MUTED_COLOR);

  invPdfAccentRule_(body);
}


function invPdfAccentRule_(body) {
  const bar = body.appendTable([['']]);
  bar.setBorderWidth(0);
  const cell = bar.getRow(0).getCell(0);
  cell.setBackgroundColor(INV_PDF_CFG.ACCENT_COLOR);
  invPdfPadCell_(cell, 1, 1, 0, 0);
  cell.editAsText().setFontSize(1).setForegroundColor(INV_PDF_CFG.ACCENT_COLOR);
}


function invPdfStyleLabelParagraph_(paragraph) {
  paragraph.editAsText()
    .setFontFamily(INV_PDF_CFG.FONT_FAMILY)
    .setFontSize(7.5)
    .setBold(true)
    .setForegroundColor(INV_PDF_CFG.MUTED_COLOR);
}


function invPdfAppendMetaLine_(cell, label, value) {
  const p = cell.appendParagraph((label ? label + ': ' : '') + (value || ''));
  p.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  p.editAsText()
    .setFontFamily(INV_PDF_CFG.FONT_FAMILY)
    .setFontSize(8)
    .setForegroundColor(INV_PDF_CFG.DARK_COLOR);
}


function invPdfPadCell_(cell, top, bottom, left, right) {
  try { cell.setPaddingTop(top); } catch (err) {}
  try { cell.setPaddingBottom(bottom); } catch (err) {}
  try { cell.setPaddingLeft(left); } catch (err) {}
  try { cell.setPaddingRight(right); } catch (err) {}
}


function invPdfNumberPlain_(value) {
  return invPdfNum_(value).toLocaleString('en-US', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
}


function invPdfNumberSignedPlain_(value) {
  const n = invPdfNum_(value);
  if (Math.abs(n) < 0.0001) return '0.00';
  return n.toLocaleString('en-US', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
}


function invPdfNumberNegative_(value) {
  const n = Math.abs(invPdfNum_(value));
  if (n < 0.0001) return '0.00';
  return '-' + n.toLocaleString('en-US', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
}


/*************************************************
 * LOGO CACHE
 *************************************************/

function invPdfGetLogoBlobSafe_() {
  try {
    return invPdfGetLogoBlob_();
  } catch (err) {
    return null;
  }
}


function invPdfGetLogoBlob_() {
  const props = PropertiesService.getScriptProperties();
  const savedId = invPdfStr_(props.getProperty(INV_PDF_CFG.SCRIPT_PROPERTY_LOGO_ID));

  if (savedId) {
    try {
      const file = DriveApp.getFileById(savedId);
      file.getName();
      return file.getBlob();
    } catch (err) {
      props.deleteProperty(INV_PDF_CFG.SCRIPT_PROPERTY_LOGO_ID);
    }
  }

  const folder = invPdfGetOrCreateFolder_();
  const existing = folder.getFilesByName(INV_PDF_CFG.LOGO_FILE_NAME);

  if (existing.hasNext()) {
    const file = existing.next();
    props.setProperty(INV_PDF_CFG.SCRIPT_PROPERTY_LOGO_ID, file.getId());
    return file.getBlob();
  }

  const response = UrlFetchApp.fetch(INV_PDF_CFG.LOGO_URL, {
    muteHttpExceptions: true,
    followRedirects: true
  });

  const code = response.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error('Could not load VNS logo. HTTP ' + code);
  }

  const blob = response.getBlob().setName(INV_PDF_CFG.LOGO_FILE_NAME);
  const logoFile = folder.createFile(blob);
  props.setProperty(INV_PDF_CFG.SCRIPT_PROPERTY_LOGO_ID, logoFile.getId());
  return logoFile.getBlob();
}


/*************************************************
 * DATA READ
 *************************************************/

/*
 * Settlement status is owned by Customer.gs. Defer to its helper so the list of
 * real-invoice statuses lives in exactly one place, and fall back to the same
 * three only if that file is somehow unavailable.
 */
function invPdfIsEffectiveSettlementStatus_(status) {
  if (typeof custIsEffectiveSettlementStatus_ === 'function') {
    return custIsEffectiveSettlementStatus_(status);
  }

  const s = invPdfStr_(status).toUpperCase();
  return s === 'ISSUED' || s === 'PARTIALLY_PAID' || s === 'PAID';
}


function invPdfRequireEffectiveSettlement_(settlement, headers, settlementKey) {
  if (!headers['Status']) {
    return;
  }

  const status =
    invPdfStr_(settlement.values[headers['Status'] - 1]).toUpperCase() ||
    'ISSUED';

  if (!invPdfIsEffectiveSettlementStatus_(status)) {
    throw new Error(
      'Settlement ' + settlementKey + ' is ' + status +
      ' and is not an issued invoice. No PDF was generated.'
    );
  }
}


function invPdfFindSettlement_(sheet, headers, settlementKey) {
  const lastRow = sheet.getLastRow();

  if (lastRow < INV_PDF_CFG.DATA_START_ROW) {
    return null;
  }

  const numRows = lastRow - INV_PDF_CFG.DATA_START_ROW + 1;
  const range = sheet.getRange(
    INV_PDF_CFG.DATA_START_ROW,
    1,
    numRows,
    sheet.getLastColumn()
  );

  const values = range.getValues();
  const display = range.getDisplayValues();

  for (let i = 0; i < values.length; i++) {
    const key = invPdfStr_(values[i][headers['Settlement Key'] - 1]);

    if (key === settlementKey) {
      return {
        rowNumber: INV_PDF_CFG.DATA_START_ROW + i,
        values: values[i],
        display: display[i]
      };
    }
  }

  return null;
}


function invPdfReadSettlementItems_(sheet, headers, settlementKey) {
  const lastRow = sheet.getLastRow();

  if (lastRow < INV_PDF_CFG.DATA_START_ROW) {
    return [];
  }

  const numRows = lastRow - INV_PDF_CFG.DATA_START_ROW + 1;
  const range = sheet.getRange(
    INV_PDF_CFG.DATA_START_ROW,
    1,
    numRows,
    sheet.getLastColumn()
  );

  const values = range.getValues();
  const display = range.getDisplayValues();
  const out = [];

  values.forEach(function(row, i) {
    if (invPdfStr_(row[headers['Settlement Key'] - 1]) !== settlementKey) {
      return;
    }

    out.push({
      rowNumber: INV_PDF_CFG.DATA_START_ROW + i,
      values: row,
      display: display[i]
    });
  });

  // Oldest transaction first on the PDF.
  out.sort(function(a, b) {
    const ad = invPdfDate_(a.values[headers['Transaction Date'] - 1]);
    const bd = invPdfDate_(b.values[headers['Transaction Date'] - 1]);

    const at = ad ? ad.getTime() : 0;
    const bt = bd ? bd.getTime() : 0;

    if (at !== bt) return at - bt;

    return invPdfStr_(a.values[headers['Settlement Item Key'] - 1])
      .localeCompare(invPdfStr_(b.values[headers['Settlement Item Key'] - 1]));
  });

  return out;
}


/*************************************************
 * DRIVE
 *************************************************/

function invPdfGetOrCreateFolder_() {
  const props = PropertiesService.getScriptProperties();
  const savedId = invPdfStr_(props.getProperty(INV_PDF_CFG.SCRIPT_PROPERTY_FOLDER_ID));

  if (savedId) {
    try {
      const folder = DriveApp.getFolderById(savedId);
      // Touch the name so invalid/deleted folders throw here.
      folder.getName();
      return folder;
    } catch (err) {
      props.deleteProperty(INV_PDF_CFG.SCRIPT_PROPERTY_FOLDER_ID);
    }
  }

  const matches = DriveApp.getFoldersByName(INV_PDF_CFG.DRIVE_FOLDER_NAME);
  let folder;

  if (matches.hasNext()) {
    folder = matches.next();
  } else {
    folder = DriveApp.createFolder(INV_PDF_CFG.DRIVE_FOLDER_NAME);
  }

  props.setProperty(INV_PDF_CFG.SCRIPT_PROPERTY_FOLDER_ID, folder.getId());
  return folder;
}


/*************************************************
 * SCHEMA HELPERS
 *************************************************/

function invPdfEnsureColumn_(sheet, header) {
  const map = invPdfHeaderMap_(sheet);
  if (map[header]) return map[header];

  const col = Math.max(sheet.getLastColumn() + 1, 1);
  sheet.getRange(INV_PDF_CFG.HEADER_ROW, col).setValue(header);
  return col;
}


function invPdfHeaderMap_(sheet) {
  if (!sheet || sheet.getLastColumn() < 1) return {};

  const values = sheet
    .getRange(INV_PDF_CFG.HEADER_ROW, 1, 1, sheet.getLastColumn())
    .getDisplayValues()[0];

  const map = {};

  values.forEach(function(value, index) {
    const name = invPdfStr_(value);
    if (name) map[name] = index + 1;
  });

  return map;
}


function invPdfRequireHeaders_(sheetName, headers, required) {
  const missing = required.filter(function(header) {
    return !headers[header];
  });

  if (missing.length) {
    throw new Error(sheetName + ' missing headers: ' + missing.join(', '));
  }
}


function invPdfRequireSheet_(ss, name) {
  const sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error('Missing required sheet: ' + name);
  return sheet;
}


/*************************************************
 * FORMAT HELPERS
 *************************************************/

function invPdfBuildFilename_(invoiceNumber, customerName, fromDisplay, toDisplay, regenerated) {
  const safeInvoice = invPdfSafeFilenamePart_(invoiceNumber || 'INVOICE');
  const safeCustomer = invPdfSafeFilenamePart_(customerName || 'CUSTOMER');
  const safeFrom = invPdfSafeFilenamePart_(fromDisplay || 'FROM');
  const safeTo = invPdfSafeFilenamePart_(toDisplay || 'TO');

  let suffix = '';

  if (regenerated) {
    suffix = '_' + Utilities.formatDate(
      new Date(),
      INV_PDF_CFG.BUSINESS_TIMEZONE,
      'yyyyMMdd_HHmmss'
    );
  }

  return safeInvoice + '_' + safeCustomer + '_' + safeFrom + '_to_' + safeTo + suffix + '.pdf';
}


function invPdfSafeFilenamePart_(value) {
  return invPdfStr_(value)
    .replace(/[\\/:*?"<>|]+/g, '-')
    .replace(/\s+/g, '_')
    .replace(/_+/g, '_')
    .replace(/^_+|_+$/g, '')
    .substring(0, 80);
}


function invPdfFormatPeriodDisplay_(displayValue) {
  const s = invPdfStr_(displayValue);
  if (!s) return '';

  // Handles dd/MM/yyyy from the current sheet screenshots.
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);

  if (m) {
    const day = Number(m[1]);
    const month = Number(m[2]);
    const year = Number(m[3]);

    const months = [
      'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
      'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'
    ];

    if (month >= 1 && month <= 12) {
      return months[month - 1] + ' ' + day + ', ' + year;
    }
  }

  return s;
}


function invPdfMoney_(value, currency) {
  return invPdfNum_(value).toFixed(2) + (currency ? ' ' + currency : '');
}


function invPdfMoneySignedFee_(value, currency) {
  const n = Math.abs(invPdfNum_(value));
  if (n < 0.0001) return '0.00' + (currency ? ' ' + currency : '');
  return '-' + n.toFixed(2) + (currency ? ' ' + currency : '');
}


function invPdfPercent_(value) {
  let n = invPdfNum_(value);
  if (n > 1) n = n / 100;
  return (n * 100).toFixed(2).replace(/\.00$/, '') + '%';
}


function invPdfDate_(value) {
  if (value instanceof Date) {
    return isNaN(value.getTime()) ? null : value;
  }

  if (!value) return null;

  const d = new Date(value);
  return isNaN(d.getTime()) ? null : d;
}


function invPdfNum_(value) {
  if (typeof value === 'number' && isFinite(value)) return value;

  let s = invPdfStr_(value);
  if (!s) return 0;

  s = s.replace(/[^\d,.\-]/g, '');

  if (s.indexOf(',') !== -1 && s.indexOf('.') !== -1) {
    if (s.lastIndexOf(',') > s.lastIndexOf('.')) {
      s = s.replace(/\./g, '').replace(',', '.');
    } else {
      s = s.replace(/,/g, '');
    }
  } else if (s.indexOf(',') !== -1 && s.indexOf('.') === -1) {
    const parts = s.split(',');
    if (parts.length === 2 && parts[1].length <= 2) {
      s = parts[0] + '.' + parts[1];
    } else {
      s = s.replace(/,/g, '');
    }
  }

  const n = Number(s);
  return isFinite(n) ? n : 0;
}


function invPdfStr_(value) {
  if (value === null || typeof value === 'undefined') return '';
  return String(value).trim();
}
