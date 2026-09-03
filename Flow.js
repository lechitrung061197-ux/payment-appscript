/************************************************************
 * PAYMENT ALL - FLOW.GS
 *
 * Row 1 = notes
 * Row 2 = exact headers
 * Row 3+ = data
 *
 * FLOW:
 *
 * PayPal API -> Paypal -> Flow -> Transaction
 * Grey Gmail -> Transaction RAW -> Flow normalize
 * Grey USDT withdrawal -> derived USDT_RECEIVED
 * Payouts -> USDT / PINGPONG OUT -> Transaction
 *
 ************************************************************/

const FLOW_CFG = Object.freeze({

  SPREADSHEET_ID: '',

  ACCOUNT_SHEET: 'Account',
  PAYPAL_SHEET: 'Paypal',
  GREY_SHEET: 'Grey',
  USDT_SHEET: 'USDT',
  PINGPONG_SHEET: 'Pingpong',
  MANUAL_IN_SHEET: 'Manual In',
  PAYOUTS_SHEET: 'Payouts',
  TRANSACTION_SHEET: 'Transaction',

  /*
   * Read-only. Flow never writes Settlements; it only needs to know whether a
   * payout is backed by a real invoice before moving treasury money.
   */
  SETTLEMENTS_SHEET: 'Settlements',

  /*
   * Provider nhận mà reconcile THẬT SỰ biết cách đối chiếu.
   *
   * Sheet Account cho phép điền Linked Provider bất kỳ, nên danh sách này là
   * thứ duy nhất phân biệt "đã khai đích đến và hệ thống tới được" với "đã khai
   * đích đến nhưng không ai đi đón". Thiếu nó thì tiền kẹt im lặng ở OVERDUE.
   *
   * Thêm provider mới vào đây CHƯA đủ: còn cần collector và bộ chuẩn hoá raw
   * của provider đó, cộng một nhánh reconcile.
   */
  RECEIVING_PROVIDERS: ['GREY'],

  /* Cùng sheet khoá với collector Grey ở project khác. */
  LOCK_SHEET: '_Sync Lock',
  LOCK_WAIT_MS: 90000,
  LOCK_STALE_MS: 600000,

  HEADER_ROW: 2,
  DATA_START_ROW: 3,

  USD_TO_USDT_RATE: 0.998,

  MATCH_MAX_DAYS: 14,
  OVERDUE_DAYS: 7,
  AMOUNT_TOLERANCE: 0.01

});


var PAYMENT_ALL_LOCK_DEPTH =
  typeof PAYMENT_ALL_LOCK_DEPTH === 'number'
    ? PAYMENT_ALL_LOCK_DEPTH
    : 0;


function paymentAllWithScriptLock_(
  label,
  fn
) {

  if (
    PAYMENT_ALL_LOCK_DEPTH > 0
  ) {

    return fn();

  }


  const lock =
    LockService.getScriptLock();


  lock.waitLock(
    30000
  );


  PAYMENT_ALL_LOCK_DEPTH++;


  try {

    return fn();

  } finally {

    PAYMENT_ALL_LOCK_DEPTH--;
    lock.releaseLock();

  }

}


/************************************************************
 * MAIN
 ************************************************************/

function syncFlow() {

  /*
   * Hai lớp khoá, mỗi lớp một việc:
   *
   * paymentAllWithScriptLock_ chặn các entry point TRONG project này giẫm nhau.
   * ScriptLock không xuyên được giữa các Apps Script project, mà collector Grey
   * chạy ở project khác và cũng ghi sheet Transaction — nên cần thêm một khoá
   * nằm trong chính spreadsheet dùng chung.
   *
   * Quan trọng vì flowWriteTransactionState_ ghi đè TOÀN BỘ vùng dữ liệu
   * Transaction: chạy chồng với một bộ ghi khác là mất dòng, không phải trùng.
   */
  return paymentAllWithScriptLock_(
    'syncFlow',
    function () {

      const lock =
        flowAcquireSyncLock_(
          flowGetSpreadsheet_()
        );

      try {

        return syncFlowCore_();

      } finally {

        flowReleaseSyncLock_(lock);

      }

    }
  );

}


/************************************************************
 * ADVISORY LOCK DÙNG CHUNG VỚI COLLECTOR
 *
 * Cùng giao thức và cùng sheet với standalone/GreyGmailCollector.gs.
 * Không dùng chung code được vì hai project khác nhau — sửa một bên
 * thì phải sửa bên kia.
 *
 * Ghi token rồi đọc lại để xác nhận mình thắng. Vẫn còn khe hở rất
 * nhỏ, nhưng đủ để hai bộ ghi-đè-toàn-sheet không chạy chồng nhau.
 ************************************************************/

function flowAcquireSyncLock_(
  ss
) {

  let sheet =
    ss.getSheetByName(
      FLOW_CFG.LOCK_SHEET
    );


  if (!sheet) {

    sheet =
      ss.insertSheet(
        FLOW_CFG.LOCK_SHEET
      );

    sheet.getRange('A1').setValue('holder');
    sheet.getRange('B1').setValue('since');
    sheet.hideSheet();

  }


  const token =
    'FLOW|' +
    Utilities.getUuid();


  const deadline =
    new Date().getTime() +
    FLOW_CFG.LOCK_WAIT_MS;


  while (
    new Date().getTime() <
    deadline
  ) {

    const holder =
      flowString_(
        sheet.getRange('A2').getValue()
      );


    const since =
      sheet.getRange('B2').getValue();


    const sinceMs =
      since instanceof Date
        ? since.getTime()
        : 0;


    const stale =
      !sinceMs ||
      (
        new Date().getTime() -
        sinceMs >
        FLOW_CFG.LOCK_STALE_MS
      );


    if (
      !holder ||
      stale
    ) {

      sheet.getRange('A2').setValue(token);
      sheet.getRange('B2').setValue(new Date());

      SpreadsheetApp.flush();

      Utilities.sleep(1200);


      if (
        flowString_(
          sheet.getRange('A2').getValue()
        ) === token
      ) {

        return {
          sheet: sheet,
          token: token
        };

      }

    }


    Utilities.sleep(3000);

  }


  throw new Error(
    'Không lấy được sync lock sau ' +
    (FLOW_CFG.LOCK_WAIT_MS / 1000) +
    's — một collector khác đang ghi Transaction. Thử lại sau.'
  );

}


/*
 * So số dòng hiện tại với số dòng lúc flowLoadTransactionState_ đọc vào.
 * Chỉ bắt được việc CHÈN THÊM dòng, nhưng đó đúng là thứ mà collector Grey
 * làm (appendRow), nên đủ cho ca thực tế.
 */
function flowAssertTransactionUnchanged_(
  sheet,
  state
) {

  const expected =
    FLOW_CFG.DATA_START_ROW +
    state.oldRowCount -
    1;


  const actual =
    sheet.getLastRow();


  if (
    state.oldRowCount > 0 &&
    actual > expected
  ) {

    throw new Error(
      'Sheet Transaction đã đổi trong lúc syncFlow chạy (' +
      expected +
      ' -> ' +
      actual +
      ' dòng). Huỷ lượt ghi để không xoá mất dòng của bộ ghi khác. Chạy lại syncFlow().'
    );

  }

}


function flowReleaseSyncLock_(
  lock
) {

  if (!lock) {
    return;
  }


  try {

    if (
      flowString_(
        lock.sheet.getRange('A2').getValue()
      ) === lock.token
    ) {

      lock.sheet.getRange('A2:B2').clearContent();

      SpreadsheetApp.flush();

    }

  } catch (ignored) {
    /* lock tự hết hạn sau LOCK_STALE_MS */
  }

}


function syncFlowCore_() {

  const startedAt = new Date();

  Logger.log('========================================');
  Logger.log('FLOW START');
  Logger.log('========================================');

  const ss = flowGetSpreadsheet_();


  /**********************************************************
   * REQUIRED SHEETS
   **********************************************************/

  const accountSheet =
    flowNeedSheet_(ss, FLOW_CFG.ACCOUNT_SHEET);

  const paypalSheet =
    flowNeedSheet_(ss, FLOW_CFG.PAYPAL_SHEET);

  const greySheet =
    flowNeedSheet_(ss, FLOW_CFG.GREY_SHEET);

  const usdtSheet =
    flowNeedSheet_(ss, FLOW_CFG.USDT_SHEET);

  const payoutsSheet =
    flowNeedSheet_(ss, FLOW_CFG.PAYOUTS_SHEET);

  const txSheet =
    flowNeedSheet_(ss, FLOW_CFG.TRANSACTION_SHEET);


  /*
   * Pingpong hiện tại optional.
   */
  const pingpongSheet =
    flowOptionalSheet_(ss, FLOW_CFG.PINGPONG_SHEET);


  /**********************************************************
   * HEADERS
   **********************************************************/

  const ah = flowHeaderMap_(accountSheet);
  const ph = flowHeaderMap_(paypalSheet);
  const gh = flowHeaderMap_(greySheet);
  const uh = flowHeaderMap_(usdtSheet);
  const poh = flowHeaderMap_(payoutsSheet);
  const th = flowHeaderMap_(txSheet);


  flowValidateHeaders_(
    ah,
    ph,
    gh,
    uh,
    poh,
    th
  );


  /**********************************************************
   * MASTER MAPS
   **********************************************************/

  const accountMap =
    flowReadAccountMap_(accountSheet, ah);

  const greyMap =
    flowReadGreyMap_(greySheet, gh);

  const usdtMap =
    flowReadUsdtMap_(usdtSheet, uh);


  const payoutAccountMaps = {

    USDT:
      flowReadProviderAccountMap_(
        usdtSheet,
        ['Tag', 'Label']
      ),

    PINGPONG:
      flowReadProviderAccountMap_(
        pingpongSheet,
        ['Label', 'Tag']
      )

  };


  /**********************************************************
   * 1. PAYPAL RAW
   **********************************************************/

  const paypalRows =
    flowReadRows_(paypalSheet);


  flowProcessPaypalSheet_(
    paypalRows,
    ph
  );


  flowWriteRows_(
    paypalSheet,
    paypalRows
  );


  /**********************************************************
   * LOAD TRANSACTION STATE
   **********************************************************/

  const txState =
    flowLoadTransactionState_(
      txSheet,
      th
    );


  /**********************************************************
   * 2. PAYPAL -> TRANSACTION
   **********************************************************/

  const paypalResult =
    flowSyncPaypalToTransaction_(
      paypalRows,
      ph,
      accountMap,
      txState,
      th
    );


  /**********************************************************
   * 3. GREY RAW -> NORMALIZED
   **********************************************************/

  const greyResult =
    flowProcessGreyRawRows_(
      txState.rows,
      th,
      greyMap
    );


  /**********************************************************
   * 4. PAYPAL RETURNS
   **********************************************************/

  const returnMatched =
    flowReconcilePaypalReturns_(
      txState.rows,
      th
    );


  /**********************************************************
   * 5. PAYPAL -> GREY
   **********************************************************/

  const paypalGreyMatched =
    flowReconcilePaypalGrey_(
      txState.rows,
      th
    );


  /**********************************************************
   * 6. GREY -> USDT
   **********************************************************/

  const usdtCreated =
    flowSyncGreyToUsdt_(
      txState,
      th,
      usdtMap
    );


  /**********************************************************
   * 7. PAYOUTS -> TRANSACTION
   **********************************************************/

  const payoutRows =
    flowReadRowsPreserve_(
      payoutsSheet
    );


  const payoutResult =
    flowSyncPayoutsToTransaction_(
      payoutRows,
      poh,
      payoutAccountMaps,
      txState,
      th,
      flowReadSettlementStatusMap_(ss)
    );


  flowWriteRowsPreserve_(
    payoutsSheet,
    payoutRows
  );


  /**********************************************************
   * 7b. MANUAL IN -> TRANSACTION
   *
   * Tiền vào ví mà không đi qua đường ống nào: có người chuyển tới, ai đó trả
   * lại, một khoản từ chỗ khác. Không có mail, không có API — nguồn duy nhất
   * là chính bạn gõ vào.
   *
   * Optional như Pingpong: workbook chưa có sheet này thì bỏ qua, đừng làm
   * hỏng cả lượt sync chỉ vì một tính năng chưa dùng tới.
   **********************************************************/

  const manualInSheet =
    flowOptionalSheet_(ss, FLOW_CFG.MANUAL_IN_SHEET);


  if (manualInSheet) {

    const manualInRows =
      flowReadRowsPreserve_(manualInSheet);


    flowSyncManualInToTransaction_(
      manualInRows,
      flowHeaderMap_(manualInSheet),
      payoutAccountMaps,
      txState,
      th
    );


    flowWriteRowsPreserve_(
      manualInSheet,
      manualInRows
    );

  }


  /**********************************************************
   * 8. FINAL PAYPAL STATUS
   **********************************************************/

  flowFinalizePaypalStatus_(
    txState.rows,
    th
  );


  /**********************************************************
   * 9. WRITE TRANSACTION
   **********************************************************/

  /*
   * Lưới cuối dưới advisory lock: nếu sheet đã dài ra kể từ lúc đọc thì có bộ
   * ghi khác vừa chèn dòng, và bản trong bộ nhớ này đã cũ. Ghi đè lúc đó là
   * xoá mất dòng của họ, nên huỷ hẳn lượt này. Lần chạy sau sẽ đọc lại và làm
   * đúng — không mất gì ngoài một lượt.
   */
  flowAssertTransactionUnchanged_(
    txSheet,
    txState
  );


  flowWriteTransactionState_(
    txSheet,
    th,
    txState
  );


  /**********************************************************
   * 10. ACCOUNT SUMMARY
   **********************************************************/

  flowUpdateAccountSummary_(
    accountSheet,
    ah,
    txState.rows,
    th
  );


  /**********************************************************
   * 11. GREY SUMMARY
   **********************************************************/

  flowUpdateGreySummary_(
    greySheet,
    gh,
    txState.rows,
    th
  );


  /**********************************************************
   * 12. USDT SUMMARY
   **********************************************************/

  flowUpdateUsdtSummary_(
    usdtSheet,
    uh,
    txState.rows,
    th
  );


  SpreadsheetApp.flush();


  /**********************************************************
   * LOG
   **********************************************************/

  Logger.log(
    'PayPal withdrawals: ' +
    paypalResult.withdrawals
  );

  Logger.log(
    'PayPal returns: ' +
    paypalResult.returns
  );

  Logger.log(
    'Grey processed: ' +
    greyResult.processed
  );

  Logger.log(
    'Grey failed: ' +
    greyResult.failed
  );

  Logger.log(
    'PayPal return matched: ' +
    returnMatched
  );

  Logger.log(
    'PayPal -> Grey matched: ' +
    paypalGreyMatched
  );

  Logger.log(
    'USDT received rows: ' +
    usdtCreated
  );

  Logger.log(
    'Payout PAID: ' +
    payoutResult.paid
  );

  Logger.log(
    'Payout DRAFT: ' +
    payoutResult.draft
  );

  Logger.log(
    'Payout CANCELLED: ' +
    payoutResult.cancelled
  );

  Logger.log(
    'Payout INVALID: ' +
    payoutResult.invalid
  );

  /* Counted apart from INVALID: this one needs a person to look, not a retype. */
  Logger.log(
    'Payout BLOCKED (settlement): ' +
    payoutResult.blocked
  );


  const elapsed =
    (
      new Date().getTime() -
      startedAt.getTime()
    ) / 1000;


  Logger.log('========================================');
  Logger.log('FLOW DONE');
  Logger.log(
    'Transaction rows: ' +
    txState.rows.length
  );
  Logger.log(
    'Time: ' +
    elapsed.toFixed(2) +
    ' sec'
  );
  Logger.log('========================================');

}


/************************************************************
 * ON EDIT
 *
 * Account:
 * PayPal Status -> Hold Balance
 *
 * Payouts:
 * Provider -> Account dropdown + Currency
 * Amount/Fee -> Total Deducted
 * Status -> syncFlow()
 ************************************************************/

function onEdit(e) {

  if (!e || !e.range) {
    return;
  }


  const sheet =
    e.range.getSheet();

  const sheetName =
    sheet.getName();

  if (sheetName === 'Customers') {
   handleCustomerEdit_(e);
   return;
  } 
  if (sheetName === 'Invoice Control') {
   handleCustomerControlEdit_(e);
   return;
  }
  /**********************************************************
   * ACCOUNT
   **********************************************************/

  if (
    sheetName ===
    FLOW_CFG.ACCOUNT_SHEET
  ) {

    if (
      e.range.getRow() <
      FLOW_CFG.DATA_START_ROW
    ) {
      return;
    }


    const h =
      flowHeaderMap_(sheet);


    if (
      !Object.prototype.hasOwnProperty.call(
        h,
        'PayPal Status'
      )
    ) {
      return;
    }


    if (
      flowRangeTouchesColumn_(
        e.range,
        h['PayPal Status'] + 1
      )
    ) {

      const firstRow =
        Math.max(
          e.range.getRow(),
          FLOW_CFG.DATA_START_ROW
        );


      const lastRow =
        e.range.getRow() +
        e.range.getNumRows() -
        1;


      for (
        let row = firstRow;
        row <= lastRow;
        row++
      ) {

        flowRefreshHoldBalanceRow_(
          sheet,
          row,
          h
        );

      }

    }


    return;
  }


  /**********************************************************
   * PAYOUTS
   **********************************************************/

  if (
    sheetName !==
    FLOW_CFG.PAYOUTS_SHEET
  ) {
    return;
  }


  if (
    e.range.getRow() <
    FLOW_CFG.DATA_START_ROW
  ) {
    return;
  }


  const h =
    flowHeaderMap_(sheet);


  const required = [
    'Provider',
    'Account Key',
    'Amount',
    'Fee',
    'Total Deducted',
    'Currency',
    'Status'
  ];


  for (
    let i = 0;
    i < required.length;
    i++
  ) {

    if (
      !Object.prototype.hasOwnProperty.call(
        h,
        required[i]
      )
    ) {
      return;
    }

  }


  const firstRow =
    Math.max(
      e.range.getRow(),
      FLOW_CFG.DATA_START_ROW
    );


  const lastRow =
    e.range.getRow() +
    e.range.getNumRows() -
    1;


  /**********************************************************
   * PROVIDER CHANGED
   **********************************************************/

  if (
    flowRangeTouchesColumn_(
      e.range,
      h['Provider'] + 1
    )
  ) {

    for (
      let row = firstRow;
      row <= lastRow;
      row++
    ) {

      /*
       * Account dropdown.
       */
      flowUpdatePayoutAccountDropdown_(
        sheet,
        row,
        h
      );


      /*
       * Auto Currency.
       */
      flowAutoSetPayoutCurrency_(
        sheet,
        row,
        h
      );

    }

  }


  /**********************************************************
   * AMOUNT / FEE CHANGED
   **********************************************************/

  const amountChanged =
    flowRangeTouchesColumn_(
      e.range,
      h['Amount'] + 1
    );


  const feeChanged =
    flowRangeTouchesColumn_(
      e.range,
      h['Fee'] + 1
    );


  if (
    amountChanged ||
    feeChanged
  ) {

    for (
      let row = firstRow;
      row <= lastRow;
      row++
    ) {

      flowAutoSetPayoutTotal_(
        sheet,
        row,
        h
      );

    }

  }


  /**********************************************************
   * STATUS CHANGED
   *
   * DRAFT / PAID / CANCELLED
   * -> chạy toàn bộ Flow ngay.
   **********************************************************/

  if (
    flowRangeTouchesColumn_(
      e.range,
      h['Status'] + 1
    )
  ) {

    SpreadsheetApp.flush();

    syncFlow();

    return;
  }


  SpreadsheetApp.flush();

}


/************************************************************
 * PAYOUTS — PROVIDER -> ACCOUNT DROPDOWN
 ************************************************************/

function flowUpdatePayoutAccountDropdown_(
  payoutSheet,
  rowNumber,
  h
) {

  const ss =
    payoutSheet.getParent();


  const provider =
    flowString_(
      payoutSheet
        .getRange(
          rowNumber,
          h['Provider'] + 1
        )
        .getValue()
    )
      .toUpperCase();


  const accountCell =
    payoutSheet.getRange(
      rowNumber,
      h['Account Key'] + 1
    );


  /*
   * Provider đổi:
   * Account Key cũ không còn đáng tin.
   */
  accountCell.clearContent();
  accountCell.clearDataValidations();


  let sourceSheetName = '';


  if (
    provider ===
    'USDT'
  ) {

    sourceSheetName =
      FLOW_CFG.USDT_SHEET;


  } else if (
    provider ===
    'PINGPONG'
  ) {

    sourceSheetName =
      FLOW_CFG.PINGPONG_SHEET;


  } else {

    return;

  }


  const accounts =
    flowGetActiveAccountKeys_(
      ss,
      sourceSheetName
    );


  if (!accounts.length) {
    return;
  }


  const rule =
    SpreadsheetApp
      .newDataValidation()
      .requireValueInList(
        accounts,
        true
      )
      .setAllowInvalid(false)
      .build();


  accountCell.setDataValidation(
    rule
  );

}


/************************************************************
 * PAYOUTS — PROVIDER -> CURRENCY
 ************************************************************/

function flowAutoSetPayoutCurrency_(
  sheet,
  row,
  h
) {

  const provider =
    flowString_(
      sheet
        .getRange(
          row,
          h['Provider'] + 1
        )
        .getValue()
    )
      .toUpperCase();


  const currency =
    flowPayoutCurrencyForProvider_(
      provider
    );


  sheet
    .getRange(
      row,
      h['Currency'] + 1
    )
    .setValue(
      currency
    );

}


/************************************************************
 * PAYOUT PROVIDER -> CURRENCY
 ************************************************************/

function flowPayoutCurrencyForProvider_(
  provider
) {

  const value =
    flowString_(provider)
      .toUpperCase();


  if (
    value ===
    'USDT'
  ) {
    return 'USDT';
  }


  if (
    value ===
    'PINGPONG'
  ) {
    return 'USD';
  }


  return '';

}


/************************************************************
 * PAYOUTS — TOTAL DEDUCTED
 *
 * Amount = receiver gets
 * Fee    = extra fee
 *
 * Total Deducted = Amount + Fee
 ************************************************************/

function flowAutoSetPayoutTotal_(
  sheet,
  row,
  h
) {

  const amountRaw =
    sheet
      .getRange(
        row,
        h['Amount'] + 1
      )
      .getValue();


  const feeRaw =
    sheet
      .getRange(
        row,
        h['Fee'] + 1
      )
      .getValue();


  const amount =
    flowNullableNumber_(
      amountRaw
    );


  let fee =
    flowNullableNumber_(
      feeRaw
    );


  /*
   * Fee blank = zero.
   */
  if (
    flowString_(feeRaw) ===
    ''
  ) {

    fee = 0;

  }


  const totalCell =
    sheet.getRange(
      row,
      h['Total Deducted'] + 1
    );


  if (
    amount === null ||
    fee === null
  ) {

    totalCell.clearContent();

    return;

  }


  totalCell.setValue(
    flowRoundMoney_(
      amount +
      fee
    )
  );

}


/************************************************************
 * PAYOUTS — ACTIVE ACCOUNT KEYS
 *
 * Empty / missing Pingpong -> []
 ************************************************************/

function flowGetActiveAccountKeys_(
  ss,
  sheetName
) {

  const sheet =
    ss.getSheetByName(
      sheetName
    );


  if (!sheet) {
    return [];
  }


  if (
    sheet.getLastColumn() < 1 ||
    sheet.getLastRow() <
      FLOW_CFG.HEADER_ROW
  ) {

    return [];

  }


  const h =
    flowHeaderMap_(sheet);


  if (
    !Object.prototype.hasOwnProperty.call(
      h,
      'Account Key'
    )
  ) {

    return [];

  }


  const lastRow =
    sheet.getLastRow();


  if (
    lastRow <
    FLOW_CFG.DATA_START_ROW
  ) {

    return [];

  }


  const rows =
    sheet
      .getRange(
        FLOW_CFG.DATA_START_ROW,
        1,
        lastRow -
          FLOW_CFG.DATA_START_ROW +
          1,
        sheet.getLastColumn()
      )
      .getValues();


  const hasStatus =
    Object.prototype.hasOwnProperty.call(
      h,
      'Status'
    );


  const result = [];


  rows.forEach(
    row => {

      const key =
        flowString_(
          row[
            h['Account Key']
          ]
        );


      if (!key) {
        return;
      }


      if (hasStatus) {

        const status =
          flowString_(
            row[
              h['Status']
            ]
          )
            .toUpperCase();


        if (
          status ===
          'DISABLED'
        ) {

          return;

        }

      }


      result.push(key);

    }
  );


  return [
    ...new Set(result)
  ];

}


/************************************************************
 * OPTIONAL MANUAL REFRESH
 ************************************************************/

function refreshAllPayoutAccountDropdowns() {

  const ss =
    flowGetSpreadsheet_();


  const sheet =
    flowNeedSheet_(
      ss,
      FLOW_CFG.PAYOUTS_SHEET
    );


  const h =
    flowHeaderMap_(sheet);


  const lastRow =
    sheet.getLastRow();


  if (
    lastRow <
    FLOW_CFG.DATA_START_ROW
  ) {
    return;
  }


  for (
    let row =
      FLOW_CFG.DATA_START_ROW;
    row <= lastRow;
    row++
  ) {

    const provider =
      flowString_(
        sheet
          .getRange(
            row,
            h['Provider'] + 1
          )
          .getValue()
      );


    if (!provider) {
      continue;
    }


    flowUpdatePayoutAccountDropdown_(
      sheet,
      row,
      h
    );


    flowAutoSetPayoutCurrency_(
      sheet,
      row,
      h
    );


    flowAutoSetPayoutTotal_(
      sheet,
      row,
      h
    );

  }


  SpreadsheetApp.flush();

}


/************************************************************
 * PAYOUTS — PHÂN LOẠI
 *
 * Khai bằng HÀM chứ không phải const top-level: Apps Script gộp mọi file
 * lại và không đảm bảo thứ tự nạp, nên một const ở file này có thể chưa
 * khởi tạo khi file khác đọc tới. Khai báo hàm thì luôn được hoist.
 *
 * Category KHÔNG tham gia phép tính số dư ví — mọi payout PAID đều trừ ví
 * như nhau. Nó chỉ phục vụ hai việc: guard CUSTOMER_PAYOUT phải có
 * Settlement Key, và sổ chủ đọc thêm ở lớp trên.
 ************************************************************/

/*
 * CUSTOMER_PAYOUT do engine hoá đơn sinh, người dùng không tự gõ — nên nó
 * có trong danh sách này nhưng bị loại khỏi dropdown và khỏi AI.
 */
function flowPayoutCategories_() {
  return [
    'CUSTOMER_PAYOUT',
    'OWNER_DRAW',
    'EXPENSE',
    'REFUND',
    'TRANSFER',
    'ADJUSTMENT',
    'OTHER'
  ];
}


/* Những giá trị người dùng và AI được phép đặt. */
function flowPayoutManualCategories_() {
  return flowPayoutCategories_().filter(
    function (c) { return c !== 'CUSTOMER_PAYOUT'; }
  );
}


/*
 * Khoản nào tiêu vào phần của chủ.
 *
 * Rỗng VÀ giá trị lạ đều trả về false một cách CỐ Ý: chưa phân loại được thì
 * chưa kết luận được, và đoán bừa sẽ thổi phồng "đã lấy ra". Cả hai trường
 * hợp được cảnh báo riêng thay vì bị gộp vào đây.
 *
 * Quy tắc này phải nằm ở ĐÚNG MỘT chỗ. Trước đó sổ chủ loại giá trị lạ ra
 * còn cảnh báo lại cộng vào, nên hai nơi báo hai số khác nhau.
 */
function flowPayoutIsOwnerCost_(category) {
  const c = flowString_(category).toUpperCase();
  if (!c || c === 'CUSTOMER_PAYOUT') {
    return false;
  }
  return flowPayoutCategories_().indexOf(c) !== -1;
}


/************************************************************
 * PAYOUTS — ROW ACTUALLY HAS A PAYOUT
 *
 * Important:
 *
 * Provider = USDT
 * Currency = USDT
 *
 * alone DOES NOT create a PO key.
 ************************************************************/

function flowPayoutRowHasInput_(
  row,
  h
) {

  const fields = [

    'Payout Key',

    'Payout Date',

    'Account Key',

    'Payee',

    'Category',

    'Amount',

    'Fee',

    'Reference',

    'Status',

    'Note'

  ];


  return fields.some(
    header => {

      if (
        !Object.prototype.hasOwnProperty.call(
          h,
          header
        )
      ) {
        return false;
      }


      const value =
        row[
          h[header]
        ];


      return (
        value !== '' &&
        value !== null
      );

    }
  );

}


/************************************************************
 * PAYOUTS -> TRANSACTION
 ************************************************************/

function flowSyncPayoutsToTransaction_(
  rows,
  h,
  providerMaps,
  state,
  th,
  settlementStatusMap
) {

  let paid = 0;
  let draft = 0;
  let cancelled = 0;
  let invalid = 0;
  let blocked = 0;


  /**********************************************************
   * FIND HIGHEST PO NUMBER
   **********************************************************/

  let maxKeyNumber = 0;


  rows.forEach(
    row => {

      const key =
        flowCell_(
          row,
          h,
          'Payout Key'
        );


      const match =
        key.match(
          /^PO-(\d+)$/i
        );


      if (!match) {
        return;
      }


      maxKeyNumber =
        Math.max(
          maxKeyNumber,
          Number(match[1])
        );

    }
  );


  /**********************************************************
   * CREATE KEY / CREATED AT
   **********************************************************/

  rows.forEach(
    row => {

      /*
       * Ignore rows that only contain default
       * Provider/Currency dropdown values.
       */
      if (
        !flowPayoutRowHasInput_(
          row,
          h
        )
      ) {

        return;

      }


      let payoutKey =
        flowCell_(
          row,
          h,
          'Payout Key'
        );


      if (!payoutKey) {

        maxKeyNumber++;


        payoutKey =
          'PO-' +
          String(
            maxKeyNumber
          )
            .padStart(
              6,
              '0'
            );


        row[
          h['Payout Key']
        ] =
          payoutKey;

      }


      if (
        !row[
          h['Created At']
        ]
      ) {

        row[
          h['Created At']
        ] =
          new Date();

      }

    }
  );


  /**********************************************************
   * DUPLICATE KEY COUNTS
   **********************************************************/

  const keyCounts = {};


  rows.forEach(
    row => {

      const key =
        flowCell_(
          row,
          h,
          'Payout Key'
        );


      if (!key) {
        return;
      }


      keyCounts[key] =
        (
          keyCounts[key] ||
          0
        ) + 1;

    }
  );


  /**********************************************************
   * PROCESS PAYOUTS
   **********************************************************/

  rows.forEach(
    row => {

      if (
        !flowPayoutRowHasInput_(
          row,
          h
        )
      ) {

        return;

      }


      const payoutKey =
        flowCell_(
          row,
          h,
          'Payout Key'
        );


      if (!payoutKey) {
        return;
      }


      const uniqueKey =
        'PAYOUT|' +
        payoutKey;


      const existingTx =
        flowFindTransactionByUnique_(
          state,
          uniqueKey
        );


      /********************************************************
       * DUPLICATE KEY
       ********************************************************/

      if (
        keyCounts[payoutKey] > 1
      ) {

        row[
          h['Validation']
        ] =
          'DUPLICATE PAYOUT KEY';


        flowSetExistingPayoutTxState_(
          existingTx,
          th,
          'INVALID'
        );


        invalid++;

        return;

      }


      /********************************************************
       * PROVIDER
       ********************************************************/

      const provider =
        flowCell_(
          row,
          h,
          'Provider'
        )
          .toUpperCase();


      if (provider) {

        row[
          h['Provider']
        ] =
          provider;

      }


      /********************************************************
       * AUTO CURRENCY
       ********************************************************/

      const expectedCurrency =
        flowPayoutCurrencyForProvider_(
          provider
        );


      if (expectedCurrency) {

        row[
          h['Currency']
        ] =
          expectedCurrency;


      } else if (!provider) {

        row[
          h['Currency']
        ] =
          '';

      }


      /********************************************************
       * AMOUNT / FEE
       ********************************************************/

      const amount =
        flowNullableNumber_(
          row[
            h['Amount']
          ]
        );


      let fee =
        flowNullableNumber_(
          row[
            h['Fee']
          ]
        );


      if (
        flowString_(
          row[
            h['Fee']
          ]
        ) ===
        ''
      ) {

        fee = 0;

      }


      if (
        amount !== null &&
        fee !== null
      ) {

        row[
          h['Total Deducted']
        ] =
          flowRoundMoney_(
            amount +
            fee
          );


      } else {

        row[
          h['Total Deducted']
        ] =
          '';

      }


      /********************************************************
       * STATUS
       ********************************************************/

      const status =
        flowCell_(
          row,
          h,
          'Status'
        )
          .toUpperCase();


      /********************************************************
       * DRAFT
       ********************************************************/

      if (
        !status ||
        status ===
        'DRAFT'
      ) {

        row[
          h['Validation']
        ] =
          'DRAFT';


        flowSetExistingPayoutTxState_(
          existingTx,
          th,
          'DRAFT'
        );


        if (existingTx) {

          row[
            h['Transaction Key']
          ] =
            flowCell_(
              existingTx,
              th,
              'Transaction Key'
            );

        }


        draft++;

        return;

      }


      /********************************************************
       * CANCELLED
       ********************************************************/

      if (
        status ===
        'CANCELLED'
      ) {

        row[
          h['Validation']
        ] =
          'CANCELLED';


        flowSetExistingPayoutTxState_(
          existingTx,
          th,
          'CANCELLED'
        );


        if (existingTx) {

          row[
            h['Transaction Key']
          ] =
            flowCell_(
              existingTx,
              th,
              'Transaction Key'
            );

        }


        cancelled++;

        return;

      }


      /********************************************************
       * UNKNOWN STATUS
       ********************************************************/

      if (
        status !==
        'PAID'
      ) {

        row[
          h['Validation']
        ] =
          'INVALID STATUS';


        flowSetExistingPayoutTxState_(
          existingTx,
          th,
          'INVALID'
        );


        invalid++;

        return;

      }


      /********************************************************
       * SETTLEMENT VALIDATION
       *
       * Everything past the check above is PAID, so this is the last gate
       * before real treasury money moves. A payout is only as real as the
       * invoice behind it: CREATING means a create that never finished,
       * FAILED and CANCELLED never happened at all.
       *
       * Skipped entirely when there is no Settlements sheet, so Flow still
       * runs standalone on a workbook without the Customer module.
       ********************************************************/

      if (settlementStatusMap) {

        const settlementKey =
          h['Settlement Key']
            ? flowCell_(
                row,
                h,
                'Settlement Key'
              )
            : '';


        const category =
          flowCell_(
            row,
            h,
            'Category'
          )
            .toUpperCase();


        let blockReason = '';


        /*
         * CUSTOMER_PAYOUT is only ever written by the invoice engine, always
         * with a Settlement Key. A blank key means the row was edited by hand,
         * which is exactly how someone would route around this guard.
         */
        if (
          !settlementKey &&
          category ===
            'CUSTOMER_PAYOUT'
        ) {

          blockReason =
            'MISSING SETTLEMENT';

        } else if (
          settlementKey &&
          !flowIsEffectiveSettlementStatus_(
            settlementStatusMap[
              settlementKey
            ]
          )
        ) {

          blockReason =
            'SETTLEMENT NOT EFFECTIVE (' +
            (
              settlementStatusMap[
                settlementKey
              ] ||
              'MISSING'
            ) +
            ')';

        }


        if (blockReason) {

          row[
            h['Validation']
          ] =
            blockReason;


          /*
           * An existing transaction already marked PAID records money that
           * really left. Leave it counted so a mistaken edit to
           * Settlement.Status cannot inflate the treasury balance; the
           * Validation cell is what tells a human to look.
           */
          const alreadyPaidOut =
            existingTx &&
            flowCell_(
              existingTx,
              th,
              'Reconcile Status'
            )
              .toUpperCase() ===
              'PAID';


          if (!alreadyPaidOut) {

            flowSetExistingPayoutTxState_(
              existingTx,
              th,
              'INVALID'
            );

          }


          blocked++;

          return;

        }

      }


      /********************************************************
       * PROVIDER VALIDATION
       ********************************************************/

      if (
        provider !==
          'USDT' &&
        provider !==
          'PINGPONG'
      ) {

        row[
          h['Validation']
        ] =
          provider
            ? 'INVALID PROVIDER'
            : 'MISSING PROVIDER';


        flowSetExistingPayoutTxState_(
          existingTx,
          th,
          'INVALID'
        );


        invalid++;

        return;

      }


      /********************************************************
       * ACCOUNT VALIDATION
       ********************************************************/

      const accountKey =
        flowCell_(
          row,
          h,
          'Account Key'
        );


      if (!accountKey) {

        row[
          h['Validation']
        ] =
          'MISSING ACCOUNT';


        flowSetExistingPayoutTxState_(
          existingTx,
          th,
          'INVALID'
        );


        invalid++;

        return;

      }


      const providerMap =
        providerMaps[
          provider
        ] ||
        {};


      const account =
        providerMap[
          accountKey
        ];


      if (!account) {

        row[
          h['Validation']
        ] =
          'ACCOUNT NOT FOUND';


        flowSetExistingPayoutTxState_(
          existingTx,
          th,
          'INVALID'
        );


        invalid++;

        return;

      }


      if (
        account.status ===
        'DISABLED'
      ) {

        row[
          h['Validation']
        ] =
          'ACCOUNT DISABLED';


        flowSetExistingPayoutTxState_(
          existingTx,
          th,
          'INVALID'
        );


        invalid++;

        return;

      }


      /********************************************************
       * AMOUNT VALIDATION
       ********************************************************/

      if (
        amount === null ||
        amount <= 0
      ) {

        row[
          h['Validation']
        ] =
          'INVALID AMOUNT';


        flowSetExistingPayoutTxState_(
          existingTx,
          th,
          'INVALID'
        );


        invalid++;

        return;

      }


      if (
        fee === null ||
        fee < 0
      ) {

        row[
          h['Validation']
        ] =
          'INVALID FEE';


        flowSetExistingPayoutTxState_(
          existingTx,
          th,
          'INVALID'
        );


        invalid++;

        return;

      }


      /********************************************************
       * PAYOUT DATE
       ********************************************************/

      let payoutDate =
        flowDate_(
          row[
            h['Payout Date']
          ]
        );


      /*
       * Nếu PAID nhưng chưa nhập date,
       * dùng thời điểm hiện tại.
       */
      if (!payoutDate) {

        payoutDate =
          new Date();


        row[
          h['Payout Date']
        ] =
          payoutDate;

      }


      /********************************************************
       * PAID AT
       ********************************************************/

      if (
        !row[
          h['Paid At']
        ]
      ) {

        row[
          h['Paid At']
        ] =
          new Date();

      }


      const totalDeducted =
        flowRoundMoney_(
          amount +
          fee
        );


      row[
        h['Total Deducted']
      ] =
        totalDeducted;


      /********************************************************
       * TRANSACTION
       ********************************************************/

      const transactionKey =
        'POTX|' +
        payoutKey;


      const type =
        provider ===
          'USDT'
          ? 'USDT_PAYOUT'
          : 'PINGPONG_PAYOUT';


      flowUpsertTransaction_(
        state,
        th,
        {

          'Transaction Key':
            transactionKey,

          'Transaction Date':
            payoutDate,

          'Provider':
            provider,

          'Account Key':
            accountKey,

          'Label':
            account.label ||
            accountKey,

          'Type':
            type,

          'Direction':
            'OUT',

          /*
           * Amount actually lost by source account.
           */
          'Gross Amount':
            totalDeducted,

          /*
           * Amount receiver gets.
           */
          'Net Amount':
            flowRoundMoney_(
              amount
            ),

          'Fee Amount':
            flowRoundMoney_(
              fee
            ),

          'Currency':
            expectedCurrency,

          'Reference':
            flowCell_(
              row,
              h,
              'Reference'
            ),

          'Counterparty Provider':
            '',

          'Counterparty Account Key':
            '',

          'Reconcile Status':
            'PAID',

          'Matched Transaction Key':
            '',

          'Source':
            'PAYOUTS',

          'Source ID':
            payoutKey,

          'Sync Time':
            new Date(),

          'Unique Key':
            uniqueKey

        }
      );


      row[
        h['Transaction Key']
      ] =
        transactionKey;


      row[
        h['Validation']
      ] =
        'OK';


      paid++;

    }
  );


  return {

    paid:
      paid,

    draft:
      draft,

    cancelled:
      cancelled,

    invalid:
      invalid,

    blocked:
      blocked

  };

}


/************************************************************
 * EXISTING PAYOUT TX STATUS
 ************************************************************/

/************************************************************
 * MANUAL IN -> TRANSACTION
 *
 * Đối xứng với Payouts nhưng chỉ một chiều: tiền RA đã có Payouts lo, sheet
 * này chỉ nhận tiền VÀO. Không trùng vai, nên không có chuyện một khoản được
 * ghi ở hai nơi.
 *
 * Không có trạng thái DRAFT: tiền đã nằm trong ví rồi, không có giai đoạn "dự
 * định" như một lệnh chi. Gõ nhầm thì đổi sang CANCELLED — dòng transaction
 * tương ứng bị đánh dấu chứ không xoá, để còn dấu vết.
 ************************************************************/

function flowSyncManualInToTransaction_(
  rows,
  h,
  providerMaps,
  state,
  th
) {

  const required = [
    'Entry Key',
    'Date',
    'Provider',
    'Account Key',
    'Amount',
    'Status'
  ];


  for (
    let i = 0;
    i < required.length;
    i++
  ) {

    if (
      !Object.prototype.hasOwnProperty.call(
        h,
        required[i]
      )
    ) {

      Logger.log(
        'Manual In: thiếu cột "' + required[i] +
        '", bỏ qua sheet này. Header phải nằm ở hàng ' +
        FLOW_CFG.HEADER_ROW + ' và viết đúng từng ký tự.'
      );

      return {
        confirmed: 0,
        cancelled: 0,
        invalid: 0
      };

    }

  }


  /********************************************************
   * CẤP KHOÁ
   ********************************************************/

  let maxKeyNumber = 0;


  rows.forEach(
    row => {

      const match =
        flowCell_(row, h, 'Entry Key')
          .match(/^ME-(\d+)$/i);


      if (match) {

        maxKeyNumber =
          Math.max(
            maxKeyNumber,
            Number(match[1])
          );

      }

    }
  );


  let confirmed = 0;
  let cancelled = 0;
  let invalid = 0;


  rows.forEach(
    row => {

      /*
       * Dòng chỉ có giá trị dropdown mặc định thì chưa phải dòng thật.
       */
      if (
        !flowManualInRowHasInput_(row, h)
      ) {
        return;
      }


      let entryKey =
        flowCell_(row, h, 'Entry Key');


      if (!entryKey) {

        maxKeyNumber++;

        entryKey =
          'ME-' +
          String(maxKeyNumber).padStart(6, '0');


        row[h['Entry Key']] = entryKey;

      }


      if (
        h['Created At'] &&
        !row[h['Created At']]
      ) {
        row[h['Created At']] = new Date();
      }


      const uniqueKey =
        'MANUAL|' + entryKey;


      const existingTx =
        flowFindTransactionByUnique_(
          state,
          uniqueKey
        );


      const fail = reason => {

        if (h['Validation']) {
          row[h['Validation']] = reason;
        }

        flowSetExistingPayoutTxState_(
          existingTx,
          th,
          'INVALID'
        );

        invalid++;

      };


      const provider =
        flowCell_(row, h, 'Provider')
          .toUpperCase();


      if (provider) {
        row[h['Provider']] = provider;
      }


      const status =
        flowCell_(row, h, 'Status')
          .toUpperCase();


      /********************************************************
       * CANCELLED
       *
       * Giữ dòng, huỷ tác dụng. flowTransactionCounts_ loại nó khỏi mọi số dư.
       ********************************************************/

      if (status === 'CANCELLED') {

        if (h['Validation']) {
          row[h['Validation']] = 'CANCELLED';
        }

        flowSetExistingPayoutTxState_(
          existingTx,
          th,
          'CANCELLED'
        );

        cancelled++;

        return;

      }


      if (status !== 'CONFIRMED') {

        fail(
          status
            ? 'INVALID STATUS'
            : 'MISSING STATUS'
        );

        return;

      }


      if (
        provider !== 'USDT' &&
        provider !== 'PINGPONG'
      ) {

        fail(
          provider
            ? 'INVALID PROVIDER'
            : 'MISSING PROVIDER'
        );

        return;

      }


      const accountKey =
        flowCell_(row, h, 'Account Key');


      if (!accountKey) {
        fail('MISSING ACCOUNT');
        return;
      }


      const account =
        (
          providerMaps[provider] ||
          {}
        )[accountKey];


      if (!account) {
        fail('ACCOUNT NOT FOUND');
        return;
      }


      if (account.status === 'DISABLED') {
        fail('ACCOUNT DISABLED');
        return;
      }


      const amount =
        flowNullableNumber_(
          row[h['Amount']]
        );


      if (
        amount === null ||
        amount <= 0
      ) {
        fail('INVALID AMOUNT');
        return;
      }


      /*
       * Không có ngày thì lấy hôm nay, giống cách Payouts xử lý — thiếu ngày
       * không đáng để chặn một khoản tiền đã thật sự vào ví.
       */
      let date =
        flowDate_(row[h['Date']]);


      if (!date) {
        date = new Date();
        row[h['Date']] = date;
      }


      const currency =
        flowPayoutCurrencyForProvider_(provider);


      if (h['Currency']) {
        row[h['Currency']] = currency;
      }


      const transactionKey =
        'MITX|' + entryKey;


      flowUpsertTransaction_(
        state,
        th,
        {

          'Transaction Key':
            transactionKey,

          'Transaction Date':
            date,

          'Provider':
            provider,

          'Account Key':
            accountKey,

          'Label':
            account.label ||
            accountKey,

          /*
           * Type riêng, KHÔNG dùng lại USDT_RECEIVED: khoản này không đến từ
           * Grey, và tổng trên sheet USDT tách "từ PayPal" với "nhận thêm"
           * đúng bằng cách nhìn vào Type này.
           */
          'Type':
            'MANUAL_IN',

          'Direction':
            'IN',

          'Gross Amount':
            amount,

          'Net Amount':
            amount,

          'Fee Amount':
            0,

          'Currency':
            currency,

          'Reference':
            flowCell_(row, h, 'From'),

          'Counterparty Provider':
            '',

          'Counterparty Account Key':
            '',

          'Reconcile Status':
            'ARRIVED',

          'Matched Transaction Key':
            '',

          'Source':
            'MANUAL IN',

          'Source ID':
            entryKey,

          'Sync Time':
            new Date(),

          'Unique Key':
            uniqueKey

        }
      );


      if (h['Transaction Key']) {
        row[h['Transaction Key']] = transactionKey;
      }


      if (h['Validation']) {
        row[h['Validation']] = 'OK';
      }


      confirmed++;

    }
  );


  Logger.log(
    'Manual In: ' + confirmed + ' confirmed · ' +
    cancelled + ' cancelled · ' + invalid + ' invalid'
  );


  return {
    confirmed: confirmed,
    cancelled: cancelled,
    invalid: invalid
  };

}


/*
 * Provider/Status có thể là giá trị dropdown mặc định trên một dòng trống, nên
 * riêng chúng không đủ để coi dòng là có dữ liệu — giống hệt lý do bên Payouts.
 */
function flowManualInRowHasInput_(
  row,
  h
) {

  return [
    'Entry Key',
    'Date',
    'Account Key',
    'Amount',
    'From',
    'Note',
    'Transaction Key'
  ].some(
    header =>

      Object.prototype.hasOwnProperty.call(
        h,
        header
      ) &&

      flowString_(
        row[h[header]]
      ) !== ''
  );

}


function flowSetExistingPayoutTxState_(
  tx,
  h,
  status
) {

  if (!tx) {
    return;
  }


  tx[
    h['Reconcile Status']
  ] =
    status;


  tx[
    h['Sync Time']
  ] =
    new Date();

}


/************************************************************
 * PAYOUT PROVIDER ACCOUNT MAP
 *
 * Pingpong may currently be empty.
 ************************************************************/

function flowReadProviderAccountMap_(
  sheet,
  labelCandidates
) {

  const result = {};


  if (!sheet) {
    return result;
  }


  if (
    sheet.getLastColumn() < 1 ||
    sheet.getLastRow() <
      FLOW_CFG.HEADER_ROW
  ) {

    return result;

  }


  const h =
    flowHeaderMap_(sheet);


  if (
    !Object.prototype.hasOwnProperty.call(
      h,
      'Account Key'
    )
  ) {

    return result;

  }


  const rows =
    flowReadRows_(sheet);


  rows.forEach(
    row => {

      const accountKey =
        flowString_(
          row[
            h['Account Key']
          ]
        );


      if (!accountKey) {
        return;
      }


      let label = '';


      for (
        let i = 0;
        i < labelCandidates.length;
        i++
      ) {

        const candidate =
          labelCandidates[i];


        if (
          !Object.prototype.hasOwnProperty.call(
            h,
            candidate
          )
        ) {
          continue;
        }


        label =
          flowString_(
            row[
              h[candidate]
            ]
          );


        if (label) {
          break;
        }

      }


      let status = '';


      if (
        Object.prototype.hasOwnProperty.call(
          h,
          'Status'
        )
      ) {

        status =
          flowString_(
            row[
              h['Status']
            ]
          )
            .toUpperCase();

      }


      result[
        accountKey
      ] = {

        label:
          label ||
          accountKey,

        status:
          status

      };

    }
  );


  return result;

}


/************************************************************
 * PAYPAL STATUS -> HOLD BALANCE
 ************************************************************/

function flowRefreshHoldBalanceRow_(
  sheet,
  rowNumber,
  h
) {

  const required = [
    'Provider',
    'PayPal Status',
    'Total Balance',
    'Withheld Balance',
    'Hold Balance'
  ];


  for (
    let i = 0;
    i < required.length;
    i++
  ) {

    if (
      !Object.prototype.hasOwnProperty.call(
        h,
        required[i]
      )
    ) {
      return;
    }

  }


  const provider =
    flowString_(
      sheet
        .getRange(
          rowNumber,
          h['Provider'] + 1
        )
        .getValue()
    )
      .toUpperCase();


  if (
    provider !==
    'PAYPAL'
  ) {
    return;
  }


  const status =
    flowString_(
      sheet
        .getRange(
          rowNumber,
          h['PayPal Status'] + 1
        )
        .getValue()
    )
      .toUpperCase();


  const totalBalance =
    flowNumber_(
      sheet
        .getRange(
          rowNumber,
          h['Total Balance'] + 1
        )
        .getValue()
    );


  const withheldBalance =
    flowNumber_(
      sheet
        .getRange(
          rowNumber,
          h['Withheld Balance'] + 1
        )
        .getValue()
    );


  const holdBalance =
    status ===
      'SUSPENDED'
      ? totalBalance
      : withheldBalance;


  sheet
    .getRange(
      rowNumber,
      h['Hold Balance'] + 1
    )
    .setValue(
      flowRoundMoney_(
        holdBalance
      )
    );

}


/************************************************************
 * ACCOUNT SUMMARY
 ************************************************************/

function flowUpdateAccountSummary_(
  sheet,
  h,
  transactionRows,
  th
) {

  const rows =
    flowReadRows_(sheet);


  rows.forEach(
    row => {

      const provider =
        flowString_(
          row[
            h['Provider']
          ]
        )
          .toUpperCase();


      if (
        provider !==
        'PAYPAL'
      ) {
        return;
      }


      const accountKey =
        flowString_(
          row[
            h['Account Key']
          ]
        );


      if (!accountKey) {
        return;
      }


      let totalWithdrawn = 0;


      transactionRows.forEach(
        tx => {

          if (
            flowCell_(
              tx,
              th,
              'Provider'
            )
              .toUpperCase() !==
              'PAYPAL'
          ) {
            return;
          }


          if (
            flowCell_(
              tx,
              th,
              'Account Key'
            ) !==
            accountKey
          ) {
            return;
          }


          if (
            flowCell_(
              tx,
              th,
              'Type'
            )
              .toUpperCase() !==
              'PAYPAL_WITHDRAWAL'
          ) {
            return;
          }


          /*
           * Returned no longer counts.
           */
          if (
            flowCell_(
              tx,
              th,
              'Reconcile Status'
            )
              .toUpperCase() ===
              'RETURNED'
          ) {
            return;
          }


          totalWithdrawn +=
            flowNumber_(
              tx[
                th['Gross Amount']
              ]
            );

        }
      );


      const status =
        flowString_(
          row[
            h['PayPal Status']
          ]
        )
          .toUpperCase();


      const totalBalance =
        flowNumber_(
          row[
            h['Total Balance']
          ]
        );


      const withheld =
        flowNumber_(
          row[
            h['Withheld Balance']
          ]
        );


      row[
        h['Total Withdrawn']
      ] =
        flowRoundMoney_(
          totalWithdrawn
        );


      row[
        h['Hold Balance']
      ] =
        flowRoundMoney_(
          status ===
            'SUSPENDED'
            ? totalBalance
            : withheld
        );

    }
  );


  flowWriteRows_(
    sheet,
    rows
  );

}


/************************************************************
 * PAYPAL RAW PROCESSING
 ************************************************************/

function flowProcessPaypalSheet_(
  rows,
  h
) {

  rows.forEach(
    row => {

      const eventCode =
        flowCell_(
          row,
          h,
          'Event Code'
        )
          .toUpperCase();


      let rawStatus =
        flowCell_(
          row,
          h,
          'Status'
        );


      if (
        Object.prototype.hasOwnProperty.call(
          h,
          'Raw Status'
        )
      ) {

        rawStatus =
          flowCell_(
            row,
            h,
            'Raw Status'
          );

      }


      const invoice =
        flowCell_(
          row,
          h,
          'Invoice ID'
        );


      const amount =
        flowNumber_(
          row[
            h['Transaction Amount']
          ]
        );


      const fee =
        flowNumber_(
          row[
            h['Fee Amount']
          ]
        );


      row[
        h['Type']
      ] =
        flowPaypalType_(
          eventCode
        );


      row[
        h['Status']
      ] =
        flowPaypalStatus_(
          rawStatus
        );


      row[
        h['Invoice Prefix']
      ] =
        flowInvoicePrefix_(
          invoice
        );


      row[
        h['Balance Impact']
      ] =
        flowRoundMoney_(
          amount +
          fee
        );

    }
  );

}


/************************************************************
 * PAYPAL TYPE
 ************************************************************/

function flowPaypalType_(
  code
) {

  const value =
    flowString_(code)
      .toUpperCase();


  if (!value) {
    return '';
  }


  const map = {

    T0000: 'GENERAL_PAYMENT',
    T0001: 'MASS_PAYMENT',
    T0002: 'SUBSCRIPTION_PAYMENT',
    T0003: 'PREAPPROVED_PAYMENT',
    T0004: 'ECHECK_PAYMENT',
    T0005: 'CREDIT_CARD_PAYMENT',
    T0006: 'EXPRESS_CHECKOUT_PAYMENT',
    T0007: 'WEBSITE_PAYMENT',
    T0008: 'POSTAGE_PAYMENT',

    T0114: 'DISPUTE_FEE',

    T0200: 'CURRENCY_CONVERSION',

    T0400: 'BANK_WITHDRAWAL',
    T0401: 'BANK_WITHDRAWAL',
    T0403: 'BANK_WITHDRAWAL',

    /*
     * T1101 là dòng PayPal ghi khi trả lại một khoản đã đi ra — gặp thực
     * tế với khoản rút bị từ chối: cùng Transaction ID, ngược dấu, Status
     * DENIED, còn dòng rút gốc vẫn nguyên PENDING. PASS 2 dựa vào cặp đó.
     */
    T1101: 'REVERSAL',

    T1105: 'HOLD_RELEASE',
    T1106: 'PAYMENT_REVERSAL',
    T1107: 'PAYMENT_REFUND',

    T1110: 'HOLD',
    T1111: 'DISPUTE_HOLD_RELEASE',

    T1201: 'CHARGEBACK',

    T1501: 'HOLD'

  };


  return (
    map[value] ||
    'UNMAPPED_' +
    value
  );

}


/************************************************************
 * PAYPAL STATUS
 ************************************************************/

function flowPaypalStatus_(
  raw
) {

  const value =
    flowString_(raw)
      .toUpperCase();


  if (!value) {
    return '';
  }


  const map = {

    S: 'SUCCESS',
    P: 'PENDING',
    V: 'REVERSED',
    D: 'DENIED'

  };


  return (
    map[value] ||
    value
  );

}


/************************************************************
 * INVOICE PREFIX
 ************************************************************/

function flowInvoicePrefix_(
  invoice
) {

  const value =
    flowString_(invoice);


  if (!value) {
    return '';
  }


  const index =
    value.indexOf('-');


  return index ===
    -1
      ? value
      : value.substring(
          0,
          index
        );

}


/************************************************************
 * PAYPAL -> TRANSACTION
 ************************************************************/

function flowSyncPaypalToTransaction_(
  rows,
  ph,
  accountMap,
  state,
  th
) {

  let withdrawalCount = 0;
  let returnCount = 0;


  const withdrawalIds = {};


  /*
   * Số tiền của từng khoản rút, để PASS 2 đối chiếu được khi dòng trả về
   * nối bằng Transaction ID chung thay vì bằng Reference ID.
   */
  const withdrawalAmounts = {};


  /**********************************************************
   * PASS 1 — WITHDRAWALS
   **********************************************************/

  rows.forEach(
    row => {

      const accountKey =
        flowCell_(
          row,
          ph,
          'Account Key'
        );


      const transactionId =
        flowCell_(
          row,
          ph,
          'Transaction ID'
        );


      const eventCode =
        flowCell_(
          row,
          ph,
          'Event Code'
        )
          .toUpperCase();


      const type =
        flowCell_(
          row,
          ph,
          'Type'
        )
          .toUpperCase();


      if (
        !flowIsPaypalWithdrawal_(
          eventCode,
          type
        )
      ) {
        return;
      }


      if (
        !accountKey ||
        !transactionId
      ) {
        return;
      }


      withdrawalCount++;


      withdrawalIds[
        accountKey +
        '|' +
        transactionId
      ] =
        true;


      withdrawalAmounts[
        accountKey +
        '|' +
        transactionId
      ] =
        flowRoundMoney_(
          Math.abs(
            flowNumber_(
              row[
                ph['Transaction Amount']
              ]
            )
          )
        );


      const account =
        accountMap[
          accountKey
        ] ||
        {};


      const uniqueKey =
        'PAYPAL|' +
        accountKey +
        '|' +
        transactionId +
        '|' +
        eventCode;


      const existing =
        flowFindTransactionByUnique_(
          state,
          uniqueKey
        );


      let counterpartyProvider =
        flowString_(
          account.linkedProvider
        )
          .toUpperCase();


      let counterpartyAccountKey =
        flowString_(
          account.linkedAccountKey
        );


      /*
       * Preserve historical mapping.
       */
      if (existing) {

        const oldProvider =
          flowCell_(
            existing,
            th,
            'Counterparty Provider'
          );


        const oldAccount =
          flowCell_(
            existing,
            th,
            'Counterparty Account Key'
          );


        if (oldProvider) {

          counterpartyProvider =
            oldProvider;

        }


        if (oldAccount) {

          counterpartyAccountKey =
            oldAccount;

        }

      }


      const paypalStatus =
        flowCell_(
          row,
          ph,
          'Status'
        )
          .toUpperCase();


      let reconcileStatus =
        (
          counterpartyProvider &&
          counterpartyAccountKey
        )
          ? 'IN_TRANSIT'
          : 'UNMATCHED';


      if (
        paypalStatus ===
          'REVERSED' ||
        paypalStatus ===
          'DENIED'
      ) {

        reconcileStatus =
          'RETURNED';

      }


      const amount =
        Math.abs(
          flowNumber_(
            row[
              ph['Transaction Amount']
            ]
          )
        );


      const fee =
        Math.abs(
          flowNumber_(
            row[
              ph['Fee Amount']
            ]
          )
        );


      flowUpsertTransaction_(
        state,
        th,
        {

          'Transaction Key':
            'PTX|' +
            accountKey +
            '|' +
            transactionId,

          'Transaction Date':
            row[
              ph['Initiation Date']
            ],

          'Provider':
            'PAYPAL',

          'Account Key':
            accountKey,

          'Label':
            flowCell_(
              row,
              ph,
              'Label'
            ),

          'Type':
            'PAYPAL_WITHDRAWAL',

          'Direction':
            'OUT',

          'Gross Amount':
            flowRoundMoney_(
              amount
            ),

          'Net Amount':
            flowRoundMoney_(
              amount
            ),

          'Fee Amount':
            flowRoundMoney_(
              fee
            ),

          'Currency':
            flowCell_(
              row,
              ph,
              'Currency'
            )
              .toUpperCase(),

          'Reference':
            transactionId,

          'Counterparty Provider':
            counterpartyProvider,

          'Counterparty Account Key':
            counterpartyAccountKey,

          'Reconcile Status':
            reconcileStatus,

          'Matched Transaction Key':
            '',

          'Source':
            'PAYPAL_RAW',

          'Source ID':
            flowCell_(
              row,
              ph,
              'Unique Key'
            ),

          'Sync Time':
            new Date(),

          'Unique Key':
            uniqueKey

        }
      );

    }
  );


  /**********************************************************
   * PASS 2 — RETURNS
   **********************************************************/

  rows.forEach(
    row => {

      const accountKey =
        flowCell_(
          row,
          ph,
          'Account Key'
        );


      const transactionId =
        flowCell_(
          row,
          ph,
          'Transaction ID'
        );


      const referenceId =
        flowCell_(
          row,
          ph,
          'Reference ID'
        );


      if (
        !accountKey ||
        !transactionId
      ) {
        return;
      }


      /*
       * PayPal nối dòng trả về với khoản rút theo HAI cách khác nhau:
       *
       *   1. Reference ID trỏ tới Transaction ID của khoản rút.
       *
       *   2. Dòng trả về dùng CHÍNH Transaction ID của khoản rút và bỏ
       *      trống Reference ID. Gặp thực tế với T1101: khoản rút và dòng
       *      trả về chung một Transaction ID, chỉ khác Event Code, và
       *      PayPal KHÔNG lật trạng thái dòng rút gốc — nó vẫn là PENDING.
       *
       * Chỉ có nhánh 1 thì khoản rút bị trả về nằm IN_TRANSIT vĩnh viễn:
       * nó vẫn giữ chỗ nguồn của khách nên tiền đó không bao giờ phân bổ
       * lại được cho lần rút sau, mà cũng không ai hay biết.
       *
       * Nhánh 2 mới nên siết chặt hơn nhánh 1: đòi số tiền khớp đúng
       * khoản rút. Cùng Transaction ID, ngược dấu, cùng số tiền thì gần
       * như không thể là thứ gì khác.
       */
      let linkedId = '';
      let linkedBySharedId = false;

      if (
        referenceId &&
        withdrawalIds[
          accountKey +
          '|' +
          referenceId
        ]
      ) {

        linkedId = referenceId;

      } else if (
        withdrawalIds[
          accountKey +
          '|' +
          transactionId
        ]
      ) {

        linkedId = transactionId;
        linkedBySharedId = true;

      }


      if (!linkedId) {
        return;
      }


      const eventCode =
        flowCell_(
          row,
          ph,
          'Event Code'
        )
          .toUpperCase();


      if (
        flowIsPaypalWithdrawal_(
          eventCode,
          flowCell_(
            row,
            ph,
            'Type'
          )
        )
      ) {
        return;
      }


      const amount =
        flowNumber_(
          row[
            ph['Transaction Amount']
          ]
        );


      if (
        amount <= 0
      ) {
        return;
      }


      /*
       * Siết nhánh Transaction ID chung: chỉ nhận khi số tiền khớp đúng
       * khoản rút. Nhánh Reference ID cũ giữ nguyên, không siết thêm, để
       * không đổi cách đọc dữ liệu lịch sử.
       */
      if (
        linkedBySharedId &&
        !flowAmountEqual_(
          withdrawalAmounts[
            accountKey +
            '|' +
            linkedId
          ],
          amount
        )
      ) {
        return;
      }


      if (linkedBySharedId) {

        Logger.log(
          '  ↩ trả về khoản rút (Transaction ID chung) · ' +
          accountKey + ' ' + transactionId +
          ' · ' + eventCode +
          ' · ' + flowRoundMoney_(amount).toFixed(2)
        );

      }


      returnCount++;


      flowUpsertTransaction_(
        state,
        th,
        {

          'Transaction Key':
            'PRT|' +
            accountKey +
            '|' +
            transactionId,

          'Transaction Date':
            row[
              ph['Initiation Date']
            ],

          'Provider':
            'PAYPAL',

          'Account Key':
            accountKey,

          'Label':
            flowCell_(
              row,
              ph,
              'Label'
            ),

          'Type':
            'PAYPAL_WITHDRAWAL_RETURN',

          'Direction':
            'IN',

          'Gross Amount':
            flowRoundMoney_(
              Math.abs(amount)
            ),

          'Net Amount':
            flowRoundMoney_(
              Math.abs(amount)
            ),

          'Fee Amount':
            flowRoundMoney_(
              Math.abs(
                flowNumber_(
                  row[
                    ph['Fee Amount']
                  ]
                )
              )
            ),

          'Currency':
            flowCell_(
              row,
              ph,
              'Currency'
            )
              .toUpperCase(),

          /*
           * Matcher ghép chính xác bằng Reference của khoản rút, mà khoản
           * rút lấy Transaction ID của nó làm Reference. Nên phải ghi
           * linkedId chứ không phải referenceId — nhánh Transaction ID
           * chung có referenceId rỗng.
           */
          'Reference':
            linkedId,

          'Counterparty Provider':
            '',

          'Counterparty Account Key':
            '',

          'Reconcile Status':
            'UNMATCHED',

          'Matched Transaction Key':
            '',

          'Source':
            'PAYPAL_RAW',

          'Source ID':
            flowCell_(
              row,
              ph,
              'Unique Key'
            ),

          'Sync Time':
            new Date(),

          'Unique Key':
            'PAYPAL|' +
            accountKey +
            '|' +
            transactionId +
            '|' +
            eventCode

        }
      );

    }
  );


  return {

    withdrawals:
      withdrawalCount,

    returns:
      returnCount

  };

}


/************************************************************
 * PAYPAL WITHDRAWAL?
 ************************************************************/

function flowIsPaypalWithdrawal_(
  eventCode,
  type
) {

  const code =
    flowString_(eventCode)
      .toUpperCase();


  const normalizedType =
    flowString_(type)
      .toUpperCase();


  return (
    code === 'T0400' ||
    code === 'T0401' ||
    code === 'T0403' ||
    normalizedType ===
      'BANK_WITHDRAWAL'
  );

}


/************************************************************
 * GREY RAW -> NORMALIZED
 ************************************************************/

function flowProcessGreyRawRows_(
  rows,
  h,
  greyMap
) {

  let processed = 0;
  let failed = 0;


  rows.forEach(
    row => {

      if (
        flowCell_(
          row,
          h,
          'Provider'
        )
          .toUpperCase() !==
          'GREY' ||

        flowCell_(
          row,
          h,
          'Source'
        )
          .toUpperCase() !==
          'GREY_GMAIL'
      ) {
        return;
      }


      const parsed =
        flowInterpretGreyRaw_(
          flowCell_(
            row,
            h,
            'Raw Event'
          ),

          flowCell_(
            row,
            h,
            'Raw Transaction Type'
          ),

          flowCell_(
            row,
            h,
            'Raw Amount Tendered'
          ),

          flowCell_(
            row,
            h,
            'Raw Amount Received'
          ),

          flowCell_(
            row,
            h,
            'Raw Channel'
          ),

          flowCell_(
            row,
            h,
            'Raw Date Time'
          )
        );


      if (!parsed) {

        failed++;

        return;

      }


      processed++;


      const accountKey =
        flowCell_(
          row,
          h,
          'Account Key'
        );


      const reference =
        flowCell_(
          row,
          h,
          'Reference'
        );


      const sourceId =
        flowCell_(
          row,
          h,
          'Source ID'
        );


      row[
        h['Transaction Key']
      ] =
        'GTX|' +
        accountKey +
        '|' +
        (
          reference ||
          sourceId
        );


      if (
        parsed.transactionDate
      ) {

        row[
          h['Transaction Date']
        ] =
          parsed.transactionDate;

      }


      row[
        h['Type']
      ] =
        parsed.type;


      row[
        h['Direction']
      ] =
        parsed.direction;


      row[
        h['Gross Amount']
      ] =
        parsed.gross;


      row[
        h['Net Amount']
      ] =
        parsed.net;


      row[
        h['Fee Amount']
      ] =
        parsed.fee;


      row[
        h['Currency']
      ] =
        parsed.currency;


      /*
       * Rebuild each Flow.
       */
      row[
        h['Reconcile Status']
      ] =
        'UNMATCHED';


      row[
        h['Matched Transaction Key']
      ] =
        '';


      /********************************************************
       * GREY FUNDING
       ********************************************************/

      if (
        parsed.type ===
        'GREY_FUNDING'
      ) {

        row[
          h['Counterparty Provider']
        ] =
          '';


        row[
          h['Counterparty Account Key']
        ] =
          '';

      }


      /********************************************************
       * GREY P2P
       ********************************************************/

      if (
        parsed.type ===
        'GREY_P2P'
      ) {

        row[
          h['Counterparty Provider']
        ] =
          'GREY';

      }


      /********************************************************
       * GREY -> USDT
       ********************************************************/

      if (
        parsed.type ===
          'GREY_WITHDRAWAL' &&
        parsed.currency ===
          'USDT'
      ) {

        row[
          h['Counterparty Provider']
        ] =
          'USDT';


        const currentDestination =
          flowCell_(
            row,
            h,
            'Counterparty Account Key'
          );


        /*
         * Preserve historical destination.
         */
        if (!currentDestination) {

          const account =
            greyMap[
              accountKey
            ] ||
            {};


          row[
            h['Counterparty Account Key']
          ] =
            flowString_(
              account.defaultUsdtAccount
            );

        }

      }

    }
  );


  return {

    processed:
      processed,

    failed:
      failed

  };

}


/************************************************************
 * INTERPRET GREY
 ************************************************************/

function flowInterpretGreyRaw_(
  rawEvent,
  rawTransactionType,
  rawTendered,
  rawReceived,
  rawChannel,
  rawDateTime
) {

  const event =
    flowString_(rawEvent)
      .toLowerCase();


  const isFunding =
    event.indexOf(
      'fund'
    ) !== -1;


  const isWithdrawal =
    event.indexOf(
      'withdraw'
    ) !== -1;


  if (
    !isFunding &&
    !isWithdrawal
  ) {

    return null;

  }


  const tendered =
    flowParseGreyRawMoney_(
      rawTendered,
      rawTransactionType
    );


  const received =
    flowParseGreyRawMoney_(
      rawReceived,
      rawTransactionType
    );


  if (
    tendered.amount === null ||
    received.amount === null
  ) {

    return null;

  }


  const currency =
    String(
      tendered.currency ||
      received.currency ||
      flowGreyCurrencyFromType_(
        rawTransactionType
      )
    )
      .toUpperCase();


  if (!currency) {
    return null;
  }


  let type = '';
  let direction = '';


  if (isFunding) {

    type =
      'GREY_FUNDING';

    direction =
      'IN';


  } else if (
    flowString_(rawChannel)
      .toLowerCase() ===
      'p2p'
  ) {

    type =
      'GREY_P2P';

    direction =
      'OUT';


  } else {

    type =
      'GREY_WITHDRAWAL';

    direction =
      'OUT';

  }


  return {

    type:
      type,

    direction:
      direction,

    gross:
      flowRoundMoney_(
        tendered.amount
      ),

    net:
      flowRoundMoney_(
        received.amount
      ),

    fee:
      flowRoundMoney_(
        Math.abs(
          tendered.amount -
          received.amount
        )
      ),

    currency:
      currency,

    transactionDate:
      flowParseGreyDate_(
        rawDateTime
      )

  };

}


/************************************************************
 * GREY MONEY
 ************************************************************/

function flowParseGreyRawMoney_(
  raw,
  transactionType
) {

  const text =
    flowString_(raw);


  if (!text) {

    return {

      amount:
        null,

      currency:
        ''

    };

  }


  let currency = '';


  if (
    /USDT/i.test(text)
  ) {

    currency =
      'USDT';


  } else if (
    /\$/.test(text) ||
    /\bUSD\b/i.test(text)
  ) {

    currency =
      'USD';


  } else {

    currency =
      flowGreyCurrencyFromType_(
        transactionType
      );

  }


  return {

    amount:
      flowNullableNumber_(text),

    currency:
      currency

  };

}


/************************************************************
 * GREY CURRENCY
 ************************************************************/

function flowGreyCurrencyFromType_(
  raw
) {

  const value =
    flowString_(raw)
      .toUpperCase();


  if (
    value.indexOf(
      'USDT'
    ) !== -1
  ) {
    return 'USDT';
  }


  if (
    value.indexOf(
      'USD'
    ) !== -1
  ) {
    return 'USD';
  }


  return '';

}


/************************************************************
 * GREY UTC DATE
 ************************************************************/

function flowParseGreyDate_(
  raw
) {

  const text =
    flowString_(raw);


  if (!text) {
    return null;
  }


  const match =
    text.match(
      /(\d{1,2})\/(\d{1,2})\/(\d{4})\s*-\s*(\d{1,2}):(\d{2})\s*(AM|PM)\s*UTC/i
    );


  if (!match) {

    return flowDate_(
      text
    );

  }


  const month =
    Number(match[1]);

  const day =
    Number(match[2]);

  const year =
    Number(match[3]);

  let hour =
    Number(match[4]);

  const minute =
    Number(match[5]);

  const ampm =
    String(match[6])
      .toUpperCase();


  if (
    ampm === 'PM' &&
    hour !== 12
  ) {

    hour += 12;

  }


  if (
    ampm === 'AM' &&
    hour === 12
  ) {

    hour = 0;

  }


  return new Date(
    Date.UTC(
      year,
      month - 1,
      day,
      hour,
      minute,
      0
    )
  );

}


/************************************************************
 * PAYPAL RETURN RECONCILIATION
 ************************************************************/

function flowReconcilePaypalReturns_(
  rows,
  h
) {

  const withdrawals =
    rows.filter(
      row =>

        flowCell_(
          row,
          h,
          'Provider'
        )
          .toUpperCase() ===
          'PAYPAL' &&

        flowCell_(
          row,
          h,
          'Type'
        )
          .toUpperCase() ===
          'PAYPAL_WITHDRAWAL'
    );


  const returns =
    rows.filter(
      row =>

        flowCell_(
          row,
          h,
          'Provider'
        )
          .toUpperCase() ===
          'PAYPAL' &&

        flowCell_(
          row,
          h,
          'Type'
        )
          .toUpperCase() ===
          'PAYPAL_WITHDRAWAL_RETURN'
    );


  let matched = 0;


  returns.forEach(
    returnRow => {

      const accountKey =
        flowCell_(
          returnRow,
          h,
          'Account Key'
        );


      const reference =
        flowCell_(
          returnRow,
          h,
          'Reference'
        );


      const currency =
        flowCell_(
          returnRow,
          h,
          'Currency'
        )
          .toUpperCase();


      const amount =
        flowNumber_(
          returnRow[
            h['Gross Amount']
          ]
        );


      const returnDate =
        flowMillis_(
          returnRow[
            h['Transaction Date']
          ]
        );


      /*
       * Exact reference first.
       */
      let candidate =
        withdrawals.find(
          withdrawal =>

            flowCell_(
              withdrawal,
              h,
              'Account Key'
            ) ===
              accountKey &&

            flowCell_(
              withdrawal,
              h,
              'Reference'
            ) ===
              reference
        );


      /*
       * Fallback amount/date.
       */
      if (!candidate) {

        const possible =
          withdrawals
            .filter(
              withdrawal => {

                const wdDate =
                  flowMillis_(
                    withdrawal[
                      h['Transaction Date']
                    ]
                  );


                return (

                  flowCell_(
                    withdrawal,
                    h,
                    'Account Key'
                  ) ===
                    accountKey &&

                  flowCell_(
                    withdrawal,
                    h,
                    'Currency'
                  )
                    .toUpperCase() ===
                    currency &&

                  flowAmountEqual_(
                    flowNumber_(
                      withdrawal[
                        h['Gross Amount']
                      ]
                    ),
                    amount
                  ) &&

                  wdDate <=
                    returnDate

                );

              }
            )
            .sort(
              (a, b) =>

                flowMillis_(
                  b[
                    h['Transaction Date']
                  ]
                ) -

                flowMillis_(
                  a[
                    h['Transaction Date']
                  ]
                )
            );


        candidate =
          possible[0] ||
          null;

      }


      if (!candidate) {
        return;
      }


      const withdrawalKey =
        flowCell_(
          candidate,
          h,
          'Transaction Key'
        );


      const returnKey =
        flowCell_(
          returnRow,
          h,
          'Transaction Key'
        );


      candidate[
        h['Reconcile Status']
      ] =
        'RETURNED';


      candidate[
        h['Matched Transaction Key']
      ] =
        returnKey;


      returnRow[
        h['Reconcile Status']
      ] =
        'MATCHED';


      returnRow[
        h['Matched Transaction Key']
      ] =
        withdrawalKey;


      matched++;

    }
  );


  return matched;

}


/************************************************************
 * PAYPAL -> GREY
 ************************************************************/

function flowReconcilePaypalGrey_(
  rows,
  h
) {

  const paypalWithdrawals =
    rows.filter(
      row =>

        flowCell_(
          row,
          h,
          'Provider'
        )
          .toUpperCase() ===
          'PAYPAL' &&

        flowCell_(
          row,
          h,
          'Type'
        )
          .toUpperCase() ===
          'PAYPAL_WITHDRAWAL'
    );


  const greyFundings =
    rows.filter(
      row =>

        flowCell_(
          row,
          h,
          'Provider'
        )
          .toUpperCase() ===
          'GREY' &&

        flowCell_(
          row,
          h,
          'Type'
        )
          .toUpperCase() ===
          'GREY_FUNDING'
    );


  paypalWithdrawals.sort(
    (a, b) =>

      flowMillis_(
        a[
          h['Transaction Date']
        ]
      ) -

      flowMillis_(
        b[
          h['Transaction Date']
        ]
      )
  );


  greyFundings.sort(
    (a, b) =>

      flowMillis_(
        a[
          h['Transaction Date']
        ]
      ) -

      flowMillis_(
        b[
          h['Transaction Date']
        ]
      )
  );


  const usedGrey = {};

  let matched = 0;


  paypalWithdrawals.forEach(
    paypalRow => {

      if (
        flowCell_(
          paypalRow,
          h,
          'Reconcile Status'
        )
          .toUpperCase() ===
          'RETURNED'
      ) {
        return;
      }


      const destinationProvider =
        flowCell_(
          paypalRow,
          h,
          'Counterparty Provider'
        )
          .toUpperCase();


      const destinationAccount =
        flowCell_(
          paypalRow,
          h,
          'Counterparty Account Key'
        );


      if (
        destinationProvider !==
          'GREY' ||
        !destinationAccount
      ) {
        return;
      }


      const paypalAmount =
        flowNumber_(
          paypalRow[
            h['Gross Amount']
          ]
        );


      const paypalCurrency =
        flowCell_(
          paypalRow,
          h,
          'Currency'
        )
          .toUpperCase();


      const paypalDate =
        flowMillis_(
          paypalRow[
            h['Transaction Date']
          ]
        );


      const maximumDate =
        paypalDate +
        (
          FLOW_CFG.MATCH_MAX_DAYS *
          24 *
          60 *
          60 *
          1000
        );


      const candidate =
        greyFundings.find(
          greyRow => {

            const greyKey =
              flowCell_(
                greyRow,
                h,
                'Transaction Key'
              );


            if (
              usedGrey[
                greyKey
              ]
            ) {
              return false;
            }


            const greyDate =
              flowMillis_(
                greyRow[
                  h['Transaction Date']
                ]
              );


            return (

              flowCell_(
                greyRow,
                h,
                'Account Key'
              ) ===
                destinationAccount &&

              flowCell_(
                greyRow,
                h,
                'Currency'
              )
                .toUpperCase() ===
                paypalCurrency &&

              flowAmountEqual_(
                flowNumber_(
                  greyRow[
                    h['Gross Amount']
                  ]
                ),
                paypalAmount
              ) &&

              greyDate >=
                paypalDate &&

              greyDate <=
                maximumDate

            );

          }
        );


      if (!candidate) {
        return;
      }


      const paypalKey =
        flowCell_(
          paypalRow,
          h,
          'Transaction Key'
        );


      const greyKey =
        flowCell_(
          candidate,
          h,
          'Transaction Key'
        );


      paypalRow[
        h['Reconcile Status']
      ] =
        'ARRIVED';


      paypalRow[
        h['Matched Transaction Key']
      ] =
        greyKey;


      candidate[
        h['Reconcile Status']
      ] =
        'MATCHED';


      candidate[
        h['Matched Transaction Key']
      ] =
        paypalKey;


      candidate[
        h['Counterparty Provider']
      ] =
        'PAYPAL';


      candidate[
        h['Counterparty Account Key']
      ] =
        flowCell_(
          paypalRow,
          h,
          'Account Key'
        );


      usedGrey[
        greyKey
      ] =
        true;


      matched++;

    }
  );


  return matched;

}


/************************************************************
 * GREY -> USDT
 ************************************************************/

function flowSyncGreyToUsdt_(
  state,
  h,
  usdtMap
) {

  const greyWithdrawals =
    state.rows.filter(
      row =>

        flowCell_(
          row,
          h,
          'Provider'
        )
          .toUpperCase() ===
          'GREY' &&

        flowCell_(
          row,
          h,
          'Type'
        )
          .toUpperCase() ===
          'GREY_WITHDRAWAL' &&

        flowCell_(
          row,
          h,
          'Currency'
        )
          .toUpperCase() ===
          'USDT' &&

        flowCell_(
          row,
          h,
          'Counterparty Provider'
        )
          .toUpperCase() ===
          'USDT'
    );


  let count = 0;


  greyWithdrawals.forEach(
    greyRow => {

      const destination =
        flowCell_(
          greyRow,
          h,
          'Counterparty Account Key'
        );


      if (
        !destination ||
        !usdtMap[
          destination
        ]
      ) {
        return;
      }


      const greyKey =
        flowCell_(
          greyRow,
          h,
          'Transaction Key'
        );


      const reference =
        flowCell_(
          greyRow,
          h,
          'Reference'
        );


      const received =
        flowNumber_(
          greyRow[
            h['Net Amount']
          ]
        );


      const usdtKey =
        'UTX|' +
        destination +
        '|' +
        (
          reference ||
          greyKey
        );


      flowUpsertTransaction_(
        state,
        h,
        {

          'Transaction Key':
            usdtKey,

          'Transaction Date':
            greyRow[
              h['Transaction Date']
            ],

          'Provider':
            'USDT',

          'Account Key':
            destination,

          'Label':
            usdtMap[
              destination
            ].tag ||
            '',

          'Type':
            'USDT_RECEIVED',

          'Direction':
            'IN',

          'Gross Amount':
            flowRoundMoney_(
              received
            ),

          'Net Amount':
            flowRoundMoney_(
              received
            ),

          'Fee Amount':
            0,

          'Currency':
            'USDT',

          'Reference':
            reference,

          'Counterparty Provider':
            'GREY',

          'Counterparty Account Key':
            flowCell_(
              greyRow,
              h,
              'Account Key'
            ),

          'Reconcile Status':
            'MATCHED',

          'Matched Transaction Key':
            greyKey,

          'Source':
            'FLOW_DERIVED',

          'Source ID':
            greyKey,

          'Sync Time':
            new Date(),

          'Unique Key':
            'USDT|' +
            destination +
            '|GREY|' +
            greyKey

        }
      );


      greyRow[
        h['Reconcile Status']
      ] =
        'ARRIVED';


      greyRow[
        h['Matched Transaction Key']
      ] =
        usdtKey;


      count++;

    }
  );


  return count;

}


/************************************************************
 * FINAL PAYPAL STATUS
 ************************************************************/

function flowFinalizePaypalStatus_(
  rows,
  h
) {

  const now =
    Date.now();


  const overdueMs =
    FLOW_CFG.OVERDUE_DAYS *
    24 *
    60 *
    60 *
    1000;


  rows.forEach(
    row => {

      if (
        flowCell_(
          row,
          h,
          'Provider'
        )
          .toUpperCase() !==
          'PAYPAL' ||

        flowCell_(
          row,
          h,
          'Type'
        )
          .toUpperCase() !==
          'PAYPAL_WITHDRAWAL'
      ) {
        return;
      }


      const current =
        flowCell_(
          row,
          h,
          'Reconcile Status'
        )
          .toUpperCase();


      if (
        current ===
          'ARRIVED' ||
        current ===
          'RETURNED'
      ) {
        return;
      }


      const destinationProvider =
        flowCell_(
          row,
          h,
          'Counterparty Provider'
        );


      const destinationAccount =
        flowCell_(
          row,
          h,
          'Counterparty Account Key'
        );


      if (
        !destinationProvider ||
        !destinationAccount
      ) {

        row[
          h['Reconcile Status']
        ] =
          'UNMATCHED';

        return;

      }


      /*
       * Đã khai đích đến nhưng reconcile không biết provider đó. Không được để
       * rơi vào IN_TRANSIT/OVERDUE: hai trạng thái đó nói rằng tiền đang trên
       * đường về và sẽ tự khớp, trong khi thực tế không có ai đi đón. Nó sẽ
       * nằm OVERDUE vĩnh viễn mà không ai biết tại sao.
       */
      if (
        FLOW_CFG.RECEIVING_PROVIDERS.indexOf(
          destinationProvider.toUpperCase()
        ) === -1
      ) {

        row[
          h['Reconcile Status']
        ] =
          'UNSUPPORTED_PROVIDER';

        return;

      }


      const date =
        flowMillis_(
          row[
            h['Transaction Date']
          ]
        );


      row[
        h['Reconcile Status']
      ] =
        (
          date &&
          now - date >
            overdueMs
        )
          ? 'OVERDUE'
          : 'IN_TRANSIT';

    }
  );

}


/************************************************************
 * WHETHER TX COUNTS IN BALANCE
 ************************************************************/

/************************************************************
 * SETTLEMENT STATUS (READ-ONLY)
 *
 * Flow reads the Settlements sheet itself rather than calling into
 * Customer.gs. A guard must not depend on the runtime behaviour of the very
 * layer it is guarding: the Customer sweeper only cancels DRAFT payouts, so it
 * cannot be the last line of defence for real treasury movement.
 *
 * Returns null when there is no usable Settlements sheet, which tells the
 * caller to skip the guard entirely and behave exactly as before.
 ************************************************************/

function flowReadSettlementStatusMap_(
  ss
) {

  const sheet =
    flowOptionalSheet_(
      ss,
      FLOW_CFG.SETTLEMENTS_SHEET
    );


  if (!sheet) {
    return null;
  }


  const h =
    flowHeaderMap_(
      sheet
    );


  if (
    !h['Settlement Key'] ||
    !h['Status']
  ) {
    return null;
  }


  const map = {};


  flowReadRows_(
    sheet
  ).forEach(
    row => {

      const key =
        flowCell_(
          row,
          h,
          'Settlement Key'
        );


      if (!key) {
        return;
      }


      map[key] =
        flowCell_(
          row,
          h,
          'Status'
        )
          .toUpperCase();

    }
  );


  return map;

}


/*
 * Settlement status is owned by Customer.gs. Defer to its helper so the list of
 * real-invoice statuses lives in exactly one place, and fall back to the same
 * three only if that file is somehow unavailable.
 */
function flowIsEffectiveSettlementStatus_(
  status
) {

  if (
    typeof custIsEffectiveSettlementStatus_ ===
    'function'
  ) {

    return custIsEffectiveSettlementStatus_(
      status
    );

  }


  const s =
    String(status || '')
      .trim()
      .toUpperCase();


  return (
    s === 'ISSUED' ||
    s === 'PARTIALLY_PAID' ||
    s === 'PAID'
  );

}


function flowTransactionCounts_(
  row,
  h
) {

  const source =
    flowCell_(
      row,
      h,
      'Source'
    )
      .toUpperCase();


  /*
   * Manual In: chỉ đếm khi còn hiệu lực. Dòng bị huỷ vẫn nằm lại trong
   * Transaction để giữ dấu vết, nên phải loại ở đây — không thì một khoản đã
   * huỷ vẫn cộng vào số dư mãi mãi.
   */
  if (
    source ===
    'MANUAL IN'
  ) {

    const state =
      flowCell_(
        row,
        h,
        'Reconcile Status'
      )
        .toUpperCase();


    return (
      state !== 'CANCELLED' &&
      state !== 'INVALID'
    );

  }


  /*
   * Non-Payout transactions work normally.
   */
  if (
    source !==
    'PAYOUTS'
  ) {
    return true;
  }


  /*
   * Payout only counts when PAID.
   */
  return (
    flowCell_(
      row,
      h,
      'Reconcile Status'
    )
      .toUpperCase() ===
      'PAID'
  );

}


/************************************************************
 * GREY SUMMARY
 ************************************************************/

function flowUpdateGreySummary_(
  sheet,
  gh,
  txRows,
  th
) {

  const rows =
    flowReadRows_(sheet);


  rows.forEach(
    row => {

      const accountKey =
        flowString_(
          row[
            gh['Account Key']
          ]
        );


      if (!accountKey) {
        return;
      }


      let totalReceivedUsd = 0;
      let totalWithdrawnUsd = 0;
      let totalP2pUsd = 0;
      let totalUsdtOut = 0;

      let totalFeesUsd = 0;
      let totalFeesUsdt = 0;

      let incomingUsd = 0;

      let lastActivity = null;


      txRows.forEach(
        tx => {

          const provider =
            flowCell_(
              tx,
              th,
              'Provider'
            )
              .toUpperCase();


          const txAccount =
            flowCell_(
              tx,
              th,
              'Account Key'
            );


          const type =
            flowCell_(
              tx,
              th,
              'Type'
            )
              .toUpperCase();


          const direction =
            flowCell_(
              tx,
              th,
              'Direction'
            )
              .toUpperCase();


          const currency =
            flowCell_(
              tx,
              th,
              'Currency'
            )
              .toUpperCase();


          const gross =
            flowNumber_(
              tx[
                th['Gross Amount']
              ]
            );


          const net =
            flowNumber_(
              tx[
                th['Net Amount']
              ]
            );


          const fee =
            flowNumber_(
              tx[
                th['Fee Amount']
              ]
            );


          /****************************************************
           * GREY ACCOUNT ITSELF
           ****************************************************/

          if (
            provider ===
              'GREY' &&
            txAccount ===
              accountKey
          ) {

            if (
              type ===
                'GREY_FUNDING' &&
              currency ===
                'USD'
            ) {

              totalReceivedUsd +=
                net;

            }


            if (
              direction ===
                'OUT' &&
              currency ===
                'USD'
            ) {

              totalWithdrawnUsd +=
                gross;

            }


            if (
              type ===
                'GREY_P2P' &&
              currency ===
                'USD'
            ) {

              totalP2pUsd +=
                gross;

            }


            if (
              direction ===
                'OUT' &&
              currency ===
                'USDT'
            ) {

              totalUsdtOut +=
                gross;

            }


            if (
              currency ===
              'USD'
            ) {

              totalFeesUsd +=
                fee;

            }


            if (
              currency ===
              'USDT'
            ) {

              totalFeesUsdt +=
                fee;

            }


            const activity =
              flowMillis_(
                tx[
                  th['Transaction Date']
                ]
              );


            if (
              activity &&
              (
                !lastActivity ||
                activity >
                  lastActivity.getTime()
              )
            ) {

              lastActivity =
                new Date(
                  activity
                );

            }

          }


          /****************************************************
           * PAYPAL IN TRANSIT TO GREY
           ****************************************************/

          if (
            provider ===
              'PAYPAL' &&

            type ===
              'PAYPAL_WITHDRAWAL' &&

            currency ===
              'USD' &&

            flowCell_(
              tx,
              th,
              'Counterparty Account Key'
            ) ===
              accountKey &&

            flowCell_(
              tx,
              th,
              'Reconcile Status'
            )
              .toUpperCase() ===
              'IN_TRANSIT'
          ) {

            incomingUsd +=
              gross;

          }

        }
      );


      const estimatedUsdConverted =
        totalUsdtOut /
        FLOW_CFG.USD_TO_USDT_RATE;


      const estimatedUsdBalance =
        totalReceivedUsd -
        totalWithdrawnUsd -
        estimatedUsdConverted;


      row[
        gh['Estimated USD Balance']
      ] =
        flowRoundMoney_(
          estimatedUsdBalance
        );


      row[
        gh['Incoming USD (Estimated)']
      ] =
        flowRoundMoney_(
          incomingUsd
        );


      row[
        gh['Total Received USD']
      ] =
        flowRoundMoney_(
          totalReceivedUsd
        );


      row[
        gh['Total Withdrawn USD']
      ] =
        flowRoundMoney_(
          totalWithdrawnUsd
        );


      row[
        gh['Total P2P Sent USD']
      ] =
        flowRoundMoney_(
          totalP2pUsd
        );


      row[
        gh['Total USDT Out']
      ] =
        flowRoundMoney_(
          totalUsdtOut
        );


      row[
        gh['Total Fees USD']
      ] =
        flowRoundMoney_(
          totalFeesUsd
        );


      row[
        gh['Total Fees USDT']
      ] =
        flowRoundMoney_(
          totalFeesUsdt
        );


      row[
        gh['Last Activity']
      ] =
        lastActivity ||
        '';

    }
  );


  flowWriteRows_(
    sheet,
    rows
  );

}


/************************************************************
 * USDT SUMMARY
 *
 * Includes:
 *
 * Grey -> USDT RECEIVED
 * Payouts -> USDT OUT
 ************************************************************/

function flowUpdateUsdtSummary_(
  sheet,
  uh,
  txRows,
  th
) {

  const rows =
    flowReadRows_(sheet);

  /*
   * Tổng toàn hệ thống, cộng qua mọi ví.
   *
   * Con số theo từng ví trả lời "đường ống đã giao về ví nào", và nó sai ngay
   * khi tiền bị xáo giữa các ví — việc hệ thống không có cách nào nhìn thấy.
   * Tổng thì miễn nhiễm, vì nó không nhìn vào ví nào cả.
   *
   * Và nó KHÔNG phải số dư ví thật. Nó là "đường ống PayPal giao về bao nhiêu,
   * đã chi ra bao nhiêu, còn lại bao nhiêu" — cả hai vế đều suy từ nguồn có
   * thật, nên đúng tuyệt đối miễn là mọi khoản chi đều có dòng Payouts. Muốn
   * biết ví thật có bao nhiêu thì mở ví; hệ thống không dùng số đó để quyết
   * định bất cứ điều gì.
   */
  let poolPipeline = 0;
  let poolManual = 0;
  let poolWithdrawn = 0;


  rows.forEach(
    row => {

      const accountKey =
        flowString_(
          row[
            uh['Account Key']
          ]
        );


      if (!accountKey) {
        return;
      }


      /*
       * Tách nguồn tiền vào: MANUAL_IN là khoản gõ tay, còn lại là tiền chảy
       * qua đường ống PayPal → Grey. Estimated Balance vẫn là tổng của cả hai;
       * việc tách chỉ để dòng tổng nói được tiền từ đâu tới.
       */
      let pipelineReceived = 0;
      let manualReceived = 0;
      let totalWithdrawn = 0;

      let lastActivity = null;


      txRows.forEach(
        tx => {

          /*
           * Cancelled / Draft payouts ignored.
           */
          if (
            !flowTransactionCounts_(
              tx,
              th
            )
          ) {
            return;
          }


          if (
            flowCell_(
              tx,
              th,
              'Provider'
            )
              .toUpperCase() !==
              'USDT' ||

            flowCell_(
              tx,
              th,
              'Account Key'
            ) !==
              accountKey ||

            flowCell_(
              tx,
              th,
              'Currency'
            )
              .toUpperCase() !==
              'USDT'
          ) {
            return;
          }


          const direction =
            flowCell_(
              tx,
              th,
              'Direction'
            )
              .toUpperCase();


          if (
            direction ===
            'IN'
          ) {

            const amount =
              flowNumber_(
                tx[
                  th['Net Amount']
                ]
              );


            if (
              flowCell_(
                tx,
                th,
                'Type'
              )
                .toUpperCase() ===
                'MANUAL_IN'
            ) {
              manualReceived += amount;
            } else {
              pipelineReceived += amount;
            }

          }


          if (
            direction ===
            'OUT'
          ) {

            /*
             * Gross = actual wallet deduction.
             */
            totalWithdrawn +=
              flowNumber_(
                tx[
                  th['Gross Amount']
                ]
              );

          }


          const activity =
            flowMillis_(
              tx[
                th['Transaction Date']
              ]
            );


          if (
            activity &&
            (
              !lastActivity ||
              activity >
                lastActivity.getTime()
            )
          ) {

            lastActivity =
              new Date(
                activity
              );

          }

        }
      );


      /* Cột này vẫn là TỔNG tiền vào, không phân biệt nguồn. */
      const totalReceived =
        pipelineReceived +
        manualReceived;


      row[
        uh['Total Received USDT']
      ] =
        flowRoundMoney_(
          totalReceived
        );


      row[
        uh['Total Withdrawn USDT']
      ] =
        flowRoundMoney_(
          totalWithdrawn
        );


      row[
        uh['Estimated Balance']
      ] =
        flowRoundMoney_(
          totalReceived -
          totalWithdrawn
        );


      row[
        uh['Last Activity']
      ] =
        lastActivity ||
        '';


      poolPipeline += pipelineReceived;
      poolManual += manualReceived;
      poolWithdrawn += totalWithdrawn;

    }
  );


  flowWriteRows_(
    sheet,
    rows
  );


  flowWriteUsdtPoolTotal_(
    sheet,
    poolPipeline,
    poolManual,
    poolWithdrawn
  );

}


/*
 * Hàng 1 là hàng ghi chú theo quy ước của workbook: header ở hàng 2, dữ liệu từ
 * hàng 3. flowWriteRows_ chỉ đụng từ DATA_START_ROW trở xuống, nên hàng 1 là
 * chỗ duy nhất ghi được mà không bao giờ bị vòng ghi dữ liệu xoá mất.
 */
function flowWriteUsdtPoolTotal_(
  sheet,
  pipeline,
  manual,
  withdrawn
) {

  const remaining =
    flowRoundMoney_(
      pipeline +
      manual -
      withdrawn
    );


  /*
   * Tách "từ PayPal" và "nhận thêm" thay vì gộp một cục: không tốn gì mà giữ
   * được con số cũ, nên vẫn biết bao nhiêu là tiền chảy qua đường ống khách
   * hàng, bao nhiêu đến từ chỗ khác.
   *
   * Dòng "nhận thêm" chỉ hiện khi có, để sheet không bị rác khi chưa dùng tới.
   */
  sheet
    .getRange(1, 1)
    .setValue(
      'TỔNG USDT TRONG HỆ THỐNG · từ PayPal ' +
      flowRoundMoney_(pipeline).toFixed(2) +
      (
        Math.abs(manual) > 0.0001
          ? ' · nhận thêm ' +
            flowRoundMoney_(manual).toFixed(2)
          : ''
      ) +
      ' · đã chi ' +
      flowRoundMoney_(withdrawn).toFixed(2) +
      ' · CÒN LẠI ' +
      remaining.toFixed(2) +
      '   (chỉ tính phần hệ thống ghi nhận, không phải số dư ví thật)'
    );

}


/************************************************************
 * ACCOUNT MAP
 ************************************************************/

/*
 * Cột Enabled không nằm trong hợp đồng header của Flow (chỉ Paypal.gs đọc nó),
 * nên phải kiểm tra sự tồn tại trước: thiếu cột thì coi như account đang bật,
 * đừng nuốt mất cảnh báo chỉ vì workbook chưa có cột đó.
 */
function flowWarnUnsupportedLink_(
  key,
  row,
  h
) {

  const provider =
    flowString_(
      row[
        h['Provider']
      ]
    )
      .toUpperCase();


  if (provider !== 'PAYPAL') {
    return;
  }


  if (
    Object.prototype.hasOwnProperty.call(
      h,
      'Enabled'
    )
  ) {

    const enabled =
      flowString_(
        row[
          h['Enabled']
        ]
      )
        .toUpperCase();


    if (
      enabled &&
      enabled !== 'TRUE' &&
      enabled !== 'YES' &&
      enabled !== '1' &&
      enabled !== 'ACTIVE'
    ) {
      return;
    }

  }


  const linked =
    flowString_(
      row[
        h['Linked Provider']
      ]
    )
      .toUpperCase();


  if (
    !linked ||
    FLOW_CFG.RECEIVING_PROVIDERS.indexOf(linked) !== -1
  ) {
    return;
  }


  Logger.log(
    'CẢNH BÁO: PayPal account "' + key +
    '" khai Linked Provider = "' + linked +
    '" mà reconcile chưa hỗ trợ. Withdrawal của nó sẽ mang trạng thái ' +
    'UNSUPPORTED_PROVIDER và tiền nằm lại Pending in PayPal cho tới khi ' +
    'có collector cho provider đó.'
  );

}


function flowReadAccountMap_(
  sheet,
  h
) {

  const result = {};


  flowReadRows_(sheet)
    .forEach(
      row => {

        const key =
          flowString_(
            row[
              h['Account Key']
            ]
          );


        if (!key) {
          return;
        }


        result[key] = {

          label:
            flowString_(
              row[
                h['Label']
              ]
            ),

          linkedProvider:
            flowString_(
              row[
                h['Linked Provider']
              ]
            ),

          linkedAccountKey:
            flowString_(
              row[
                h['Linked Account Key']
              ]
            )

        };


        /*
         * Cảnh báo ngay khi vừa cấu hình sai, đừng đợi tới lần rút đầu tiên
         * rồi mới phát hiện tiền kẹt. Reconcile chỉ biết đường tới các
         * provider trong FLOW_CFG.RECEIVING_PROVIDERS.
         *
         * Chỉ soi account PayPal đang bật: Linked Provider của dòng Grey hay
         * của account đã tắt không nói lên điều gì về đường rút đang chạy, báo
         * chúng chỉ tạo nhiễu rồi người ta thôi đọc log.
         */
        flowWarnUnsupportedLink_(
          key,
          row,
          h
        );

      }
    );

  return result;

}


/************************************************************
 * GREY MAP
 ************************************************************/

function flowReadGreyMap_(
  sheet,
  h
) {

  const result = {};


  flowReadRows_(sheet)
    .forEach(
      row => {

        const key =
          flowString_(
            row[
              h['Account Key']
            ]
          );


        if (!key) {
          return;
        }


        result[key] = {

          label:
            flowString_(
              row[
                h['Label']
              ]
            ),

          defaultUsdtAccount:
            flowString_(
              row[
                h['Default USDT Account']
              ]
            )

        };

      }
    );


  return result;

}


/************************************************************
 * USDT MAP
 ************************************************************/

function flowReadUsdtMap_(
  sheet,
  h
) {

  const result = {};


  flowReadRows_(sheet)
    .forEach(
      row => {

        const key =
          flowString_(
            row[
              h['Account Key']
            ]
          );


        if (!key) {
          return;
        }


        result[key] = {

          tag:
            flowString_(
              row[
                h['Tag']
              ]
            ),

          status:
            flowString_(
              row[
                h['Status']
              ]
            )
              .toUpperCase()

        };

      }
    );


  return result;

}


/************************************************************
 * TRANSACTION STATE
 ************************************************************/

function flowLoadTransactionState_(
  sheet,
  h
) {

  const lastRow =
    sheet.getLastRow();


  const lastColumn =
    sheet.getLastColumn();


  let rows = [];


  if (
    lastRow >=
      FLOW_CFG.DATA_START_ROW &&
    lastColumn > 0
  ) {

    rows =
      sheet
        .getRange(
          FLOW_CFG.DATA_START_ROW,
          1,
          lastRow -
            FLOW_CFG.DATA_START_ROW +
            1,
          lastColumn
        )
        .getValues()
        .filter(
          row =>
            row.some(
              value =>
                value !== '' &&
                value !== null
            )
        );

  }


  const keyMap = {};


  rows.forEach(
    (row, index) => {

      const key =
        flowString_(
          row[
            h['Unique Key']
          ]
        );


      if (key) {

        keyMap[key] =
          index;

      }

    }
  );


  return {

    rows:
      rows,

    keyMap:
      keyMap,

    columnCount:
      lastColumn,

    oldRowCount:
      Math.max(
        0,
        lastRow -
          FLOW_CFG.DATA_START_ROW +
          1
      )

  };

}


/************************************************************
 * FIND TX BY UNIQUE KEY
 ************************************************************/

function flowFindTransactionByUnique_(
  state,
  uniqueKey
) {

  if (
    !Object.prototype.hasOwnProperty.call(
      state.keyMap,
      uniqueKey
    )
  ) {
    return null;
  }


  return state.rows[
    state.keyMap[
      uniqueKey
    ]
  ];

}


/************************************************************
 * UPSERT TRANSACTION
 ************************************************************/

function flowUpsertTransaction_(
  state,
  h,
  record
) {

  const uniqueKey =
    record[
      'Unique Key'
    ];


  if (!uniqueKey) {

    throw new Error(
      'Transaction thiếu Unique Key.'
    );

  }


  let row;


  if (
    Object.prototype.hasOwnProperty.call(
      state.keyMap,
      uniqueKey
    )
  ) {

    row =
      state.rows[
        state.keyMap[
          uniqueKey
        ]
      ];


  } else {

    row =
      new Array(
        state.columnCount
      )
        .fill('');


    state.rows.push(row);


    state.keyMap[
      uniqueKey
    ] =
      state.rows.length -
      1;

  }


  Object.keys(record)
    .forEach(
      header => {

        if (
          !Object.prototype.hasOwnProperty.call(
            h,
            header
          )
        ) {
          return;
        }


        row[
          h[header]
        ] =
          record[
            header
          ];

      }
    );


  return row;

}


/************************************************************
 * WRITE TRANSACTION — NEWEST FIRST
 ************************************************************/

function flowWriteTransactionState_(
  sheet,
  h,
  state
) {

  const dateIndex =
    h[
      'Transaction Date'
    ];


  state.rows.sort(
    (a, b) =>

      flowMillis_(
        b[
          dateIndex
        ]
      ) -

      flowMillis_(
        a[
          dateIndex
        ]
      )
  );


  const clearRows =
    Math.max(
      state.oldRowCount,
      state.rows.length
    );


  if (
    clearRows > 0
  ) {

    sheet
      .getRange(
        FLOW_CFG.DATA_START_ROW,
        1,
        clearRows,
        state.columnCount
      )
      .clearContent();

  }


  if (
    state.rows.length > 0
  ) {

    sheet
      .getRange(
        FLOW_CFG.DATA_START_ROW,
        1,
        state.rows.length,
        state.columnCount
      )
      .setValues(
        state.rows
      );

  }

}


/************************************************************
 * READ DATA ROWS
 ************************************************************/

function flowReadRows_(
  sheet
) {

  if (!sheet) {
    return [];
  }


  const lastRow =
    sheet.getLastRow();


  const lastColumn =
    sheet.getLastColumn();


  if (
    lastRow <
      FLOW_CFG.DATA_START_ROW ||
    lastColumn < 1
  ) {

    return [];

  }


  return sheet
    .getRange(
      FLOW_CFG.DATA_START_ROW,
      1,
      lastRow -
        FLOW_CFG.DATA_START_ROW +
        1,
      lastColumn
    )
    .getValues()
    .filter(
      row =>
        row.some(
          value =>
            value !== '' &&
            value !== null
        )
    );

}


/************************************************************
 * PAYOUTS READ — PRESERVE POSITIONS
 ************************************************************/

function flowReadRowsPreserve_(
  sheet
) {

  const lastRow =
    sheet.getLastRow();


  const lastColumn =
    sheet.getLastColumn();


  if (
    lastRow <
      FLOW_CFG.DATA_START_ROW ||
    lastColumn < 1
  ) {

    return [];

  }


  return sheet
    .getRange(
      FLOW_CFG.DATA_START_ROW,
      1,
      lastRow -
        FLOW_CFG.DATA_START_ROW +
        1,
      lastColumn
    )
    .getValues();

}


/************************************************************
 * PAYOUTS WRITE — PRESERVE POSITIONS
 ************************************************************/

function flowWriteRowsPreserve_(
  sheet,
  rows
) {

  if (!rows.length) {
    return;
  }


  sheet
    .getRange(
      FLOW_CFG.DATA_START_ROW,
      1,
      rows.length,
      sheet.getLastColumn()
    )
    .setValues(rows);

}


/************************************************************
 * GENERAL WRITE
 ************************************************************/

function flowWriteRows_(
  sheet,
  rows
) {

  const columns =
    sheet.getLastColumn();


  if (
    columns < 1
  ) {
    return;
  }


  const currentRows =
    Math.max(
      0,
      sheet.getLastRow() -
        FLOW_CFG.DATA_START_ROW +
        1
    );


  const clearRows =
    Math.max(
      currentRows,
      rows.length
    );


  if (
    clearRows > 0
  ) {

    sheet
      .getRange(
        FLOW_CFG.DATA_START_ROW,
        1,
        clearRows,
        columns
      )
      .clearContent();

  }


  if (
    rows.length > 0
  ) {

    sheet
      .getRange(
        FLOW_CFG.DATA_START_ROW,
        1,
        rows.length,
        columns
      )
      .setValues(
        rows
      );

  }

}


/************************************************************
 * VALIDATE HEADERS
 ************************************************************/

function flowValidateHeaders_(
  ah,
  ph,
  gh,
  uh,
  poh,
  th
) {

  flowRequire_(
    ah,
    [
      'Account Key',
      'Label',
      'Provider',

      'Linked Provider',
      'Linked Account Key',

      'Total Balance',
      'Available Balance',
      'Withheld Balance',

      'Total Withdrawn',
      'Hold Balance',

      'Currency',

      'PayPal Status'
    ],
    'Account'
  );


  flowRequire_(
    ph,
    [
      'Account Key',
      'Label',

      'Initiation Date',
      'Updated Date',

      'Transaction ID',
      'Reference ID',

      'Event Code',

      'Type',
      'Status',

      'Invoice ID',
      'Invoice Prefix',

      'Buyer Name',
      'Buyer Email',

      'Item',

      'Transaction Amount',
      'Fee Amount',
      'Balance Impact',

      'Currency',

      'Unique Key'
    ],
    'Paypal'
  );


  flowRequire_(
    gh,
    [
      'Account Key',
      'Label',
      'Status',

      'Default USDT Account',

      'Estimated USD Balance',
      'Incoming USD (Estimated)',

      'Total Received USD',
      'Total Withdrawn USD',
      'Total P2P Sent USD',
      'Total USDT Out',

      'Total Fees USD',
      'Total Fees USDT',

      'Last Activity'
    ],
    'Grey'
  );


  flowRequire_(
    uh,
    [
      'Account Key',
      'Tag',
      'Status',

      'Estimated Balance',

      'Total Received USDT',
      'Total Withdrawn USDT',

      'Last Activity'
    ],
    'USDT'
  );


  flowRequire_(
    poh,
    [
      'Payout Key',
      'Payout Date',

      'Provider',
      'Account Key',

      'Payee',
      'Category',

      'Amount',
      'Fee',
      'Total Deducted',

      'Currency',

      'Reference',

      'Status',

      'Validation',

      'Transaction Key',

      'Created At',
      'Paid At',

      'Note'
    ],
    'Payouts'
  );


  flowRequire_(
    th,
    [
      'Transaction Key',
      'Transaction Date',

      'Provider',
      'Account Key',
      'Label',

      'Type',
      'Direction',

      'Gross Amount',
      'Net Amount',
      'Fee Amount',

      'Currency',
      'Reference',

      'Counterparty Provider',
      'Counterparty Account Key',

      'Reconcile Status',
      'Matched Transaction Key',

      'Source',
      'Source ID',

      'Sync Time',
      'Unique Key',

      'Raw Event',
      'Raw Transaction Type',

      'Raw Amount Tendered',
      'Raw Amount Received',

      'Raw Recipient',
      'Raw Channel',
      'Raw Date Time'
    ],
    'Transaction'
  );

}


/************************************************************
 * REQUIRE HEADERS
 ************************************************************/

function flowRequire_(
  map,
  required,
  sheetName
) {

  const missing =
    required.filter(
      header =>
        !Object.prototype.hasOwnProperty.call(
          map,
          header
        )
    );


  if (
    missing.length
  ) {

    throw new Error(
      'Sheet "' +
      sheetName +
      '" thiếu header: ' +
      missing.join(', ')
    );

  }

}


/************************************************************
 * GET SPREADSHEET
 ************************************************************/

function flowGetSpreadsheet_() {

  if (
    FLOW_CFG.SPREADSHEET_ID
  ) {

    return SpreadsheetApp.openById(
      FLOW_CFG.SPREADSHEET_ID
    );

  }


  const ss =
    SpreadsheetApp.getActiveSpreadsheet();


  if (!ss) {

    throw new Error(
      'Không có active spreadsheet.'
    );

  }


  return ss;

}


/************************************************************
 * REQUIRED SHEET
 ************************************************************/

function flowNeedSheet_(
  ss,
  name
) {

  const sheet =
    ss.getSheetByName(
      name
    );


  if (!sheet) {

    throw new Error(
      'Không tìm thấy sheet "' +
      name +
      '".'
    );

  }


  return sheet;

}


/************************************************************
 * OPTIONAL SHEET
 ************************************************************/

function flowOptionalSheet_(
  ss,
  name
) {

  return ss.getSheetByName(
    name
  );

}


/************************************************************
 * HEADER MAP
 *
 * Empty sheet -> {}
 ************************************************************/

function flowHeaderMap_(
  sheet
) {

  const map = {};


  if (!sheet) {
    return map;
  }


  const lastColumn =
    sheet.getLastColumn();


  const lastRow =
    sheet.getLastRow();


  if (
    lastColumn < 1 ||
    lastRow <
      FLOW_CFG.HEADER_ROW
  ) {

    return map;

  }


  const values =
    sheet
      .getRange(
        FLOW_CFG.HEADER_ROW,
        1,
        1,
        lastColumn
      )
      .getValues()[0];


  values.forEach(
    (value, index) => {

      const header =
        flowString_(value);


      if (!header) {
        return;
      }


      if (
        Object.prototype.hasOwnProperty.call(
          map,
          header
        )
      ) {

        throw new Error(
          'Duplicate header "' +
          header +
          '" trong sheet "' +
          sheet.getName() +
          '".'
        );

      }


      map[
        header
      ] =
        index;

    }
  );


  return map;

}


/************************************************************
 * CELL
 ************************************************************/

function flowCell_(
  row,
  h,
  header
) {

  if (
    !Object.prototype.hasOwnProperty.call(
      h,
      header
    )
  ) {

    return '';

  }


  return flowString_(
    row[
      h[header]
    ]
  );

}


/************************************************************
 * EDIT RANGE TOUCHES COLUMN?
 ************************************************************/

function flowRangeTouchesColumn_(
  range,
  column
) {

  const first =
    range.getColumn();


  const last =
    first +
    range.getNumColumns() -
    1;


  return (
    column >= first &&
    column <= last
  );

}


/************************************************************
 * STRING
 ************************************************************/

function flowString_(
  value
) {

  return String(
    value === null ||
    value === undefined
      ? ''
      : value
  )
    .trim();

}


/************************************************************
 * NUMBER
 ************************************************************/

function flowNumber_(
  value
) {

  const parsed =
    flowNullableNumber_(
      value
    );


  return parsed ===
    null
      ? 0
      : parsed;

}


/************************************************************
 * NULLABLE NUMBER
 ************************************************************/

/*
 * Không dùng instanceof: nó so theo constructor của từng realm, nên hỏng khi
 * chạy trong sandbox test — và một phép kiểm mà test không xác minh được thì
 * chẳng bảo vệ được gì. toString đúng ở mọi realm.
 */
function flowIsDate_(value) {
  return Object.prototype.toString.call(value) === '[object Date]';
}


function flowNullableNumber_(
  value
) {

  if (
    typeof value ===
    'number'
  ) {

    return isNaN(value)
      ? null
      : value;

  }


  /*
   * Ngày tháng không phải số tiền. Ô định dạng kiểu ngày chứa số 0 sẽ về đây
   * dưới dạng Date, và bóc chữ số khỏi chuỗi ngày cho ra rác. null nghĩa là
   * "không có giá trị", đúng hơn hẳn một con số bịa.
   */
  if (flowIsDate_(value)) {
    return null;
  }


  let text =
    flowString_(value);


  if (!text) {
    return null;
  }


  text =
    text.replace(
      /\s/g,
      ''
    );


  text =
    text.replace(
      /[^0-9,.\-]/g,
      ''
    );


  if (!text) {
    return null;
  }


  /*
   * 1676,05 -> 1676.05
   * 1,676.05 -> 1676.05
   */
  if (
    text.indexOf(',') !== -1 &&
    text.indexOf('.') === -1
  ) {

    text =
      text.replace(
        /,/g,
        '.'
      );


  } else {

    text =
      text.replace(
        /,/g,
        ''
      );

  }


  const number =
    Number(text);


  return isNaN(number)
    ? null
    : number;

}


/************************************************************
 * ROUND MONEY
 ************************************************************/

function flowRoundMoney_(
  value
) {

  return Math.round(
    Number(
      value ||
      0
    ) *
    100
  ) / 100;

}


/************************************************************
 * AMOUNT EQUAL
 ************************************************************/

function flowAmountEqual_(
  first,
  second
) {

  return (
    Math.abs(
      Number(first) -
      Number(second)
    ) <=
    FLOW_CFG.AMOUNT_TOLERANCE
  );

}


/************************************************************
 * DATE
 ************************************************************/

function flowDate_(
  value
) {

  if (!value) {
    return null;
  }


  if (
    value instanceof Date
  ) {

    return isNaN(
      value.getTime()
    )
      ? null
      : value;

  }


  const date =
    new Date(value);


  return isNaN(
    date.getTime()
  )
    ? null
    : date;

}


/************************************************************
 * MILLIS
 ************************************************************/

function flowMillis_(
  value
) {

  const date =
    flowDate_(value);


  return date
    ? date.getTime()
    : 0;

}
