/************************************************************
 * PAYMENT ALL - GREY GMAIL CENTRAL COLLECTOR
 *
 * MÔ HÌNH:
 *
 * Grey account 1 ─┐
 * Grey account 2 ─┼── auto forward ──> Gmail trung tâm
 * Grey account 3 ─┘
 *
 *                          ↓
 *
 *                   GreyGmail.gs
 *
 *                          ↓
 *
 *                    Transaction
 *
 *
 * GREY.GS CHỈ LẤY RAW DATA.
 *
 * KHÔNG:
 * - tính Type
 * - tính Direction
 * - tính Gross / Net
 * - tính Fee
 * - normalize Currency
 * - map USDT
 * - reconcile
 * - tính balance
 * - tính summary
 *
 * TẤT CẢ BUSINESS LOGIC DO Flow.gs.
 ************************************************************/


const GG_CFG = Object.freeze({

  /**********************************************************
   * PAYMENT ALL SPREADSHEET
   **********************************************************/

  SPREADSHEET_ID:
    '1hBxL9CfFjSgoGuam-uxdiokyy6e35u5phR_M7KI10ew',


  GREY_SHEET:
    'Grey',

  TRANSACTION_SHEET:
    'Transaction',


  HEADER_ROW:
    2,

  DATA_START_ROW:
    3,


  /**********************************************************
   * GREY SENDER
   **********************************************************/

  SENDER:
    'hello@grey.co',


  /**********************************************************
   * GMAIL SEARCH
   *
   * Hiện dùng full search + dedupe.
   * An toàn nhất cho giai đoạn này.
   **********************************************************/

  BATCH_SIZE:
    100,

  MAX_THREADS:
    3000

});


/************************************************************
 * RAW FIELDS STORED IN TRANSACTION
 ************************************************************/

const GG_RAW_HEADERS = Object.freeze([

  'Raw Event',

  'Raw Transaction Type',

  'Raw Amount Tendered',

  'Raw Amount Received',

  'Raw Recipient',

  'Raw Channel',

  'Raw Date Time'

]);


/************************************************************
 * ENTRY
 *
 * CHẠY FUNCTION NÀY
 ************************************************************/

function syncGreyGmailRaw() {

  /**********************************************************
   * VALIDATE CONFIG
   **********************************************************/

  if (
    !GG_CFG.SPREADSHEET_ID ||
    GG_CFG.SPREADSHEET_ID ===
      'DAN_SHEET_ID_PAYMENT_ALL_VAO_DAY'
  ) {

    throw new Error(
      'Chưa điền GG_CFG.SPREADSHEET_ID.'
    );

  }


  /**********************************************************
   * OPEN PAYMENT ALL
   **********************************************************/

  const ss =
    SpreadsheetApp.openById(
      GG_CFG.SPREADSHEET_ID
    );


  const greySheet =
    ss.getSheetByName(
      GG_CFG.GREY_SHEET
    );


  const txSheet =
    ss.getSheetByName(
      GG_CFG.TRANSACTION_SHEET
    );


  if (!greySheet) {

    throw new Error(
      'Không tìm thấy sheet Grey.'
    );

  }


  if (!txSheet) {

    throw new Error(
      'Không tìm thấy sheet Transaction.'
    );

  }


  /**********************************************************
   * ENSURE RAW COLUMNS
   **********************************************************/

  ggEnsureRawHeaders_(
    txSheet
  );


  /**********************************************************
   * HEADERS
   **********************************************************/

  const greyHeaders =
    ggHeaderMap_(
      greySheet
    );


  const txHeaders =
    ggHeaderMap_(
      txSheet
    );


  ggRequireHeaders_(
    greyHeaders,
    [
      'Account Key',
      'Label',
      'Status',
      'Email',
      'Last Sync',
      'Sync Status'
    ],
    'Grey'
  );


  ggRequireHeaders_(
    txHeaders,
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


  /**********************************************************
   * LOAD ALL GREY ACCOUNTS
   *
   * Mapping:
   *
   * original email
   *      ↓
   * Account Key
   **********************************************************/

  const greyAccounts =
    ggLoadGreyAccounts_(
      greySheet,
      greyHeaders
    );


  if (
    !greyAccounts.accounts.length
  ) {

    throw new Error(
      'Không có Grey account ACTIVE nào có Email.'
    );

  }


  Logger.log(
    'Grey accounts loaded: ' +
    greyAccounts.accounts.length
  );


  /**********************************************************
   * MARK SYNCING
   **********************************************************/

  greyAccounts.accounts.forEach(
    account => {

      ggSetField_(
        greySheet,
        account.rowNumber,
        greyHeaders,
        'Sync Status',
        'SYNCING'
      );

    }
  );


  try {

    /********************************************************
     * LOAD TRANSACTION
     ********************************************************/

    const state =
      ggLoadTransactionState_(
        txSheet,
        txHeaders
      );


    /********************************************************
     * COUNTERS PER GREY ACCOUNT
     ********************************************************/

    const stats =
      {};


    greyAccounts.accounts.forEach(
      account => {

        stats[
          account.accountKey
        ] = {

          mapped:
            0,

          created:
            0,

          updated:
            0

        };

      }
    );


    /********************************************************
     * SEARCH GREY MAIL IN CENTRAL GMAIL
     ********************************************************/

    const query =
      'from:' +
      GG_CFG.SENDER;


    Logger.log(
      'Gmail query: ' +
      query
    );


    let start =
      0;


    let scanned =
      0;


    let transactionMails =
      0;


    let ignored =
      0;


    let unmapped =
      0;


    let ambiguous =
      0;


    /********************************************************
     * SEARCH LOOP
     ********************************************************/

    while (
      start <
      GG_CFG.MAX_THREADS
    ) {

      const threads =
        GmailApp.search(
          query,
          start,
          GG_CFG.BATCH_SIZE
        );


      if (
        !threads.length
      ) {

        break;

      }


      threads.forEach(
        thread => {

          const messages =
            thread.getMessages();


          messages.forEach(
            message => {

              /************************************************
               * VERIFY ACTUAL FROM
               ************************************************/

              const from =
                ggString_(
                  message.getFrom()
                )
                  .toLowerCase();


              if (
                from.indexOf(
                  GG_CFG.SENDER
                ) === -1
              ) {

                return;

              }


              scanned++;


              /************************************************
               * SUBJECT
               ************************************************/

              const subject =
                ggString_(
                  message.getSubject()
                );


              /************************************************
               * BODY
               ************************************************/

              const plainBody =
                message.getPlainBody() ||
                '';


              const htmlBody =
                ggHtmlToText_(
                  message.getBody() ||
                  ''
                );


              const combinedBody =
                ggNormalizeText_(
                  plainBody +
                  '\n' +
                  htmlBody
                );


              /************************************************
               * TRANSACTION EMAIL?
               ************************************************/

              if (
                !ggIsTransactionMail_(
                  subject,
                  combinedBody
                )
              ) {

                ignored++;

                return;

              }


              transactionMails++;


              /************************************************
               * FIND ORIGINAL GREY ACCOUNT
               *
               * Đây là phần quan trọng cho central Gmail.
               ************************************************/

              const resolved =
                ggResolveGreyAccount_(
                  message,
                  greyAccounts
                );


              /************************************************
               * NO MATCH
               ************************************************/

              if (
                resolved.status ===
                'UNMAPPED'
              ) {

                unmapped++;


                Logger.log(
                  'UNMAPPED MAIL | ' +
                  message.getId() +
                  ' | ' +
                  subject +
                  ' | recipients=' +
                  resolved.debug
                );


                return;

              }


              /************************************************
               * MULTIPLE GREY EMAILS MATCHED
               *
               * Không gán bừa.
               ************************************************/

              if (
                resolved.status ===
                'AMBIGUOUS'
              ) {

                ambiguous++;


                Logger.log(
                  'AMBIGUOUS MAIL | ' +
                  message.getId() +
                  ' | matched=' +
                  resolved.matches.join(
                    ', '
                  )
                );


                return;

              }


              const account =
                resolved.account;


              /************************************************
               * EXTRACT RAW TRANSACTION DATA
               ************************************************/

              const raw =
                ggExtractRaw_(
                  subject,
                  combinedBody
                );


              const messageId =
                String(
                  message.getId()
                );


              /************************************************
               * UNIQUE KEY
               *
               * Gmail Message ID là stable identity.
               *
               * Không gắn Account Key vào Unique Key để nếu
               * mapping email thay đổi vẫn không duplicate.
               ************************************************/

              const uniqueKey =
                'GREY|GMAIL|' +
                messageId;


              /************************************************
               * COLLECTOR-OWNED RECORD ONLY
               ************************************************/

              const record = {

                /**********************************************
                 * ORIGINAL GMAIL DATE
                 *
                 * Flow sẽ replace bằng Raw Date Time nếu
                 * parse được.
                 **********************************************/

                'Transaction Date':
                  message.getDate(),


                /**********************************************
                 * IDENTIFICATION
                 **********************************************/

                'Provider':
                  'GREY',

                'Account Key':
                  account.accountKey,

                'Label':
                  account.label,


                /**********************************************
                 * REFERENCE RAW
                 **********************************************/

                'Reference':
                  raw.reference,


                /**********************************************
                 * SOURCE
                 **********************************************/

                'Source':
                  'GREY_GMAIL',

                'Source ID':
                  messageId,

                'Sync Time':
                  new Date(),

                'Unique Key':
                  uniqueKey,


                /**********************************************
                 * RAW EMAIL DATA
                 **********************************************/

                'Raw Event':
                  subject,

                'Raw Transaction Type':
                  raw.transactionType,

                'Raw Amount Tendered':
                  raw.amountTendered,

                'Raw Amount Received':
                  raw.amountReceived,

                'Raw Recipient':
                  raw.recipient,

                'Raw Channel':
                  raw.channel,

                'Raw Date Time':
                  raw.dateTime

              };


              /************************************************
               * UPSERT
               ************************************************/

              const result =
                ggUpsertRaw_(
                  state,
                  txHeaders,
                  record
                );


              stats[
                account.accountKey
              ].mapped++;


              if (
                result ===
                'NEW'
              ) {

                stats[
                  account.accountKey
                ].created++;

              } else {

                stats[
                  account.accountKey
                ].updated++;

              }

            }
          );

        }
      );


      start +=
        threads.length;


      if (
        threads.length <
        GG_CFG.BATCH_SIZE
      ) {

        break;

      }

    }


    /********************************************************
     * WRITE TRANSACTION
     ********************************************************/

    ggWriteTransactionState_(
      txSheet,
      txHeaders,
      state
    );


    /********************************************************
     * UPDATE EACH GREY ACCOUNT SYNC STATUS
     ********************************************************/

    const now =
      new Date();


    greyAccounts.accounts.forEach(
      account => {

        const accountStats =
          stats[
            account.accountKey
          ];


        ggSetField_(
          greySheet,
          account.rowNumber,
          greyHeaders,
          'Last Sync',
          now
        );


        ggSetField_(
          greySheet,
          account.rowNumber,
          greyHeaders,
          'Sync Status',

          'RAW OK · ' +
          accountStats.mapped +
          ' tx'
        );

      }
    );


    SpreadsheetApp.flush();


    /********************************************************
     * LOG
     ********************************************************/

    Logger.log(
      '========================================'
    );


    Logger.log(
      'GREY CENTRAL COLLECTOR DONE'
    );


    Logger.log(
      'Scanned: ' +
      scanned
    );


    Logger.log(
      'Transaction mails: ' +
      transactionMails
    );


    Logger.log(
      'Ignored: ' +
      ignored
    );


    Logger.log(
      'Unmapped: ' +
      unmapped
    );


    Logger.log(
      'Ambiguous: ' +
      ambiguous
    );


    greyAccounts.accounts.forEach(
      account => {

        const s =
          stats[
            account.accountKey
          ];


        Logger.log(
          account.accountKey +
          ' | ' +
          account.email +
          ' | mapped=' +
          s.mapped +
          ' | new=' +
          s.created +
          ' | updated=' +
          s.updated
        );

      }
    );


    Logger.log(
      '========================================'
    );


  } catch (error) {

    /********************************************************
     * ERROR STATUS
     ********************************************************/

    greyAccounts.accounts.forEach(
      account => {

        ggSetField_(
          greySheet,
          account.rowNumber,
          greyHeaders,
          'Sync Status',

          'ERROR: ' +
          ggTruncate_(
            error.message ||
            String(error),
            250
          )
        );

      }
    );


    SpreadsheetApp.flush();


    throw error;

  }

}


/************************************************************
 ************************************************************
 *
 * GREY ACCOUNT MAPPING
 *
 ************************************************************
 ************************************************************/


/************************************************************
 * LOAD GREY ACCOUNTS FROM SHEET
 *
 * Grey:
 *
 * Account Key | Label | Status | Email
 ************************************************************/

function ggLoadGreyAccounts_(
  sheet,
  headers
) {

  const lastRow =
    sheet.getLastRow();


  const accounts =
    [];


  const emailMap =
    {};


  if (
    lastRow <
    GG_CFG.DATA_START_ROW
  ) {

    return {

      accounts:
        accounts,

      emailMap:
        emailMap

    };

  }


  const rows =
    sheet
      .getRange(
        GG_CFG.DATA_START_ROW,
        1,
        lastRow -
          GG_CFG.DATA_START_ROW +
          1,
        sheet.getLastColumn()
      )
      .getValues();


  rows.forEach(
    (row, index) => {

      const accountKey =
        ggString_(
          row[
            headers[
              'Account Key'
            ]
          ]
        );


      if (!accountKey) {

        return;

      }


      const status =
        ggString_(
          row[
            headers[
              'Status'
            ]
          ]
        )
          .toUpperCase();


      if (
        status ===
        'DISABLED'
      ) {

        return;

      }


      const rawEmail =
        ggString_(
          row[
            headers[
              'Email'
            ]
          ]
        );


      const email =
        ggNormalizeEmail_(
          rawEmail
        );


      /*
       * Không có Email thì central collector
       * không thể map account.
       */

      if (!email) {

        Logger.log(
          'Grey account thiếu Email: ' +
          accountKey
        );


        return;

      }


      /********************************************************
       * DUPLICATE EMAIL = nguy hiểm
       *
       * Không cho 2 Grey account dùng cùng Email.
       ********************************************************/

      if (
        Object.prototype
          .hasOwnProperty.call(
            emailMap,
            email
          )
      ) {

        throw new Error(
          'Duplicate Grey Email "' +
          email +
          '" ở account ' +
          emailMap[
            email
          ].accountKey +
          ' và ' +
          accountKey
        );

      }


      const account = {

        rowNumber:
          GG_CFG.DATA_START_ROW +
          index,

        accountKey:
          accountKey,

        label:
          ggString_(
            row[
              headers[
                'Label'
              ]
            ]
          ),

        email:
          email

      };


      accounts.push(
        account
      );


      emailMap[
        email
      ] =
        account;

    }
  );


  return {

    accounts:
      accounts,

    emailMap:
      emailMap

  };

}


/************************************************************
 * RESOLVE MESSAGE -> GREY ACCOUNT
 *
 * Check:
 *
 * To
 * Cc
 * Bcc
 * Delivered-To
 * X-Original-To
 * Envelope-To
 * X-Forwarded-To
 * Resent-To
 * raw RFC822 headers
 ************************************************************/

function ggResolveGreyAccount_(
  message,
  greyAccounts
) {

  const headerParts =
    [];


  /**********************************************************
   * STANDARD GMAIL VALUES
   **********************************************************/

  headerParts.push(
    ggSafeStringCall_(
      function() {
        return message.getTo();
      }
    )
  );


  headerParts.push(
    ggSafeStringCall_(
      function() {
        return message.getCc();
      }
    )
  );


  headerParts.push(
    ggSafeStringCall_(
      function() {
        return message.getBcc();
      }
    )
  );


  /**********************************************************
   * FORWARD / DELIVERY HEADERS
   **********************************************************/

  const headerNames = [

    'To',

    'Cc',

    'Delivered-To',

    'X-Original-To',

    'Envelope-To',

    'X-Envelope-To',

    'X-Forwarded-To',

    'Resent-To',

    'Original-Recipient',

    'Final-Recipient'

  ];


  headerNames.forEach(
    headerName => {

      headerParts.push(
        ggSafeStringCall_(
          function() {

            return message.getHeader(
              headerName
            );

          }
        )
      );

    }
  );


  /**********************************************************
   * RAW RFC822 HEADERS
   *
   * Chỉ lấy phần header trước blank line.
   *
   * Không search body để tránh match nhầm email
   * nằm trong nội dung transaction.
   **********************************************************/

  const rawHeader =
    ggGetRawHeaderBlock_(
      message
    );


  if (rawHeader) {

    headerParts.push(
      rawHeader
    );

  }


  /**********************************************************
   * COMBINED HEADER TEXT
   **********************************************************/

  const headerText =
    headerParts
      .filter(
        Boolean
      )
      .join(
        '\n'
      );


  /**********************************************************
   * FIND ALL GREY EMAIL MATCHES
   **********************************************************/

  const matches =
    [];


  greyAccounts.accounts.forEach(
    account => {

      if (
        ggHeaderContainsEmail_(
          headerText,
          account.email
        )
      ) {

        matches.push(
          account
        );

      }

    }
  );


  /**********************************************************
   * EXACTLY ONE
   **********************************************************/

  if (
    matches.length ===
    1
  ) {

    return {

      status:
        'OK',

      account:
        matches[0],

      matches:
        [
          matches[0].email
        ],

      debug:
        headerText

    };

  }


  /**********************************************************
   * MORE THAN ONE
   *
   * Không được tự đoán.
   **********************************************************/

  if (
    matches.length >
    1
  ) {

    return {

      status:
        'AMBIGUOUS',

      account:
        null,

      matches:
        matches.map(
          item =>
            item.email
        ),

      debug:
        headerText

    };

  }


  /**********************************************************
   * NO ACCOUNT
   **********************************************************/

  return {

    status:
      'UNMAPPED',

    account:
      null,

    matches:
      [],

    debug:
      ggTruncate_(
        headerText
          .replace(
            /\s+/g,
            ' '
          ),
        500
      )

  };

}


/************************************************************
 * GET RAW EMAIL HEADER BLOCK
 ************************************************************/

function ggGetRawHeaderBlock_(
  message
) {

  try {

    const raw =
      String(
        message.getRawContent() ||
        ''
      );


    if (!raw) {

      return '';

    }


    /*
     * RFC822:
     *
     * headers
     *
     * blank line
     *
     * body
     */

    let splitIndex =
      raw.indexOf(
        '\r\n\r\n'
      );


    let separatorLength =
      4;


    if (
      splitIndex === -1
    ) {

      splitIndex =
        raw.indexOf(
          '\n\n'
        );


      separatorLength =
        2;

    }


    if (
      splitIndex === -1
    ) {

      /*
       * Safety:
       * không lấy toàn raw body.
       */

      return raw.substring(
        0,
        Math.min(
          raw.length,
          12000
        )
      );

    }


    return raw.substring(
      0,
      splitIndex +
      separatorLength
    );


  } catch (error) {

    return '';

  }

}


/************************************************************
 * EXACT EMAIL MATCH INSIDE HEADER TEXT
 ************************************************************/

function ggHeaderContainsEmail_(
  headerText,
  email
) {

  const source =
    String(
      headerText || ''
    )
      .toLowerCase();


  const target =
    String(
      email || ''
    )
      .toLowerCase()
      .trim();


  if (!target) {

    return false;

  }


  /*
   * Simple exact substring is safe enough because
   * full email address contains @domain.
   *
   * Additional boundaries avoid:
   *
   * foo@domain.com matching
   * xfoo@domain.com
   */

  const escaped =
    ggRegexEscape_(
      target
    );


  const regex =
    new RegExp(
      '(^|[^a-z0-9._%+\\-])' +
      escaped +
      '([^a-z0-9._%+\\-]|$)',
      'i'
    );


  return regex.test(
    source
  );

}


/************************************************************
 ************************************************************
 *
 * TRANSACTION EMAIL DETECTION
 *
 ************************************************************
 ************************************************************/


function ggIsTransactionMail_(
  subject,
  body
) {

  const text =
    (
      String(
        subject || ''
      ) +
      '\n' +
      String(
        body || ''
      )
    )
      .toLowerCase();


  /**********************************************************
   * FUNDING
   **********************************************************/

  if (
    text.indexOf(
      'wallet has been funded'
    ) !== -1
  ) {

    return true;

  }


  if (
    text.indexOf(
      'wallet funding successful'
    ) !== -1
  ) {

    return true;

  }


  /**********************************************************
   * WITHDRAWAL
   **********************************************************/

  if (
    text.indexOf(
      'withdrawal successful'
    ) !== -1
  ) {

    return true;

  }


  if (
    text.indexOf(
      'withdrawal was successful'
    ) !== -1
  ) {

    return true;

  }


  /**********************************************************
   * FALLBACK
   **********************************************************/

  if (
    text.indexOf(
      'transaction type'
    ) !== -1 &&
    (
      text.indexOf(
        'amount tendered'
      ) !== -1 ||
      text.indexOf(
        'amount received'
      ) !== -1
    )
  ) {

    return true;

  }


  return false;

}


/************************************************************
 ************************************************************
 *
 * RAW GREY EXTRACTION
 *
 ************************************************************
 ************************************************************/


function ggExtractRaw_(
  subject,
  body
) {

  const text =
    ggNormalizeText_(
      body
    );


  return {

    transactionType:
      ggFindGreyField_(
        text,
        'Transaction Type'
      ),

    amountTendered:
      ggFindGreyField_(
        text,
        'Amount tendered'
      ),

    amountReceived:
      ggFindGreyField_(
        text,
        'Amount received'
      ),

    recipient:
      ggFindGreyField_(
        text,
        'Recipient'
      ),

    channel:
      ggFindGreyField_(
        text,
        'Channel'
      ),

    reference:
      ggFindGreyField_(
        text,
        'Reference'
      ),

    dateTime:
      ggFindGreyField_(
        text,
        'Date & Time'
      )

  };

}


/************************************************************
 * KNOWN GREY EMAIL FIELD LABELS
 *
 * Used to stop previous field before next field.
 ************************************************************/

function ggGreyKnownLabels_() {

  return [

    'Transaction Type',

    'Amount tendered',

    'Amount received',

    'Network',

    'Phone number',

    'Memo',

    'Recipient',

    'Channel',

    'Reference',

    'Date & Time'

  ];

}


/************************************************************
 * ROBUST GREY FIELD EXTRACTOR
 *
 * Handles flattened email:
 *
 * Amount received: USDT499.00 Network: BEP20
 * Reference: ABC Date & Time: 08/24/2026...
 *
 * Correctly outputs:
 *
 * USDT499.00
 * ABC
 * 08/24/2026...
 ************************************************************/

function ggFindGreyField_(
  text,
  label
) {

  const source =
    String(
      text || ''
    )
      .replace(
        /\u00a0/g,
        ' '
      )
      .replace(
        /\r/g,
        '\n'
      )
      .replace(
        /[ \t]+/g,
        ' '
      );


  const lowerSource =
    source.toLowerCase();


  const lowerLabel =
    String(
      label || ''
    )
      .toLowerCase();


  /**********************************************************
   * FIRST OCCURRENCE OF LABEL
   **********************************************************/

  const labelIndex =
    lowerSource.indexOf(
      lowerLabel
    );


  if (
    labelIndex ===
    -1
  ) {

    return '';

  }


  /**********************************************************
   * VALUE START
   **********************************************************/

  let valueStart =
    labelIndex +
    label.length;


  /**********************************************************
   * Skip separators after label
   **********************************************************/

  while (
    valueStart <
    source.length
  ) {

    const char =
      source.charAt(
        valueStart
      );


    if (
      char === ':' ||
      char === '-' ||
      char === ' ' ||
      char === '\t' ||
      char === '\n'
    ) {

      valueStart++;

      continue;

    }


    break;

  }


  /**********************************************************
   * FIND NEAREST NEXT KNOWN LABEL
   **********************************************************/

  const labels =
    ggGreyKnownLabels_();


  let valueEnd =
    source.length;


  labels.forEach(
    nextLabel => {

      if (
        nextLabel
          .toLowerCase() ===
        lowerLabel
      ) {

        return;

      }


      const nextIndex =
        lowerSource.indexOf(
          nextLabel
            .toLowerCase(),
          valueStart
        );


      if (
        nextIndex !==
          -1 &&
        nextIndex <
          valueEnd
      ) {

        valueEnd =
          nextIndex;

      }

    }
  );


  /**********************************************************
   * EXTRACT
   **********************************************************/

  let value =
    source.substring(
      valueStart,
      valueEnd
    );


  /**********************************************************
   * Stop at newline if newline occurs before another field.
   **********************************************************/

  const newlineIndex =
    value.indexOf(
      '\n'
    );


  if (
    newlineIndex !==
    -1
  ) {

    value =
      value.substring(
        0,
        newlineIndex
      );

  }


  return String(
    value || ''
  )
    .replace(
      /\s+/g,
      ' '
    )
    .trim();

}


/************************************************************
 ************************************************************
 *
 * RAW COLUMN SETUP
 *
 ************************************************************
 ************************************************************/


function ggEnsureRawHeaders_(
  sheet
) {

  let lastColumn =
    Math.max(
      1,
      sheet.getLastColumn()
    );


  const existing =
    sheet
      .getRange(
        GG_CFG.HEADER_ROW,
        1,
        1,
        lastColumn
      )
      .getValues()[0]
      .map(
        value =>
          ggString_(
            value
          )
      );


  GG_RAW_HEADERS.forEach(
    header => {

      if (
        existing.indexOf(
          header
        ) !==
        -1
      ) {

        return;

      }


      lastColumn++;


      sheet
        .getRange(
          GG_CFG.HEADER_ROW,
          lastColumn
        )
        .setValue(
          header
        );


      existing.push(
        header
      );

    }
  );


  SpreadsheetApp.flush();

}


/************************************************************
 ************************************************************
 *
 * TRANSACTION STATE
 *
 ************************************************************
 ************************************************************/


function ggLoadTransactionState_(
  sheet,
  headers
) {

  const lastRow =
    sheet.getLastRow();


  const lastColumn =
    sheet.getLastColumn();


  let rows =
    [];


  if (
    lastRow >=
    GG_CFG.DATA_START_ROW
  ) {

    rows =
      sheet
        .getRange(
          GG_CFG.DATA_START_ROW,
          1,
          lastRow -
            GG_CFG.DATA_START_ROW +
            1,
          lastColumn
        )
        .getValues();


    rows =
      rows.filter(
        row =>
          row.some(
            value =>
              value !== '' &&
              value !== null
          )
      );

  }


  const keyMap =
    {};


  const sourceMap =
    {};


  rows.forEach(
    (row, index) => {

      const uniqueKey =
        ggString_(
          row[
            headers[
              'Unique Key'
            ]
          ]
        );


      if (uniqueKey) {

        keyMap[
          uniqueKey
        ] =
          index;

      }


      /********************************************************
       * Also index by Source + Gmail Message ID.
       *
       * This migrates old rows whose Unique Key was:
       *
       * GREY|GR001|GMAIL|xxxx
       *
       * to new:
       *
       * GREY|GMAIL|xxxx
       ********************************************************/

      const source =
        ggString_(
          row[
            headers[
              'Source'
            ]
          ]
        )
          .toUpperCase();


      const sourceId =
        ggString_(
          row[
            headers[
              'Source ID'
            ]
          ]
        );


      if (
        source ===
          'GREY_GMAIL' &&
        sourceId
      ) {

        sourceMap[
          source +
          '|' +
          sourceId
        ] =
          index;

      }

    }
  );


  return {

    rows:
      rows,

    keyMap:
      keyMap,

    sourceMap:
      sourceMap,

    columnCount:
      lastColumn,

    oldRowCount:
      Math.max(
        0,
        lastRow -
          GG_CFG.DATA_START_ROW +
          1
      )

  };

}


/************************************************************
 * UPSERT RAW GREY EVENT
 *
 * Lookup priority:
 *
 * 1. new Unique Key
 * 2. existing Gmail Source ID
 *
 * Therefore old GR001 key will not duplicate.
 ************************************************************/

function ggUpsertRaw_(
  state,
  headers,
  record
) {

  const uniqueKey =
    record[
      'Unique Key'
    ];


  const sourceKey =
    'GREY_GMAIL|' +
    record[
      'Source ID'
    ];


  if (!uniqueKey) {

    throw new Error(
      'Grey Gmail row thiếu Unique Key.'
    );

  }


  let row =
    null;


  let rowIndex =
    -1;


  let result =
    'NEW';


  /**********************************************************
   * NEW KEY EXISTS
   **********************************************************/

  if (
    Object.prototype
      .hasOwnProperty.call(
        state.keyMap,
        uniqueKey
      )
  ) {

    rowIndex =
      state.keyMap[
        uniqueKey
      ];


    row =
      state.rows[
        rowIndex
      ];


    result =
      'UPDATED';

  }


  /**********************************************************
   * OLD KEY / SAME GMAIL MESSAGE ID
   **********************************************************/

  if (
    !row &&
    Object.prototype
      .hasOwnProperty.call(
        state.sourceMap,
        sourceKey
      )
  ) {

    rowIndex =
      state.sourceMap[
        sourceKey
      ];


    row =
      state.rows[
        rowIndex
      ];


    result =
      'UPDATED';


    /*
     * Remove old key index if different.
     */

    const oldKey =
      ggString_(
        row[
          headers[
            'Unique Key'
          ]
        ]
      );


    if (
      oldKey &&
      oldKey !==
        uniqueKey
    ) {

      delete state.keyMap[
        oldKey
      ];

    }

  }


  /**********************************************************
   * NEW ROW
   **********************************************************/

  if (!row) {

    row =
      new Array(
        state.columnCount
      )
        .fill('');


    state.rows.push(
      row
    );


    rowIndex =
      state.rows.length -
      1;


    result =
      'NEW';

  }


  /**********************************************************
   * COLLECTOR ONLY UPDATES RAW/SOURCE FIELDS
   **********************************************************/

  Object.keys(
    record
  )
    .forEach(
      header => {

        if (
          !Object.prototype
            .hasOwnProperty.call(
              headers,
              header
            )
        ) {

          return;

        }


        row[
          headers[
            header
          ]
        ] =
          record[
            header
          ];

      }
    );


  /**********************************************************
   * UPDATE INDEXES
   **********************************************************/

  state.keyMap[
    uniqueKey
  ] =
    rowIndex;


  state.sourceMap[
    sourceKey
  ] =
    rowIndex;


  return result;

}


/************************************************************
 * WRITE TRANSACTION
 ************************************************************/

function ggWriteTransactionState_(
  sheet,
  headers,
  state
) {

  const dateIndex =
    headers[
      'Transaction Date'
    ];


  /**********************************************************
   * NEWEST FIRST
   **********************************************************/

  state.rows.sort(
    (a, b) =>

      ggMillis_(
        b[
          dateIndex
        ]
      ) -

      ggMillis_(
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
    clearRows >
    0
  ) {

    sheet
      .getRange(
        GG_CFG.DATA_START_ROW,
        1,
        clearRows,
        state.columnCount
      )
      .clearContent();

  }


  if (
    state.rows.length >
    0
  ) {

    sheet
      .getRange(
        GG_CFG.DATA_START_ROW,
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
 ************************************************************
 *
 * HEADER HELPERS
 *
 ************************************************************
 ************************************************************/


function ggHeaderMap_(
  sheet
) {

  const row =
    sheet
      .getRange(
        GG_CFG.HEADER_ROW,
        1,
        1,
        sheet.getLastColumn()
      )
      .getValues()[0];


  const result =
    {};


  row.forEach(
    (value, index) => {

      const header =
        ggString_(
          value
        );


      if (!header) {

        return;

      }


      if (
        Object.prototype
          .hasOwnProperty.call(
            result,
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


      result[
        header
      ] =
        index;

    }
  );


  return result;

}


function ggRequireHeaders_(
  map,
  required,
  sheetName
) {

  const missing =
    required.filter(
      header =>
        !Object.prototype
          .hasOwnProperty.call(
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
 ************************************************************
 *
 * TEXT / EMAIL HELPERS
 *
 ************************************************************
 ************************************************************/


function ggNormalizeEmail_(
  value
) {

  const text =
    ggString_(
      value
    )
      .toLowerCase();


  if (!text) {

    return '';

  }


  /*
   * Handles:
   *
   * user@example.com
   *
   * Name <user@example.com>
   */

  const match =
    text.match(
      /[a-z0-9.!#$%&'*+/=?^_`{|}~-]+@[a-z0-9.-]+\.[a-z]{2,}/i
    );


  return match
    ? match[0]
        .toLowerCase()
        .trim()
    : '';

}


function ggNormalizeText_(
  text
) {

  return String(
    text || ''
  )
    .replace(
      /\r/g,
      '\n'
    )
    .replace(
      /\u00a0/g,
      ' '
    )
    .replace(
      /[ \t]+/g,
      ' '
    )
    .replace(
      /\n[ \t]+/g,
      '\n'
    )
    .replace(
      /\n{3,}/g,
      '\n\n'
    );

}


function ggHtmlToText_(
  html
) {

  return String(
    html || ''
  )

    .replace(
      /<br\s*\/?>/gi,
      '\n'
    )

    .replace(
      /<\/(p|div|td|tr|li|table|h1|h2|h3)>/gi,
      '\n'
    )

    .replace(
      /<[^>]+>/g,
      ' '
    )

    .replace(
      /&nbsp;/gi,
      ' '
    )

    .replace(
      /&amp;/gi,
      '&'
    )

    .replace(
      /&lt;/gi,
      '<'
    )

    .replace(
      /&gt;/gi,
      '>'
    )

    .replace(
      /&#39;/gi,
      "'"
    )

    .replace(
      /&quot;/gi,
      '"'
    );

}


function ggRegexEscape_(
  value
) {

  return String(
    value || ''
  )
    .replace(
      /[.*+?^${}()|[\]\\]/g,
      '\\$&'
    );

}


function ggSafeStringCall_(
  fn
) {

  try {

    return String(
      fn() ||
      ''
    );

  } catch (error) {

    return '';

  }

}


/************************************************************
 ************************************************************
 *
 * GENERIC HELPERS
 *
 ************************************************************
 ************************************************************/


function ggSetField_(
  sheet,
  row,
  headers,
  header,
  value
) {

  sheet
    .getRange(
      row,
      headers[
        header
      ] + 1
    )
    .setValue(
      value
    );

}


function ggString_(
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


function ggMillis_(
  value
) {

  if (!value) {

    return 0;

  }


  const date =
    value instanceof Date
      ? value
      : new Date(
          value
        );


  return isNaN(
    date.getTime()
  )
    ? 0
    : date.getTime();

}


function ggTruncate_(
  value,
  maxLength
) {

  const text =
    String(
      value || ''
    );


  return text.length <=
    maxLength
      ? text
      : text.substring(
          0,
          maxLength
        );

}