/************************************************************
 * GREY GMAIL COLLECTOR — STANDALONE
 * ==========================================================
 * KHÔNG thuộc project Payment All.
 *
 * Mỗi Gmail nhận mail Grey = một Apps Script project riêng.
 * Copy trọn file này vào project đó, sửa đúng khối CONFIG
 * bên dưới, chạy setupGreyDailyTrigger() một lần là xong.
 *
 * Dùng được cho cả hai kiểu hộp thư: hộp thư riêng của từng
 * Grey account, và hộp thư trung tâm nhận forward từ nhiều
 * account. Account của mỗi email được suy ra từ header người
 * nhận rồi đối chiếu cột Email của sheet Grey, nên không phụ
 * thuộc vào việc khai đúng account key trong config.
 *
 * NHIỆM VỤ DUY NHẤT: Gmail -> Transaction (dữ liệu thô).
 * Không tính Type / Direction / Gross / Net / Fee / Currency,
 * không normalize, không reconcile, không tính summary.
 * Toàn bộ phần đó là việc của Flow.gs bên Payment All.
 *
 * ==========================================================
 * BA THỨ FILE NÀY BẢO ĐẢM
 * ==========================================================
 *
 * 1. CHỐNG TRÙNG
 *    Khoá chính là Reference của chính Grey, không phải Gmail
 *    message ID. Cùng một giao dịch Grey về hai hộp thư khác
 *    nhau sẽ là hai email khác nhau với hai message ID khác
 *    nhau — khoá theo message ID không thể gộp chúng lại,
 *    khoá theo Reference thì được.
 *
 * 2. CHỐNG GHI ĐÈ
 *    Chỉ sửa đúng dòng của mình, và chỉ đúng những cột mình
 *    sở hữu. Không bao giờ đọc-rồi-ghi-lại cả sheet, không
 *    sắp xếp lại sheet. Cột do Flow sở hữu (Type, Direction,
 *    Gross, Net, Fee, Currency, Counterparty, Reconcile...)
 *    không bị đụng tới.
 *
 * 3. KHÔNG XOÁ DỮ LIỆU CŨ
 *    Không có một lệnh clearContent() nào trong file này.
 *    Chỉ append dòng mới hoặc vá ô đã có.
 *
 * ==========================================================
 * CHẠY CÙNG LÚC VỚI syncFlow
 * ==========================================================
 * syncFlow() CÓ tham gia advisory lock bên dưới: cùng sheet
 * '_Sync Lock', cùng ô A2/B2, cùng cách ghi token rồi đọc lại
 * để xác nhận. Nên hai bên loại trừ nhau đúng như mong muốn.
 *
 * Ngoài ra syncFlow còn một lưới nữa: trước khi ghi nó kiểm
 * lại số dòng sheet Transaction, thấy dài ra so với lúc đọc
 * thì bỏ hẳn lượt thay vì ghi đè mất dòng collector vừa thêm.
 *
 * Nhưng syncFlow vẫn ghi đè TOÀN BỘ vùng dữ liệu Transaction,
 * nên đụng nhau tuy an toàn mà vẫn tốn: một bên phải chờ tới
 * 90 giây rồi có thể bỏ lượt. Vẫn nên xếp lịch trigger lệch
 * hẳn khỏi giờ chạy syncFlow — vì thông lượng, không phải vì
 * an toàn.
 ************************************************************/


/************************************************************
 * CONFIG — SỬA PHẦN NÀY CHO TỪNG GMAIL
 ************************************************************/

const GG_CFG = Object.freeze({

  // Payment All thật. Giống nhau ở mọi project.
  SPREADSHEET_ID: '1YherTvOxpqkmW5XJ_ydn_U-5dQLFeHhFn2Hpm4XvWos',

  // Account được suy ra từ header người nhận của từng email, đối chiếu với
  // cột Email của sheet Grey. Nhờ vậy một file dùng chung được cho cả hộp thư
  // riêng của từng Grey account lẫn hộp thư trung tâm nhận forward.
  //
  // Giá trị dưới đây CHỈ là phương án dự phòng, dùng khi không tra được
  // account nào từ header. Hộp thư riêng thì điền; hộp thư trung tâm nên
  // để trống, vì đoán bừa còn tệ hơn bỏ sót.
  FALLBACK_ACCOUNT_KEY: 'GR001',

  SENDER: 'hello@grey.co',

  // Đọc cả mail forward TAY (bấm nút Forward), không chỉ mail Grey gửi thẳng
  // và mail chuyển tiếp tự động.
  //
  // Cần bật ở hộp thư trung tâm nhận mail cũ từ những hộp mình không có quyền
  // chạy script: chuyển tiếp tự động chỉ áp dụng cho mail đến SAU khi bật, nên
  // toàn bộ lịch sử phải forward tay.
  //
  // Hộp thư riêng của một Grey account thì tắt đi cho nhanh — ở đó không có
  // mail forward nào.
  READ_FORWARDED: true,


  // Chỉ quét mail trong ngần này ngày. Chạy hằng ngày thì 7 là dư biên:
  // mail về muộn, Gmail lỗi, hay trigger nghỉ vài ngày đều bù được.
  // Đặt 0 = quét toàn bộ (chỉ dùng cho lần nạp đầu tiên).
  LOOKBACK_DAYS: 7,

  // Giờ chạy hằng ngày. Xếp lệch nhau giữa các project:
  //   GR001 -> HOUR 1,  GR002 -> HOUR 2,  GR003 -> HOUR 3
  // nearMinute() có thể lệch ±15 phút nên đừng để cách nhau dưới 30 phút.
  TRIGGER_HOUR: 1,
  TRIGGER_NEAR_MINUTE: 0,
  TIMEZONE: 'Asia/Ho_Chi_Minh',

  /*********** Từ đây trở xuống thường không cần sửa ***********/

  GREY_SHEET: 'Grey',
  TRANSACTION_SHEET: 'Transaction',
  LOCK_SHEET: '_Sync Lock',

  HEADER_ROW: 2,
  DATA_START_ROW: 3,

  BATCH_SIZE: 100,
  MAX_THREADS: 3000,

  LOCK_WAIT_MS: 90000,
  LOCK_STALE_MS: 600000,

  HANDLER: 'syncGreyGmailRaw'
});


const GG_RAW_HEADERS = Object.freeze([
  'Raw Event',
  'Raw Transaction Type',
  'Raw Amount Tendered',
  'Raw Amount Received',
  'Raw Recipient',
  'Raw Channel',
  'Raw Date Time'
]);


/*
 * Cột collector được phép ghi khi CẬP NHẬT một dòng đã có.
 * Mọi cột khác — kể cả Transaction Date — để nguyên, vì Flow
 * thay Transaction Date bằng Raw Date Time khi parse được, và
 * ghi đè lại là phá kết quả của Flow.
 */
const GG_UPDATABLE = Object.freeze([
  'Reference',
  'Raw Event',
  'Raw Transaction Type',
  'Raw Amount Tendered',
  'Raw Amount Received',
  'Raw Recipient',
  'Raw Channel',
  'Raw Date Time'
]);


/************************************************************
 * TRIGGER
 ************************************************************/

function setupGreyDailyTrigger() {
  if (GG_CFG.TRIGGER_HOUR < 0 || GG_CFG.TRIGGER_HOUR > 23) {
    throw new Error('TRIGGER_HOUR phải từ 0 đến 23.');
  }

  const removed = removeGreyDailyTrigger();

  ScriptApp.newTrigger(GG_CFG.HANDLER)
    .timeBased()
    .atHour(GG_CFG.TRIGGER_HOUR)
    .nearMinute(GG_CFG.TRIGGER_NEAR_MINUTE)
    .everyDays(1)
    .inTimezone(GG_CFG.TIMEZONE)
    .create();

  Logger.log('Trigger đã tạo — mỗi ngày gần ' +
    String(GG_CFG.TRIGGER_HOUR).padStart(2, '0') + ':' +
    String(GG_CFG.TRIGGER_NEAR_MINUTE).padStart(2, '0') +
    ' (' + GG_CFG.TIMEZONE + '). Đã xoá ' + removed + ' trigger cũ.');
}


function removeGreyDailyTrigger() {
  let removed = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === GG_CFG.HANDLER) {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });
  return removed;
}


/************************************************************
 * ENTRY POINTS
 ************************************************************/

function syncGreyGmailRaw() {
  return ggRun_(GG_CFG.LOOKBACK_DAYS);
}

/* Lần nạp đầu tiên. Đừng đặt vào trigger: thời gian chạy tăng mãi. */
function syncGreyGmailRawFull() {
  return ggRun_(0);
}


function ggRun_(lookbackDays) {
  const started = new Date();
  const ss = SpreadsheetApp.openById(GG_CFG.SPREADSHEET_ID);

  const greySheet = ss.getSheetByName(GG_CFG.GREY_SHEET);
  const txSheet = ss.getSheetByName(GG_CFG.TRANSACTION_SHEET);
  if (!greySheet) throw new Error('Không tìm thấy sheet Grey.');
  if (!txSheet) throw new Error('Không tìm thấy sheet Transaction.');

  ggEnsureRawHeaders_(txSheet);

  const greyHeaders = ggHeaderMap_(greySheet);
  const txHeaders = ggHeaderMap_(txSheet);

  ggRequireHeaders_(greyHeaders, ['Account Key', 'Label', 'Status', 'Email', 'Last Sync', 'Sync Status'], 'Grey');
  ggRequireHeaders_(txHeaders, [
    'Transaction Date', 'Provider', 'Account Key', 'Label',
    'Reference', 'Source', 'Source ID', 'Sync Time', 'Unique Key'
  ].concat(GG_RAW_HEADERS), 'Transaction');

  const greyAccounts = ggLoadGreyAccounts_(greySheet, greyHeaders);
  if (!greyAccounts.accounts.length) {
    throw new Error('Sheet Grey không có account nào bật kèm cột Email.');
  }

  const fallback = ggFallbackAccount_(greyAccounts);

  const lock = ggAcquireLock_(ss);

  /* Chỉ những account thật sự có mail trong lượt này mới được đánh dấu. */
  const touched = {};
  const markSyncing = function (acc) {
    if (touched[acc.accountKey]) return;
    touched[acc.accountKey] = acc;
    ggSetField_(greySheet, acc.rowNumber, greyHeaders, 'Sync Status', 'SYNCING');
  };
  if (fallback) markSyncing(fallback);

  try {
    const stats = ggCollect_(txSheet, txHeaders, greyAccounts, fallback, lookbackDays, markSyncing);

    const now = new Date();
    Object.keys(touched).forEach(function (key) {
      const acc = touched[key];
      const per = stats.perAccount[key] || { created: 0, updated: 0, unchanged: 0 };
      ggSetField_(greySheet, acc.rowNumber, greyHeaders, 'Last Sync', now);
      ggSetField_(greySheet, acc.rowNumber, greyHeaders, 'Sync Status',
        'RAW OK · ' + per.created + ' new · ' + per.updated + ' upd · ' +
        per.unchanged + ' same' +
        (stats.unmapped ? ' · ' + stats.unmapped + ' UNMAPPED' : '') +
        (stats.ambiguous ? ' · ' + stats.ambiguous + ' AMBIGUOUS' : ''));
    });

    SpreadsheetApp.flush();

    const secs = ((new Date().getTime() - started.getTime()) / 1000).toFixed(1);
    Logger.log('GREY COLLECTOR DONE in ' + secs + 's · scanned ' + stats.scanned +
      ' · tx mails ' + stats.transactionMails +
      ' · new ' + stats.created + ' · updated ' + stats.updated +
      ' · unchanged ' + stats.unchanged +
      ' · no reference ' + stats.noReference +
      ' · unmapped ' + stats.unmapped +
      ' · ambiguous ' + stats.ambiguous +
      ' · forwarded ' + stats.forwarded +
      ' · ignored ' + stats.ignored);
    Object.keys(stats.perPass).forEach(function (label) {
      const t = stats.perPass[label];
      Logger.log('  lượt ' + label + ': Gmail trả ' + t.matched +
        ' · không phải Grey ' + t.notGrey +
        ' · xét ' + t.scanned +
        ' · bỏ qua ' + t.ignored +
        ' · nhận ' + t.kept);
    });
    Object.keys(stats.perAccount).forEach(function (key) {
      const p = stats.perAccount[key];
      Logger.log('  ' + key + ': ' + p.created + ' new · ' + p.updated + ' upd · ' + p.unchanged + ' same');
    });

    return stats;

  } catch (err) {
    try {
      Object.keys(touched).forEach(function (key) {
        ggSetField_(greySheet, touched[key].rowNumber, greyHeaders, 'Sync Status',
          'ERROR: ' + ggTruncate_(err.message || String(err), 300));
      });
      SpreadsheetApp.flush();
    } catch (ignored) { /* đừng để việc ghi trạng thái che mất lỗi gốc */ }
    throw err;

  } finally {
    ggReleaseLock_(lock);
  }
}


function ggFallbackAccount_(greyAccounts) {
  const key = ggString_(GG_CFG.FALLBACK_ACCOUNT_KEY);
  if (!key) return null;
  const acc = greyAccounts.byKey[key];
  if (!acc) {
    throw new Error('FALLBACK_ACCOUNT_KEY "' + key + '" không có trong sheet Grey (hoặc đang DISABLED / thiếu Email).');
  }
  return acc;
}


/************************************************************
 * COLLECT
 ************************************************************/

function ggCollect_(txSheet, txHeaders, greyAccounts, fallback, lookbackDays, markSyncing) {
  const index = ggBuildIndex_(txSheet, txHeaders);
  const since = lookbackDays > 0 ? ' newer_than:' + lookbackDays + 'd' : '';

  const stats = {
    scanned: 0, transactionMails: 0, ignored: 0,
    created: 0, updated: 0, unchanged: 0, noReference: 0,
    unmapped: 0, ambiguous: 0, forwarded: 0, perAccount: {},
    /* Tách theo lượt, nếu không thì một lượt về tay trắng nhìn y hệt lượt kia
       làm hết việc — và không có cách nào biết lượt nào hỏng. */
    perPass: {}
  };

  /*
   * Cùng một email không được xử lý hai lần dù lọt vào cả hai lượt. Hai truy
   * vấn hiện đã loại trừ nhau, nhưng đây là thứ không nên phụ thuộc vào cú
   * pháp tìm kiếm của Gmail.
   */
  const seen = {};

  /*
   * LƯỢT 1 — mail Grey gửi thẳng, hoặc chuyển tiếp TỰ ĐỘNG.
   *
   * Chuyển tiếp tự động gửi lại nguyên mail: From vẫn là Grey, To vẫn là địa
   * chỉ account gốc, chỉ thêm Delivered-To/X-Forwarded-To. Nên nó đi đúng
   * đường này, không cần gì đặc biệt.
   */
  ggCollectPass_({
    label: 'direct',
    query: 'from:' + GG_CFG.SENDER + since,
    isGrey: function (message) {
      return String(message.getFrom() || '').toLowerCase()
        .indexOf(GG_CFG.SENDER) !== -1;
    },
    resolve: function (message) {
      return ggResolveGreyAccount_(message, greyAccounts);
    },
    cleanSubject: function (subject) { return subject; }
  }, txSheet, txHeaders, index, greyAccounts, fallback, markSyncing, stats, seen);

  /*
   * LƯỢT 2 — chuyển tiếp TAY.
   *
   * Bấm nút Forward tạo ra một mail mới: From thành người forward, và địa chỉ
   * Grey gốc chỉ còn trong khối trích dẫn ở thân mail. Cả hai thứ lượt 1 dựa
   * vào đều mất.
   *
   * Cần đường này cho mail cũ ở những hộp thư mình không có quyền chạy script
   * — chuyển tiếp tự động chỉ áp dụng cho mail đến sau khi bật.
   */
  if (GG_CFG.READ_FORWARDED) {
    ggCollectPass_({
      label: 'forwarded',
      /*
       * Chỉ loại mail Grey gửi thẳng — lượt 1 đã lo rồi. Ngoài ra không thu
       * hẹp gì thêm.
       *
       * Từng có chế độ tìm 'hello@grey.co' trong nội dung cho gọn. Bỏ, vì nó
       * dựa vào cách Gmail đánh chỉ mục địa chỉ email: chạy thật thì truy vấn
       * đó trả về vài mail không liên quan và KHÔNG trả về mail forward nào.
       * Một tuỳ chọn mà một giá trị của nó âm thầm làm mất dữ liệu thì không
       * phải tuỳ chọn.
       *
       * Hộp thư quá lớn để quét hết thì tắt READ_FORWARDED, đừng thu hẹp bằng
       * một truy vấn không đáng tin.
       */
      query: '-from:' + GG_CFG.SENDER + since,
      isGrey: function (message, bodyText) {
        return ggForwardHeaderBlock_(bodyText)
          .toLowerCase().indexOf(GG_CFG.SENDER) !== -1;
      },
      resolve: function (message, bodyText) {
        return ggResolveFromText_(ggForwardHeaderBlock_(bodyText), greyAccounts);
      },
      cleanSubject: ggStripForwardPrefix_,
      countAs: 'forwarded'
    }, txSheet, txHeaders, index, greyAccounts, fallback, markSyncing, stats, seen);
  }

  return stats;
}


function ggCollectPass_(pass, txSheet, txHeaders, index, greyAccounts, fallback, markSyncing, stats, seen) {
  Logger.log('Gmail query [' + pass.label + ']: ' + pass.query);

  let start = 0;
  while (start < GG_CFG.MAX_THREADS) {
    const threads = GmailApp.search(pass.query, start, GG_CFG.BATCH_SIZE);
    if (!threads.length) break;

    threads.forEach(function (thread) {
      thread.getMessages().forEach(function (message) {
        const messageId = String(message.getId());
        if (seen[messageId]) return;

        const rawSubject = String(message.getSubject() || '').trim();
        const bodyText = ggNormalizeText_(
          (message.getPlainBody() || '') + '\n' + ggHtmlToText_(message.getBody() || '')
        );

        const tally = stats.perPass[pass.label] ||
          (stats.perPass[pass.label] = { matched: 0, notGrey: 0, scanned: 0, ignored: 0, kept: 0 });

        tally.matched++;

        if (!pass.isGrey(message, bodyText)) {
          tally.notGrey++;
          /*
           * Cửa này im lặng là chỗ khó chẩn đoán nhất: mail bị loại mà không
           * biết vì nó thật sự không liên quan, hay vì phần cắt khối hỏng.
           */
          if (pass.label === 'forwarded' && tally.notGrey <= 3) {
            Logger.log('  loại [forwarded]: "' + ggTruncate_(rawSubject, 60) +
              '" · khối cắt ra ' +
              JSON.stringify(ggTruncate_(ggForwardHeaderBlock_(bodyText).replace(/\s+/g, ' '), 160)));
          }
          return;
        }

        seen[messageId] = true;
        stats.scanned++;
        tally.scanned++;

        const subject = pass.cleanSubject(rawSubject);

        if (!ggIsTransactionMail_(subject, bodyText)) {
          stats.ignored++;
          tally.ignored++;
          /* Vài mẫu đầu là đủ để thấy đang loại nhầm hay loại đúng. */
          if (tally.ignored <= 3) {
            Logger.log('  bỏ qua [' + pass.label + ']: không phải mail giao dịch — "' +
              ggTruncate_(subject, 70) + '"');
          }
          return;
        }
        stats.transactionMails++;
        tally.kept++;

        /*
         * Account đến từ địa chỉ người nhận, không từ config. Nhiều khớp thì
         * bỏ qua chứ không đoán: đây là quy kết tiền theo account, gán sai
         * còn tệ hơn bỏ sót.
         */
        const resolved = pass.resolve(message, bodyText);
        let account = resolved.account;

        if (resolved.status === 'AMBIGUOUS') {
          stats.ambiguous++;
          Logger.log('AMBIGUOUS [' + pass.label + ']: khớp nhiều account ' +
            resolved.matches.join(', ') + ' — subject "' + ggTruncate_(subject, 80) + '"');
          return;
        }

        if (!account) {
          if (!fallback) {
            stats.unmapped++;
            Logger.log('UNMAPPED [' + pass.label + ']: không tra được account — subject "' +
              ggTruncate_(subject, 80) + '" · nguồn tra ' + ggTruncate_(resolved.debug, 200));
            return;
          }
          account = fallback;
        }

        markSyncing(account);

        const raw = ggExtractRaw_(bodyText);
        if (!raw.reference) stats.noReference++;
        if (pass.countAs) stats[pass.countAs]++;

        const outcome = ggUpsertRow_(txSheet, txHeaders, index, {
          account: account,
          messageId: messageId,
          subject: subject,
          gmailDate: message.getDate(),
          raw: raw
        });

        stats[outcome]++;
        const per = stats.perAccount[account.accountKey] ||
          (stats.perAccount[account.accountKey] = { created: 0, updated: 0, unchanged: 0 });
        per[outcome]++;
      });
    });

    start += threads.length;
    if (threads.length < GG_CFG.BATCH_SIZE) break;
  }
}


/*
 * Khối trích dẫn đầu một mail forward tay, ví dụ:
 *
 *   ---------- Forwarded message ---------
 *   Từ: Grey <hello@grey.co>
 *   Date: ...
 *   Subject: Wallet Funding Successful
 *   To: <hangntb7@gmail.com>
 *
 * Cắt tại dòng trống đầu tiên. Giữ đúng kỷ luật của bản đọc header thật: chỉ
 * dò địa chỉ trong khối này, KHÔNG dò cả thân mail — thân mail giao dịch Grey
 * có chứa địa chỉ email và sẽ khớp nhầm.
 *
 * Không parse theo nhãn. Gmail dịch nhãn theo ngôn ngữ người forward — bản
 * tiếng Việt ghi "Từ:" thay cho "From:" — nên chỉ hỏi "địa chỉ Grey nào xuất
 * hiện ở đây", y như cách làm với header thật. Miễn nhiễm ngôn ngữ.
 */
function ggForwardHeaderBlock_(bodyText) {
  const text = String(bodyText || '');

  const marker = text.search(/-{3,}\s*forwarded message\s*-{3,}/i);

  /*
   * Không thấy dấu hiệu forward thì lùi về vùng quanh địa chỉ Grey. Vẫn là
   * một cửa sổ hẹp chứ không phải cả thân mail, nên không mở đường cho khớp
   * nhầm; chỉ là bớt cứng nhắc khi Gmail đổi cách hiển thị.
   */
  const anchor = marker !== -1
    ? marker
    : text.toLowerCase().indexOf(GG_CFG.SENDER);

  if (anchor === -1) return '';

  const from = Math.max(0, anchor - (marker !== -1 ? 0 : 200));
  const rest = text.substring(from);

  const blank = rest.search(/\n[ \t]*\n/);
  const end = blank !== -1 ? blank : Math.min(rest.length, 800);

  return rest.substring(0, end);
}


function ggStripForwardPrefix_(subject) {
  return String(subject || '').replace(/^\s*(fwd?|vs|tr)\s*:\s*/i, '').trim();
}


/************************************************************
 * IDENTITY
 *
 * Reference là danh tính thật của giao dịch Grey. Cùng một
 * giao dịch về nhiều hộp thư vẫn cho cùng Reference, nên đây
 * là thứ duy nhất gộp được chúng lại.
 *
 * Không có Reference thì lùi về message ID: vẫn ổn định trong
 * phạm vi một hộp thư, chỉ là không gộp được liên hộp thư.
 ************************************************************/

function ggUniqueKey_(reference, accountKey, messageId) {
  const ref = ggString_(reference).toUpperCase();
  if (ref) return 'GREY|REF|' + ref;
  return 'GREY|MSG|' + accountKey + '|' + messageId;
}


/*
 * Ba đường tra để tìm dòng đã tồn tại:
 *   byRef  — mạnh nhất, gộp được liên hộp thư
 *   byKey  — Unique Key nguyên văn
 *   byMsg  — Source ID, bắt được cả dòng do bản collector cũ ghi
 *
 * Nhờ byMsg mà dòng cũ được nâng cấp tại chỗ sang khoá Reference
 * thay vì sinh thêm một dòng trùng.
 */
function ggBuildIndex_(sheet, headers) {
  const index = { byRef: {}, byKey: {}, byMsg: {}, columnCount: sheet.getLastColumn() };
  const lastRow = sheet.getLastRow();
  if (lastRow < GG_CFG.DATA_START_ROW) return index;

  const values = sheet.getRange(
    GG_CFG.DATA_START_ROW, 1,
    lastRow - GG_CFG.DATA_START_ROW + 1,
    index.columnCount
  ).getValues();

  values.forEach(function (row, i) {
    const rowNumber = GG_CFG.DATA_START_ROW + i;
    const provider = ggString_(row[headers['Provider']]).toUpperCase();
    const source = ggString_(row[headers['Source']]).toUpperCase();

    const key = ggString_(row[headers['Unique Key']]);
    if (key) index.byKey[key] = rowNumber;

    if (provider !== 'GREY') return;

    const ref = ggString_(row[headers['Reference']]).toUpperCase();
    if (ref && !index.byRef[ref]) index.byRef[ref] = rowNumber;

    if (source === 'GREY_GMAIL') {
      const sourceId = ggString_(row[headers['Source ID']]);
      if (sourceId) index.byMsg[sourceId] = rowNumber;
    }
  });

  return index;
}


/************************************************************
 * UPSERT — chỉ append hoặc vá ô, không bao giờ ghi lại cả sheet
 ************************************************************/

function ggUpsertRow_(sheet, headers, index, ctx) {
  const ref = ggString_(ctx.raw.reference).toUpperCase();
  const uniqueKey = ggUniqueKey_(ctx.raw.reference, ctx.account.accountKey, ctx.messageId);

  const collected = {
    'Reference': ctx.raw.reference,
    'Raw Event': ctx.subject,
    'Raw Transaction Type': ctx.raw.transactionType,
    'Raw Amount Tendered': ctx.raw.amountTendered,
    'Raw Amount Received': ctx.raw.amountReceived,
    'Raw Recipient': ctx.raw.recipient,
    'Raw Channel': ctx.raw.channel,
    'Raw Date Time': ctx.raw.dateTime
  };

  const rowNumber =
    (ref && index.byRef[ref]) ||
    index.byKey[uniqueKey] ||
    index.byMsg[ctx.messageId] ||
    0;

  /* ---------- Dòng mới ---------- */
  if (!rowNumber) {
    const row = new Array(index.columnCount).fill('');
    const put = function (header, value) {
      if (Object.prototype.hasOwnProperty.call(headers, header)) {
        row[headers[header]] = value;
      }
    };

    put('Provider', 'GREY');
    put('Account Key', ctx.account.accountKey);
    put('Label', ctx.account.label);
    /* Chỉ đặt lúc tạo. Flow sẽ thay bằng Raw Date Time nếu parse được. */
    put('Transaction Date', ctx.gmailDate);
    put('Source', 'GREY_GMAIL');
    put('Source ID', ctx.messageId);
    put('Sync Time', new Date());
    put('Unique Key', uniqueKey);
    Object.keys(collected).forEach(function (h) { put(h, collected[h]); });

    sheet.appendRow(row);

    const newRow = sheet.getLastRow();
    if (ref) index.byRef[ref] = newRow;
    index.byKey[uniqueKey] = newRow;
    index.byMsg[ctx.messageId] = newRow;
    return 'created';
  }

  /* ---------- Dòng đã có ---------- */
  const current = sheet.getRange(rowNumber, 1, 1, index.columnCount).getValues()[0];

  const changes = [];
  GG_UPDATABLE.forEach(function (header) {
    if (!Object.prototype.hasOwnProperty.call(headers, header)) return;
    const now = ggString_(collected[header]);
    /* Đừng xoá giá trị đã có bằng một lần bóc tách hụt. */
    if (!now) return;
    if (ggString_(current[headers[header]]) === now) return;
    changes.push({ col: headers[header] + 1, value: collected[header] });
  });

  /* Nâng dòng cũ lên khoá Reference để lần sau gộp được. */
  if (ggString_(current[headers['Unique Key']]) !== uniqueKey) {
    changes.push({ col: headers['Unique Key'] + 1, value: uniqueKey });
  }

  /*
   * Quét lại cùng một email mà không có gì đổi thì không ghi gì cả.
   * Đây là thứ giữ cho lookback hằng ngày gần như miễn phí.
   */
  if (!changes.length) return 'unchanged';

  changes.push({ col: headers['Sync Time'] + 1, value: new Date() });
  changes.forEach(function (c) { sheet.getRange(rowNumber, c.col).setValue(c.value); });

  if (ref) index.byRef[ref] = rowNumber;
  index.byKey[uniqueKey] = rowNumber;
  return 'updated';
}


/************************************************************
 * ADVISORY LOCK
 *
 * ScriptLock không xuyên được giữa các Apps Script project,
 * nên khoá phải nằm trong chính spreadsheet dùng chung.
 *
 * Ghi token rồi đọc lại để xác nhận mình thắng. Vẫn còn một
 * khe hở rất nhỏ, nhưng đủ để các bên không giẫm nhau.
 *
 * syncFlow() bên Payment All dùng CÙNG giao thức này trên
 * cùng sheet và cùng ô, nên collector và Flow loại trừ nhau.
 * Đổi tên sheet, đổi ô, hay đổi cách ghi token ở một bên là
 * phá vỡ điều đó trong im lặng — hai bên sẽ tưởng mình đang
 * giữ khoá trong khi bên kia vẫn chạy.
 ************************************************************/

function ggAcquireLock_(ss) {
  let sheet = ss.getSheetByName(GG_CFG.LOCK_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(GG_CFG.LOCK_SHEET);
    sheet.getRange('A1').setValue('holder');
    sheet.getRange('B1').setValue('since');
    sheet.hideSheet();
  }

  const token = (ggString_(GG_CFG.FALLBACK_ACCOUNT_KEY) || 'GREY') + '|' + Utilities.getUuid();
  const deadline = new Date().getTime() + GG_CFG.LOCK_WAIT_MS;

  while (new Date().getTime() < deadline) {
    const holder = ggString_(sheet.getRange('A2').getValue());
    const since = sheet.getRange('B2').getValue();
    const sinceMs = since instanceof Date ? since.getTime() : 0;
    const stale = !sinceMs || (new Date().getTime() - sinceMs > GG_CFG.LOCK_STALE_MS);

    if (!holder || stale) {
      sheet.getRange('A2').setValue(token);
      sheet.getRange('B2').setValue(new Date());
      SpreadsheetApp.flush();
      Utilities.sleep(1200);
      if (ggString_(sheet.getRange('A2').getValue()) === token) {
        return { sheet: sheet, token: token };
      }
    }

    Utilities.sleep(3000);
  }

  throw new Error('Không lấy được sync lock sau ' +
    (GG_CFG.LOCK_WAIT_MS / 1000) + 's — có collector khác đang chạy. Thử lại sau.');
}


function ggReleaseLock_(lock) {
  if (!lock) return;
  try {
    if (ggString_(lock.sheet.getRange('A2').getValue()) === lock.token) {
      lock.sheet.getRange('A2:B2').clearContent();
      SpreadsheetApp.flush();
    }
  } catch (ignored) { /* lock tự hết hạn sau LOCK_STALE_MS */ }
}


/************************************************************
 * PHÂN LOẠI MAIL — chỉ để loại mail quảng cáo / bảo mật
 ************************************************************/

function ggIsTransactionMail_(subject, body) {
  const text = (String(subject || '') + '\n' + String(body || '')).toLowerCase();

  if (text.indexOf('wallet has been funded') !== -1) return true;
  if (text.indexOf('wallet funding successful') !== -1) return true;
  if (text.indexOf('withdrawal successful') !== -1) return true;
  if (text.indexOf('withdrawal was successful') !== -1) return true;

  return text.indexOf('transaction type') !== -1 &&
    (text.indexOf('amount tendered') !== -1 || text.indexOf('amount received') !== -1);
}


/************************************************************
 * BÓC TÁCH THÔ — không parse số, không parse tiền tệ
 ************************************************************/

function ggExtractRaw_(bodyText) {
  return {
    transactionType: ggFindGreyField_(bodyText, 'Transaction Type'),
    amountTendered: ggFindGreyField_(bodyText, 'Amount tendered'),
    amountReceived: ggFindGreyField_(bodyText, 'Amount received'),
    recipient: ggFindGreyField_(bodyText, 'Recipient'),
    channel: ggFindGreyField_(bodyText, 'Channel'),
    reference: ggFindGreyField_(bodyText, 'Reference'),
    dateTime: ggFindGreyField_(bodyText, 'Date & Time')
  };
}


function ggGreyKnownLabels_() {
  /*
   * Danh sách này là RANH GIỚI cắt giá trị, không phải danh sách trường cần
   * lấy. Thiếu một nhãn ở đây là giá trị đứng trước nó bị nuốt sang — 'Sender'
   * xuất hiện trong mail funding và từng thiếu, khiến 'Amount received' chỉ
   * đúng nhờ may mắn có xuống dòng cắt hộ.
   */
  return ['Transaction Type', 'Amount tendered', 'Amount received', 'Network',
    'Phone number', 'Memo', 'Sender', 'Recipient', 'Channel', 'Reference', 'Date & Time'];
}


/*
 * Gmail hay ép HTML thành một dòng dài:
 *   Amount received: USDT499.00 Network: BEP20 Reference: ABC123
 * nên giá trị phải cắt tại nhãn kế tiếp, không dựa vào xuống dòng.
 */
function ggFindGreyField_(text, label) {
  const source = String(text || '')
    .replace(/ /g, ' ')
    .replace(/\r/g, '\n')
    .replace(/[ \t]+/g, ' ');

  const lowerSource = source.toLowerCase();
  const lowerLabel = String(label || '').toLowerCase();

  const labelIndex = lowerSource.indexOf(lowerLabel);
  if (labelIndex === -1) return '';

  let valueStart = labelIndex + label.length;
  while (valueStart < source.length) {
    const ch = source.charAt(valueStart);
    if (ch === ':' || ch === '-' || ch === ' ' || ch === '\t' || ch === '\n') { valueStart++; continue; }
    break;
  }

  let valueEnd = source.length;
  ggGreyKnownLabels_().forEach(function (next) {
    if (String(next).toLowerCase() === lowerLabel) return;
    const at = lowerSource.indexOf(next.toLowerCase(), valueStart);
    if (at !== -1 && at < valueEnd) valueEnd = at;
  });

  let value = source.substring(valueStart, valueEnd);
  const nl = value.indexOf('\n');
  if (nl !== -1) value = value.substring(0, nl);

  return String(value || '').replace(/\s+/g, ' ').trim();
}


/************************************************************
 * SHEET HELPERS
 ************************************************************/

function ggEnsureRawHeaders_(sheet) {
  let lastColumn = Math.max(1, sheet.getLastColumn());
  const existing = sheet.getRange(GG_CFG.HEADER_ROW, 1, 1, lastColumn)
    .getValues()[0].map(function (v) { return String(v || '').trim(); });

  let added = 0;
  GG_RAW_HEADERS.forEach(function (header) {
    if (existing.indexOf(header) !== -1) return;
    lastColumn++;
    sheet.getRange(GG_CFG.HEADER_ROW, lastColumn).setValue(header);
    existing.push(header);
    added++;
  });

  if (added) SpreadsheetApp.flush();
}


/************************************************************
 * NHẬN DIỆN GREY ACCOUNT
 *
 * Port từ bản mail.js cũ (commit 4e46b08), đã chạy thật với
 * mail Grey. Đối chiếu email người nhận với cột Email của
 * sheet Grey.
 ************************************************************/

function ggLoadGreyAccounts_(sheet, headers) {
  const state = { accounts: [], byKey: {}, byEmail: {} };
  const lastRow = sheet.getLastRow();
  if (lastRow < GG_CFG.DATA_START_ROW) return state;

  const rows = sheet.getRange(
    GG_CFG.DATA_START_ROW, 1,
    lastRow - GG_CFG.DATA_START_ROW + 1,
    sheet.getLastColumn()
  ).getValues();

  rows.forEach(function (row, i) {
    const accountKey = ggString_(row[headers['Account Key']]);
    if (!accountKey) return;
    if (ggString_(row[headers['Status']]).toUpperCase() === 'DISABLED') return;

    const email = ggNormalizeEmail_(row[headers['Email']]);
    if (!email) {
      /* Không có Email thì không bao giờ tra ra account này. */
      Logger.log('CẢNH BÁO: Grey account "' + accountKey + '" đang bật nhưng thiếu cột Email — mail của nó sẽ không tra được.');
      return;
    }

    /* Hai account dùng chung một email thì mọi mail đều thành AMBIGUOUS. */
    if (state.byEmail[email]) {
      throw new Error('Hai Grey account dùng chung Email "' + email + '": ' +
        state.byEmail[email].accountKey + ' và ' + accountKey + '.');
    }

    const account = {
      rowNumber: GG_CFG.DATA_START_ROW + i,
      accountKey: accountKey,
      label: ggString_(row[headers['Label']]),
      email: email
    };

    state.accounts.push(account);
    state.byKey[accountKey] = account;
    state.byEmail[email] = account;
  });

  return state;
}


/*
 * Gom mọi chỗ có thể chứa địa chỉ người nhận gốc. Với mail được
 * forward, To là hộp thư trung tâm còn Delivered-To / X-Forwarded-To
 * mới giữ địa chỉ Grey account ban đầu.
 *
 * Chỉ đọc khối header, KHÔNG đọc body: body mail giao dịch Grey có
 * thể chứa địa chỉ email và sẽ khớp nhầm.
 */
function ggResolveGreyAccount_(message, greyAccounts) {
  const parts = [];

  parts.push(ggSafeStringCall_(function () { return message.getTo(); }));
  parts.push(ggSafeStringCall_(function () { return message.getCc(); }));
  parts.push(ggSafeStringCall_(function () { return message.getBcc(); }));

  ['To', 'Cc', 'Delivered-To', 'X-Original-To', 'Envelope-To', 'X-Envelope-To',
   'X-Forwarded-To', 'Resent-To', 'Original-Recipient', 'Final-Recipient'
  ].forEach(function (name) {
    parts.push(ggSafeStringCall_(function () { return message.getHeader(name); }));
  });

  parts.push(ggGetRawHeaderBlock_(message));

  return ggResolveFromText_(parts.filter(Boolean).join('\n'), greyAccounts);
}


/*
 * Dò địa chỉ Grey account trong một đoạn văn bản đã được giới hạn sẵn — khối
 * header thật, hoặc khối trích dẫn của mail forward tay.
 *
 * Người gọi chịu trách nhiệm đưa vào đúng phạm vi. Đưa cả thân mail vào là
 * mở đường cho khớp nhầm, vì thân mail giao dịch Grey có chứa địa chỉ email.
 */
function ggResolveFromText_(headerText, greyAccounts) {
  const matches = greyAccounts.accounts.filter(function (acc) {
    return ggHeaderContainsEmail_(headerText, acc.email);
  });

  if (matches.length === 1) {
    return { status: 'OK', account: matches[0], matches: [matches[0].email], debug: '' };
  }

  if (matches.length > 1) {
    return {
      status: 'AMBIGUOUS', account: null,
      matches: matches.map(function (m) { return m.email; }), debug: ''
    };
  }

  return {
    status: 'UNMAPPED', account: null, matches: [],
    debug: ggTruncate_(String(headerText || '').replace(/\s+/g, ' '), 500)
  };
}


/* Chỉ phần header RFC822, cắt tại dòng trống đầu tiên. */
function ggGetRawHeaderBlock_(message) {
  try {
    const raw = String(message.getRawContent() || '');
    if (!raw) return '';

    let at = raw.indexOf('\r\n\r\n');
    let sep = 4;
    if (at === -1) { at = raw.indexOf('\n\n'); sep = 2; }

    /* Không tìm thấy ranh giới thì cắt cứng, đừng nuốt cả body. */
    if (at === -1) return raw.substring(0, Math.min(raw.length, 12000));

    return raw.substring(0, at + sep);
  } catch (error) {
    return '';
  }
}


/*
 * Ràng buộc biên hai đầu để foo@x.com không khớp xfoo@x.com.
 */
function ggHeaderContainsEmail_(headerText, email) {
  const target = String(email || '').toLowerCase().trim();
  if (!target) return false;

  const regex = new RegExp(
    '(^|[^a-z0-9._%+\\-])' + ggRegexEscape_(target) + '([^a-z0-9._%+\\-]|$)', 'i');

  return regex.test(String(headerText || '').toLowerCase());
}


/* Nhận cả "user@x.com" lẫn "Name <user@x.com>". */
function ggNormalizeEmail_(value) {
  const text = ggString_(value).toLowerCase();
  if (!text) return '';
  const m = text.match(/[a-z0-9.!#$%&'*+/=?^_`{|}~-]+@[a-z0-9.-]+\.[a-z]{2,}/i);
  return m ? m[0].toLowerCase().trim() : '';
}


function ggRegexEscape_(value) {
  return String(value || '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}


/* getHeader() không có ở mọi phiên bản runtime. */
function ggSafeStringCall_(fn) {
  try {
    return String(fn() || '');
  } catch (error) {
    return '';
  }
}


/* Header map 0-based: dùng row[headers[h]], ghi thì headers[h] + 1. */
function ggHeaderMap_(sheet) {
  const row = sheet.getRange(GG_CFG.HEADER_ROW, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  row.forEach(function (value, index) {
    const header = ggString_(value);
    if (!header) return;
    if (Object.prototype.hasOwnProperty.call(map, header)) {
      throw new Error('Header trùng "' + header + '" trong sheet "' + sheet.getName() + '".');
    }
    map[header] = index;
  });
  return map;
}


function ggRequireHeaders_(map, required, sheetName) {
  const missing = required.filter(function (h) {
    return !Object.prototype.hasOwnProperty.call(map, h);
  });
  if (missing.length) {
    throw new Error('Sheet "' + sheetName + '" thiếu header: ' + missing.join(', '));
  }
}


function ggSetField_(sheet, rowNumber, headers, header, value) {
  if (!Object.prototype.hasOwnProperty.call(headers, header)) return;
  sheet.getRange(rowNumber, headers[header] + 1).setValue(value);
}


/************************************************************
 * TEXT HELPERS
 ************************************************************/

function ggNormalizeText_(text) {
  return String(text || '')
    .replace(/\r/g, '\n')
    .replace(/ /g, ' ')
    .replace(/[ \t]+/g, ' ')
    .replace(/\n[ \t]+/g, '\n')
    .replace(/\n{3,}/g, '\n\n');
}


function ggHtmlToText_(html) {
  return String(html || '')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/(p|div|td|tr|li|table|h1|h2|h3)>/gi, '\n')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/&#39;/gi, "'")
    .replace(/&quot;/gi, '"');
}


function ggString_(value) {
  return String(value === null || value === undefined ? '' : value).trim();
}


function ggTruncate_(value, maxLength) {
  const text = String(value || '');
  return text.length <= maxLength ? text : text.substring(0, maxLength);
}
