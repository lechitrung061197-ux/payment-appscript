/**
 * ============================================================================
 * TRIGGERS.GS — PAYMENT ALL
 * ============================================================================
 * Lịch chạy tự động cho project Payment All.
 *
 * KHÔNG bao gồm collector Grey: nó chạy ở các project Gmail riêng và có lịch
 * riêng trong standalone/GreyGmailCollector.gs.
 *
 * KHÔNG BAO GIỜ tự động tạo invoice. Chuỗi dừng ở syncCustomers(). Toàn bộ lớp
 * an toàn preview-trước-create dựa trên việc có người nhìn số trước khi bấm.
 *
 * ==========================================================================
 * VÌ SAO CHIA LÀM HAI MỐC
 * ==========================================================================
 * Apps Script giết execution ở 6 phút. Gộp cả bốn việc vào một lần chạy là tự
 * đẩy mình tới giới hạn đó. Tách làm hai, cách nhau một tiếng:
 *
 *   mốc 1: syncPayPalAllAccounts()          — gọi mạng, chậm nhất
 *   mốc 2: syncFlow() rồi syncCustomers()   — chỉ đọc/ghi sheet
 *
 * Nếu mốc 1 hỏng thì mốc 2 vẫn chạy trên dữ liệu cũ, không kẹt.
 *
 * ==========================================================================
 * XẾP LỆCH VỚI COLLECTOR GREY
 * ==========================================================================
 * syncFlow() ghi đè toàn bộ vùng dữ liệu sheet Transaction. Nó có tham gia
 * advisory lock chung với collector, nhưng lock chỉ tránh giẫm nhau chứ không
 * làm hai bên nhanh hơn — đụng nhau nghĩa là một bên phải chờ tới 90 giây rồi
 * có thể bỏ lượt. Nên vẫn đặt giờ lệch hẳn khỏi các project collector.
 *
 * atHour(h) chạy đâu đó TRONG khoảng một tiếng đó, không đúng phút.
 * ============================================================================
 */

const TRIG_CFG = Object.freeze({
  TIMEZONE: 'Asia/Ho_Chi_Minh',

  /* Hai lần mỗi ngày. Mỗi phần tử là một mốc: collector chạy giờ đầu,
     sync chạy giờ sau. Collector Grey nên nằm ở các giờ 1–3. */
  RUNS: [
    { collect: 6, sync: 7 },
    { collect: 18, sync: 19 }
  ],

  COLLECT_HANDLER: 'triggerCollectPayPal',
  SYNC_HANDLER: 'triggerSyncFlowAndCustomers'
});


/*********************** HÀM CHẠY ************************/

function triggerCollectPayPal() {
  return syncPayPalAllAccounts();
}


/*
 * Flow phải xong trước Customer: Flow chuẩn hoá withdrawal PayPal vào
 * Transaction và đối chiếu với provider nhận, còn Customer chỉ đọc kết quả đó.
 * Chạy ngược thứ tự thì Customer phân bổ trên dữ liệu của hôm qua.
 */
function triggerSyncFlowAndCustomers() {
  const started = new Date();

  try {
    syncFlow();
    syncCustomers();
  } catch (err) {
    /*
     * Hỏng là lúc cần biết nhất. Nếu chỉ gửi báo cáo khi thành công thì đúng
     * lúc hệ thống chết lại là lúc im lặng nhất — và im lặng thì không phân
     * biệt được với "hôm nay không có gì mới".
     */
    trigNotify_('❌ PAYMENT ALL · SYNC LỖI\n\n' +
      (err && err.message ? err.message : String(err)) +
      '\n\nDữ liệu giữ nguyên ở lần sync trước. Xem Execution log.');
    throw err;
  }

  /*
   * Dashboard chỉ đọc lại kết quả của hai bước trên nên phải đứng sau, và
   * nằm NGOÀI try chính: nó hỏng thì dữ liệu vẫn đúng, không được để một
   * lượt sync đã thành công trông như thất bại.
   *
   * Nhưng vẫn phải báo — dashboard đứng im hiện số của hôm qua mà không ai
   * biết thì tệ hơn là không có dashboard.
   */
  try {
    updateDashboard();
  } catch (dashError) {
    Logger.log('Không dựng được dashboard: ' + dashError);
    trigNotify_('⚠️ PAYMENT ALL · DASHBOARD LỖI\n\n' +
      (dashError && dashError.message ? dashError.message : String(dashError)) +
      '\n\nSync vẫn thành công, số liệu trên các sheet là đúng. ' +
      'Chỉ tab Dashboard đang hiện số của lần chạy trước.');
  }

  Logger.log(
    'TRIGGER SYNC DONE in ' +
    ((new Date().getTime() - started.getTime()) / 1000).toFixed(1) + 's'
  );

  if (trigShouldReport_(started)) {
    /*
     * Báo cáo là tiện ích, không phải nhiệm vụ. Nó hỏng thì cũng không được
     * làm một lượt sync đã thành công trông như thất bại.
     */
    try {
      trigNotify_(trigBuildReport_(SpreadsheetApp.getActive()));
    } catch (reportError) {
      Logger.log('Không gửi được báo cáo: ' + reportError);
    }
  }
}


/*
 * Một tin mỗi ngày, ở lượt sync đầu tiên. Lượt tối gần như luôn giống lượt
 * sáng, mà tin nhắn lặp lại là thứ người ta bắt đầu lướt qua — lúc đó nó mất
 * luôn tác dụng cảnh báo.
 *
 * Nới ±1 giờ vì atHour() chạy đâu đó trong khoảng một tiếng, không đúng phút.
 */
function trigShouldReport_(now) {
  const hour = Number(
    Utilities.formatDate(now, TRIG_CFG.TIMEZONE, 'H')
  );
  return Math.abs(hour - TRIG_CFG.RUNS[0].sync) <= 1;
}


/*********************** CÀI ĐẶT ************************/

function setupPaymentAllTriggers() {
  const removed = removePaymentAllTriggers();

  TRIG_CFG.RUNS.forEach(function (run) {
    [[run.collect, TRIG_CFG.COLLECT_HANDLER], [run.sync, TRIG_CFG.SYNC_HANDLER]]
      .forEach(function (pair) {
        const hour = pair[0], handler = pair[1];
        if (hour < 0 || hour > 23) throw new Error('Giờ phải từ 0 đến 23: ' + hour);
        ScriptApp.newTrigger(handler)
          .timeBased()
          .atHour(hour)
          .everyDays(1)
          .inTimezone(TRIG_CFG.TIMEZONE)
          .create();
      });
  });

  const lines = TRIG_CFG.RUNS.map(function (r) {
    return '  ' + String(r.collect).padStart(2, '0') + 'h PayPal → ' +
           String(r.sync).padStart(2, '0') + 'h Flow+Customer';
  }).join('\n');

  Logger.log('Đã xoá ' + removed + ' trigger cũ. Lịch mới (' + TRIG_CFG.TIMEZONE + '):\n' + lines);
  Logger.log('Không có trigger nào tạo invoice — đó vẫn là thao tác tay.');
}


/* Xoá trước khi tạo, nếu không chạy setup hai lần là nhân đôi lịch. */
function removePaymentAllTriggers() {
  let removed = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    const fn = t.getHandlerFunction();
    if (fn === TRIG_CFG.COLLECT_HANDLER || fn === TRIG_CFG.SYNC_HANDLER) {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });
  return removed;
}


/*********************** BÁO CÁO HẰNG NGÀY ************************/
/*
 * Một tin Telegram mỗi sáng, sau khi sync xong.
 *
 * Giá trị lớn nhất của nó không phải mấy con số — bạn hỏi bot lúc nào cũng
 * được. Nó nằm ở chỗ IM LẶNG TRỞ THÀNH TÍN HIỆU: sáng nào cũng có tin, một
 * hôm không thấy là biết ngay có gì đó hỏng. Còn phải nhớ mà đi hỏi thì đúng
 * lúc hệ thống chết lại là lúc dễ quên nhất.
 *
 * Chỉ ĐỌC những gì lượt sync vừa tính xong, không tính lại. Đặc biệt không
 * gọi validateCustomerSystem() — nó chạy syncCustomers_() rồi dựng lại toàn
 * bộ, tức là làm lại y nguyên công việc vừa xong. Các phép kiểm đắt như
 * custReportUnattributedSales_ để dành cho lúc chạy validate bằng tay.
 */

const TRIG_REPORT = Object.freeze({
  /* Collector Grey chạy hằng ngày, nên im quá 2 ngày là đáng ngờ. */
  GREY_STALE_DAYS: 2,
  MAX_ISSUES: 8
});


/* Chạy tay để xem thử tin nhắn trông thế nào, không cần đợi tới sáng. */
function sendDailyReport() {
  const text = trigBuildReport_(SpreadsheetApp.getActive());
  Logger.log(text);
  trigNotify_(text);
  return text;
}


function trigBuildReport_(ss) {
  const lines = [];
  const issues = [];

  lines.push('📊 PAYMENT ALL · ' +
    Utilities.formatDate(new Date(), TRIG_CFG.TIMEZONE, 'dd/MM/yyyy'));
  lines.push('');

  /********** Khách hàng **********/

  const shCustomers = ss.getSheetByName(CUST_CFG.CUSTOMERS_SHEET);
  if (shCustomers) {
    const h = custHeaderMap_(shCustomers);
    custReadRowsObjects_(shCustomers, h).forEach(function (c) {
      const key = custStr_(c['Customer Key']);
      if (!key) return;
      if (custUpper_(c['Status']) === 'DISABLED') return;

      /*
       * Tên dễ theo dõi hơn CUS001. Lùi về khoá khi chưa đặt tên, vì một dòng
       * không có nhãn nào thì tệ hơn hẳn một dòng mang khoá.
       */
      const who = custStr_(c['Customer Name']) || key;

      lines.push(who + ' — chưa invoice ' + trigMoney_(c['Outstanding Amount']) +
        ' · đang về ' + trigMoney_(c['Pending in PayPal']));

      /* Hai khoản khách nợ lại: nêu riêng chứ đừng trộn vào dòng trên. */
      const advance = custNum_(c['Advance Outstanding']);
      if (advance > CUST_CFG.EPSILON) {
        issues.push(who + ' còn nợ ứng ' + trigMoney_(advance));
      }
      const owed = custNum_(c['Refund Owed']);
      if (owed > CUST_CFG.EPSILON) {
        issues.push(who + ' nợ lại do refund muộn ' + trigMoney_(owed));
      }
    });
  }

  /********** Ngân quỹ **********/

  lines.push('');

  const shUsdt = ss.getSheetByName(FLOW_CFG.USDT_SHEET);
  if (shUsdt) {
    /* Dòng tổng do syncFlow vừa ghi vào hàng ghi chú. */
    const total = custStr_(shUsdt.getRange(1, 1).getValue());
    const m = total.match(/CÒN LẠI\s+([\d.]+)/);
    /* Flow ghi số thô; đưa qua cùng bộ định dạng để cả tin nhắn đọc một kiểu. */
    lines.push('USDT trong hệ thống: ' +
      (m ? trigMoney_(m[1]) : '(chưa có, chạy syncFlow)'));
  }

  const draft = trigDraftPayoutTotal_(ss);
  if (draft > CUST_CFG.EPSILON) {
    lines.push('Payout DRAFT chờ trả: ' + trigMoney_(draft));
  }

  /********** Những thứ cần người nhìn **********/

  trigCollectStuckMoney_(ss, issues);
  trigCollectOverdueWithdrawals_(ss, issues);
  trigCollectForeignCurrency_(ss, issues);
  trigCollectUnknownEvents_(ss, issues);
  trigCollectDraftOverUsdt_(ss, issues);
  trigCollectPayoutCategory_(ss, issues);
  trigCollectOwnerOverdraw_(ss, issues);
  trigCollectStaleGrey_(ss, issues);

  lines.push('');
  if (!issues.length) {
    lines.push('✓ Không có gì bất thường');
  } else {
    lines.push('⚠ CẦN XEM');
    issues.slice(0, TRIG_REPORT.MAX_ISSUES).forEach(function (issue) {
      lines.push('• ' + issue);
    });
    if (issues.length > TRIG_REPORT.MAX_ISSUES) {
      lines.push('• … và ' + (issues.length - TRIG_REPORT.MAX_ISSUES) +
        ' mục nữa — chạy validateCustomerSystem()');
    }
  }

  return lines.join('\n');
}


/* Tiền đã duyệt nhưng chưa trả — nghĩa vụ đang treo. */
function trigDraftPayoutTotal_(ss) {
  const sheet = ss.getSheetByName(CUST_CFG.PAYOUTS_SHEET);
  if (!sheet) return 0;

  const h = custHeaderMap_(sheet);
  if (!h['Status'] || !h['Amount']) return 0;

  let total = 0;
  custReadRowsObjects_(sheet, h).forEach(function (p) {
    if (custUpper_(p['Status']) !== 'DRAFT') return;
    total += custNum_(p['Amount']);
  });
  return custRound2_(total);
}


/*
 * Tiền sẽ không bao giờ tự về: provider nhận chưa được hỗ trợ. Đây là lỗi cấu
 * hình chứ không phải chậm trễ, nên phải hiện thành số tiền cụ thể.
 */
function trigCollectStuckMoney_(ss, issues) {
  const sheet = ss.getSheetByName(CUST_CFG.ALLOCATION_SHEET);
  if (!sheet) return;

  const h = custHeaderMap_(sheet);
  if (!h['Allocation Status']) return;

  const stuck = {};
  custReadRowsObjects_(sheet, h).forEach(function (a) {
    if (custUpper_(a['Allocation Status']) !== 'UNSUPPORTED_PROVIDER') return;
    const provider = custUpper_(a['Receiving Provider']) || '(chưa rõ)';
    stuck[provider] = custRound2_((stuck[provider] || 0) + custNum_(a['Allocated Amount']));
  });

  Object.keys(stuck).forEach(function (provider) {
    issues.push('UNSUPPORTED_PROVIDER · ' + provider + ' ' +
      trigMoney_(stuck[provider]) + ' đang kẹt');
  });
}


/*
 * Khoản rút quá hạn mà chưa tới nơi nhận. Hai khả năng, và cả hai đều cần
 * người nhìn: hoặc nơi nhận chưa gửi mail về, hoặc PayPal đã trả tiền lại
 * theo cách hệ thống chưa nhận ra.
 *
 * Chừng nào khoản rút chưa được đánh RETURNED thì nó vẫn giữ chỗ nguồn của
 * khách, nên tiền đó không phân bổ lại được cho lần rút sau. Số kẹt chứ
 * không phải số sai — nhưng tự nó không gỡ được.
 */
function trigCollectOverdueWithdrawals_(ss, issues) {
  const sheet = ss.getSheetByName(FLOW_CFG.TRANSACTION_SHEET);
  if (!sheet) return;

  const h = custHeaderMap_(sheet);
  if (!h['Reconcile Status'] || !h['Type']) return;

  let count = 0;
  let total = 0;

  custReadRowsObjects_(sheet, h).forEach(function (t) {
    if (custUpper_(t['Type']) !== 'PAYPAL_WITHDRAWAL') return;
    if (custUpper_(t['Reconcile Status']) !== 'OVERDUE') return;
    count++;
    total += custNum_(t['Net Amount']);
  });

  if (count) {
    issues.push(count + ' khoản rút OVERDUE · ' + trigMoney_(total) +
      ' chưa tới nơi nhận');
  }
}


/*
 * Payout đã duyệt nhiều hơn USDT hệ thống ghi nhận. Không tab nào thấy được
 * điều này vì hai con số nằm ở hai bảng khác nhau.
 *
 * Chưa chắc là sai: USDT!A1 vốn ghi rõ đó là "phần hệ thống ghi nhận", ví
 * thật có thể còn dư. Nhưng nếu đúng là thiếu thì trả hết DRAFT sẽ âm quỹ,
 * nên đây là thứ phải biết TRƯỚC khi bấm trả chứ không phải sau.
 */
function trigCollectDraftOverUsdt_(ss, issues) {
  const usdt = dashUsdtRemaining_(ss);
  if (usdt === null) return;

  const draft = trigDraftPayoutTotal_(ss);
  if (draft - usdt <= CUST_CFG.EPSILON) return;

  issues.push('Payout DRAFT ' + trigMoney_(draft) + ' > USDT còn ' +
    trigMoney_(usdt) + ' · thiếu ' + trigMoney_(draft - usdt));
}


/*
 * Payout đã chi mà không phân loại. Category không tham gia phép tính số dư
 * ví, nên dòng thiếu nó vẫn trừ tiền bình thường — chỉ là không ai biết tiền
 * đó đi vào việc gì, và sổ chủ cố ý KHÔNG đoán bừa mà để riêng ra.
 *
 * Chỉ cảnh báo, không chặn: chặn nghĩa là transaction không được sinh, và số
 * dư ví đổi ngay trên tiền thật chỉ vì một ô bỏ trống.
 */
function trigCollectPayoutCategory_(ss, issues) {
  const sheet = ss.getSheetByName(CUST_CFG.PAYOUTS_SHEET);
  if (!sheet) return;

  const h = custHeaderMap_(sheet);
  if (!h['Status'] || !h['Amount'] || !h['Category']) return;

  const known = flowPayoutCategories_();
  let blank = 0, odd = 0, total = 0;
  const oddNames = {};

  custReadRowsObjects_(sheet, h).forEach(function (p) {
    if (custUpper_(p['Status']) !== 'PAID') return;

    const category = custUpper_(p['Category']);
    if (category && known.indexOf(category) !== -1) return;

    if (category) { odd++; oddNames[category] = true; } else { blank++; }
    total += custNum_(p['Amount']) + custNum_(p['Fee']);
  });

  if (!blank && !odd) return;

  const parts = [];
  if (blank) parts.push(blank + ' dòng để trống');
  if (odd) parts.push(odd + ' dòng lạ (' + Object.keys(oddNames).join(', ') + ')');

  issues.push('Payout đã chi chưa phân loại · ' + parts.join(' · ') +
    ' · ' + trigMoney_(total));
}


/*
 * Đã lấy ra khỏi ví nhiều hơn phần phí dịch vụ kiếm được.
 *
 * Đây thường là lời giải thích trực tiếp cho việc ví không đủ trả nợ khách:
 * tiền của khách đã bị tiêu sang việc khác. Không phải lỗi kỹ thuật, nhưng
 * là thứ phải biết trước khi hứa trả ai đồng nào.
 */
function trigCollectOwnerOverdraw_(ss, issues) {
  const settlements = ss.getSheetByName(CUST_CFG.SETTLEMENTS_SHEET);
  const payouts = ss.getSheetByName(CUST_CFG.PAYOUTS_SHEET);
  if (!settlements || !payouts) return;

  const sh = custHeaderMap_(settlements);
  if (!sh['Payout Fee Amount'] || !sh['Status']) return;

  let earned = 0;
  custReadRowsObjects_(settlements, sh).forEach(function (s) {
    if (!custIsEffectiveSettlementStatus_(s['Status'])) return;
    earned += custNum_(s['Payout Fee Amount']);
  });

  const ph = custHeaderMap_(payouts);
  if (!ph['Status'] || !ph['Amount'] || !ph['Category']) return;

  let drawn = 0;
  custReadRowsObjects_(payouts, ph).forEach(function (p) {
    if (custUpper_(p['Status']) !== 'PAID') return;
    if (!flowPayoutIsOwnerCost_(p['Category'])) return;
    drawn += custNum_(p['Amount']) + custNum_(p['Fee']);
  });

  const over = custRound2_(drawn - earned);
  if (over <= CUST_CFG.EPSILON) return;

  issues.push('Đã lấy ra ' + trigMoney_(drawn) + ' nhưng phí kiếm được mới ' +
    trigMoney_(earned) + ' · vượt ' + trigMoney_(over));
}


/*
 * Dòng sổ khách còn mang đơn vị ngoài USD. Bình thường PayPal đổi ngoại tệ
 * bằng một cặp T0200 và sổ ghi thẳng số USD đo được; còn ngoại tệ nghĩa là
 * KHÔNG tìm thấy cặp đó, nên con số đang là tiền gốc bị coi như USD.
 *
 * Cả chuỗi tính phía sau không có phép đổi tiền nào, nên đây là chỗ duy
 * nhất một đơn vị lạ có thể lọt vào mà vẫn trông như bình thường.
 */
function trigCollectForeignCurrency_(ss, issues) {
  const sheet = ss.getSheetByName(CUST_CFG.LEDGER_SHEET);
  if (!sheet) return;

  const h = custHeaderMap_(sheet);
  if (!h['Currency']) return;

  const seen = {};

  custReadRowsObjects_(sheet, h).forEach(function (r) {
    const cur = custUpper_(r['Currency']);
    if (!cur || cur === 'USD') return;
    if (!seen[cur]) seen[cur] = { n: 0, total: 0 };
    seen[cur].n++;
    seen[cur].total += custNum_(r['Net Amount']);
  });

  Object.keys(seen).sort().forEach(function (cur) {
    issues.push('Sổ khách còn ' + seen[cur].n + ' dòng ' + cur + ' · ' +
      trigMoney_(seen[cur].total) + ' — chưa quy đổi được sang USD');
  });
}


/*
 * Mã sự kiện PayPal chưa từng được ánh xạ. Phần lớn là vô hại, nhưng đây
 * chính là chỗ một loại giao dịch mới đi vào hệ thống mà không ai đọc —
 * nên thà kêu thừa còn hơn im lặng bỏ qua tiền.
 */
function trigCollectUnknownEvents_(ss, issues) {
  const sheet = ss.getSheetByName(FLOW_CFG.PAYPAL_SHEET);
  if (!sheet) return;

  const h = custHeaderMap_(sheet);
  if (!h['Type']) return;

  const seen = {};

  custReadRowsObjects_(sheet, h).forEach(function (row) {
    const type = custUpper_(row['Type']);
    if (type.indexOf('UNMAPPED_') !== 0) return;

    const code = type.substring('UNMAPPED_'.length);
    if (!seen[code]) seen[code] = { n: 0, total: 0 };
    seen[code].n++;
    seen[code].total += custNum_(row['Transaction Amount']);
  });

  Object.keys(seen).sort().forEach(function (code) {
    issues.push('Mã PayPal chưa biết ' + code + ' · ' + seen[code].n +
      ' dòng · ' + trigMoney_(seen[code].total));
  });
}


/*
 * Collector Grey chạy ở project khác, nên nếu nó chết thì bên này không hay
 * biết gì — dữ liệu chỉ đơn giản ngừng mới. Cột Last Sync là cách duy nhất
 * nhìn thấy điều đó từ đây.
 */
function trigCollectStaleGrey_(ss, issues) {
  const sheet = ss.getSheetByName(FLOW_CFG.GREY_SHEET);
  if (!sheet) return;

  const h = custHeaderMap_(sheet);
  if (!h['Last Sync']) return;

  const now = new Date().getTime();

  custReadRowsObjects_(sheet, h).forEach(function (row) {
    const key = custStr_(row['Account Key']);
    if (!key) return;
    if (custUpper_(row['Status']) === 'DISABLED') return;

    const last = custDate_(row['Last Sync']);
    if (!last) {
      issues.push(key + ' chưa sync lần nào');
      return;
    }

    const days = Math.floor((now - last.getTime()) / 86400000);
    if (days >= TRIG_REPORT.GREY_STALE_DAYS) {
      issues.push(key + ' chưa sync ' + days + ' ngày');
    }
  });
}


/*
 * Gửi cho đúng chủ bot. TELEGRAM_ALLOWED_USER_ID cũng chính là chat id của
 * cuộc trò chuyện riêng, nên không cần lưu thêm gì.
 *
 * Thiếu cấu hình thì ghi log rồi thôi — báo cáo không có thì tiếc, chứ không
 * đáng để ném lỗi ra ngoài.
 */
function trigNotify_(text) {
  if (typeof payTgSendLongMessage_ !== 'function') {
    Logger.log('TELEGRAM_AI.gs chưa có trong project — bỏ qua báo cáo.');
    return false;
  }

  const chatId = String(
    PropertiesService.getScriptProperties()
      .getProperty('TELEGRAM_ALLOWED_USER_ID') || ''
  ).trim();

  if (!chatId) {
    Logger.log('Chưa có TELEGRAM_ALLOWED_USER_ID — bỏ qua báo cáo.');
    return false;
  }

  payTgSendLongMessage_(chatId, text);
  return true;
}


/* 1181.35 -> "1.181,35" — đọc quen mắt hơn trong tin nhắn. */
function trigMoney_(value) {
  const n = custNum_(value);
  const parts = Math.abs(n).toFixed(2).split('.');
  const whole = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, '.');
  return (n < 0 ? '-' : '') + whole + ',' + parts[1];
}


/* Xem lịch đang chạy mà không đụng gì. */
function listPaymentAllTriggers() {
  const all = ScriptApp.getProjectTriggers();
  if (!all.length) { Logger.log('Chưa có trigger nào.'); return []; }
  const out = all.map(function (t) { return t.getHandlerFunction(); });
  Logger.log('Trigger hiện có: ' + out.join(', '));
  return out;
}
