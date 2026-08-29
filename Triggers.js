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
  syncFlow();
  syncCustomers();
  Logger.log(
    'TRIGGER SYNC DONE in ' +
    ((new Date().getTime() - started.getTime()) / 1000).toFixed(1) + 's'
  );
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


/* Xem lịch đang chạy mà không đụng gì. */
function listPaymentAllTriggers() {
  const all = ScriptApp.getProjectTriggers();
  if (!all.length) { Logger.log('Chưa có trigger nào.'); return []; }
  const out = all.map(function (t) { return t.getHandlerFunction(); });
  Logger.log('Trigger hiện có: ' + out.join(', '));
  return out;
}
