/**
 * ============================================================================
 * DASHBOARD.GS — PAYMENT ALL
 * ============================================================================
 * Một trang trả lời hai câu hỏi hay được hỏi nhất: "tiền đang ở đâu" và "có
 * gì sai không". Báo cáo Telegram đẩy tới bạn mỗi sáng; trang này thì ngược
 * lại — mở ra xem lúc nào cũng được, qua menu Payment All.
 *
 * ==========================================================================
 * VÌ SAO LÀ DIALOG HTML CHỨ KHÔNG PHẢI MỘT TAB SHEET
 * ==========================================================================
 * Biểu đồ nhúng của Sheets chạy trên engine Google Charts đời cũ: chỉnh được
 * màu và font, nhưng khoảng trắng, độ dày, chú thích thì không. Trần thẩm mỹ
 * của nó chỉ tới đó. HTML thì toàn quyền, và số liệu vẫn sống.
 *
 * ==========================================================================
 * VÌ SAO KHÔNG DÙNG doGet
 * ==========================================================================
 * doGet đang phục vụ webhook Telegram và web app đó để công khai. Đưa số liệu
 * tài chính lên đó là hở ra ngoài. Dialog thì chỉ ai mở được sheet mới xem
 * được — đúng bằng quyền họ vốn đã có.
 *
 * ==========================================================================
 * LỊCH SỬ
 * ==========================================================================
 * `Dashboard History` ghi MỘT dòng mỗi ngày, khoá theo ngày nên chạy lại
 * trong cùng ngày là ghi đè chứ không thêm dòng. Không vẽ được xu hướng cho
 * thứ chưa từng ghi lại, nên nó bắt đầu tích từ hôm nay và phải vài tuần mới
 * đủ dày. Trang HTML cố ý KHÔNG vẽ đường khi mới có một chấm.
 * ============================================================================
 */

const DASH_CFG = Object.freeze({
  SHEET: 'Dashboard',
  HISTORY_SHEET: 'Dashboard History',
  DIALOG_FILE: 'DashboardDialog',
  DIALOG_TITLE: 'Bảng điều khiển',
  WIDTH: 1280,
  HEIGHT: 820,

  /* Giữ đúng hơn một năm. Dài hơn thế thì đọc bằng mắt không ra gì. */
  HISTORY_KEEP_DAYS: 400,
  TREND_POINTS: 90,

  MAX_DRAFTS: 15,
  MAX_ALERTS: 8,

  TIMEZONE: 'Asia/Ho_Chi_Minh'
});


/*********************** ĐIỂM VÀO ************************/

/*
 * Mở trang. Số liệu nhồi thẳng vào HTML lúc dựng template nên trang hiện ra
 * là có số ngay, không phải chờ một vòng google.script.run rồi nháy trắng.
 */
function showDashboard() {
  const payload = dashCollectMetrics_(SpreadsheetApp.getActive());

  const t = HtmlService.createTemplateFromFile(DASH_CFG.DIALOG_FILE);
  t.payload = JSON.stringify(payload);

  SpreadsheetApp.getUi().showModalDialog(
    t.evaluate().setWidth(DASH_CFG.WIDTH).setHeight(DASH_CFG.HEIGHT),
    DASH_CFG.DIALOG_TITLE
  );
}


/*
 * Chạy trong chuỗi trigger. Việc DUY NHẤT còn phải làm ngoài giờ mở trang là
 * ghi ảnh chụp hằng ngày — trang HTML dựng lúc mở nên không cần chuẩn bị gì.
 */
function updateDashboard() {
  const ss = SpreadsheetApp.getActive();
  const m = dashCollectMetrics_(ss);
  const history = dashUpsertHistory_(ss, m);
  dashWriteSignpost_(ss, m);

  Logger.log('DASHBOARD SNAPSHOT · ' + history.length + ' ngày lịch sử');
  return history.length;
}


/*********************** SỐ LIỆU ************************/

function dashCollectMetrics_(ss) {
  const now = new Date();
  const m = {
    at: Utilities.formatDate(now, DASH_CFG.TIMEZONE, 'HH:mm · dd/MM/yyyy'),
    buckets: { pending: 0, outstanding: 0, invoiced: 0, total: 0 },
    customers: [], chain: null, drafts: [], draftTotal: 0,
    unassigned: [], unassignedTotal: 0,
    paypal: [], treasury: [], health: { types: [], reconcile: [] },
    usdt: null, holdBalance: 0, feeEarned: 0, owed: 0,
    owner: {
      earned: 0, drawn: 0, remaining: 0,
      byCategory: [], byPayee: [],
      unclassified: { n: 0, total: 0 }
    },
    history: [], alerts: []
  };

  dashReadCustomers_(ss, m);
  /* Settlements phải đứng trước Payouts: sổ chủ dùng feeEarned nó tính ra. */
  dashReadSettlements_(ss, m);
  dashReadPayouts_(ss, m);
  dashFinishOwnerAccount_(m);
  dashReadAccounts_(ss, m);
  dashReadTreasury_(ss, m);
  dashReadUnassigned_(ss, m);
  dashReadHealth_(ss, m);
  dashReadHistory_(ss, m);

  /*
   * Dùng lại đúng các collector của báo cáo Telegram. Một nguồn sự thật cho
   * cảnh báo: sửa một chỗ thì cả hai nơi cùng đổi, không bao giờ lệch nhau.
   */
  [
    trigCollectStuckMoney_, trigCollectOverdueWithdrawals_,
    trigCollectForeignCurrency_, trigCollectUnknownEvents_,
    trigCollectDraftOverUsdt_, trigCollectPayoutCategory_,
    trigCollectOwnerOverdraw_, trigCollectStaleGrey_
  ].forEach(function (collect) { collect(ss, m.alerts); });

  if (m.unassignedTotal > CUST_CFG.EPSILON) {
    m.alerts.push(dashMoney_(m.unassignedTotal) + ' chưa gắn khách nào · ' +
      m.unassigned.length + ' prefix: ' +
      m.unassigned.map(function (u) { return u.prefix; }).join(' · '));
  }

  return m;
}


/*
 * Ba rổ KHÔNG chồng lấn, cộng lại đúng bằng toàn bộ tiền khách hệ thống biết
 * tới. Lưu ý `Outstanding Amount` nghĩa là available TRỪ invoiced, tức "đã về
 * nhưng chưa xuất hoá đơn" — KHÔNG phải "đã xuất hoá đơn mà chưa trả". Số nợ
 * thật nằm ở `Outstanding Payable` bên Settlements.
 */
function dashReadCustomers_(ss, m) {
  const sh = ss.getSheetByName(CUST_CFG.CUSTOMERS_SHEET);
  if (!sh) return;

  const h = custHeaderMap_(sh);
  custReadRowsObjects_(sh, h).forEach(function (c) {
    const name = custStr_(c['Customer Name']) || custStr_(c['Customer Key']);
    if (!name) return;
    if (custUpper_(c['Status']) === 'DISABLED') return;

    const row = {
      name: name,
      prefixes: custStr_(c['Prefixes']),
      pending: custNum_(c['Pending in PayPal']),
      outstanding: custNum_(c['Outstanding Amount']),
      invoiced: custNum_(c['Invoiced Amount']),
      feePct: custNum_(c['Payout Fee %']),
      advance: custNum_(c['Advance Outstanding']),
      invoices: 0, owed: 0,
      last: dashShortDate_(c['Last Settlement'])
    };
    row.total = custRound2_(row.pending + row.outstanding + row.invoiced);

    m.customers.push(row);
    m.buckets.pending += row.pending;
    m.buckets.outstanding += row.outstanding;
    m.buckets.invoiced += row.invoiced;
  });

  ['pending', 'outstanding', 'invoiced'].forEach(function (k) {
    m.buckets[k] = custRound2_(m.buckets[k]);
  });
  m.buckets.total = custRound2_(
    m.buckets.pending + m.buckets.outstanding + m.buckets.invoiced);

  m.customers.sort(function (a, b) { return b.total - a.total; });
}


/*
 * Chuỗi chi phí cộng dồn mọi hoá đơn hiệu lực.
 *
 * Hai vị trí của hoàn trả được SUY RA từ chính hai đẳng thức, không đoán:
 *
 *   hoàn trả trong kỳ   = gross − phí PayPal − phí nhận − khả dụng
 *   hoàn trả kết chuyển = khả dụng − đổi − rút − trước phí dịch vụ
 *
 * Vì hoàn trả phát sinh trong kỳ thì đã trừ vào ledger nên nằm trước "khả
 * dụng", còn khoản kết chuyển từ kỳ trước thì trừ sau chi phí đường ống.
 * Suy ra như vậy thì chuỗi luôn đóng đúng, kể cả khi cơ chế đổi.
 */
function dashReadSettlements_(ss, m) {
  const sh = ss.getSheetByName(CUST_CFG.SETTLEMENTS_SHEET);
  if (!sh) return;

  const h = custHeaderMap_(sh);
  const sum = {
    gross: 0, paypalFees: 0, receivingFees: 0, available: 0,
    conversion: 0, transfer: 0, netBefore: 0, payoutFee: 0,
    amountDue: 0, amountPaid: 0, outstandingPayable: 0, invoiceCount: 0
  };

  const byCustomer = {};

  custReadRowsObjects_(sh, h).forEach(function (s) {
    if (!custIsEffectiveSettlementStatus_(s['Status'])) return;

    sum.invoiceCount++;
    sum.gross += custNum_(s['Gross Sales']);
    sum.paypalFees += custNum_(s['PayPal Fees']);
    sum.receivingFees += custNum_(s['Receiving Fees']);
    sum.available += custNum_(s['Available Amount']);
    sum.conversion += custNum_(s['Conversion Cost']);
    sum.transfer += custNum_(s['Transfer Cost']);
    sum.netBefore += custNum_(s['Net Before Payout Fee']);
    sum.payoutFee += custNum_(s['Payout Fee Amount']);
    sum.amountDue += custNum_(s['Amount Due']);
    sum.amountPaid += custNum_(s['Amount Paid']);
    sum.outstandingPayable += custNum_(s['Outstanding Payable']);

    const who = custStr_(s['Customer Name']) || custStr_(s['Customer Key']);
    if (!byCustomer[who]) byCustomer[who] = { n: 0, owed: 0 };
    byCustomer[who].n++;
    byCustomer[who].owed += custNum_(s['Outstanding Payable']);
  });

  Object.keys(sum).forEach(function (k) {
    if (k !== 'invoiceCount') sum[k] = custRound2_(sum[k]);
  });

  sum.refundInPeriod = custRound2_(
    sum.gross - sum.paypalFees - sum.receivingFees - sum.available);
  sum.refundCarry = custRound2_(
    sum.available - sum.conversion - sum.transfer - sum.netBefore);

  m.chain = sum;
  m.feeEarned = sum.payoutFee;
  m.owed = sum.outstandingPayable;

  m.customers.forEach(function (c) {
    const b = byCustomer[c.name];
    if (!b) return;
    c.invoices = b.n;
    c.owed = custRound2_(b.owed);
  });
}


/*
 * Một lượt đọc Payouts phục vụ hai việc: danh sách DRAFT chờ trả, và sổ chủ.
 *
 * Sổ chủ cộng bằng Amount + Fee chứ không phải Amount, vì đó mới là số thật
 * rời khỏi ví — giống hệt cách flowUpdateUsdtSummary_ trừ bằng Gross.
 *
 * DRAFT không tính vào "đã lấy ra": nó là ý định, chưa phải tiền đi.
 */
function dashReadPayouts_(ss, m) {
  const sh = ss.getSheetByName(CUST_CFG.PAYOUTS_SHEET);
  if (!sh) return;

  const h = custHeaderMap_(sh);
  if (!h['Status'] || !h['Amount']) return;

  const known = flowPayoutCategories_();
  const byCategory = {};
  const byPayee = {};

  custReadRowsObjects_(sh, h).forEach(function (p) {
    const status = custUpper_(p['Status']);
    const amount = custNum_(p['Amount']);

    if (status === 'DRAFT') {
      m.draftTotal += amount;
      m.drafts.push({
        key: custStr_(p['Payout Key']),
        payee: custStr_(p['Payee']),
        ref: custStr_(p['Reference']),
        amount: amount
      });
      return;
    }

    if (status !== 'PAID') return;

    const left = custRound2_(amount + custNum_(p['Fee']));
    const category = custUpper_(p['Category']);

    if (!category || known.indexOf(category) === -1) {
      m.owner.unclassified.n++;
      m.owner.unclassified.total = custRound2_(m.owner.unclassified.total + left);
      return;
    }

    if (!flowPayoutIsOwnerCost_(category)) return;

    m.owner.drawn = custRound2_(m.owner.drawn + left);
    byCategory[category] = custRound2_((byCategory[category] || 0) + left);

    const payee = custStr_(p['Payee']) || '(không ghi)';
    byPayee[payee] = custRound2_((byPayee[payee] || 0) + left);
  });

  m.draftTotal = custRound2_(m.draftTotal);
  m.drafts.sort(function (a, b) { return b.amount - a.amount; });
  m.drafts = m.drafts.slice(0, DASH_CFG.MAX_DRAFTS);

  const pack = function (obj) {
    return Object.keys(obj)
      .map(function (k) { return { name: k, amount: obj[k] }; })
      .sort(function (a, b) { return b.amount - a.amount; });
  };

  m.owner.byCategory = pack(byCategory);
  m.owner.byPayee = pack(byPayee);
}


/*
 * Sổ chủ: phí dịch vụ kiếm được, trừ đi mọi khoản đã lấy ra khỏi ví mà
 * không phải trả khách. Người dùng đã chốt gộp chung chi phí vận hành với
 * tiền tự rút — nên chỉ có MỘT con số "còn được lấy".
 *
 * Phần kiếm được lấy từ chuỗi hoá đơn đã tính sẵn ở dashReadSettlements_,
 * nên không đọc lại Settlements lần nữa.
 */
function dashFinishOwnerAccount_(m) {
  m.owner.earned = m.feeEarned;
  m.owner.remaining = custRound2_(m.owner.earned - m.owner.drawn);
}


function dashReadAccounts_(ss, m) {
  const sh = ss.getSheetByName(FLOW_CFG.ACCOUNT_SHEET);
  if (!sh) return;

  const h = custHeaderMap_(sh);
  custReadRowsObjects_(sh, h).forEach(function (a) {
    if (custUpper_(a['Provider']) !== 'PAYPAL') return;

    const hold = custNum_(a['Hold Balance']);
    m.holdBalance += hold;

    const withdrawn = custNum_(a['Total Withdrawn']);
    let state = 'ACTIVE';
    if (custUpper_(a['PayPal Status']) === 'SUSPENDED') state = 'HOLD';
    else if (!custBool_(a['Enabled'])) state = 'TẮT';
    else if (withdrawn <= CUST_CFG.EPSILON) state = 'CHƯA RÚT';

    m.paypal.push({
      key: custStr_(a['Account Key']),
      label: custStr_(a['Label']),
      balance: custNum_(a['Total Balance']),
      withdrawn: withdrawn,
      state: state
    });
  });

  m.holdBalance = custRound2_(m.holdBalance);
  m.paypal.sort(function (a, b) { return b.balance - a.balance; });
}


function dashReadTreasury_(ss, m) {
  [
    { sheet: FLOW_CFG.GREY_SHEET, kind: 'Grey USD',
      bal: 'Estimated USD Balance', got: 'Total Received USD' },
    { sheet: FLOW_CFG.USDT_SHEET, kind: 'Ví USDT',
      bal: 'Estimated Balance', got: 'Total Received USDT' }
  ].forEach(function (src) {
    const sh = ss.getSheetByName(src.sheet);
    if (!sh) return;
    const h = custHeaderMap_(sh);
    custReadRowsObjects_(sh, h).forEach(function (r) {
      const key = custStr_(r['Account Key']);
      if (!key) return;
      m.treasury.push({
        key: key,
        kind: custStr_(r['Tag']) || src.kind,
        balance: custNum_(r[src.bal]),
        received: custNum_(r[src.got]),
        state: custUpper_(r['Status']) || 'ACTIVE',
        sync: dashShortDate_(r['Last Sync'] || r['Last Activity'])
      });
    });
  });

  m.usdt = dashUsdtRemaining_(ss);
}


/*
 * Prefix có tiền thật trên PayPal nhưng chưa khai cho khách nào. Không phải
 * lỗi code — chỉ là tiền chưa vào được sổ, và nó im lặng nên phải nói ra.
 */
function dashReadUnassigned_(ss, m) {
  const sh = ss.getSheetByName(FLOW_CFG.PAYPAL_SHEET);
  if (!sh) return;

  const known = {};
  const cs = ss.getSheetByName(CUST_CFG.CUSTOMERS_SHEET);
  if (cs) {
    const ch = custHeaderMap_(cs);
    custReadRowsObjects_(cs, ch).forEach(function (c) {
      custStr_(c['Prefixes']).split(',').forEach(function (p) {
        const k = custUpper_(p);
        if (k) known[k] = true;
      });
    });
  }

  const h = custHeaderMap_(sh);
  if (!h['Invoice Prefix']) return;

  const acc = {};
  custReadRowsObjects_(sh, h).forEach(function (r) {
    const p = custUpper_(r['Invoice Prefix']);
    if (!p || known[p]) return;
    acc[p] = custRound2_((acc[p] || 0) + custNum_(r['Transaction Amount']));
  });

  Object.keys(acc).forEach(function (p) {
    m.unassignedTotal += acc[p];
    m.unassigned.push({ prefix: p, amount: acc[p] });
  });

  m.unassignedTotal = custRound2_(m.unassignedTotal);
  m.unassigned.sort(function (a, b) { return b.amount - a.amount; });
}


function dashReadHealth_(ss, m) {
  const sh = ss.getSheetByName(FLOW_CFG.TRANSACTION_SHEET);
  if (!sh) return;

  const h = custHeaderMap_(sh);
  const types = {}, recon = {};

  custReadRowsObjects_(sh, h).forEach(function (t) {
    const ty = custUpper_(t['Type']);
    const rc = custUpper_(t['Reconcile Status']);
    if (ty) types[ty] = (types[ty] || 0) + 1;
    if (rc) recon[rc] = (recon[rc] || 0) + 1;
  });

  const pack = function (obj) {
    return Object.keys(obj)
      .map(function (k) { return { name: k, n: obj[k] }; })
      .sort(function (a, b) { return b.n - a.n; });
  };

  m.health.types = pack(types);
  m.health.reconcile = pack(recon);
}


function dashReadHistory_(ss, m) {
  const sh = ss.getSheetByName(DASH_CFG.HISTORY_SHEET);
  if (!sh) return;

  const last = sh.getLastRow();
  if (last < 3) return;

  sh.getRange(3, 1, last - 2, 9).getValues().forEach(function (r) {
    const d = dashDateKey_(r[0]);
    if (!d) return;
    m.history.push({
      d: d.substring(5).replace('-', '/'),
      pending: custNum_(r[1]), outstanding: custNum_(r[2]),
      invoiced: custNum_(r[3]), usdt: custNum_(r[4])
    });
  });

  m.history.sort(function (a, b) { return a.d < b.d ? -1 : 1; });
  if (m.history.length > DASH_CFG.TREND_POINTS) {
    m.history = m.history.slice(-DASH_CFG.TREND_POINTS);
  }
}


/* `USDT!A1` là dòng chữ do Flow ghi; con số nằm sau "CÒN LẠI". */
function dashUsdtRemaining_(ss) {
  const sh = ss.getSheetByName(FLOW_CFG.USDT_SHEET);
  if (!sh) return null;
  const text = custStr_(sh.getRange(1, 1).getValue());
  const found = text.match(/CÒN LẠI\s+(-?[\d.]+)/);
  return found ? custRound2_(Number(found[1])) : null;
}


/*********************** LỊCH SỬ ************************/

/*
 * Khoá theo NGÀY chứ không theo lần chạy. Sync chạy hai lượt mỗi ngày, mà
 * hai chấm cách nhau 12 tiếng trên cùng đường xu hướng chỉ làm biểu đồ răng
 * cưa chứ không thêm thông tin nào.
 */
function dashUpsertHistory_(ss, m) {
  const headers = ['Date', 'Pending in PayPal', 'Chưa xuất HĐ', 'Đã xuất HĐ',
    'USDT còn lại', 'Còn nợ khách', 'Hold Balance',
    'Phí kiếm được', 'Đã lấy ra'];

  let sh = ss.getSheetByName(DASH_CFG.HISTORY_SHEET);
  if (!sh) {
    sh = ss.insertSheet(DASH_CFG.HISTORY_SHEET);
    sh.getRange(1, 1).setValue(
      'Ảnh chụp mỗi ngày một dòng, do updateDashboard() ghi. Đừng sửa tay — ' +
      'chạy lại trong ngày sẽ ghi đè dòng của ngày đó.');
    sh.getRange(2, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(2);
  }

  const today = Utilities.formatDate(new Date(), DASH_CFG.TIMEZONE, 'yyyy-MM-dd');
  const record = [today, m.buckets.pending, m.buckets.outstanding,
    m.buckets.invoiced, m.usdt === null ? '' : m.usdt, m.owed, m.holdBalance,
    m.owner.earned, m.owner.drawn];

  const last = sh.getLastRow();
  let target = 0;

  if (last >= 3) {
    const keys = sh.getRange(3, 1, last - 2, 1).getValues();
    for (let i = 0; i < keys.length; i++) {
      if (dashDateKey_(keys[i][0]) === today) { target = 3 + i; break; }
    }
  }

  if (!target) target = Math.max(last, 2) + 1;
  sh.getRange(target, 1, 1, record.length).setValues([record]);

  const total = sh.getLastRow();
  if (total < 3) return [];

  const rows = sh.getRange(3, 1, total - 2, headers.length).getValues()
    .filter(function (r) { return dashDateKey_(r[0]); });

  if (rows.length > DASH_CFG.HISTORY_KEEP_DAYS) {
    const drop = rows.length - DASH_CFG.HISTORY_KEEP_DAYS;
    sh.deleteRows(3, drop);
    return rows.slice(drop);
  }

  return rows;
}


/* Ô ngày có thể là Date hoặc chuỗi, tuỳ ai chạm vào sheet gần nhất. */
function dashDateKey_(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]') {
    return Utilities.formatDate(value, DASH_CFG.TIMEZONE, 'yyyy-MM-dd');
  }
  return custStr_(value).substring(0, 10);
}


function dashShortDate_(value) {
  const d = custDate_(value);
  return d ? Utilities.formatDate(d, DASH_CFG.TIMEZONE, 'dd/MM') : '';
}


/* 1181.35 -> "1.181,35" — đọc quen mắt hơn trong câu cảnh báo. */
function dashMoney_(value) {
  const n = custNum_(value);
  const parts = Math.abs(n).toFixed(2).split('.');
  return (n < 0 ? '-' : '') +
    parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, '.') + ',' + parts[1];
}


/*********************** BIỂN CHỈ ĐƯỜNG ************************/

/*
 * Tab Dashboard cũ giờ chỉ còn một dòng chỉ đường. Cố ý KHÔNG xoá tab: xoá
 * sheet của người khác là việc không hoàn tác được, mà giữ lại thì cũng
 * không tốn gì.
 */
function dashWriteSignpost_(ss, m) {
  let sh = ss.getSheetByName(DASH_CFG.SHEET);
  if (!sh) return;

  sh.getCharts().forEach(function (c) { sh.removeChart(c); });
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart();
  sh.clear();

  sh.getRange(1, 1).setValue('BẢNG ĐIỀU KHIỂN ĐÃ CHUYỂN SANG TRANG RIÊNG')
    .setFontSize(13).setFontWeight('bold');
  sh.getRange(2, 1).setValue('Mở bằng menu  Payment All  ›  Bảng điều khiển')
    .setFontSize(11);
  sh.getRange(4, 1).setValue(
    'Ảnh chụp hằng ngày vẫn được ghi ở sheet "' + DASH_CFG.HISTORY_SHEET +
    '" để dựng biểu đồ xu hướng. Cập nhật gần nhất: ' + m.at)
    .setFontSize(10).setFontColor('#6d7d92');
  sh.setColumnWidth(1, 640);
}
