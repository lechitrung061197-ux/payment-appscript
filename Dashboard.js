/**
 * ============================================================================
 * DASHBOARD.GS — PAYMENT ALL
 * ============================================================================
 * Một trang trả lời hai câu hỏi hay được hỏi nhất: "tiền đang ở đâu" và "có
 * gì sai không". Báo cáo Telegram đẩy tới bạn mỗi sáng; trang này thì ngược
 * lại — mở ra xem lúc nào cũng được.
 *
 * ==========================================================================
 * VÌ SAO CÓ SHEET DỮ LIỆU ẨN
 * ==========================================================================
 * Biểu đồ nhúng của Sheets neo vào một VÙNG Ô. Nếu vùng đó nằm ngay trên
 * trang hiển thị, mỗi lần số khách tăng giảm là vùng đổi kích thước và biểu
 * đồ đọc sai cột, hoặc tệ hơn là đè lên phần bên dưới.
 *
 * Nên toàn bộ nguồn của biểu đồ nằm ở sheet ẩn `_Dashboard Data`, mỗi khối
 * có kích thước CỐ ĐỊNH, ô thừa để trống. Trang hiển thị thuần trình bày.
 *
 * ==========================================================================
 * VÌ SAO XOÁ SẠCH BIỂU ĐỒ RỒI VẼ LẠI
 * ==========================================================================
 * insertChart() luôn tạo đối tượng MỚI. Gọi lại mỗi lần sync mà không dọn
 * thì sau một tuần có mười bốn biểu đồ chồng lên nhau, và cái nhìn thấy là
 * cái cũ nhất. Dọn trước khi vẽ là cách duy nhất giữ đúng một bộ.
 *
 * ==========================================================================
 * LỊCH SỬ
 * ==========================================================================
 * `Dashboard History` ghi MỘT dòng mỗi ngày, khoá theo ngày nên chạy lại
 * trong cùng ngày là ghi đè chứ không thêm dòng. Không vẽ được xu hướng cho
 * thứ chưa từng ghi lại, nên nó bắt đầu tích từ hôm nay và phải vài tuần
 * mới đủ dày để có nghĩa.
 * ============================================================================
 */

const DASH_CFG = Object.freeze({
  SHEET: 'Dashboard',
  DATA_SHEET: '_Dashboard Data',
  HISTORY_SHEET: 'Dashboard History',

  /* Giữ đúng một năm. Xu hướng dài hơn thế thì đọc bằng mắt không ra gì. */
  HISTORY_KEEP_DAYS: 400,

  /* Số điểm cuối đưa vào biểu đồ xu hướng. */
  TREND_POINTS: 90,

  TIMEZONE: 'Asia/Ho_Chi_Minh',

  /* Cùng bảng màu với tài liệu dòng tiền: xanh = tiền, đất nung = chi phí
     và cảnh báo. Hai màu mang nghĩa, phần còn lại là xám trung tính. */
  COLOR: Object.freeze({
    INK: '#14202e',
    MUTED: '#6b7787',
    PAPER: '#f3f5f8',
    CARD: '#ffffff',
    LINE: '#d3dae4',
    MONEY: '#0d6e63',
    MONEY_SOFT: '#4c9c92',
    MONEY_PALE: '#a8ccc7',
    COST: '#a8503a',
    COST_PALE: '#e8d5cf'
  }),

  LAYOUT: Object.freeze({
    WIDTH: 10,
    ROW_KPI_LABEL: 4,
    ROW_KPI_VALUE: 5,
    ROW_SECTION_WHERE: 7,
    ROW_CHART_TOP: 8,
    ROW_SECTION_CUSTOMER: 27,
    ROW_TABLE_HEAD: 28,
    ROW_TABLE_FIRST: 29,
    ROW_SECTION_TREND: 35,
    ROW_CHART_TREND: 36,
    ROW_SECTION_ALERT: 55,
    ROW_ALERT_FIRST: 56,
    MAX_ALERTS: 10
  })
});


/*********************** ĐIỂM VÀO ************************/

/*
 * Chạy tay lúc nào cũng được. Chuỗi trigger gọi nó sau syncCustomers, vì
 * nó chỉ ĐỌC kết quả của Flow và Customer chứ không tự tính lại gì.
 */
function updateDashboard() {
  const ss = SpreadsheetApp.getActive();
  const started = new Date();

  const metrics = dashCollectMetrics_(ss);
  const history = dashUpsertHistory_(ss, metrics);

  dashWriteDataBlocks_(ss, metrics, history);
  dashRender_(ss, metrics);

  Logger.log(
    'DASHBOARD DONE in ' +
    ((new Date().getTime() - started.getTime()) / 1000).toFixed(1) + 's · ' +
    history.length + ' ngày lịch sử'
  );
}


/*********************** SỐ LIỆU ************************/

/*
 * Ba rổ KHÔNG chồng lấn, cộng lại đúng bằng toàn bộ tiền khách hệ thống
 * biết tới:
 *
 *   kẹt ở PayPal   — chưa rút ra được
 *   chưa xuất HĐ   — đã về nơi nhận, đủ điều kiện nhưng chưa lên hoá đơn
 *   đã xuất HĐ     — đã nằm trên một hoá đơn
 *
 * Lưu ý `Outstanding Amount` nghĩa là available TRỪ invoiced, tức là "đã về
 * nhưng chưa xuất hoá đơn" — KHÔNG phải "đã xuất hoá đơn mà chưa trả". Nhầm
 * chỗ này là vẽ sai cả biểu đồ.
 */
function dashCollectMetrics_(ss) {
  const m = {
    at: new Date(),
    pending: 0, available: 0, invoiced: 0, outstanding: 0,
    usdt: null, draft: 0, paid: 0, holdBalance: 0,
    customers: [], alerts: []
  };

  const cs = ss.getSheetByName(CUST_CFG.CUSTOMERS_SHEET);
  if (cs) {
    const h = custHeaderMap_(cs);
    custReadRowsObjects_(cs, h).forEach(function (c) {
      const name = custStr_(c['Customer Name']) || custStr_(c['Customer Key']);
      if (!name) return;
      const row = {
        name: name,
        pending: custNum_(c['Pending in PayPal']),
        outstanding: custNum_(c['Outstanding Amount']),
        invoiced: custNum_(c['Invoiced Amount'])
      };
      row.total = custRound2_(row.pending + row.outstanding + row.invoiced);
      m.customers.push(row);
      m.pending += row.pending;
      m.outstanding += row.outstanding;
      m.invoiced += row.invoiced;
      m.available += custNum_(c['Available Amount']);
    });
  }

  ['pending', 'available', 'invoiced', 'outstanding'].forEach(function (k) {
    m[k] = custRound2_(m[k]);
  });

  /* Khách nhiều tiền nhất lên trước — biểu đồ cột đọc mới ra thứ tự. */
  m.customers.sort(function (a, b) { return b.total - a.total; });

  m.usdt = dashUsdtRemaining_(ss);

  const ps = ss.getSheetByName(CUST_CFG.PAYOUTS_SHEET);
  if (ps) {
    const h = custHeaderMap_(ps);
    if (h['Status'] && h['Amount']) {
      custReadRowsObjects_(ps, h).forEach(function (p) {
        const st = custUpper_(p['Status']);
        if (st === 'DRAFT') m.draft += custNum_(p['Amount']);
        else if (st === 'PAID') m.paid += custNum_(p['Amount']);
      });
    }
  }
  m.draft = custRound2_(m.draft);
  m.paid = custRound2_(m.paid);

  const acc = ss.getSheetByName(FLOW_CFG.ACCOUNT_SHEET);
  if (acc) {
    const h = custHeaderMap_(acc);
    if (h['Hold Balance']) {
      custReadRowsObjects_(acc, h).forEach(function (a) {
        if (custUpper_(a['Provider']) !== 'PAYPAL') return;
        m.holdBalance += custNum_(a['Hold Balance']);
      });
    }
  }
  m.holdBalance = custRound2_(m.holdBalance);

  /*
   * Dùng lại đúng các collector của báo cáo Telegram. Một nguồn sự thật cho
   * cảnh báo: sửa một chỗ thì cả hai nơi cùng đổi, không bao giờ lệch nhau.
   *
   * Gọi thẳng chứ không tra qua tên chuỗi — TRIGGERS.gs vắng mặt thì hỏng
   * cả dashboard, mà im lặng bỏ qua cảnh báo còn tệ hơn là báo lỗi.
   */
  [
    trigCollectStuckMoney_, trigCollectOverdueWithdrawals_,
    trigCollectForeignCurrency_, trigCollectUnknownEvents_,
    trigCollectDraftOverUsdt_, trigCollectStaleGrey_
  ].forEach(function (collect) { collect(ss, m.alerts); });

  return m;
}


/* `USDT!A1` là dòng chữ do Flow ghi; con số nằm sau "CÒN LẠI". */
function dashUsdtRemaining_(ss) {
  const sh = ss.getSheetByName(FLOW_CFG.USDT_SHEET);
  if (!sh) return null;
  const text = custStr_(sh.getRange(1, 1).getValue());
  const m = text.match(/CÒN LẠI\s+(-?[\d.]+)/);
  return m ? custRound2_(Number(m[1])) : null;
}


/*********************** LỊCH SỬ ************************/

/*
 * Khoá theo NGÀY chứ không theo lần chạy. Sync chạy hai lượt mỗi ngày, mà
 * hai chấm cách nhau 12 tiếng trên cùng một đường xu hướng thì chỉ làm biểu
 * đồ răng cưa chứ không thêm thông tin nào.
 */
function dashUpsertHistory_(ss, m) {
  const headers = ['Date', 'Pending in PayPal', 'Chưa xuất HĐ', 'Đã xuất HĐ',
    'USDT còn lại', 'Payout DRAFT', 'Hold Balance'];

  let sh = ss.getSheetByName(DASH_CFG.HISTORY_SHEET);
  if (!sh) {
    sh = ss.insertSheet(DASH_CFG.HISTORY_SHEET);
    sh.getRange(1, 1).setValue(
      'Ảnh chụp mỗi ngày một dòng, do updateDashboard() ghi. ' +
      'Đừng sửa tay — chạy lại trong ngày sẽ ghi đè dòng của ngày đó.');
    sh.getRange(2, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(2);
  }

  const today = Utilities.formatDate(m.at, DASH_CFG.TIMEZONE, 'yyyy-MM-dd');
  const record = [today, m.pending, m.outstanding, m.invoiced,
    m.usdt === null ? '' : m.usdt, m.draft, m.holdBalance];

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

  /* Đọc lại để trả về đúng thứ đã nằm trên sheet, kể cả dòng vừa ghi. */
  const total = sh.getLastRow();
  if (total < 3) return [];

  const rows = sh.getRange(3, 1, total - 2, headers.length).getValues()
    .filter(function (r) { return dashDateKey_(r[0]); })
    .sort(function (a, b) {
      return dashDateKey_(a[0]) < dashDateKey_(b[0]) ? -1 : 1;
    });

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


/*********************** NGUỒN CỦA BIỂU ĐỒ ************************/

/*
 * Ba khối, mỗi khối một kích thước CỐ ĐỊNH. Ô không dùng tới thì xoá trắng
 * chứ không thu hẹp vùng — vùng co giãn chính là thứ làm biểu đồ đọc lệch
 * cột sau vài lần chạy.
 */
function dashWriteDataBlocks_(ss, m, history) {
  let sh = ss.getSheetByName(DASH_CFG.DATA_SHEET);
  if (!sh) {
    sh = ss.insertSheet(DASH_CFG.DATA_SHEET);
    sh.getRange(1, 1).setValue(
      'Nguồn dữ liệu của biểu đồ Dashboard. Sheet này cố ý bị ẩn — ' +
      'sửa tay ở đây sẽ bị ghi đè ở lần updateDashboard() kế tiếp.');
  }

  const MAX_CUST = 12;

  /* Khối 1 — ba rổ tiền, cho biểu đồ tròn. A3:B6 */
  sh.getRange(3, 1, 4, 2).clearContent();
  sh.getRange(3, 1, 4, 2).setValues([
    ['Rổ', 'Số tiền'],
    ['Kẹt ở PayPal', m.pending],
    ['Đã về, chưa xuất HĐ', m.outstanding],
    ['Đã xuất hoá đơn', m.invoiced]
  ]);

  /* Khối 2 — theo khách, cột chồng. D3:G(3+MAX_CUST) */
  const custBlock = [['Khách', 'Kẹt ở PayPal', 'Chưa xuất HĐ', 'Đã xuất HĐ']];
  for (let i = 0; i < MAX_CUST; i++) {
    const c = m.customers[i];
    custBlock.push(c ? [c.name, c.pending, c.outstanding, c.invoiced] : ['', '', '', '']);
  }
  sh.getRange(3, 4, MAX_CUST + 1, 4).clearContent();
  sh.getRange(3, 4, MAX_CUST + 1, 4).setValues(custBlock);

  /* Khối 3 — xu hướng. I3:M(3+TREND_POINTS) */
  const N = DASH_CFG.TREND_POINTS;
  const tail = history.slice(-N);
  const trend = [['Ngày', 'Kẹt ở PayPal', 'Chưa xuất HĐ', 'Đã xuất HĐ', 'USDT còn lại']];
  for (let i = 0; i < N; i++) {
    const r = tail[i];
    trend.push(r ? [dashDateKey_(r[0]), r[1], r[2], r[3], r[4]] : ['', '', '', '', '']);
  }
  sh.getRange(3, 9, N + 1, 5).clearContent();
  sh.getRange(3, 9, N + 1, 5).setValues(trend);

  sh.hideSheet();
  return sh;
}


/*********************** TRÌNH BÀY ************************/

function dashRender_(ss, m) {
  const L = DASH_CFG.LAYOUT;
  const C = DASH_CFG.COLOR;

  let sh = ss.getSheetByName(DASH_CFG.SHEET);
  if (!sh) sh = ss.insertSheet(DASH_CFG.SHEET);

  /*
   * Dọn sạch rồi dựng lại. Trang này không chứa dữ liệu gốc nào — mọi thứ
   * đều suy ra được — nên xoá trắng là an toàn và tránh được tàn dư của
   * lần chạy trước khi số dòng co lại.
   *
   * breakApart() là bắt buộc: clear() KHÔNG gỡ ô đã gộp. Lần sau số khách
   * đổi, merge chồng lên một vùng gộp cũ có biên khác là ném lỗi giữa chừng.
   */
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart();
  sh.clear();
  dashRemoveCharts_(sh);

  sh.setHiddenGridlines(true);
  sh.getRange(1, 1, sh.getMaxRows(), L.WIDTH).setBackground(C.PAPER);

  /* Tiêu đề */
  sh.getRange(1, 1, 1, L.WIDTH).merge()
    .setValue('PAYMENT ALL')
    .setFontSize(22).setFontWeight('bold').setFontColor(C.INK)
    .setVerticalAlignment('middle');
  sh.setRowHeight(1, 44);

  sh.getRange(2, 1, 1, L.WIDTH).merge()
    .setValue('Cập nhật ' +
      Utilities.formatDate(m.at, DASH_CFG.TIMEZONE, 'HH:mm · dd/MM/yyyy') +
      '   ·   số liệu do syncFlow và syncCustomers dựng, trang này chỉ đọc lại')
    .setFontSize(10).setFontColor(C.MUTED);

  /* Bốn ô số lớn */
  const kpi = [
    ['Kẹt ở PayPal', m.pending, C.COST],
    ['Đã về, chưa xuất HĐ', m.outstanding, C.MONEY],
    ['USDT còn lại', m.usdt, C.MONEY],
    ['Payout DRAFT chờ trả', m.draft, C.COST]
  ];

  kpi.forEach(function (k, i) {
    const col = 1 + i * 2;
    sh.getRange(L.ROW_KPI_LABEL, col, 1, 2).merge()
      .setValue(k[0].toUpperCase())
      .setFontSize(8).setFontColor(C.MUTED).setFontWeight('bold');
    const cell = sh.getRange(L.ROW_KPI_VALUE, col, 1, 2).merge()
      .setValue(k[1] === null ? '—' : k[1])
      .setFontSize(19).setFontWeight('bold').setFontColor(k[2])
      .setVerticalAlignment('middle');
    if (k[1] !== null) cell.setNumberFormat('#,##0.00');
  });
  sh.setRowHeight(L.ROW_KPI_VALUE, 36);

  dashSection_(sh, L.ROW_SECTION_WHERE, 'TIỀN ĐANG Ở ĐÂU');

  /* Bảng theo khách */
  dashSection_(sh, L.ROW_SECTION_CUSTOMER, 'THEO KHÁCH');

  const head = ['Khách', 'Kẹt ở PayPal', 'Đã về, chưa xuất HĐ', 'Đã xuất hoá đơn', 'Tổng'];
  sh.getRange(L.ROW_TABLE_HEAD, 1, 1, head.length).setValues([head])
    .setFontSize(9).setFontWeight('bold').setFontColor(C.MUTED)
    .setBorder(null, null, true, null, null, null, C.LINE,
      SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(L.ROW_TABLE_HEAD, 2, 1, 4).setHorizontalAlignment('right');

  const body = m.customers.map(function (c) {
    return [c.name, c.pending, c.outstanding, c.invoiced, c.total];
  });
  body.push(['TỔNG', m.pending, m.outstanding, m.invoiced,
    custRound2_(m.pending + m.outstanding + m.invoiced)]);

  if (body.length) {
    const rng = sh.getRange(L.ROW_TABLE_FIRST, 1, body.length, 5);
    rng.setValues(body).setFontSize(10).setFontColor(C.INK);
    sh.getRange(L.ROW_TABLE_FIRST, 2, body.length, 4)
      .setNumberFormat('#,##0.00').setHorizontalAlignment('right');
    sh.getRange(L.ROW_TABLE_FIRST + body.length - 1, 1, 1, 5)
      .setFontWeight('bold')
      .setBorder(true, null, null, null, null, null, C.LINE,
        SpreadsheetApp.BorderStyle.SOLID);
  }

  dashSection_(sh, L.ROW_SECTION_TREND, 'XU HƯỚNG THEO NGÀY');

  /* Cảnh báo */
  dashSection_(sh, L.ROW_SECTION_ALERT, 'CẦN XEM');

  if (!m.alerts.length) {
    sh.getRange(L.ROW_ALERT_FIRST, 1, 1, L.WIDTH).merge()
      .setValue('✓  Không có gì bất thường')
      .setFontSize(11).setFontColor(C.MONEY);
  } else {
    const shown = m.alerts.slice(0, L.MAX_ALERTS);
    shown.forEach(function (a, i) {
      sh.getRange(L.ROW_ALERT_FIRST + i, 1, 1, L.WIDTH).merge()
        .setValue('⚠  ' + a)
        .setFontSize(11).setFontColor(C.COST);
    });
    if (m.alerts.length > shown.length) {
      sh.getRange(L.ROW_ALERT_FIRST + shown.length, 1, 1, L.WIDTH).merge()
        .setValue('… và ' + (m.alerts.length - shown.length) +
          ' mục nữa — chạy validateCustomerSystem()')
        .setFontSize(10).setFontColor(C.MUTED);
    }
  }

  sh.setColumnWidth(1, 190);
  for (let c = 2; c <= L.WIDTH; c++) sh.setColumnWidth(c, 122);

  dashBuildCharts_(ss, sh, m);
  sh.setActiveSelection('A1');
}


function dashSection_(sheet, row, label) {
  const C = DASH_CFG.COLOR;
  sheet.getRange(row, 1, 1, DASH_CFG.LAYOUT.WIDTH).merge()
    .setValue(label)
    .setFontSize(9).setFontWeight('bold').setFontColor(C.MUTED)
    .setVerticalAlignment('bottom');
  sheet.setRowHeight(row, 30);
}


/*********************** BIỂU ĐỒ ************************/

function dashRemoveCharts_(sheet) {
  sheet.getCharts().forEach(function (c) { sheet.removeChart(c); });
}


function dashBuildCharts_(ss, sheet, m) {
  const data = ss.getSheetByName(DASH_CFG.DATA_SHEET);
  if (!data) return;

  const L = DASH_CFG.LAYOUT;
  const C = DASH_CFG.COLOR;
  const MAX_CUST = 12;
  const N = DASH_CFG.TREND_POINTS;

  /* 1 — ba rổ tiền */
  sheet.insertChart(
    sheet.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(data.getRange(3, 1, 4, 2))
      .setPosition(L.ROW_CHART_TOP, 1, 0, 0)
      .setNumHeaders(1)
      .setOption('title', 'Tiền khách, theo nơi đang nằm')
      .setOption('titleTextStyle', { color: C.INK, fontSize: 13, bold: true })
      .setOption('pieHole', 0.55)
      .setOption('colors', [C.COST, C.MONEY_SOFT, C.MONEY])
      .setOption('legend', { position: 'labeled', textStyle: { color: C.MUTED, fontSize: 11 } })
      .setOption('pieSliceText', 'value')
      .setOption('backgroundColor', C.CARD)
      .setOption('chartArea', { left: 10, top: 46, width: '92%', height: '78%' })
      .setOption('width', 560)
      .setOption('height', 340)
      .build()
  );

  /* 2 — theo khách, cột chồng: cùng ba rổ, nhìn được ai đang kẹt ở đâu */
  sheet.insertChart(
    sheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(data.getRange(3, 4, MAX_CUST + 1, 4))
      .setPosition(L.ROW_CHART_TOP, 6, 0, 0)
      .setNumHeaders(1)
      .setOption('title', 'Từng khách, tiền đang ở đâu')
      .setOption('titleTextStyle', { color: C.INK, fontSize: 13, bold: true })
      .setOption('isStacked', true)
      .setOption('colors', [C.COST, C.MONEY_SOFT, C.MONEY])
      .setOption('legend', { position: 'top', textStyle: { color: C.MUTED, fontSize: 11 } })
      .setOption('backgroundColor', C.CARD)
      .setOption('hAxis', { textStyle: { color: C.MUTED, fontSize: 11 } })
      .setOption('vAxis', {
        textStyle: { color: C.MUTED, fontSize: 10 },
        gridlines: { color: C.LINE }, minorGridlines: { count: 0 }
      })
      .setOption('chartArea', { left: 64, top: 64, width: '88%', height: '68%' })
      .setOption('width', 560)
      .setOption('height', 340)
      .build()
  );

  /* 3 — xu hướng. Một điểm thì chưa ra đường, phải đợi lịch sử dày lên. */
  sheet.insertChart(
    sheet.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(data.getRange(3, 9, N + 1, 5))
      .setPosition(L.ROW_CHART_TREND, 1, 0, 0)
      .setNumHeaders(1)
      .setOption('title', 'Diễn biến theo ngày')
      .setOption('titleTextStyle', { color: C.INK, fontSize: 13, bold: true })
      .setOption('colors', [C.COST, C.MONEY_SOFT, C.MONEY, C.INK])
      .setOption('legend', { position: 'top', textStyle: { color: C.MUTED, fontSize: 11 } })
      .setOption('backgroundColor', C.CARD)
      .setOption('curveType', 'none')
      .setOption('pointSize', 3)
      .setOption('lineWidth', 2)
      .setOption('hAxis', { textStyle: { color: C.MUTED, fontSize: 10 }, showTextEvery: 7 })
      .setOption('vAxis', {
        textStyle: { color: C.MUTED, fontSize: 10 },
        gridlines: { color: C.LINE }, minorGridlines: { count: 0 }
      })
      .setOption('chartArea', { left: 64, top: 56, width: '90%', height: '72%' })
      .setOption('width', 1140)
      .setOption('height', 320)
      .build()
  );
}
