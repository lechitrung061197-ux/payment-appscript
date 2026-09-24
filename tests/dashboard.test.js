/*
 * Dashboard: số học chuỗi tiền, ba rổ, cảnh báo chéo — bằng số thật của
 * workbook. Và chạy thật phần JS của trang HTML, vì lỗi runtime ở đó thì
 * người dùng chỉ thấy một trang trắng, không có thông báo nào.
 */
const fs = require('fs');
const vm = require('vm');
const path = require('path');

const REPO = path.join(__dirname, '..');

let pass = 0, fail = 0;
function ok(name, cond, extra) {
  if (cond) { pass++; console.log('  ✓ ' + name); }
  else { fail++; console.log('  ✗ ' + name + (extra ? '  → ' + extra : '')); }
}
function eq(name, actual, expected, tol) {
  const t = tol === undefined ? 0.005 : tol;
  const good = typeof expected === 'number'
    ? Math.abs(actual - expected) <= t
    : actual === expected;
  ok(name, good, 'được ' + JSON.stringify(actual) + ', mong ' + JSON.stringify(expected));
}

/* ---------------- số thật, workbook 02/09/2026 ---------------- */

const CUSTOMERS = [
  { 'Customer Key': 'CUS001', 'Customer Name': 'HUI', 'Status': 'ACTIVE',
    'Prefixes': 'SHOPSH, VASC, BUDD, BERMSH', 'Payout Fee %': 0.05,
    'Pending in PayPal': 1419.58, 'Outstanding Amount': 0, 'Invoiced Amount': 4129.19,
    'Available Amount': 4129.19, 'Advance Outstanding': 0, 'Last Settlement': '' },
  { 'Customer Key': 'CUS002', 'Customer Name': 'DITECH', 'Status': 'ACTIVE',
    'Prefixes': 'BOSSX, BEAC, THEWSH', 'Payout Fee %': 0.05,
    'Pending in PayPal': 295.68, 'Outstanding Amount': 0, 'Invoiced Amount': 1740.14,
    'Available Amount': 1740.14, 'Advance Outstanding': 0, 'Last Settlement': '' },
  { 'Customer Key': 'CUS003', 'Customer Name': 'ASON', 'Status': 'ACTIVE',
    'Prefixes': 'USAH', 'Payout Fee %': 0,
    'Pending in PayPal': 25.90, 'Outstanding Amount': 684.39, 'Invoiced Amount': 0,
    'Available Amount': 684.39, 'Advance Outstanding': 0, 'Last Settlement': '' }
];

/* Gross, Refunds, PPFee, Recv, Avail, Conv, Transfer, NetBefore, Fee, Due, Paid, Out */
const S = [
  ['HUI', 1885.92, 34.98, 80.15, 14.13, 1756.66, 3.51, 3.51, 1749.64, 87.48, 1662.16, 1662.16, 0],
  ['HUI', 1782.69, 0, 70.94, 14.54, 1697.21, 3.39, 3.39, 1690.43, 84.52, 1605.91, 1605.91, 0],
  ['DITECH', 312.03, 0, 14.26, 2.42, 295.35, 0.59, 1.00, 293.76, 14.69, 279.07, 0, 279.07],
  ['DITECH', 929.00, 0, 38.68, 7.14, 883.18, 1.77, 1.77, 879.64, 43.98, 835.66, 835.66, 0],
  ['DITECH', 271.55, 0, 10.56, 2.32, 258.67, 0.52, 1.00, 257.15, 12.86, 244.29, 0, 244.29],
  ['HUI', 194.90, 0, 8.20, 6.00, 180.70, 0.45, 1.00, 179.25, 8.96, 170.29, 170.29, 0],
  ['HUI', 523.68, 0, 22.64, 6.42, 494.62, 1.24, 1.24, 492.14, 24.61, 467.53, 0, 467.53],
  ['DITECH', 323.83, 1.47, 13.80, 7.09, 302.94, 0.76, 1.00, 299.71, 14.99, 284.72, 0, 284.72]
];
const SETTLEMENTS = S.map((r, i) => ({
  'Settlement Key': 'SET-' + (i + 1), 'Invoice Number': 'INV-' + (i + 1),
  'Customer Key': r[0] === 'HUI' ? 'CUS001' : 'CUS002', 'Customer Name': r[0],
  'Gross Sales': r[1], 'Refunds': r[2], 'PayPal Fees': r[3], 'Receiving Fees': r[4],
  'Available Amount': r[5], 'Conversion Cost': r[6], 'Transfer Cost': r[7],
  'Net Before Payout Fee': r[8], 'Payout Fee Amount': r[9], 'Amount Due': r[10],
  'Amount Paid': r[11], 'Outstanding Payable': r[12],
  'Status': r[11] > 0 && r[12] === 0 ? 'PAID' : 'ISSUED', 'Issued At': ''
}));

/* Category/Fee đúng như sheet thật ngày 02/09: 3 dòng trống, 1 dòng lạ. */
const P_ = (k, payee, ref, amt, st, cat, fee, settlement) => ({
  'Payout Key': k, 'Payee': payee, 'Reference': ref, 'Amount': amt,
  'Status': st, 'Category': cat || '', 'Fee': fee || 0,
  'Settlement Key': settlement || ''
});
const PAYOUTS = [
  P_('PO-000003', 'DITECH', 'INV-000003', 279.07, 'DRAFT', 'CUSTOMER_PAYOUT'),
  P_('PO-000005', 'DITECH', 'INV-000005', 244.29, 'DRAFT', 'CUSTOMER_PAYOUT'),
  P_('PO-000011', 'HUI', 'INV-000007', 467.53, 'DRAFT', 'CUSTOMER_PAYOUT'),
  P_('PO-000012', 'DITECH', 'INV-000008', 284.72, 'DRAFT', 'CUSTOMER_PAYOUT'),
  P_('PO-000001', 'HUI', 'INV-000001', 1662.16, 'PAID', 'CUSTOMER_PAYOUT'),
  P_('PO-000006', 'TRUNG', '', 100.00, 'PAID', ''),
  P_('PO-000007', 'TRUNG', '', 24.00, 'PAID', ''),
  P_('PO-000008', 'ASON', '', 466.00, 'PAID', '', 2.00),
  P_('PO-000009', 'ASON', '', 708.00, 'PAID', 'MANUAL_PAYOUT')
];

/*
 * Cùng dữ liệu nhưng đã phân loại xong. Thay THEO KHOÁ chứ không lọc theo
 * "có category hay chưa" — PO-000009 vốn đã mang MANUAL_PAYOUT, lọc kiểu đó
 * là giữ lại bản cũ rồi thêm bản mới, thành đếm 708 hai lần.
 */
const FIXED_ = {
  'PO-000006': P_('PO-000006', 'TRUNG', '', 100.00, 'PAID', 'OWNER_DRAW'),
  'PO-000007': P_('PO-000007', 'TRUNG', '', 24.00, 'PAID', 'OWNER_DRAW'),
  /* Bù trừ phần ghi thừa vào ví mặc định, không phải tiền tiêu. */
  'PO-000008': P_('PO-000008', 'ASON', '', 466.00, 'PAID', 'ADJUSTMENT', 2.00),
  'PO-000009': P_('PO-000009', 'ASON', '', 708.00, 'PAID', 'ADJUSTMENT')
};
const PAYOUTS_FIXED = PAYOUTS.map(p => FIXED_[p['Payout Key']] || p);

const ACCOUNTS = [
  { 'Account Key': 'PP1', 'Label': 'bellaboho', 'Provider': 'PAYPAL', 'Enabled': true,
    'PayPal Status': 'SUSPENDED', 'Total Balance': 289.85, 'Hold Balance': 289.85, 'Total Withdrawn': 4931.62 },
  { 'Account Key': 'PP2', 'Label': 'socialchics', 'Provider': 'PAYPAL', 'Enabled': true,
    'PayPal Status': 'SUSPENDED', 'Total Balance': 1589.43, 'Hold Balance': 1589.43, 'Total Withdrawn': 1525.53 },
  { 'Account Key': 'PP4', 'Label': 'bambini', 'Provider': 'PAYPAL', 'Enabled': true,
    'PayPal Status': '', 'Total Balance': 204.18, 'Hold Balance': 0, 'Total Withdrawn': 0 },
  { 'Account Key': 'GR001', 'Label': 'grey', 'Provider': 'GREY', 'Enabled': true,
    'PayPal Status': '', 'Total Balance': 0, 'Hold Balance': 999, 'Total Withdrawn': 0 }
];

const GREY = [
  { 'Account Key': 'GR001', 'Tag': '', 'Status': 'ACTIVE', 'Estimated USD Balance': 87.28,
    'Total Received USD': 5380.17, 'Last Sync': '' }
];
const USDT_ROWS = [
  { 'Account Key': 'USDT1', 'Tag': 'Nhat', 'Status': 'ACTIVE', 'Estimated Balance': 1078.06,
    'Total Received USDT': 6652.08, 'Last Activity': '' }
];

const PAYPAL_ROWS = [
  { 'Invoice Prefix': 'SHOPSH', 'Transaction Amount': 100 },
  { 'Invoice Prefix': 'TTPEZ', 'Transaction Amount': 240.90 },
  { 'Invoice Prefix': 'JUST', 'Transaction Amount': 130.28 },
  { 'Invoice Prefix': 'PRAN', 'Transaction Amount': 78.05 },
  { 'Invoice Prefix': 'MASTSH', 'Transaction Amount': 34.98 }
];

const TX = [
  { 'Type': 'PAYPAL_WITHDRAWAL', 'Reconcile Status': 'ARRIVED', 'Net Amount': 100 },
  { 'Type': 'PAYPAL_WITHDRAWAL', 'Reconcile Status': 'RETURNED', 'Net Amount': 860.4 },
  { 'Type': 'GREY_FUNDING', 'Reconcile Status': 'MATCHED', 'Net Amount': 50 }
];

const USDT_A1 = 'TỔNG USDT TRONG HỆ THỐNG · từ PayPal 6652.08 · đã chi 5574.02 ' +
  '· CÒN LẠI 1078.06   (chỉ tính phần hệ thống ghi nhận)';

function build(over) {
  const o = over || {};
  const data = Object.assign({
    'Customers': CUSTOMERS, 'Settlements': SETTLEMENTS, 'Payouts': PAYOUTS,
    'Account': ACCOUNTS, 'Grey': GREY, 'USDT': USDT_ROWS, 'Paypal': PAYPAL_ROWS,
    'Transaction': TX, 'Customer Allocation': [], 'Customer Ledger': []
  }, o.sheets || {});

  const usdtText = o.usdt !== undefined ? o.usdt : USDT_A1;
  const history = o.history || [];

  const sb = {
    console, Logger: { log: () => {} },
    Utilities: {
      formatDate: (d, tz, fmt) => {
        const p = n => String(n).padStart(2, '0');
        if (fmt === 'yyyy-MM-dd') return d.getFullYear() + '-' + p(d.getMonth() + 1) + '-' + p(d.getDate());
        if (fmt === 'dd/MM') return p(d.getDate()) + '/' + p(d.getMonth() + 1);
        if (fmt === 'H') return String(d.getHours());
        return 'HH:mm · dd/MM/yyyy';
      }
    },
    SpreadsheetApp: { getActive: () => ss },
    Date, Math, JSON, Object, Array, String, Number, Error, RegExp, isNaN, parseFloat, parseInt
  };

  const ss = {
    getSheetByName: name => {
      if (name === 'Dashboard History') {
        if (!history.length) return null;
        return {
          __name: name,
          getLastRow: () => history.length + 2,
          getRange: () => ({ getValues: () => history })
        };
      }
      if (!data[name]) return null;
      return {
        __name: name,
        getRange: () => ({ getValue: () => (name === 'USDT' ? usdtText : '') })
      };
    }
  };

  vm.createContext(sb);
  ['Flow.js', 'Customer.gs.js', 'Triggers.js', 'Dashboard.js'].forEach(f =>
    vm.runInContext(fs.readFileSync(path.join(REPO, f), 'utf8'), sb, { filename: f }));

  /* custHeaderMap_ thật trả về map 1-BASED (m[k]=i+1) nên guard `if(h['X'])`
     an toàn cả với cột A. Stub phải bắt chước đúng vậy. */
  sb.__data = data;
  vm.runInContext(
    'custHeaderMap_ = function(sheet){' +
    '  var rows = __data[sheet.__name] || [], m = {};' +
    '  if (rows.length) Object.keys(rows[0]).forEach(function(k,i){ m[k]=i+1; });' +
    '  return m; };' +
    'custReadRowsObjects_ = function(sheet){ return __data[sheet.__name] || []; };', sb);

  return sb;
}

function metrics(over) {
  const sb = build(over);
  return vm.runInContext('dashCollectMetrics_(SpreadsheetApp.getActive())', sb);
}

/* ════════════════════════════════════════════════════════ */
console.log('\n1 · chuỗi tiền phải ĐÓNG ĐÚNG ở mọi mốc');
{
  const c = metrics().chain;
  eq('số hoá đơn hiệu lực', c.invoiceCount, 8);
  eq('tiền hàng', c.gross, 6223.60);
  eq('phí PayPal', c.paypalFees, 259.23);
  eq('phí Grey nhận', c.receivingFees, 60.06);
  eq('tiền khả dụng', c.available, 5869.33);

  /* Hai vị trí của hoàn trả được SUY RA, không đoán. Tổng hai phần phải
     bằng đúng tổng Refunds trên sheet. */
  eq('hoàn trả TRONG KỲ (suy ra)', c.refundInPeriod, 34.98);
  eq('hoàn trả KẾT CHUYỂN (suy ra)', c.refundCarry, 1.47);
  eq('hai phần cộng lại = tổng Refunds', c.refundInPeriod + c.refundCarry, 36.45);

  eq('mốc 1 đóng: gross − pp − recv − refund = available',
     c.gross - c.paypalFees - c.receivingFees - c.refundInPeriod, c.available);
  eq('mốc 2 đóng: available − conv − transfer − carry = netBefore',
     c.available - c.conversion - c.transfer - c.refundCarry, c.netBefore);
  eq('mốc 3 đóng: netBefore − payoutFee = amountDue',
     c.netBefore - c.payoutFee, c.amountDue);
  eq('mốc 4 đóng: paid + outstanding = amountDue',
     c.amountPaid + c.outstandingPayable, c.amountDue);

  eq('phí dịch vụ đã thu', c.payoutFee, 292.09);
  eq('còn nợ khách', c.outstandingPayable, 1275.61);
}

console.log('\n2 · ba rổ không chồng lấn');
{
  const m = metrics();
  eq('kẹt PayPal', m.buckets.pending, 1741.16);
  eq('chưa xuất HĐ', m.buckets.outstanding, 684.39);
  eq('đã xuất HĐ', m.buckets.invoiced, 5869.33);
  eq('tổng', m.buckets.total, 8294.88);
  eq('đã xuất HĐ khớp Available Amount của Settlements',
     m.buckets.invoiced, m.chain.available);
  eq('còn nợ khách = tổng DRAFT', m.owed, m.draftTotal);
}

console.log('\n3 · khách, DRAFT, prefix chưa gắn');
{
  const m = metrics();
  eq('HUI đứng đầu', m.customers[0].name, 'HUI');
  eq('HUI 4 hoá đơn', m.customers[0].invoices, 4);
  eq('HUI còn nợ', m.customers[0].owed, 467.53);
  eq('DITECH còn nợ', m.customers[1].owed, 808.08);
  eq('ASON chưa hoá đơn nào', m.customers[2].invoices, 0);

  eq('4 lệnh DRAFT', m.drafts.length, 4);
  eq('DRAFT sắp giảm dần', m.drafts[0].amount, 467.53);
  eq('tổng DRAFT', m.draftTotal, 1275.61);

  eq('4 prefix chưa gắn', m.unassigned.length, 4);
  eq('tổng chưa gắn', m.unassignedTotal, 484.21);
  eq('lớn nhất là TTPEZ', m.unassigned[0].prefix, 'TTPEZ');
  ok('SHOPSH đã gắn nên KHÔNG xuất hiện',
     m.unassigned.every(u => u.prefix !== 'SHOPSH'));

  eq('PP2 mang trạng thái HOLD', m.paypal[0].state, 'HOLD');
  eq('PP4 chưa rút lần nào', m.paypal.filter(a => a.key === 'PP4')[0].state, 'CHƯA RÚT');
  eq('Hold Balance chỉ cộng PAYPAL', m.holdBalance, 1879.28);
  ok('ngân quỹ gồm cả Grey và USDT', m.treasury.length === 2);
}

console.log('\n4 · cảnh báo');
{
  const m = metrics();
  const draftAlert = m.alerts.filter(a => a.indexOf('Payout DRAFT') === 0);
  eq('kêu DRAFT vượt USDT', draftAlert.length, 1);
  ok('nêu số thiếu 197,55', draftAlert[0].indexOf('197,55') !== -1, draftAlert[0]);
  ok('kêu tiền chưa gắn khách',
     m.alerts.some(a => a.indexOf('484,21') !== -1 && a.indexOf('chưa gắn') !== -1));
}
{
  const m = metrics({ sheets: { 'Payouts': [{ 'Payout Key': 'X', 'Payee': 'A', 'Reference': '', 'Amount': 500, 'Status': 'DRAFT' }] } });
  eq('DRAFT < USDT thì im', m.alerts.filter(a => a.indexOf('Payout DRAFT') === 0).length, 0);
}
{
  const m = metrics({ usdt: 'chưa chạy syncFlow' });
  eq('A1 chưa có dòng tổng → usdt null', m.usdt, null);
  eq('và không đoán bừa', m.alerts.filter(a => a.indexOf('Payout DRAFT') === 0).length, 0);
}

console.log('\n5 · trang HTML chạy thật, không ném lỗi');
{
  const raw = fs.readFileSync(path.join(REPO, 'DashboardDialog.html'), 'utf8');
  ok('có chỗ nhồi payload', raw.indexOf('<?!= payload ?>') !== -1);

  const script = raw.substring(raw.lastIndexOf('<script>') + 8, raw.lastIndexOf('</script>'));

  const render = (m) => {
    let out = '';
    const doc = { getElementById: () => ({ set innerHTML(v) { out = v; } }) };
    const s = { document: doc, Math, JSON, Object, Array, String, Number, isNaN, console };
    vm.createContext(s);
    vm.runInContext(script.replace('<?!= payload ?>', JSON.stringify(m)), s);
    return out;
  };

  const html = render(metrics());
  ok('dựng ra HTML', html.length > 3000, html.length + ' ký tự');
  ok('có chuỗi tiền', html.indexOf('Chuỗi tiền') !== -1);
  ok('số hiện đúng định dạng Việt', html.indexOf('6.223,60') !== -1);
  ok('có bảng DRAFT', html.indexOf('PO-000011') !== -1);
  ok('có prefix chưa gắn', html.indexOf('TTPEZ') !== -1);
  ok('mới 0 ngày lịch sử thì KHÔNG vẽ đường', html.indexOf('<svg') === -1);
  ok('không lọt chuỗi undefined/NaN vào trang',
     html.indexOf('undefined') === -1 && html.indexOf('NaN') === -1);

  /* Đủ 3 ngày thì mới vẽ */
  const hist = [
    ['2026-09-01', 1700, 700, 5800, 1100, 1200, 0],
    ['2026-09-02', 1741.16, 684.39, 5869.33, 1078.06, 1275.61, 1879.28],
    ['2026-09-03', 1600, 690, 5900, 1050, 1300, 1879.28]
  ];
  const html2 = render(metrics({ history: hist }));
  ok('đủ 3 ngày thì vẽ đường', html2.indexOf('<svg') !== -1);
  ok('có 4 đường', (html2.match(/<path /g) || []).length === 4);

  /* Sổ rỗng hoàn toàn cũng không được vỡ */
  const empty = metrics({ sheets: {
    'Customers': [], 'Settlements': [], 'Payouts': [], 'Paypal': [], 'Transaction': []
  } });
  const html3 = render(empty);
  ok('sổ rỗng vẫn dựng được trang', html3.length > 500);
  ok('sổ rỗng không có NaN', html3.indexOf('NaN') === -1);
}

console.log('\n6 · sổ chủ');
{
  const sb = build();
  const own = (cat) => { sb.__c = cat; return vm.runInContext('flowPayoutIsOwnerCost_(__c)', sb); };
  eq('CUSTOMER_PAYOUT không phải phần chủ', own('CUSTOMER_PAYOUT'), false);
  eq('rỗng → false (chưa phân loại thì chưa kết luận)', own(''), false);
  eq('OWNER_DRAW', own('OWNER_DRAW'), true);
  eq('EXPENSE', own('EXPENSE'), true);
  eq('chữ thường vẫn nhận', own('expense'), true);
  /* Giá trị lạ KHÔNG được coi là đã lấy — nếu không, sổ chủ và cảnh báo
     sẽ báo hai con số khác nhau cho cùng một sheet. */
  eq('giá trị lạ MANUAL_PAYOUT → false', own('MANUAL_PAYOUT'), false);

  /* Chỉ OWNER_DRAW và EXPENSE mới là chủ tiêu. Ca thật: 1.176,00 bù trừ ghi
     ví từng bị tính thành chủ tiêu, kéo sổ chủ xuống âm oan. */
  eq('ADJUSTMENT không phải chủ tiêu', own('ADJUSTMENT'), false);
  eq('TRANSFER không phải chủ tiêu', own('TRANSFER'), false);
  eq('REFUND không phải chủ tiêu', own('REFUND'), false);
  eq('OTHER không kết luận', own('OTHER'), false);

  /* Settlement Key thắng Category. Trên sheet thật có dòng
     "Auto-created from INV-000009" bị đổi thành MANUAL_PAYOUT. */
  const own2 = (cat, sk) => { sb.__c = cat; sb.__k = sk;
    return vm.runInContext('flowPayoutIsOwnerCost_(__c, __k)', sb); };
  eq('có Settlement Key thì EXPENSE cũng không phải phần chủ',
     own2('EXPENSE', 'SET-000009'), false);
  eq('có Settlement Key + MANUAL_PAYOUT', own2('MANUAL_PAYOUT', 'SET-000009'), false);
  eq('không Settlement Key thì EXPENSE là phần chủ', own2('EXPENSE', ''), true);

  /* Chưa phân loại: KHÔNG được đoán bừa vào "đã lấy ra" */
  const m = metrics();
  eq('kiếm được', m.owner.earned, 292.09);
  eq('đã lấy ra (chỉ dòng đã phân loại)', m.owner.drawn, 0);
  eq('4 dòng chưa phân loại', m.owner.unclassified.n, 4);
  eq('tổng chưa phân loại = 1.300,00 (gồm Fee)', m.owner.unclassified.total, 1300.00);

  /* Sau khi điền category */
  const f = metrics({ sheets: { 'Payouts': PAYOUTS_FIXED } });
  eq('chỉ TRUNG là chủ rút', f.owner.drawn, 124.00);
  eq('còn được lấy → DƯƠNG', f.owner.remaining, 168.09);
  eq('hết dòng chưa phân loại', f.owner.unclassified.n, 0);
  eq('OWNER_DRAW gộp đúng', f.owner.byCategory.filter(x => x.name === 'OWNER_DRAW')[0].amount, 124.00);
  eq('1.176,00 bù trừ KHÔNG lọt vào', f.owner.byCategory.length, 1);
  eq('theo người nhận chỉ còn TRUNG', f.owner.byPayee[0].name, 'TRUNG');
  eq('TRUNG 124,00', f.owner.byPayee[0].amount, 124.00);

  /* Ca thật trên sheet 03/09: payout của hoá đơn ASON bị gõ nhầm Category.
     Không được tính là chủ rút, cũng không được kêu "chưa phân loại". */
  const asonCase = metrics({ sheets: { 'Payouts': PAYOUTS_FIXED.concat([
    P_('PO-000013', 'ASON', 'INV-000009', 680.97, 'PAID', 'MANUAL_PAYOUT', 0, 'SET-000009')
  ]) } });
  eq('payout gắn hoá đơn không vào sổ chủ', asonCase.owner.drawn, 124.00);
  eq('và không bị kêu chưa phân loại', asonCase.owner.unclassified.n, 0);
  eq('cảnh báo phân loại cũng im',
     asonCase.alerts.filter(a => a.indexOf('chưa phân loại') !== -1).length, 0);

  /* DRAFT không bao giờ tính là đã lấy */
  const d2 = metrics({ sheets: { 'Payouts': [
    P_('X1', 'TRUNG', '', 500, 'DRAFT', 'OWNER_DRAW')
  ] } });
  eq('DRAFT không tính vào đã lấy', d2.owner.drawn, 0);
  eq('DRAFT vẫn vào danh sách chờ', d2.drafts.length, 1);
}

console.log('\n7 · cảnh báo phân loại và rút quá');
{
  const m = metrics();
  const cat = m.alerts.filter(a => a.indexOf('chưa phân loại') !== -1);
  eq('kêu payout chưa phân loại', cat.length, 1);
  ok('nêu 3 dòng trống', cat[0].indexOf('3 dòng để trống') !== -1, cat[0]);
  ok('nêu tên giá trị lạ MANUAL_PAYOUT', cat[0].indexOf('MANUAL_PAYOUT') !== -1, cat[0]);
  ok('nêu tổng 1.300,00', cat[0].indexOf('1.300,00') !== -1, cat[0]);

  /* Chưa phân loại thì chưa kết luận rút quá */
  eq('chưa kêu rút quá khi còn dòng trống',
     m.alerts.filter(a => a.indexOf('Đã lấy ra') === 0).length, 0);

  const f = metrics({ sheets: { 'Payouts': PAYOUTS_FIXED } });
  eq('phân loại xong thì hết kêu thiếu',
     f.alerts.filter(a => a.indexOf('chưa phân loại') !== -1).length, 0);
  eq('và KHÔNG còn kêu rút quá — bù trừ không phải tiêu',
     f.alerts.filter(a => a.indexOf('Đã lấy ra') === 0).length, 0);

  /* Rút quá thật thì vẫn phải kêu */
  const real = metrics({ sheets: { 'Payouts': [
    P_('X1', 'TRUNG', '', 400, 'PAID', 'OWNER_DRAW')
  ] } });
  const over = real.alerts.filter(a => a.indexOf('Đã lấy ra') === 0);
  eq('kêu rút quá', over.length, 1);
  ok('nêu vượt 107,91', over[0].indexOf('107,91') !== -1, over[0]);

  /* Lấy ít hơn kiếm được thì im */
  const okCase = metrics({ sheets: { 'Payouts': [
    P_('X1', 'TRUNG', '', 50, 'PAID', 'OWNER_DRAW')
  ] } });
  eq('lấy ít hơn kiếm được → im', okCase.alerts.filter(a => a.indexOf('Đã lấy ra') === 0).length, 0);
  eq('còn được lấy', okCase.owner.remaining, 242.09);
}

console.log('\n8 · trang HTML với sổ chủ');
{
  const raw = fs.readFileSync(path.join(REPO, 'DashboardDialog.html'), 'utf8');
  const script = raw.substring(raw.lastIndexOf('<script>') + 8, raw.lastIndexOf('</script>'));
  const render = (m) => {
    let out = '';
    const doc = { getElementById: () => ({ set innerHTML(v) { out = v; } }) };
    const s2 = { document: doc, Math, JSON, Object, Array, String, Number, isNaN, console };
    vm.createContext(s2);
    vm.runInContext(script.replace('<?!= payload ?>', JSON.stringify(m)), s2);
    return out;
  };

  const html = render(metrics({ sheets: { 'Payouts': PAYOUTS_FIXED } }));
  ok('có khối Sổ chủ', html.indexOf('Sổ chủ') !== -1);
  ok('hiện còn được lấy 168,09', html.indexOf('168,09') !== -1);
  ok('không cảnh báo lấy vào tiền khách', html.indexOf('tiền của khách') === -1);
  ok('có breakdown OWNER_DRAW', html.indexOf('OWNER_DRAW') !== -1);
  ok('không NaN', html.indexOf('NaN') === -1);

  const negative = render(metrics({ sheets: { 'Payouts': [
    P_('X1', 'TRUNG', '', 400, 'PAID', 'OWNER_DRAW')
  ] } }));
  ok('rút quá thật thì hiện số âm', negative.indexOf('-107,91') !== -1);
  ok('và cảnh báo lấy vào tiền khách', negative.indexOf('tiền của khách') !== -1);

  const html2 = render(metrics());
  ok('còn dòng chưa phân loại thì nói ra', html2.indexOf('chưa phân loại') !== -1);

  const empty = metrics({ sheets: {
    'Customers': [], 'Settlements': [], 'Payouts': [], 'Paypal': [], 'Transaction': []
  } });
  const html3 = render(empty);
  ok('sổ rỗng vẫn dựng được', html3.length > 500);
  ok('sổ rỗng không NaN', html3.indexOf('NaN') === -1);
}

console.log('\n' + (fail ? '✗ ' + fail + ' FAIL' : '✓ tất cả') + ' · ' + pass + ' pass\n');
process.exit(fail ? 1 : 0);
