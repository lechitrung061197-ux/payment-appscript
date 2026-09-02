/*
 * Dashboard: số học ba rổ và cảnh báo chéo, chạy bằng số thật của workbook.
 * Phần vẽ biểu đồ cần Apps Script nên không test được ở đây — cái test được
 * là con số đưa vào biểu đồ, và đó mới là chỗ sai được.
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

/* ---- số thật lấy từ workbook 02/09 ---- */
const CUSTOMERS = [
  { 'Customer Key': 'CUS001', 'Customer Name': 'HUI', 'Status': 'ACTIVE',
    'Pending in PayPal': 1419.58, 'Outstanding Amount': 0.00,
    'Invoiced Amount': 4129.19, 'Available Amount': 4129.19 },
  { 'Customer Key': 'CUS002', 'Customer Name': 'DITECH', 'Status': 'ACTIVE',
    'Pending in PayPal': 295.68, 'Outstanding Amount': 0.00,
    'Invoiced Amount': 1740.14, 'Available Amount': 1740.14 },
  { 'Customer Key': 'CUS003', 'Customer Name': 'ASON', 'Status': 'ACTIVE',
    'Pending in PayPal': 25.90, 'Outstanding Amount': 684.39,
    'Invoiced Amount': 0.00, 'Available Amount': 684.39 }
];

const PAYOUTS = [
  { 'Status': 'DRAFT', 'Amount': 700.00 }, { 'Status': 'DRAFT', 'Amount': 575.61 },
  { 'Status': 'PAID', 'Amount': 5572.02 }
];

const ACCOUNTS = [
  { 'Provider': 'PAYPAL', 'Account Key': 'PP1', 'Hold Balance': 0 },
  { 'Provider': 'PAYPAL', 'Account Key': 'PP2', 'Hold Balance': 1589.43 },
  { 'Provider': 'GREY', 'Account Key': 'GR001', 'Hold Balance': 999 }
];

const USDT_A1 = 'TỔNG USDT TRONG HỆ THỐNG · từ PayPal 6652.08 · đã chi 5574.02 ' +
  '· CÒN LẠI 1078.06   (chỉ tính phần hệ thống ghi nhận, không phải số dư ví thật)';

function build(overrides) {
  const data = Object.assign({
    'Customers': CUSTOMERS, 'Payouts': PAYOUTS, 'Account': ACCOUNTS,
    'Customer Allocation': [], 'Customer Ledger': [], 'Transaction': [],
    'Paypal': [], 'Grey': []
  }, overrides || {});

  const usdtText = (overrides && overrides.__usdt !== undefined) ? overrides.__usdt : USDT_A1;

  const sb = {
    console,
    Logger: { log: () => {} },
    Utilities: {
      formatDate: (d, tz, fmt) => {
        const p = n => String(n).padStart(2, '0');
        if (fmt === 'yyyy-MM-dd') return d.getFullYear() + '-' + p(d.getMonth() + 1) + '-' + p(d.getDate());
        if (fmt === 'H') return String(d.getHours());
        return d.toISOString();
      }
    },
    SpreadsheetApp: { getActive: () => ss },
    Date, Math, JSON, Object, Array, String, Number, Error, RegExp, isNaN, parseFloat, parseInt
  };

  const ss = {
    getSheetByName: name => {
      if (name === 'USDT') {
        return usdtText === null ? null
          : { getRange: () => ({ getValue: () => usdtText }) };
      }
      return data[name] ? { __name: name } : null;
    }
  };

  vm.createContext(sb);
  ['Flow.js', 'Customer.gs.js', 'Triggers.js', 'Dashboard.js'].forEach(f =>
    vm.runInContext(fs.readFileSync(path.join(REPO, f), 'utf8'), sb, { filename: f }));

  /*
   * Stub SAU khi nạp file: khai báo function ghi đè thuộc tính sandbox.
   *
   * custHeaderMap_ thật trả về map 1-BASED (m[k]=i+1), nên guard kiểu
   * `if (h['Status'])` an toàn cả với cột A. Stub phải bắt chước đúng vậy,
   * nếu không thì guard rớt và test báo lỗi giả.
   */
  sb.__data = data;
  vm.runInContext(
    'custHeaderMap_ = function(sheet){' +
    '  var rows = __data[sheet.__name] || [], m = {};' +
    '  if (rows.length) Object.keys(rows[0]).forEach(function(k,i){ m[k]=i+1; });' +
    '  return m; };' +
    'custReadRowsObjects_ = function(sheet){ return __data[sheet.__name] || []; };',
    sb);

  return sb;
}

/* ════════════════════════════════════════════════════════ */
console.log('\n1 · dashCollectMetrics_ · ba rổ không chồng lấn');
{
  const sb = build();
  const m = vm.runInContext('dashCollectMetrics_(SpreadsheetApp.getActive())', sb);

  eq('kẹt ở PayPal', m.pending, 1741.16);
  eq('đã về, chưa xuất HĐ', m.outstanding, 684.39);
  eq('đã xuất hoá đơn', m.invoiced, 5869.33);
  eq('available', m.available, 6553.72);

  /* Bất biến: Outstanding nghĩa là available TRỪ invoiced, nên hai rổ sau
     cộng lại phải bằng đúng available. Sai chỗ này là biểu đồ đếm hai lần. */
  eq('outstanding + invoiced = available', m.outstanding + m.invoiced, m.available);

  eq('USDT đọc từ A1', m.usdt, 1078.06);
  eq('payout DRAFT', m.draft, 1275.61);
  eq('payout PAID', m.paid, 5572.02);
  eq('Hold Balance chỉ cộng PAYPAL', m.holdBalance, 1589.43);

  eq('3 khách', m.customers.length, 3);
  eq('sắp xếp giảm dần, HUI đứng đầu', m.customers[0].name, 'HUI');
  eq('tổng của HUI', m.customers[0].total, 5548.77);
  eq('ASON đứng cuối', m.customers[2].name, 'ASON');
}

console.log('\n2 · cảnh báo DRAFT vượt USDT');
{
  const sb = build();
  const m = vm.runInContext('dashCollectMetrics_(SpreadsheetApp.getActive())', sb);
  const hit = m.alerts.filter(a => a.indexOf('Payout DRAFT') === 0);
  eq('có đúng 1 cảnh báo DRAFT', hit.length, 1);
  ok('nêu đúng số thiếu 197,55', hit[0] && hit[0].indexOf('197,55') !== -1, hit[0]);
}
{
  /* DRAFT thấp hơn USDT thì phải im */
  const sb = build({ 'Payouts': [{ 'Status': 'DRAFT', 'Amount': 500 }] });
  const m = vm.runInContext('dashCollectMetrics_(SpreadsheetApp.getActive())', sb);
  eq('DRAFT < USDT → không cảnh báo', m.alerts.filter(a => a.indexOf('Payout DRAFT') === 0).length, 0);
}
{
  /* Bằng đúng nhau cũng không được kêu */
  const sb = build({ 'Payouts': [{ 'Status': 'DRAFT', 'Amount': 1078.06 }] });
  const m = vm.runInContext('dashCollectMetrics_(SpreadsheetApp.getActive())', sb);
  eq('DRAFT = USDT → không cảnh báo', m.alerts.filter(a => a.indexOf('Payout DRAFT') === 0).length, 0);
}
{
  /* Chưa chạy syncFlow thì A1 chưa có dòng tổng — không được đoán bừa */
  const sb = build({ __usdt: 'chưa có gì' });
  const m = vm.runInContext('dashCollectMetrics_(SpreadsheetApp.getActive())', sb);
  eq('A1 không đọc được → usdt null', m.usdt, null);
  eq('và không cảnh báo bừa', m.alerts.filter(a => a.indexOf('Payout DRAFT') === 0).length, 0);
}

console.log('\n3 · dashDateKey_ · ô ngày là Date hay chuỗi đều được');
{
  const sb = build();
  const k = v => { sb.__v = v; return vm.runInContext('dashDateKey_(__v)', sb); };
  eq('Date', k(new Date(2026, 8, 2, 9, 49)), '2026-09-02');
  eq('chuỗi yyyy-MM-dd', k('2026-09-02'), '2026-09-02');
  eq('chuỗi có giờ kèm', k('2026-09-02 09:49:00'), '2026-09-02');
  eq('rỗng', k(''), '');
  eq('null', k(null), '');
}

console.log('\n4 · dashUsdtRemaining_ · số âm và định dạng lạ');
{
  const cases = [
    ['CÒN LẠI 1078.06 (ghi chú)', 1078.06],
    ['... · CÒN LẠI 0.00   (x)', 0],
    ['CÒN LẠI -12.50', -12.5],
    ['không có từ khoá', null],
    ['', null]
  ];
  cases.forEach(c => {
    const sb = build({ __usdt: c[0] });
    const v = vm.runInContext('dashUsdtRemaining_(SpreadsheetApp.getActive())', sb);
    eq(JSON.stringify(c[0].substring(0, 24)), v, c[1]);
  });
}

console.log('\n' + (fail ? '✗ ' + fail + ' FAIL' : '✓ tất cả') + ' · ' + pass + ' pass\n');
process.exit(fail ? 1 : 0);
