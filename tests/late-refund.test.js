/*
 * LATE_REFUND_ADJUSTMENT — chống tái diễn bug trừ oan tiền khách.
 *
 * Bug: dòng SALE ghi `PayPal Net Amount` = net ĐẦY ĐỦ của sale, nên một sale
 * chẻ qua hai hoá đơn sinh hai dòng cùng mang trọn net. Code cũ CỘNG chúng
 * lại, ra invoicedNet gấp đôi, và totalLate = 2×net − net = net — dựng ra
 * một khoản hoàn tiền bằng cả sale dù Refund Amount = 0.
 *
 * Số dùng ở đây lấy nguyên từ workbook ngày 24/09/2026.
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

/* ---- dựng sandbox, stub ba hàm đọc sheet ---- */
function build(sheets) {
  const sb = {
    console, Logger: { log: () => {} },
    Utilities: { formatDate: () => '' },
    SpreadsheetApp: { getActive: () => ({}) },
    Date, Math, JSON, Object, Array, String, Number, Error, RegExp,
    isNaN, parseFloat, parseInt
  };
  vm.createContext(sb);
  ['Flow.js', 'Customer.gs.js'].forEach(f =>
    vm.runInContext(fs.readFileSync(path.join(REPO, f), 'utf8'), sb, { filename: f }));

  /* Stub SAU khi nạp file: khai báo hàm ghi đè thuộc tính sandbox. */
  sb.__sheets = sheets;
  vm.runInContext(
    'custRequireSheet_ = function(ss, name){ return {__name:name}; };' +
    'custHeaderMap_ = function(sheet){ return {__sheet: sheet.__name}; };' +
    'custReadRowsObjects_ = function(sheet){ return __sheets[sheet.__name] || []; };', sb);
  return sb;
}

function state(sheets, ledgerRows) {
  const sb = build(sheets);
  sb.__ledgers = ledgerRows;
  return vm.runInContext('custBuildRefundOwedState_({}, __ledgers)', sb);
}

/* ---- tiện dựng dữ liệu ---- */
const LED = (key, gross, fee, refund) => ({
  'Ledger Key': key, 'Customer Key': 'CUS001', 'Currency': 'USD',
  'Gross Amount': gross, 'PayPal Fee': fee, 'Refund Amount': refund,
  'Net Amount': Math.max(0, +(gross - fee - refund).toFixed(2)),
  'PayPal Transaction ID': key.split('|').pop()
});
const SALE = (setKey, ledgerKey, net, included, refund, createdAt) => ({
  'Settlement Key': setKey, 'Customer Ledger Key': ledgerKey, 'Line Type': 'SALE',
  'PayPal Net Amount': net, 'Amount Included': included,
  'Refund Amount': refund || 0, 'Created At': createdAt || '2026-09-01'
});
const ADJ = (setKey, ledgerKey, take) => ({
  'Settlement Key': setKey, 'Customer Ledger Key': ledgerKey,
  'Line Type': 'LATE_REFUND_ADJUSTMENT',
  'Refund Amount': take, 'Amount Included': -take, 'Created At': '2026-09-20'
});
const SETS = keys => keys.map(k => ({ 'Settlement Key': k, 'Status': 'ISSUED', 'Issued At': '2026-09-01' }));

/* ══════════════════════════════════════════════════════════ */
console.log('\n1 · ca thật gây bug — sale chẻ qua hai hoá đơn, KHÔNG refund');
{
  /* CL|PP7|6LS91823D47072213 · gross 101,98 · fee 3,24 · refund 0 · net 98,74
     SET-000015 included 75,09 · SET-000016 included 21,75 (tổng 96,84;
     chênh 1,90 là phí Grey nhận, hoàn toàn bình thường). */
  const LK = 'CL|PP7|6LS91823D47072213';
  const st = state({
    'Settlements': SETS(['SET-000015', 'SET-000016']),
    'Settlement Items': [
      SALE('SET-000015', LK, 98.74, 75.09, 0, '2026-09-18'),
      SALE('SET-000016', LK, 98.74, 21.75, 0, '2026-09-22')
    ]
  }, [LED(LK, 101.98, 3.24, 0)]);

  ok('KHÔNG dựng late refund nào', !st[LK], JSON.stringify(st[LK] || {}));
}
{
  /* Hai ca còn lại trên sheet thật */
  [['CL|PP5|8TE50299KR4673314', 46.98, 2.13, 44.85, 15.02, 27.61],
   ['CL|PP7|7R9031746M3443201', 32.98, 2.14, 30.84, 1.98, 28.02]].forEach(c => {
    const st = state({
      'Settlements': SETS(['A', 'B']),
      'Settlement Items': [
        SALE('A', c[0], c[3], c[4], 0, '2026-09-01'),
        SALE('B', c[0], c[3], c[5], 0, '2026-09-10')
      ]
    }, [LED(c[0], c[1], c[2], 0)]);
    ok(c[0] + ' sạch', !st[c[0]], JSON.stringify(st[c[0]] || {}));
  });
}

console.log('\n2 · refund THẬT vẫn phải bị thu lại');
{
  /* Sale 62,97 lên hoá đơn khi chưa refund, sau đó refund toàn bộ.
     Ảnh chụp đầu = 30,28 (net lúc đó); trueNet giờ = 62,97−32,69−62,97. */
  const LK = 'CL|PP1|8WM41424GK0172044';
  const st = state({
    'Settlements': SETS(['SET-000002']),
    'Settlement Items': [SALE('SET-000002', LK, 30.28, 30.28, 0, '2026-08-10')]
  }, [LED(LK, 62.97, 32.69, 62.97)]);

  ok('có dựng late refund', !!st[LK]);
  eq('totalLate = ảnh chụp − trueNet', st[LK].totalLate, 62.97);
  eq('chưa thu đồng nào', st[LK].consumed, 0);
  eq('còn phải thu', st[LK].outstanding, 62.97);
}

console.log('\n3 · refund một phần, rồi refund thêm');
{
  const LK = 'CL|PPX|PARTIAL';
  /* Ảnh chụp lúc đầu: net 100. Refund 30 → totalLate 30. */
  const a = state({
    'Settlements': SETS(['S1']),
    'Settlement Items': [SALE('S1', LK, 100, 100, 0, '2026-08-01')]
  }, [LED(LK, 110, 10, 30)]);
  eq('lần đầu thu 30', a[LK].totalLate, 30);
  eq('outstanding 30', a[LK].outstanding, 30);

  /* Đã thu 30 ở S2, giờ refund lên 50 → chỉ thu thêm 20. */
  const b = state({
    'Settlements': SETS(['S1', 'S2']),
    'Settlement Items': [
      SALE('S1', LK, 100, 100, 0, '2026-08-01'),
      ADJ('S2', LK, 30)
    ]
  }, [LED(LK, 110, 10, 50)]);
  eq('totalLate lên 50', b[LK].totalLate, 50);
  eq('đã thu 30', b[LK].consumed, 30);
  eq('chỉ thu thêm 20', b[LK].outstanding, 20);
  eq('không thu thừa', b[LK].overConsumed, 0);
}

console.log('\n4 · đã thu đủ thì không thu nữa');
{
  const LK = 'CL|PPX|DONE';
  const st = state({
    'Settlements': SETS(['S1', 'S2']),
    'Settlement Items': [
      SALE('S1', LK, 100, 100, 0, '2026-08-01'),
      ADJ('S2', LK, 40)
    ]
  }, [LED(LK, 110, 10, 40)]);
  eq('totalLate 40', st[LK].totalLate, 40);
  eq('consumed 40', st[LK].consumed, 40);
  eq('outstanding 0 — không thu lần hai', st[LK].outstanding, 0);
}

console.log('\n5 · phần đã trừ oan phải lộ ra ở overConsumed');
{
  /* Đúng hiện trạng sheet sau khi sửa: adj 98,74 đã nằm trong hoá đơn cũ,
     nhưng totalLate giờ = 0. Chênh phải hiện ra để còn biết mà trả lại. */
  const LK = 'CL|PP7|6LS91823D47072213';
  const st = state({
    'Settlements': SETS(['SET-000015', 'SET-000016', 'SET-000017']),
    'Settlement Items': [
      SALE('SET-000015', LK, 98.74, 75.09, 0, '2026-09-18'),
      SALE('SET-000016', LK, 98.74, 21.75, 0, '2026-09-22'),
      ADJ('SET-000017', LK, 98.74)
    ]
  }, [LED(LK, 101.98, 3.24, 0)]);

  ok('vẫn giữ bản ghi vì đã có consumed', !!st[LK]);
  eq('totalLate về 0', st[LK].totalLate, 0);
  eq('consumed 98,74', st[LK].consumed, 98.74);
  eq('outstanding 0 — không trừ tiếp', st[LK].outstanding, 0);
  eq('overConsumed nêu đúng số trừ oan', st[LK].overConsumed, 98.74);
}

console.log('\n6 · sale bình thường một hoá đơn, không refund → im lặng');
{
  const LK = 'CL|PPX|NORMAL';
  const st = state({
    'Settlements': SETS(['S1']),
    'Settlement Items': [SALE('S1', LK, 100, 98, 0, '2026-08-01')]
  }, [LED(LK, 110, 10, 0)]);
  ok('không dựng gì', !st[LK], JSON.stringify(st[LK] || {}));
}
{
  /* Chưa từng lên hoá đơn */
  const LK = 'CL|PPX|NEVER';
  const st = state({ 'Settlements': [], 'Settlement Items': [] },
    [LED(LK, 110, 10, 40)]);
  ok('chưa invoice thì không phải late refund', !st[LK]);
}

console.log('\n' + (fail ? '✗ ' + fail + ' FAIL' : '✓ tất cả') + ' · ' + pass + ' pass\n');
process.exit(fail ? 1 : 0);
