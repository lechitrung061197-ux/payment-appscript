/*
 * Đối chứng bốn thay đổi bằng chính số thật lấy từ workbook.
 * Chạy: node test_fix.js
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

function load(files) {
  const sb = {
    console,
    Logger: { log: () => {} },
    Date, Math, JSON, Object, Array, String, Number, Error, RegExp, isNaN, parseFloat, parseInt
  };
  vm.createContext(sb);
  files.forEach(f => vm.runInContext(fs.readFileSync(path.join(REPO, f), 'utf8'), sb, { filename: f }));
  return sb;
}

/* ════════════════════════════════════════════════════════════════
 * 1 · flowPaypalType_ — mã đã đặt tên, và mã cố ý để nguyên
 * ════════════════════════════════════════════════════════════════ */
console.log('\n1 · flowPaypalType_');
{
  const sb = load(['Flow.js']);
  const t = c => vm.runInContext('flowPaypalType_(' + JSON.stringify(c) + ')', sb);

  eq('T1101 → REVERSAL', t('T1101'), 'REVERSAL');
  eq('T0200 → CURRENCY_CONVERSION', t('T0200'), 'CURRENCY_CONVERSION');
  eq('T1110 → HOLD', t('T1110'), 'HOLD');
  eq('T1105 → HOLD_RELEASE', t('T1105'), 'HOLD_RELEASE');
  eq('T1201 → CHARGEBACK', t('T1201'), 'CHARGEBACK');
  eq('T1501 → HOLD', t('T1501'), 'HOLD');
  eq('T0021 CỐ Ý để UNMAPPED_ (chưa biết nghĩa)', t('T0021'), 'UNMAPPED_T0021');
  eq('T0400 không đổi', t('T0400'), 'BANK_WITHDRAWAL');
  eq('T0006 không đổi', t('T0006'), 'EXPRESS_CHECKOUT_PAYMENT');
}

/* ════════════════════════════════════════════════════════════════
 * 2 · PASS 2 — dòng trả về nối bằng Transaction ID chung
 * ════════════════════════════════════════════════════════════════ */
console.log('\n2 · flowSyncPaypalToTransaction_ · PASS 2');
{
  const sb = load(['Flow.js']);

  const PH = ['Account Key','Label','Initiation Date','Transaction ID','Reference ID',
    'Event Code','Type','Status','Transaction Amount','Fee Amount','Currency','Unique Key'];
  const ph = {}; PH.forEach((h, i) => ph[h] = i);

  const TH = ['Transaction Key','Transaction Date','Provider','Account Key','Label','Type',
    'Direction','Gross Amount','Net Amount','Fee Amount','Currency','Reference',
    'Counterparty Provider','Counterparty Account Key','Reconcile Status',
    'Matched Transaction Key','Source','Source ID','Sync Time','Unique Key'];
  const th = {}; TH.forEach((h, i) => th[h] = i);

  const D = new Date(2026, 8, 1, 7, 11, 40);

  function row(o) {
    const r = new Array(PH.length).fill('');
    Object.keys(o).forEach(k => r[ph[k]] = o[k]);
    return r;
  }

  function run(rows) {
    const state = { rows: [], keyMap: {}, columnCount: TH.length };
    const accountMap = { PP2: { linkedProvider: 'GREY', linkedAccountKey: 'GR002' } };
    sb.__rows = rows; sb.__ph = ph; sb.__am = accountMap; sb.__st = state; sb.__th = th;
    const res = vm.runInContext(
      'flowSyncPaypalToTransaction_(__rows, __ph, __am, __st, __th)', sb);
    const byKey = {};
    state.rows.forEach(r => byKey[r[th['Transaction Key']]] = r);
    return { res, state, byKey };
  }

  /* Ca thật: T0400 PENDING + T1101 DENIED, cùng Transaction ID, Reference ID rỗng */
  {
    const { res, byKey } = run([
      row({ 'Account Key':'PP2','Initiation Date':D,'Transaction ID':'8S221609XN298812P',
            'Reference ID':'','Event Code':'T0400','Type':'BANK_WITHDRAWAL','Status':'PENDING',
            'Transaction Amount':-860.4,'Fee Amount':0,'Currency':'USD','Unique Key':'u1' }),
      row({ 'Account Key':'PP2','Initiation Date':D,'Transaction ID':'8S221609XN298812P',
            'Reference ID':'','Event Code':'T1101','Type':'REVERSAL','Status':'DENIED',
            'Transaction Amount':860.4,'Fee Amount':0,'Currency':'USD','Unique Key':'u2' })
    ]);

    eq('sinh đúng 1 dòng trả về', res.returns, 1);
    const ret = byKey['PRT|PP2|8S221609XN298812P'];
    ok('có dòng PRT', !!ret);
    if (ret) {
      eq('Type', ret[th['Type']], 'PAYPAL_WITHDRAWAL_RETURN');
      eq('Direction', ret[th['Direction']], 'IN');
      eq('Gross', ret[th['Gross Amount']], 860.4);
      /* Điểm mấu chốt: matcher ghép bằng Reference, nên phải là ID khoản rút */
      eq('Reference = Transaction ID khoản rút', ret[th['Reference']], '8S221609XN298812P');
    }
    const wd = byKey['PTX|PP2|8S221609XN298812P'];
    eq('khoản rút vẫn IN_TRANSIT ở PASS 1 (matcher mới đặt RETURNED)',
       wd[th['Reconcile Status']], 'IN_TRANSIT');
  }

  /* Chống dương tính giả: cùng Transaction ID nhưng KHÁC số tiền */
  {
    const { res } = run([
      row({ 'Account Key':'PP2','Initiation Date':D,'Transaction ID':'X1',
            'Reference ID':'','Event Code':'T0400','Type':'BANK_WITHDRAWAL','Status':'PENDING',
            'Transaction Amount':-860.4,'Fee Amount':0,'Currency':'USD','Unique Key':'u1' }),
      row({ 'Account Key':'PP2','Initiation Date':D,'Transaction ID':'X1',
            'Reference ID':'','Event Code':'T1101','Type':'REVERSAL','Status':'DENIED',
            'Transaction Amount':100.00,'Fee Amount':0,'Currency':'USD','Unique Key':'u2' })
    ]);
    eq('số tiền lệch → KHÔNG tạo dòng trả về', res.returns, 0);
  }

  /* Nhánh Reference ID cũ phải y nguyên */
  {
    const { res, byKey } = run([
      row({ 'Account Key':'PP2','Initiation Date':D,'Transaction ID':'W9',
            'Reference ID':'','Event Code':'T0400','Type':'BANK_WITHDRAWAL','Status':'SUCCESS',
            'Transaction Amount':-500,'Fee Amount':0,'Currency':'USD','Unique Key':'u1' }),
      row({ 'Account Key':'PP2','Initiation Date':D,'Transaction ID':'R9',
            'Reference ID':'W9','Event Code':'T1106','Type':'PAYMENT_REVERSAL','Status':'SUCCESS',
            'Transaction Amount':500,'Fee Amount':0,'Currency':'USD','Unique Key':'u2' })
    ]);
    eq('nhánh Reference ID vẫn chạy', res.returns, 1);
    eq('Reference giữ nguyên', byKey['PRT|PP2|R9'][th['Reference']], 'W9');
  }

  /* Khoản rút bình thường, không có dòng đảo nào */
  {
    const { res, byKey } = run([
      row({ 'Account Key':'PP2','Initiation Date':D,'Transaction ID':'OK1',
            'Reference ID':'','Event Code':'T0400','Type':'BANK_WITHDRAWAL','Status':'SUCCESS',
            'Transaction Amount':-438.55,'Fee Amount':0,'Currency':'USD','Unique Key':'u1' })
    ]);
    eq('không sinh dòng trả về', res.returns, 0);
    eq('khoản rút vẫn IN_TRANSIT', byKey['PTX|PP2|OK1'][th['Reconcile Status']], 'IN_TRANSIT');
  }
}

/* ════════════════════════════════════════════════════════════════
 * 3 · custResolveCurrencyConversion_ — số thật đơn JUST-2205255265
 * ════════════════════════════════════════════════════════════════ */
console.log('\n3 · custResolveCurrencyConversion_ · đơn JUST-2205255265');
{
  const sb = load(['Customer.gs.js']);

  const root = { currency: 'GBP', transactionAmount: 99.95, feeAmount: -5.38 };
  const events = [
    root,
    { eventCode: 'T0200', currency: 'GBP', transactionAmount: -94.57, feeAmount: 0 },
    { eventCode: 'T0200', currency: 'USD', transactionAmount: 124.90, feeAmount: 0 }
  ];

  sb.__root = root; sb.__events = events;
  const conv = vm.runInContext('custResolveCurrencyConversion_(__root, __events)', sb);

  ok('tìm được cặp quy đổi', !!conv);
  if (conv) {
    eq('tiền đích USD', conv.currency, 'USD');
    eq('tỷ giá 124,90 ÷ 94,57', conv.rate, 1.3207147, 1e-6);

    const r2 = v => Math.round(v * 100) / 100;
    const gross = r2(99.95 * conv.rate);
    const fee = r2(5.38 * conv.rate);
    const net = r2(gross - fee);

    eq('gross → 132,01 USD', gross, 132.01);
    eq('phí PayPal → 7,11 USD', fee, 7.11);
    /* Chốt chặn: net phải bằng ĐÚNG số PayPal đo được ở chân dương */
    eq('net = 124,90 (khớp chân dương T0200)', net, 124.90);
    eq('thiếu bao nhiêu nếu không sửa', r2(net - (99.95 - 5.38)), 30.33);
  }

  /* Thiếu chân → không đổi, giữ nguyên cách cũ */
  sb.__events = [root, { eventCode: 'T0200', currency: 'GBP', transactionAmount: -94.57 }];
  eq('chỉ có chân âm → null', vm.runInContext('custResolveCurrencyConversion_(__root, __events)', sb), null);

  sb.__events = [root];
  eq('không có T0200 → null', vm.runInContext('custResolveCurrencyConversion_(__root, __events)', sb), null);

  /* Giao dịch USD thuần không bị đụng */
  sb.__root = { currency: 'USD', transactionAmount: 57.97 };
  sb.__events = [sb.__root];
  eq('đơn USD thuần → null', vm.runInContext('custResolveCurrencyConversion_(__root, __events)', sb), null);
}

/* ════════════════════════════════════════════════════════════════
 * 4 · END-TO-END: PASS 2 → flowReconcilePaypalReturns_ → RETURNED
 *     Đây là mắt xích quyết định. Không có nó thì cả bản sửa vô nghĩa.
 * ════════════════════════════════════════════════════════════════ */
console.log('\n4 · end-to-end · khoản rút bị từ chối phải thành RETURNED');
{
  const sb = load(['Flow.js']);

  const PH = ['Account Key','Label','Initiation Date','Transaction ID','Reference ID',
    'Event Code','Type','Status','Transaction Amount','Fee Amount','Currency','Unique Key'];
  const ph = {}; PH.forEach((h, i) => ph[h] = i);
  const TH = ['Transaction Key','Transaction Date','Provider','Account Key','Label','Type',
    'Direction','Gross Amount','Net Amount','Fee Amount','Currency','Reference',
    'Counterparty Provider','Counterparty Account Key','Reconcile Status',
    'Matched Transaction Key','Source','Source ID','Sync Time','Unique Key'];
  const th = {}; TH.forEach((h, i) => th[h] = i);

  function row(o) { const r = new Array(PH.length).fill(''); Object.keys(o).forEach(k => r[ph[k]] = o[k]); return r; }
  const D1 = new Date(2026, 8, 1, 7, 11, 40);
  const D2 = new Date(2026, 7, 31, 7, 58, 12);
  const D3 = new Date(2026, 7, 28, 7, 26, 17);

  const rows = [
    /* hai khoản bị từ chối — số thật */
    row({ 'Account Key':'PP2','Initiation Date':D1,'Transaction ID':'8S221609XN298812P','Reference ID':'',
          'Event Code':'T0400','Type':'BANK_WITHDRAWAL','Status':'PENDING','Transaction Amount':-860.4,
          'Fee Amount':0,'Currency':'USD','Unique Key':'a1' }),
    row({ 'Account Key':'PP2','Initiation Date':D1,'Transaction ID':'8S221609XN298812P','Reference ID':'',
          'Event Code':'T1101','Type':'REVERSAL','Status':'DENIED','Transaction Amount':860.4,
          'Fee Amount':0,'Currency':'USD','Unique Key':'a2' }),
    row({ 'Account Key':'PP2','Initiation Date':D2,'Transaction ID':'86F559825B1562918','Reference ID':'',
          'Event Code':'T0400','Type':'BANK_WITHDRAWAL','Status':'PENDING','Transaction Amount':-542.43,
          'Fee Amount':0,'Currency':'USD','Unique Key':'b1' }),
    row({ 'Account Key':'PP2','Initiation Date':D2,'Transaction ID':'86F559825B1562918','Reference ID':'',
          'Event Code':'T1101','Type':'REVERSAL','Status':'DENIED','Transaction Amount':542.43,
          'Fee Amount':0,'Currency':'USD','Unique Key':'b2' }),
    /* khoản đã về Grey thật — CHỐT CHẶN, không được đụng */
    row({ 'Account Key':'PP2','Initiation Date':D3,'Transaction ID':'97H54784RJ853401X','Reference ID':'',
          'Event Code':'T0400','Type':'BANK_WITHDRAWAL','Status':'SUCCESS','Transaction Amount':-438.55,
          'Fee Amount':0,'Currency':'USD','Unique Key':'c1' })
  ];

  const state = { rows: [], keyMap: {}, columnCount: TH.length };
  sb.__rows = rows; sb.__ph = ph; sb.__st = state; sb.__th = th;
  sb.__am = { PP2: { linkedProvider: 'GREY', linkedAccountKey: 'GR002' } };
  vm.runInContext('flowSyncPaypalToTransaction_(__rows, __ph, __am, __st, __th)', sb);

  /* khoản 97H đã ARRIVED từ trước, mô phỏng đúng trạng thái workbook */
  state.rows.forEach(r => {
    if (r[th['Transaction Key']] === 'PTX|PP2|97H54784RJ853401X') r[th['Reconcile Status']] = 'ARRIVED';
  });

  sb.__trows = state.rows;
  const matched = vm.runInContext('flowReconcilePaypalReturns_(__trows, __th)', sb);

  const byKey = {};
  state.rows.forEach(r => byKey[r[th['Transaction Key']]] = r);

  eq('ghép được 2 cặp', matched, 2);
  eq('860,40 → RETURNED', byKey['PTX|PP2|8S221609XN298812P'][th['Reconcile Status']], 'RETURNED');
  eq('542,43 → RETURNED', byKey['PTX|PP2|86F559825B1562918'][th['Reconcile Status']], 'RETURNED');
  ok('có cross-link sang dòng trả về',
     byKey['PTX|PP2|8S221609XN298812P'][th['Matched Transaction Key']] === 'PRT|PP2|8S221609XN298812P');
  eq('dòng trả về → MATCHED', byKey['PRT|PP2|8S221609XN298812P'][th['Reconcile Status']], 'MATCHED');

  /* CHỐT CHẶN */
  eq('438,55 vẫn ARRIVED — thay đổi KHÔNG lan', byKey['PTX|PP2|97H54784RJ853401X'][th['Reconcile Status']], 'ARRIVED');
  ok('không sinh dòng trả về thừa cho khoản đã tới', !byKey['PRT|PP2|97H54784RJ853401X']);

  const returned = state.rows.filter(r => r[th['Reconcile Status']] === 'RETURNED')
    .reduce((s, r) => s + r[th['Net Amount']], 0);
  eq('tổng tiền được gỡ kẹt', Math.round(returned * 100) / 100, 1402.83);
}

console.log('\n' + (fail ? '✗ ' + fail + ' FAIL' : '✓ tất cả') + ' · ' + pass + ' pass\n');
process.exit(fail ? 1 : 0);
