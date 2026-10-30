/* Tiền kế hoạch — tổng hợp dữ liệu đầu vào → bảng tiền kế hoạch → kế hoạch sản xuất.
   Dữ liệu lưu cục bộ (IndexedDB) để chạy offline; đồng bộ đám mây tuỳ chọn qua Supabase. */
'use strict';

/* ================= Danh sách chọn nhanh cố định ================= */
const REASONS = ['Khuyến mãi', 'Tung hàng mới', 'Sự kiện / Hội chợ', 'Điều chỉnh của Sales', 'Đơn hàng lớn', 'Thu hồi / Đổi hàng', 'Khác'];
const MAT_TYPES = ['NVL', 'Phụ gia', 'Bao bì'];
const PLAN_STATUS = ['Nháp', 'Đã duyệt', 'Đang SX', 'Hoàn thành', 'Huỷ'];
const UNITS = ['lon', 'hộp', 'gói', 'chai', 'thùng', 'kg', 'g', 'lít', 'cái', 'cuộn'];

/* ================= Định nghĩa nguồn dữ liệu (mẫu chuẩn) ================= */
// alias: các tên cột thường gặp (không dấu) để tự nhận diện khi nhập Excel
const SOURCES = {
  forecast: {
    label: 'Forecast (WMS)', short: 'Forecast', key: 'forecast', file: 'Mau_Forecast_WMS.xlsx',
    hint: 'File Excel forecast do WMS cung cấp: số lượng dự báo theo kỳ, mã sản phẩm và kênh.',
    fields: [
      { k: 'period', label: 'Kỳ (tháng)', type: 'period', def: 'period', alias: ['ky', 'thang', 'period', 'month', 'ky ke hoach', 'thang ke hoach'] },
      { k: 'product', label: 'Mã SP', type: 'product', req: true, alias: ['ma sp', 'ma san pham', 'sku', 'item', 'item code', 'ma hang', 'product', 'product code', 'ma tp', 'ma thanh pham'] },
      { k: 'channel', label: 'Kênh', type: 'channel', def: 'channel', alias: ['kenh', 'channel', 'kenh ban', 'kenh phan phoi'] },
      { k: 'qty', label: 'SL dự báo', type: 'num', req: true, alias: ['sl', 'so luong', 'forecast', 'qty', 'quantity', 'du bao', 'sl du bao', 'fc'] },
    ],
    sample: [['2026-11', 'SP001', 'MT', 12000], ['2026-11', 'SP001', 'GT', 8000]],
  },
  fgStock: {
    label: 'Tồn kho thành phẩm', short: 'Tồn TP', key: 'fgStock', file: 'Mau_Ton_Thanh_Pham.xlsx',
    hint: 'Tồn kho thành phẩm khả dụng tại thời điểm chốt số liệu.',
    fields: [
      { k: 'product', label: 'Mã SP', type: 'product', req: true, alias: ['ma sp', 'ma san pham', 'sku', 'item', 'ma hang', 'product', 'ma tp'] },
      { k: 'warehouse', label: 'Kho', type: 'warehouse', def: 'warehouse', alias: ['kho', 'warehouse', 'ma kho', 'location'] },
      { k: 'qty', label: 'SL tồn', type: 'num', req: true, alias: ['ton', 'sl ton', 'so luong ton', 'ton kho', 'qty', 'stock', 'on hand', 'kha dung'] },
    ],
    sample: [['SP001', 'Kho TP Bình Dương', 5200]],
  },
  rmStock: {
    label: 'Tồn kho NVL - phụ gia', short: 'Tồn NVL', key: 'rmStock', file: 'Mau_Ton_NVL_Phu_Gia.xlsx',
    hint: 'Tồn nguyên vật liệu, phụ gia, bao bì và lượng đang về (PO đã đặt).',
    fields: [
      { k: 'material', label: 'Mã NVL/PG', type: 'material', req: true, alias: ['ma nvl', 'ma vat tu', 'ma nguyen lieu', 'material', 'material code', 'ma pg', 'ma phu gia', 'ma hang'] },
      { k: 'warehouse', label: 'Kho', type: 'warehouse', def: 'warehouse', alias: ['kho', 'warehouse', 'ma kho'] },
      { k: 'qty', label: 'SL tồn', type: 'num', req: true, alias: ['ton', 'sl ton', 'so luong ton', 'ton kho', 'qty', 'stock', 'on hand'] },
      { k: 'incoming', label: 'Đang về', type: 'num', alias: ['dang ve', 'sl dang ve', 'po', 'incoming', 'on order', 'hang dang ve'] },
    ],
    sample: [['NVL01', 'Kho NVL', 18000, 5000]],
  },
  adjust: {
    label: 'Cân chỉnh kênh', short: 'Cân chỉnh', key: 'adjust', file: 'Mau_Can_Chinh_Kenh.xlsx',
    hint: 'Điều chỉnh tăng (+) / giảm (−) số lượng theo kênh so với forecast. Không bắt buộc.',
    fields: [
      { k: 'period', label: 'Kỳ (tháng)', type: 'period', def: 'period', alias: ['ky', 'thang', 'period', 'month'] },
      { k: 'product', label: 'Mã SP', type: 'product', req: true, alias: ['ma sp', 'ma san pham', 'sku', 'item', 'ma hang', 'product'] },
      { k: 'channel', label: 'Kênh', type: 'channel', def: 'channel', alias: ['kenh', 'channel', 'kenh ban'] },
      { k: 'qty', label: 'SL điều chỉnh (±)', type: 'num', req: true, alias: ['dieu chinh', 'sl dieu chinh', 'can chinh', 'adjust', 'adjustment', 'qty', 'so luong', 'tang giam'] },
      { k: 'reason', label: 'Lý do', type: 'reason', alias: ['ly do', 'reason', 'nguyen nhan'] },
      { k: 'note', label: 'Ghi chú', type: 'text', alias: ['ghi chu', 'note', 'notes', 'dien giai'] },
    ],
    sample: [['2026-11', 'SP001', 'MT', 1500, 'Khuyến mãi', 'CTKM tháng 11']],
  },
  products: {
    label: 'Danh mục sản phẩm', short: 'Sản phẩm', key: 'products', file: 'Mau_Danh_Muc_SP.xlsx',
    fields: [
      { k: 'code', label: 'Mã SP', type: 'text', req: true, alias: ['ma sp', 'ma san pham', 'sku', 'ma hang', 'code'] },
      { k: 'name', label: 'Tên SP', type: 'text', req: true, alias: ['ten sp', 'ten san pham', 'ten hang', 'name', 'description'] },
      { k: 'unit', label: 'ĐVT', type: 'unit', alias: ['dvt', 'don vi', 'don vi tinh', 'unit', 'uom'] },
      { k: 'group', label: 'Nhóm', type: 'text', alias: ['nhom', 'nhom hang', 'group', 'category'] },
      { k: 'batch', label: 'Lô SX (SL/mẻ)', type: 'num', alias: ['lo', 'lo sx', 'batch', 'batch size', 'me', 'moq'] },
      { k: 'safetyDays', label: 'Tồn AT (ngày)', type: 'num', alias: ['ton an toan', 'safety', 'safety days', 'ngay ton an toan', 'ton at'] },
      { k: 'line', label: 'Line SX', type: 'line', alias: ['line', 'day chuyen', 'line sx'] },
    ],
    sample: [['SP001', 'Sữa bột trẻ em 900g', 'lon', 'Sữa bột', 2400, 10, 'L1']],
  },
  materials: {
    label: 'Danh mục NVL - phụ gia', short: 'NVL/PG', key: 'materials', file: 'Mau_Danh_Muc_NVL.xlsx',
    fields: [
      { k: 'code', label: 'Mã NVL/PG', type: 'text', req: true, alias: ['ma nvl', 'ma vat tu', 'ma pg', 'code', 'material'] },
      { k: 'name', label: 'Tên', type: 'text', req: true, alias: ['ten', 'ten nvl', 'ten vat tu', 'name', 'description'] },
      { k: 'type', label: 'Loại', type: 'matType', alias: ['loai', 'type', 'nhom'] },
      { k: 'unit', label: 'ĐVT', type: 'unit', alias: ['dvt', 'don vi', 'unit', 'uom'] },
    ],
    sample: [['NVL01', 'Sữa bột nguyên kem', 'NVL', 'kg']],
  },
  bom: {
    label: 'Định mức (BOM)', short: 'BOM', key: 'bom', file: 'Mau_Dinh_Muc_BOM.xlsx',
    fields: [
      { k: 'product', label: 'Mã SP', type: 'product', req: true, alias: ['ma sp', 'ma san pham', 'sku', 'product'] },
      { k: 'material', label: 'Mã NVL/PG', type: 'material', req: true, alias: ['ma nvl', 'ma vat tu', 'ma pg', 'material'] },
      { k: 'qty', label: 'Định mức / 1 ĐVT SP', type: 'num', req: true, alias: ['dinh muc', 'qty', 'so luong', 'usage', 'qty per'] },
    ],
    sample: [['SP001', 'NVL01', 0.62]],
  },
};
const INPUT_KEYS = ['forecast', 'fgStock', 'rmStock', 'adjust'];
const MASTER_KEYS = ['products', 'materials', 'bom', 'lists'];

/* ================= Dữ liệu mẫu (đánh dấu rõ là mẫu) ================= */
function sampleState() {
  const products = [
    { code: 'SP001', name: 'Sữa bột trẻ em 900g', unit: 'lon', group: 'Sữa bột', batch: 2400, safetyDays: 10, line: 'L1' },
    { code: 'SP002', name: 'Sữa bột trẻ em 400g', unit: 'lon', group: 'Sữa bột', batch: 3000, safetyDays: 10, line: 'L1' },
    { code: 'SP003', name: 'Ngũ cốc dinh dưỡng 500g', unit: 'hộp', group: 'Ngũ cốc', batch: 2000, safetyDays: 7, line: 'L2' },
    { code: 'SP004', name: 'Bột ăn dặm gạo sữa 200g', unit: 'hộp', group: 'Ăn dặm', batch: 2500, safetyDays: 7, line: 'L2' },
    { code: 'SP005', name: 'Sữa uống liền 180ml (thùng 48)', unit: 'thùng', group: 'Sữa nước', batch: 1500, safetyDays: 5, line: 'L3' },
  ];
  const materials = [
    { code: 'NVL01', name: 'Sữa bột nguyên kem', type: 'NVL', unit: 'kg' },
    { code: 'NVL02', name: 'Đường tinh luyện', type: 'NVL', unit: 'kg' },
    { code: 'NVL03', name: 'Bột ngũ cốc hỗn hợp', type: 'NVL', unit: 'kg' },
    { code: 'NVL04', name: 'Dầu thực vật', type: 'NVL', unit: 'kg' },
    { code: 'PG01', name: 'Premix vitamin & khoáng', type: 'Phụ gia', unit: 'kg' },
    { code: 'PG02', name: 'DHA dạng bột', type: 'Phụ gia', unit: 'kg' },
    { code: 'PG03', name: 'Hương vani', type: 'Phụ gia', unit: 'kg' },
    { code: 'BB01', name: 'Lon thiếc 900g', type: 'Bao bì', unit: 'cái' },
    { code: 'BB02', name: 'Lon thiếc 400g', type: 'Bao bì', unit: 'cái' },
    { code: 'BB03', name: 'Hộp giấy 500g', type: 'Bao bì', unit: 'cái' },
    { code: 'BB04', name: 'Hộp giấy 200g', type: 'Bao bì', unit: 'cái' },
    { code: 'BB05', name: 'Hộp giấy tiệt trùng 180ml', type: 'Bao bì', unit: 'cái' },
    { code: 'BB06', name: 'Thùng carton 48 hộp', type: 'Bao bì', unit: 'cái' },
  ];
  const bom = [
    ['SP001', 'NVL01', 0.62], ['SP001', 'NVL02', 0.08], ['SP001', 'NVL04', 0.12], ['SP001', 'PG01', 0.009], ['SP001', 'PG02', 0.004], ['SP001', 'BB01', 1],
    ['SP002', 'NVL01', 0.275], ['SP002', 'NVL02', 0.035], ['SP002', 'NVL04', 0.055], ['SP002', 'PG01', 0.004], ['SP002', 'PG02', 0.002], ['SP002', 'BB02', 1],
    ['SP003', 'NVL03', 0.36], ['SP003', 'NVL01', 0.08], ['SP003', 'NVL02', 0.05], ['SP003', 'PG01', 0.003], ['SP003', 'PG03', 0.001], ['SP003', 'BB03', 1],
    ['SP004', 'NVL03', 0.12], ['SP004', 'NVL01', 0.06], ['SP004', 'PG01', 0.002], ['SP004', 'BB04', 1],
    ['SP005', 'NVL01', 0.6], ['SP005', 'NVL02', 0.3], ['SP005', 'PG03', 0.01], ['SP005', 'BB05', 48], ['SP005', 'BB06', 1],
  ].map(([product, material, qty]) => ({ product, material, qty }));
  const fc = {
    SP001: { MT: 9000, GT: 7000, Online: 2500, B2B: 0 },
    SP002: { MT: 6000, GT: 6500, Online: 1800, B2B: 0 },
    SP003: { MT: 5000, GT: 3500, Online: 900, B2B: 1200 },
    SP004: { MT: 3800, GT: 2600, Online: 1500, B2B: 0 },
    SP005: { MT: 2200, GT: 1800, Online: 300, B2B: 900 },
  };
  const forecast = [];
  ['2026-11', '2026-12'].forEach((period, i) => {
    Object.entries(fc).forEach(([product, ch]) => Object.entries(ch).forEach(([channel, q]) => {
      if (q) forecast.push({ period, product, channel, qty: Math.round(q * (i ? 1.12 : 1)) });
    }));
  });
  return {
    meta: { sample: true, updated: {} },
    products, materials, bom,
    channels: ['MT', 'GT', 'Online', 'B2B'],
    warehouses: ['Kho TP Bình Dương', 'Kho TP Hà Nội', 'Kho NVL'],
    lines: [
      { code: 'L1', name: 'Line đóng lon', capacity: 6000 },
      { code: 'L2', name: 'Line đóng hộp giấy', capacity: 5000 },
      { code: 'L3', name: 'Line UHT', capacity: 1200 },
    ],
    forecast,
    fgStock: [
      { product: 'SP001', warehouse: 'Kho TP Bình Dương', qty: 4200 }, { product: 'SP001', warehouse: 'Kho TP Hà Nội', qty: 1600 },
      { product: 'SP002', warehouse: 'Kho TP Bình Dương', qty: 9800 }, { product: 'SP003', warehouse: 'Kho TP Bình Dương', qty: 2100 },
      { product: 'SP004', warehouse: 'Kho TP Hà Nội', qty: 8600 }, { product: 'SP005', warehouse: 'Kho TP Bình Dương', qty: 1300 },
    ],
    rmStock: [
      ['NVL01', 18000, 6000], ['NVL02', 5200, 0], ['NVL03', 5200, 0], ['NVL04', 3800, 0], ['PG01', 320, 0], ['PG02', 70, 0], ['PG03', 90, 0],
      ['BB01', 16000, 0], ['BB02', 6000, 10000], ['BB03', 14000, 0], ['BB04', 4000, 0], ['BB05', 200000, 120000], ['BB06', 7000, 0],
    ].map(([material, qty, incoming]) => ({ material, warehouse: 'Kho NVL', qty, incoming })),
    adjust: [
      { period: '2026-11', product: 'SP001', channel: 'MT', qty: 1500, reason: 'Khuyến mãi', note: 'CTKM 11.11' },
      { period: '2026-11', product: 'SP003', channel: 'Online', qty: 600, reason: 'Sự kiện / Hội chợ', note: 'Live sale' },
      { period: '2026-11', product: 'SP004', channel: 'GT', qty: -400, reason: 'Điều chỉnh của Sales', note: '' },
    ],
    plan: [],
    settings: { period: '2026-11', cloud: { url: '', key: '', workspace: 'nha-may-1', table: 'tkh_state', auto: false } },
  };
}
function emptyState() {
  const s = sampleState();
  return { ...s, meta: { sample: false, updated: {} }, products: [], materials: [], bom: [], forecast: [], fgStock: [], rmStock: [], adjust: [], plan: [] };
}

/* ================= Lưu trữ cục bộ (IndexedDB, dự phòng localStorage) ================= */
const DB = {
  db: null,
  open() {
    return new Promise((res) => {
      try {
        const r = indexedDB.open('tien-ke-hoach', 1);
        r.onupgradeneeded = () => r.result.createObjectStore('kv');
        r.onsuccess = () => { this.db = r.result; res(); };
        r.onerror = () => res();
      } catch { res(); }
    });
  },
  get(k) {
    return new Promise((res) => {
      if (!this.db) { try { res(JSON.parse(localStorage.getItem('tkh:' + k))); } catch { res(null); } return; }
      const q = this.db.transaction('kv').objectStore('kv').get(k);
      q.onsuccess = () => res(q.result ?? null); q.onerror = () => res(null);
    });
  },
  set(k, v) {
    return new Promise((res) => {
      if (!this.db) { try { localStorage.setItem('tkh:' + k, JSON.stringify(v)); } catch {} res(); return; }
      const t = this.db.transaction('kv', 'readwrite'); t.objectStore('kv').put(v, k);
      t.oncomplete = res; t.onerror = res;
    });
  },
};

let state;
const ui = { view: 'home', src: 'forecast', master: 'products', filter: {}, preFilter: 'all', importer: null };
let saveTimer;
function save(touchKey) {
  if (touchKey) { state.meta.updated[touchKey] = new Date().toISOString(); }
  clearTimeout(saveTimer);
  saveTimer = setTimeout(() => {
    DB.set('state', state);
    if (state.settings.cloud.auto && cloudReady()) cloudPush(true);
  }, 300);
}

/* ================= Tiện ích ================= */
const $ = (s, el = document) => el.querySelector(s);
const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const pad = (n) => String(n).padStart(2, '0');
const norm = (s) => String(s ?? '').normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[đĐ]/g, 'd').toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
const fmt = (n, d = 0) => (Number(n) || 0).toLocaleString('vi-VN', { maximumFractionDigits: d });
const fmtQ = (n) => { const v = Number(n) || 0; return fmt(v, Math.abs(v) < 100 && v % 1 ? 3 : 0); };
function num(v) {
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  let s = String(v ?? '').trim().replace(/\s/g, '');
  if (!s) return 0;
  if (/^-?\d{1,3}([.,]\d{3})+$/.test(s)) s = s.replace(/[.,]/g, '');
  else s = s.replace(',', '.');
  const n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}
function toPeriod(v) {
  if (v == null || v === '') return '';
  if (v instanceof Date && !isNaN(v)) return `${v.getFullYear()}-${pad(v.getMonth() + 1)}`;
  if (typeof v === 'number' && v > 20000 && v < 80000) { const d = new Date(Math.round((v - 25569) * 864e5)); return `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}`; }
  const s = String(v).trim();
  let m = s.match(/^(\d{4})[-/.](\d{1,2})/); if (m) return `${m[1]}-${pad(m[2])}`;
  m = s.match(/(\d{1,2})[-/.](\d{4})/); if (m) return `${m[2]}-${pad(m[1])}`;
  return s;
}
const periodLabel = (p) => { const m = /^(\d{4})-(\d{2})$/.exec(p); return m ? `T${+m[2]}/${m[1]}` : p; };
const sum = (arr, f) => arr.reduce((a, r) => a + num(f(r)), 0);
const byCode = (arr, k = 'code') => Object.fromEntries(arr.map((r) => [r[k], r]));
const relTime = (iso) => {
  if (!iso) return 'chưa cập nhật';
  const d = (Date.now() - new Date(iso)) / 864e5;
  if (d < 1 / 24) return 'vừa xong';
  if (d < 1) return `${Math.round(d * 24)} giờ trước`;
  return `${Math.round(d)} ngày trước`;
};
const dateStr = (d) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
function firstWorkday(period) {
  const d = new Date(`${period}-01T00:00:00`);
  if (isNaN(d)) return dateStr(new Date());
  while (d.getDay() === 0) d.setDate(d.getDate() + 1);
  return dateStr(d);
}
function addWorkdays(start, n) { // trả về ngày kết thúc sau n ngày làm việc (bỏ Chủ nhật)
  const d = new Date(`${start}T00:00:00`); if (isNaN(d) || n <= 0) return start;
  let left = n - 1;
  while (d.getDay() === 0) d.setDate(d.getDate() + 1);
  while (left > 0) { d.setDate(d.getDate() + 1); if (d.getDay() !== 0) left--; }
  return dateStr(d);
}
function workdaysInMonth(period) {
  const [y, m] = period.split('-').map(Number); if (!y || !m) return 26;
  let c = 0; const d = new Date(y, m - 1, 1);
  while (d.getMonth() === m - 1) { if (d.getDay() !== 0) c++; d.setDate(d.getDate() + 1); }
  return c;
}
const fmtDate = (s) => { const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s || ''); return m ? `${m[3]}/${m[2]}` : (s || '—'); };

/* ================= Danh sách xổ xuống ================= */
function periods() {
  const set = new Set(state.forecast.map((r) => r.period).concat(state.adjust.map((r) => r.period), state.plan.map((r) => r.period)));
  const d = new Date(); d.setDate(1);
  for (let i = -1; i < 7; i++) { const x = new Date(d.getFullYear(), d.getMonth() + i, 1); set.add(`${x.getFullYear()}-${pad(x.getMonth() + 1)}`); }
  return [...set].filter((p) => /^\d{4}-\d{2}$/.test(p)).sort();
}
function optionsFor(type) {
  switch (type) {
    case 'period': return periods().map((p) => [p, periodLabel(p)]);
    case 'channel': return state.channels.map((c) => [c, c]);
    case 'warehouse': return state.warehouses.map((c) => [c, c]);
    case 'reason': return REASONS.map((c) => [c, c]);
    case 'matType': return MAT_TYPES.map((c) => [c, c]);
    case 'unit': return UNITS.map((c) => [c, c]);
    case 'line': return state.lines.map((l) => [l.code, `${l.code} · ${l.name}`]);
    case 'status': return PLAN_STATUS.map((c) => [c, c]);
    default: return null;
  }
}
function selectHTML(id, opts, val, attrs = '', blank) {
  const has = opts.some(([v]) => v === val);
  return `<select ${id ? `id="${id}"` : ''} ${attrs}>${blank != null ? `<option value="">${esc(blank)}</option>` : ''}${!has && val ? `<option value="${esc(val)}" selected>${esc(val)} (ngoài DS)</option>` : ''}${opts.map(([v, l]) => `<option value="${esc(v)}" ${v === val ? 'selected' : ''}>${esc(l)}</option>`).join('')}</select>`;
}
// datalist cho mã SP / mã NVL: gõ vài ký tự là lọc nhanh
function datalists() {
  return `<datalist id="dl-product">${state.products.map((p) => `<option value="${esc(p.code)}">${esc(p.name)}</option>`).join('')}</datalist>
  <datalist id="dl-material">${state.materials.map((p) => `<option value="${esc(p.code)}">${esc(p.name)} (${esc(p.type)})</option>`).join('')}</datalist>`;
}

/* ================= Tính bảng tiền kế hoạch ================= */
function computePre(period) {
  const pMap = byCode(state.products);
  const fcCh = {}, adjCh = {};
  const add = (o, p, c, q) => { (o[p] ||= {}); o[p][c] = (o[p][c] || 0) + q; };
  state.forecast.forEach((r) => { if (r.period === period) add(fcCh, r.product, r.channel || '(chưa rõ)', num(r.qty)); });
  state.adjust.forEach((r) => { if (r.period === period) add(adjCh, r.product, r.channel || '(chưa rõ)', num(r.qty)); });
  const fg = {}; state.fgStock.forEach((r) => { fg[r.product] = (fg[r.product] || 0) + num(r.qty); });
  const codes = new Set([...state.products.map((p) => p.code), ...Object.keys(fcCh), ...Object.keys(adjCh)]);
  const bomBy = {}; state.bom.forEach((b) => (bomBy[b.product] ||= []).push(b));

  let rows = [...codes].map((code) => {
    const p = pMap[code];
    const fc = Object.values(fcCh[code] || {}).reduce((a, b) => a + b, 0);
    const adj = Object.values(adjCh[code] || {}).reduce((a, b) => a + b, 0);
    const demand = Math.max(0, fc + adj);
    const stock = fg[code] || 0;
    const safety = Math.round(demand * num(p?.safetyDays) / 30);
    const need = Math.max(0, demand + safety - stock);
    const batch = num(p?.batch);
    const prop = need > 0 ? (batch > 0 ? Math.ceil(need / batch) * batch : Math.ceil(need)) : 0;
    return { code, p, name: p?.name || '', unknown: !p, fcCh: fcCh[code] || {}, adjCh: adjCh[code] || {}, fc, adj, demand, stock, safety, need, batch, prop, bom: bomBy[code] || [] };
  }).filter((r) => r.fc || r.adj || r.prop || r.stock);

  // Nhu cầu NVL từ đề xuất SX so với tồn + đang về
  const mat = materialNeeds(rows.map((r) => ({ product: r.code, qty: r.prop })));
  rows.forEach((r) => {
    r.short = r.prop > 0 ? r.bom.filter((b) => mat.short.has(b.material)).map((b) => b.material) : [];
    r.status = r.prop <= 0 ? 'ok' : r.short.length ? 'bad' : 'need';
  });
  rows.sort((a, b) => ({ bad: 0, need: 1, ok: 2 }[a.status] - { bad: 0, need: 1, ok: 2 }[b.status]) || a.code.localeCompare(b.code));
  return { rows, mat };
}
function materialNeeds(items) {
  const bomBy = {}; state.bom.forEach((b) => (bomBy[b.product] ||= []).push(b));
  const req = {}, noBom = new Set();
  items.forEach(({ product, qty }) => {
    if (!(qty > 0)) return;
    if (!bomBy[product]) noBom.add(product);
    (bomBy[product] || []).forEach((b) => { req[b.material] = (req[b.material] || 0) + qty * num(b.qty); });
  });
  const stock = {}, inc = {};
  state.rmStock.forEach((r) => { stock[r.material] = (stock[r.material] || 0) + num(r.qty); inc[r.material] = (inc[r.material] || 0) + num(r.incoming); });
  const mMap = byCode(state.materials);
  const list = Object.keys(req).map((code) => {
    const have = (stock[code] || 0) + (inc[code] || 0);
    const gap = have - req[code];
    return { code, m: mMap[code], req: req[code], stock: stock[code] || 0, inc: inc[code] || 0, gap, short: gap < -1e-9, onlyWithInc: gap >= 0 && (stock[code] || 0) < req[code] };
  }).sort((a, b) => a.gap - b.gap);
  return { list, short: new Set(list.filter((x) => x.short).map((x) => x.code)), noBom };
}
const STATUS_CHIP = { ok: '<span class="chip ok">Đủ tồn</span>', need: '<span class="chip warn">Cần SX</span>', bad: '<span class="chip bad">Thiếu NVL</span>' };

/* ================= Điều hướng ================= */
const ICON = {
  home: '<path d="M3 11l9-7 9 7"/><path d="M5 10v10h14V10"/>',
  input: '<path d="M12 3v12"/><path d="M7 10l5 5 5-5"/><path d="M4 19h16"/>',
  preplan: '<rect x="3" y="4" width="18" height="16" rx="2"/><path d="M3 10h18M9 10v10"/>',
  plan: '<rect x="3" y="5" width="18" height="16" rx="2"/><path d="M8 3v4M16 3v4M3 11h18"/><path d="M8 15h4"/>',
  master: '<path d="M4 6h16M4 12h16M4 18h10"/>',
  settings: '<path d="M12 8a4 4 0 100 8 4 4 0 000-8z"/><path d="M4 12h2M18 12h2M12 4v2M12 18v2M6.3 6.3l1.4 1.4M16.3 16.3l1.4 1.4M6.3 17.7l1.4-1.4M16.3 7.7l1.4-1.4"/>',
};
const VIEWS = [
  ['home', 'Tổng quan', 'Tổng quan'],
  ['input', 'Nhập liệu', 'Nguồn dữ liệu đầu vào'],
  ['preplan', 'Tiền KH', 'Bảng tiền kế hoạch'],
  ['plan', 'Kế hoạch SX', 'Kế hoạch sản xuất'],
  ['master', 'Danh mục', 'Danh mục & mã chuẩn'],
  ['settings', 'Cài đặt', 'Lưu trữ & đồng bộ'],
];
function renderNav() {
  const btn = ([k, short]) => `<button data-go="${k}" ${ui.view === k ? 'aria-current="page"' : ''}><svg viewBox="0 0 24 24">${ICON[k]}</svg><span>${short}</span></button>`;
  $('#tabbar').innerHTML = VIEWS.map(btn).join('');
  $('#nav').innerHTML = VIEWS.map(btn).join('');
}
function go(view) {
  ui.view = view; ui.importer = null;
  try { localStorage.setItem('tkh:view', view); } catch {}
  render();
  $('#view').scrollTop = 0; window.scrollTo(0, 0);
}
function render() {
  const v = VIEWS.find((x) => x[0] === ui.view) || VIEWS[0];
  $('#viewEyebrow').textContent = v[1];
  $('#viewTitle').textContent = v[0] === 'home' ? 'Tiền kế hoạch' : v[2];
  renderNav();
  const ps = periods();
  if (!ps.includes(state.settings.period)) state.settings.period = ps[0];
  $('#periodSel').innerHTML = ps.map((p) => `<option value="${p}" ${p === state.settings.period ? 'selected' : ''}>${periodLabel(p)}</option>`).join('');
  const f = { home: viewHome, input: viewInput, preplan: viewPre, plan: viewPlan, master: viewMaster, settings: viewSettings }[ui.view] || viewHome;
  $('#view').innerHTML = f() + datalists();
}

/* ================= Màn hình: Tổng quan ================= */
function sampleBanner() {
  if (!state.meta.sample) return '';
  return `<div class="notice warn banner"><span><b>Đang dùng dữ liệu mẫu.</b> Các mã SP, NVL và số liệu chỉ để minh hoạ. Nhập file thật hoặc xoá mẫu khi bắt đầu dùng.</span>
  <button class="btn sm danger" data-act="clearSample">Xoá dữ liệu mẫu</button></div>`;
}
function sourceStatus(key) {
  const per = state.settings.period, arr = state[key];
  const rows = key === 'forecast' || key === 'adjust' ? arr.filter((r) => r.period === per).length : arr.length;
  const upd = state.meta.updated[key];
  const age = upd ? (Date.now() - new Date(upd)) / 864e5 : null;
  let chip;
  if (key === 'adjust') chip = rows ? '<span class="chip ok">Có điều chỉnh</span>' : '<span class="chip neutral">Không có</span>';
  else if (!rows) chip = '<span class="chip bad">Chưa có</span>';
  else if (age != null && age > 7 && key !== 'forecast') chip = '<span class="chip warn">Cũ &gt; 7 ngày</span>';
  else chip = '<span class="chip ok">Sẵn sàng</span>';
  return { rows, upd, chip };
}
function viewHome() {
  const per = state.settings.period;
  const { rows, mat } = computePre(per);
  const needRows = rows.filter((r) => r.prop > 0);
  const planRows = state.plan.filter((r) => r.period === per);
  const approved = planRows.filter((r) => r.status !== 'Nháp' && r.status !== 'Huỷ').length;
  const srcs = INPUT_KEYS.map((k) => {
    const s = sourceStatus(k);
    return `<div class="src-row"><b>${SOURCES[k].label}</b>${s.chip}<small>${fmt(s.rows)} dòng${k === 'forecast' || k === 'adjust' ? ' trong kỳ ' + periodLabel(per) : ''} · ${relTime(s.upd)}</small></div>`;
  }).join('');
  return `${sampleBanner()}
  <div class="kpis">
    <div class="kpi"><span>Nhu cầu kỳ ${periodLabel(per)}</span><b>${fmt(sum(rows, (r) => r.demand))}</b><small>forecast + cân chỉnh</small></div>
    <div class="kpi work"><span>Đề xuất sản xuất</span><b>${fmt(sum(needRows, (r) => r.prop))}</b><small>${needRows.length} mã SP cần SX</small></div>
    <div class="kpi ${mat.short.size ? 'alert' : ''}"><span>NVL - phụ gia thiếu</span><b>${mat.short.size}</b><small>${mat.short.size ? 'cần đặt thêm hàng' : 'đủ cho đề xuất'}</small></div>
    <div class="kpi"><span>Lệnh SX trong kỳ</span><b>${planRows.length}</b><small>${approved} đã duyệt</small></div>
  </div>
  <section class="panel">
    <div class="panel-head"><div><h2>Quy trình</h2><p>Ba bước từ dữ liệu thô đến kế hoạch sản xuất.</p></div></div>
    <div class="flow">
      <button class="step" data-go="input"><span class="n">1</span><span><b>Nhập nguồn dữ liệu</b><small>Forecast WMS, tồn TP, tồn NVL - phụ gia, cân chỉnh kênh</small></span></button>
      <button class="step" data-go="preplan"><span class="n">2</span><span><b>Xem bảng tiền kế hoạch</b><small>Tự tổng hợp nhu cầu, tồn, tồn an toàn, đề xuất SX</small></span></button>
      <button class="step" data-go="plan"><span class="n">3</span><span><b>Lập kế hoạch SX</b><small>Phân line, ngày chạy, duyệt và xuất Excel</small></span></button>
    </div>
  </section>
  <div class="grid-2">
    <section class="panel">
      <div class="panel-head"><div><h2>Tình trạng dữ liệu đầu vào</h2><p>Kỳ ${periodLabel(per)}</p></div><button class="btn sm ghost" data-go="input">Nhập liệu</button></div>
      <div class="sources">${srcs}</div>
    </section>
    <section class="panel">
      <div class="panel-head"><div><h2>Mã cần sản xuất</h2><p>Theo bảng tiền kế hoạch</p></div><button class="btn sm ghost" data-go="preplan">Xem bảng</button></div>
      ${needRows.length ? `<div class="table-wrap"><table class="stack"><thead><tr><th>Mã SP</th><th>Tên</th><th class="num">Đề xuất</th><th>Trạng thái</th></tr></thead><tbody>
      ${needRows.slice(0, 8).map((r) => `<tr><td class="code" data-l="Mã SP">${esc(r.code)}</td><td class="name" data-l="Tên">${esc(r.name) || '<span class="unknown">Chưa có trong danh mục</span>'}</td><td class="num" data-l="Đề xuất">${fmt(r.prop)}</td><td data-l="Trạng thái">${STATUS_CHIP[r.status]}</td></tr>`).join('')}
      </tbody></table></div>` : '<div class="empty">Chưa có mã nào cần sản xuất trong kỳ này.</div>'}
    </section>
  </div>`;
}

/* ================= Màn hình: Nhập liệu ================= */
function viewInput() {
  const k = ui.src, S = SOURCES[k];
  const seg = INPUT_KEYS.map((x) => {
    const s = sourceStatus(x);
    return `<button data-src="${x}" aria-pressed="${x === k}">${SOURCES[x].short} · ${fmt(s.rows)}</button>`;
  }).join('');
  return `${sampleBanner()}
  <div class="seg" role="group" aria-label="Chọn nguồn dữ liệu">${seg}</div>
  ${importPanel(k)}
  ${k === 'adjust' ? adjustForm() : ''}
  ${gridPanel(k, S.label, S.hint)}`;
}
function importPanel(k) {
  const S = SOURCES[k], imp = ui.importer && ui.importer.key === k ? ui.importer : null;
  let body;
  if (!imp) {
    body = `<label class="drop" id="drop" for="fileIn">
      <b>Chọn hoặc kéo thả file Excel / CSV</b>
      <small class="muted">${esc(S.hint || '')}</small>
      <input type="file" id="fileIn" data-key="${k}" accept=".xlsx,.xls,.xlsm,.csv">
    </label>`;
  } else body = importerHTML(imp);
  return `<section class="panel">
    <div class="panel-head"><div><h2>Nhập ${esc(S.label)}</h2><p>Theo mẫu chuẩn. Tên cột được tự nhận diện, có thể chọn lại bằng danh sách.</p></div>
    <button class="btn sm ghost" data-act="template" data-key="${k}">Tải file mẫu</button></div>
    ${body}
  </section>`;
}
function importerHTML(imp) {
  const S = SOURCES[imp.key];
  const colOpts = imp.headers.map((h, i) => [String(i), h || `Cột ${i + 1}`]);
  const maps = S.fields.map((f) => {
    const dflt = f.def ? `<label class="field"><span>Mặc định nếu thiếu</span>${selectHTML('', optionsFor(f.type), imp.defaults[f.k] || '', `data-imp-def="${f.k}"`, '—')}</label>` : '';
    return `<div class="field-group" style="display:flex;flex-direction:column;gap:6px"><label class="field ${f.req ? 'req' : ''}"><span>${esc(f.label)}</span>${selectHTML('', colOpts, imp.map[f.k] ?? '', `data-imp-map="${f.k}"`, '(không có cột này)')}</label>${dflt}</div>`;
  }).join('');
  const sheetSel = imp.sheets.length > 1 ? `<label class="field" style="max-width:260px"><span>Sheet</span>${selectHTML('', imp.sheets.map((s) => [s, s]), imp.sheet, 'data-imp-sheet')}</label>` : '';
  const hdrSel = `<label class="field" style="max-width:160px"><span>Dòng tiêu đề</span>${selectHTML('', [1, 2, 3, 4, 5, 6].map((n) => [String(n), `Dòng ${n}`]), String(imp.headerRow + 1), 'data-imp-hdr')}</label>`;
  const res = buildImportRows(imp);
  const missing = S.fields.filter((f) => f.req && (imp.map[f.k] == null || imp.map[f.k] === ''));
  const preview = res.rows.slice(0, 6);
  return `<div class="notice info"><b>${esc(imp.fileName)}</b> · ${fmt(imp.data.length - imp.headerRow - 1)} dòng dữ liệu</div>
  <div class="row">${sheetSel}${hdrSel}</div>
  <div class="map-grid">${maps}</div>
  ${missing.length ? `<div class="notice bad">Chưa chọn cột bắt buộc: ${missing.map((f) => esc(f.label)).join(', ')}.</div>` : ''}
  ${res.unknown.length ? `<div class="notice warn"><b>${res.unknown.length} mã chưa có trong danh mục:</b> <span class="mono">${res.unknown.slice(0, 12).map(esc).join(', ')}${res.unknown.length > 12 ? '…' : ''}</span>. Vẫn nhập được, các mã này sẽ được đánh dấu đỏ để bổ sung danh mục sau.</div>` : ''}
  ${res.skipped ? `<div class="notice">Bỏ qua ${fmt(res.skipped)} dòng trống hoặc thiếu mã.</div>` : ''}
  <div class="table-wrap"><table><thead><tr>${S.fields.map((f) => `<th>${esc(f.label)}</th>`).join('')}</tr></thead>
  <tbody>${preview.map((r) => `<tr>${S.fields.map((f) => `<td class="${f.type === 'num' ? 'num' : ''}">${esc(f.type === 'num' ? fmtQ(r[f.k]) : r[f.k])}</td>`).join('')}</tr>`).join('') || `<tr><td colspan="${S.fields.length}" class="muted">Không có dòng hợp lệ.</td></tr>`}</tbody></table></div>
  <div class="row">
    <label class="field" style="min-width:220px"><span>Cách nhập</span>${selectHTML('impMode', imp.key === 'forecast' || imp.key === 'adjust'
      ? [['replacePeriod', 'Thay dữ liệu các kỳ trong file'], ['replace', 'Thay toàn bộ'], ['append', 'Thêm vào dữ liệu hiện có']]
      : [['replace', 'Thay toàn bộ'], ['append', 'Thêm vào dữ liệu hiện có']], imp.mode, 'data-imp-mode')}</label>
    <span style="flex:1"></span>
    <button class="btn ghost" data-act="impCancel">Huỷ</button>
    <button class="btn" data-act="impCommit" ${missing.length || !res.rows.length ? 'disabled' : ''}>Nhập ${fmt(res.rows.length)} dòng</button>
  </div>`;
}
function guessMap(fields, headers) {
  const nh = headers.map(norm), map = {};
  fields.forEach((f) => {
    let i = nh.findIndex((h) => f.alias.includes(h));
    if (i < 0) i = nh.findIndex((h) => h && f.alias.some((a) => a.length > 2 && h.includes(a)));
    if (i >= 0 && !Object.values(map).includes(String(i))) map[f.k] = String(i);
  });
  return map;
}
function loadSheet(imp) {
  const ws = imp.wb.Sheets[imp.sheet];
  imp.data = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '' });
  if (imp.headerRow == null) { // tìm dòng tiêu đề: dòng đầu khớp nhiều alias nhất
    let best = 0, bestScore = -1;
    imp.data.slice(0, 6).forEach((r, i) => { const sc = Object.keys(guessMap(SOURCES[imp.key].fields, r.map(String))).length; if (sc > bestScore) { best = i; bestScore = sc; } });
    imp.headerRow = best;
  }
  imp.headers = (imp.data[imp.headerRow] || []).map((h) => String(h).trim());
  imp.map = guessMap(SOURCES[imp.key].fields, imp.headers);
}
function buildImportRows(imp) {
  const S = SOURCES[imp.key], out = [], unknownSet = new Set();
  let skipped = 0;
  const pSet = new Set(state.products.map((p) => p.code)), mSet = new Set(state.materials.map((m) => m.code));
  imp.data.slice(imp.headerRow + 1).forEach((raw) => {
    const r = {};
    S.fields.forEach((f) => {
      const ci = imp.map[f.k]; let v = ci != null && ci !== '' ? raw[+ci] : '';
      if (f.type === 'num') v = num(v);
      else if (f.type === 'period') v = toPeriod(v) || imp.defaults[f.k] || '';
      else { v = String(v ?? '').trim(); if (!v && imp.defaults[f.k]) v = imp.defaults[f.k]; }
      r[f.k] = v;
    });
    const keyF = S.fields.find((f) => f.req && f.type !== 'num');
    if (!r[keyF.k]) { skipped++; return; }
    S.fields.forEach((f) => {
      if (f.type === 'product' && r[f.k] && pSet.size && !pSet.has(r[f.k])) unknownSet.add(r[f.k]);
      if (f.type === 'material' && r[f.k] && mSet.size && !mSet.has(r[f.k])) unknownSet.add(r[f.k]);
    });
    out.push(r);
  });
  return { rows: out, unknown: [...unknownSet], skipped };
}
async function handleFile(file, key) {
  if (!window.XLSX) { toast('Chưa tải được thư viện đọc Excel. Kiểm tra kết nối mạng rồi mở lại app.'); return; }
  try {
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: 'array', cellDates: true });
    const imp = { key, wb, fileName: file.name, sheets: wb.SheetNames, sheet: wb.SheetNames[0], headerRow: null, defaults: {}, mode: key === 'forecast' || key === 'adjust' ? 'replacePeriod' : 'replace' };
    SOURCES[key].fields.forEach((f) => { if (f.def === 'period') imp.defaults[f.k] = state.settings.period; });
    loadSheet(imp);
    ui.importer = imp; render();
  } catch (e) { toast('Không đọc được file: ' + e.message); }
}
function commitImport() {
  const imp = ui.importer; if (!imp) return;
  const { rows } = buildImportRows(imp);
  const k = imp.key;
  if (imp.mode === 'replace') state[k] = rows;
  else if (imp.mode === 'replacePeriod') { const ps = new Set(rows.map((r) => r.period)); state[k] = state[k].filter((r) => !ps.has(r.period)).concat(rows); }
  else state[k] = state[k].concat(rows);
  // Tự thêm kênh / kho mới vào danh sách chọn nhanh
  rows.forEach((r) => {
    if (r.channel && !state.channels.includes(r.channel)) state.channels.push(r.channel);
    if (r.warehouse && !state.warehouses.includes(r.warehouse)) state.warehouses.push(r.warehouse);
  });
  ui.importer = null; save(k); render();
  toast(`Đã nhập ${fmt(rows.length)} dòng ${SOURCES[k].label}`);
}
function downloadTemplate(k) {
  if (!window.XLSX) { toast('Chưa tải được thư viện Excel.'); return; }
  const S = SOURCES[k];
  const ws = XLSX.utils.aoa_to_sheet([S.fields.map((f) => f.label), ...S.sample]);
  ws['!cols'] = S.fields.map((f) => ({ wch: Math.max(12, f.label.length + 4) }));
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, 'Du lieu');
  const guide = [['Cột', 'Bắt buộc', 'Ghi chú'], ...S.fields.map((f) => [f.label, f.req ? 'Có' : 'Không', {
    period: 'Dạng YYYY-MM hoặc MM/YYYY', num: 'Số', product: 'Mã SP theo danh mục chuẩn', material: 'Mã NVL/phụ gia theo danh mục chuẩn',
    channel: 'Một trong: ' + state.channels.join(', '), reason: 'Một trong: ' + REASONS.join(', '), matType: MAT_TYPES.join(' / '), line: 'Mã line: ' + state.lines.map((l) => l.code).join(', '),
  }[f.type] || ''])];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(guide), 'Huong dan');
  XLSX.writeFile(wb, S.file);
}
function adjustForm() {
  const per = state.settings.period;
  return `<section class="panel">
    <div class="panel-head"><div><h2>Thêm cân chỉnh nhanh</h2><p>Số dương là tăng, số âm là giảm so với forecast.</p></div></div>
    <form class="form-grid" id="adjForm">
      <label class="field"><span>Kỳ</span>${selectHTML('adjPeriod', optionsFor('period'), per)}</label>
      <label class="field"><span>Mã SP</span><input id="adjProduct" list="dl-product" required placeholder="Gõ mã hoặc chọn"></label>
      <label class="field"><span>Kênh</span>${selectHTML('adjChannel', optionsFor('channel'), state.channels[0])}</label>
      <label class="field"><span>SL (±)</span><input id="adjQty" type="number" step="any" required placeholder="vd 1500 hoặc -400"></label>
      <label class="field"><span>Lý do</span>${selectHTML('adjReason', optionsFor('reason'), REASONS[0])}</label>
      <label class="field"><span>Ghi chú</span><input id="adjNote" placeholder="Tuỳ chọn"></label>
      <button class="btn" type="submit">Thêm</button>
    </form>
  </section>`;
}

/* ================= Bảng dữ liệu sửa trực tiếp ================= */
function cellInput(key, i, f, v) {
  const a = `data-k="${key}" data-i="${i}" data-f="${f.k}" aria-label="${esc(f.label)}"`;
  if (f.type === 'num') return `<input type="number" step="any" value="${esc(v)}" ${a}>`;
  if (f.type === 'product') return `<input list="dl-product" value="${esc(v)}" ${a} class="mono">`;
  if (f.type === 'material') return `<input list="dl-material" value="${esc(v)}" ${a} class="mono">`;
  const opts = optionsFor(f.type);
  if (opts) return selectHTML('', opts, v, a, '—');
  return `<input value="${esc(v)}" ${a}>`;
}
const LIMIT = 200;
function gridPanel(key, title, hint) {
  const S = SOURCES[key], all = state[key], q = norm(ui.filter[key] || '');
  const per = state.settings.period, periodScoped = key === 'forecast' || key === 'adjust';
  const pMap = byCode(state.products), mMap = byCode(state.materials);
  let idx = all.map((r, i) => i);
  if (periodScoped && !ui.filter[key + ':all']) idx = idx.filter((i) => all[i].period === per);
  if (q) idx = idx.filter((i) => norm(Object.values(all[i]).join(' ')).includes(q));
  const shown = idx.slice(0, LIMIT);
  const nameOf = (r) => {
    const c = r.product || r.material; if (!c) return '';
    const n = r.product ? pMap[c]?.name : mMap[c]?.name;
    return n ? `<span class="sub">${esc(n)}</span>` : '<span class="sub unknown">Chưa có trong danh mục</span>';
  };
  return `<section class="panel">
    <div class="panel-head"><div><h2>${esc(title)}</h2><p>${fmt(idx.length)} dòng${periodScoped && !ui.filter[key + ':all'] ? ' trong kỳ ' + periodLabel(per) : ''}${hint ? '' : ''}</p></div>
      <div class="row">
        ${periodScoped ? `<label class="row" style="gap:6px;font-size:13px"><input type="checkbox" data-allp="${key}" ${ui.filter[key + ':all'] ? 'checked' : ''} style="min-height:auto"> Mọi kỳ</label>` : ''}
        <input type="search" placeholder="Tìm…" value="${esc(ui.filter[key] || '')}" data-filter="${key}" style="width:150px">
        <button class="btn sm ghost" data-act="addRow" data-key="${key}">+ Thêm dòng</button>
        ${all.length ? `<button class="btn sm danger" data-act="clearKey" data-key="${key}">Xoá hết</button>` : ''}
      </div></div>
    ${shown.length ? `<div class="table-wrap"><table><thead><tr>${S.fields.map((f) => `<th>${esc(f.label)}</th>`).join('')}<th></th></tr></thead><tbody>
    ${shown.map((i) => `<tr>${S.fields.map((f, j) => `<td>${cellInput(key, i, f, all[i][f.k])}${(f.type === 'product' || f.type === 'material') && j <= 1 ? nameOf(all[i]) : ''}</td>`).join('')}<td><button class="icon-btn" data-del="${key}" data-i="${i}" aria-label="Xoá dòng">×</button></td></tr>`).join('')}
    </tbody></table></div>${idx.length > LIMIT ? `<p class="table-foot">Đang hiện ${LIMIT}/${fmt(idx.length)} dòng. Dùng ô tìm để lọc.</p>` : ''}`
    : `<div class="empty"><b>Chưa có dữ liệu</b><span>Nhập file Excel theo mẫu ở trên hoặc bấm “+ Thêm dòng”.</span></div>`}
  </section>`;
}

/* ================= Màn hình: Bảng tiền kế hoạch ================= */
function viewPre() {
  const per = state.settings.period;
  const { rows, mat } = computePre(per);
  const chs = [...new Set([...state.channels, ...rows.flatMap((r) => Object.keys(r.fcCh).concat(Object.keys(r.adjCh)))])];
  const f = ui.preFilter, q = norm(ui.filter.pre || '');
  const list = rows.filter((r) => (f === 'all' || r.status === f) && (!q || norm(r.code + ' ' + r.name).includes(q)));
  const cnt = (s) => rows.filter((r) => r.status === s).length;
  const chCell = (r, c) => {
    const a = r.fcCh[c] || 0, d = r.adjCh[c] || 0;
    return `<td class="num" data-l="${esc(c)}">${a ? fmt(a) : '<span class="muted">–</span>'}${d ? `<span class="adj ${d > 0 ? 'up' : 'down'}"> ${d > 0 ? '+' : ''}${fmt(d)}</span>` : ''}</td>`;
  };
  const T = (k) => sum(list, (r) => r[k]);
  return `${sampleBanner()}
  <section class="panel">
    <div class="panel-head"><div><h2>Bảng tiền kế hoạch · ${periodLabel(per)}</h2>
      <p>Nhu cầu = Forecast + Cân chỉnh. Cần SX = Nhu cầu + Tồn an toàn − Tồn TP. Đề xuất làm tròn lên theo lô sản xuất.</p></div>
      <div class="row"><button class="btn sm ghost" data-act="exportPre">Xuất Excel</button><button class="btn sm tag" data-act="genPlan">Tạo kế hoạch SX →</button></div></div>
    <div class="row">
      <label class="field" style="min-width:170px"><span>Lọc trạng thái</span>${selectHTML('preFilter', [['all', `Tất cả (${rows.length})`], ['need', `Cần SX (${cnt('need')})`], ['bad', `Thiếu NVL (${cnt('bad')})`], ['ok', `Đủ tồn (${cnt('ok')})`]], f)}</label>
      <label class="field" style="min-width:170px"><span>Tìm mã / tên</span><input type="search" data-filter="pre" value="${esc(ui.filter.pre || '')}" placeholder="vd SP001"></label>
    </div>
    ${list.length ? `<div class="table-wrap"><table class="stack"><thead><tr>
      <th>Mã SP</th><th>Tên SP</th>${chs.map((c) => `<th class="num">${esc(c)}</th>`).join('')}
      <th class="num">Forecast</th><th class="num">Cân chỉnh</th><th class="num">Nhu cầu</th><th class="num">Tồn TP</th><th class="num">Tồn AT</th><th class="num">Cần SX</th><th class="num">Lô</th><th class="num prop">Đề xuất SX</th><th>Trạng thái</th>
    </tr></thead><tbody>
    ${list.map((r) => `<tr>
      <td class="code" data-l="Mã SP">${esc(r.code)}</td>
      <td class="name" data-l="Tên SP">${r.unknown ? '<span class="unknown">Chưa có trong danh mục</span>' : esc(r.name)}${r.short.length ? `<span class="sub unknown">Thiếu: ${r.short.map(esc).join(', ')}</span>` : ''}</td>
      ${chs.map((c) => chCell(r, c)).join('')}
      <td class="num" data-l="Forecast">${fmt(r.fc)}</td>
      <td class="num" data-l="Cân chỉnh">${r.adj ? `<span class="adj ${r.adj > 0 ? 'up' : 'down'}">${r.adj > 0 ? '+' : ''}${fmt(r.adj)}</span>` : '–'}</td>
      <td class="num" data-l="Nhu cầu"><b>${fmt(r.demand)}</b></td>
      <td class="num" data-l="Tồn TP">${fmt(r.stock)}</td>
      <td class="num" data-l="Tồn AT">${fmt(r.safety)}${r.p ? `<span class="sub">${fmt(r.p.safetyDays)} ngày</span>` : ''}</td>
      <td class="num" data-l="Cần SX">${fmt(r.need)}</td>
      <td class="num" data-l="Lô">${r.batch ? fmt(r.batch) : '–'}</td>
      <td class="num prop" data-l="Đề xuất SX">${fmt(r.prop)}</td>
      <td data-l="Trạng thái">${STATUS_CHIP[r.status]}</td></tr>`).join('')}
    <tr class="total"><td data-l="">Tổng</td><td data-l=""></td>${chs.map((c) => `<td class="num" data-l="${esc(c)}">${fmt(sum(list, (r) => r.fcCh[c]))}</td>`).join('')}
      <td class="num" data-l="Forecast">${fmt(T('fc'))}</td><td class="num" data-l="Cân chỉnh">${fmt(T('adj'))}</td><td class="num" data-l="Nhu cầu">${fmt(T('demand'))}</td><td class="num" data-l="Tồn TP">${fmt(T('stock'))}</td><td class="num" data-l="Tồn AT">${fmt(T('safety'))}</td><td class="num" data-l="Cần SX">${fmt(T('need'))}</td><td data-l=""></td><td class="num prop" data-l="Đề xuất SX">${fmt(T('prop'))}</td><td data-l=""></td></tr>
    </tbody></table></div>` : `<div class="empty"><b>Chưa có dữ liệu cho kỳ ${periodLabel(per)}</b><span>Nhập forecast từ WMS và tồn kho để app tự tổng hợp.</span><button class="btn sm" data-go="input">Đi tới Nhập liệu</button></div>`}
  </section>
  ${matPanel(mat, 'Kiểm tra NVL - phụ gia cho đề xuất', 'Nhu cầu theo định mức (BOM) × đề xuất SX, so với tồn + đang về.')}`;
}
function matPanel(mat, title, sub) {
  const noBom = [...mat.noBom];
  return `<section class="panel">
    <div class="panel-head"><div><h2>${title}</h2><p>${sub}</p></div>
    <span>${mat.short.size ? `<span class="chip bad">${mat.short.size} mã thiếu</span>` : '<span class="chip ok">Đủ NVL</span>'}</span></div>
    ${noBom.length ? `<div class="notice warn">Chưa có định mức (BOM) cho: <span class="mono">${noBom.map(esc).join(', ')}</span>. Bổ sung ở Danh mục › BOM.</div>` : ''}
    ${mat.list.length ? `<div class="table-wrap"><table class="stack"><thead><tr><th>Mã</th><th>Tên</th><th>Loại</th><th class="num">Cần dùng</th><th class="num">Tồn</th><th class="num">Đang về</th><th class="num">Chênh lệch</th><th>Trạng thái</th></tr></thead><tbody>
    ${mat.list.map((x) => `<tr><td class="code" data-l="Mã">${esc(x.code)}</td><td class="name" data-l="Tên">${x.m ? esc(x.m.name) : '<span class="unknown">Chưa có trong danh mục</span>'}</td><td data-l="Loại">${esc(x.m?.type || '')}</td>
      <td class="num" data-l="Cần dùng">${fmtQ(x.req)} <span class="muted">${esc(x.m?.unit || '')}</span></td><td class="num" data-l="Tồn">${fmtQ(x.stock)}</td><td class="num" data-l="Đang về">${fmtQ(x.inc)}</td>
      <td class="num ${x.short ? 'unknown' : ''}" data-l="Chênh lệch">${x.gap > 0 ? '+' : ''}${fmtQ(x.gap)}</td>
      <td data-l="Trạng thái">${x.short ? '<span class="chip bad">Thiếu</span>' : x.onlyWithInc ? '<span class="chip warn">Chờ hàng về</span>' : '<span class="chip ok">Đủ</span>'}</td></tr>`).join('')}
    </tbody></table></div>` : '<div class="empty">Chưa có nhu cầu NVL.</div>'}
  </section>`;
}
function exportPre() {
  if (!window.XLSX) { toast('Chưa tải được thư viện Excel.'); return; }
  const per = state.settings.period, { rows, mat } = computePre(per);
  const chs = [...new Set([...state.channels, ...rows.flatMap((r) => Object.keys(r.fcCh))])];
  const head = ['Mã SP', 'Tên SP', ...chs.map((c) => 'FC ' + c), 'Forecast', 'Cân chỉnh', 'Nhu cầu', 'Tồn TP', 'Tồn an toàn', 'Cần SX', 'Lô SX', 'Đề xuất SX', 'Trạng thái', 'NVL thiếu'];
  const st = { ok: 'Đủ tồn', need: 'Cần SX', bad: 'Thiếu NVL' };
  const aoa = [head, ...rows.map((r) => [r.code, r.name, ...chs.map((c) => r.fcCh[c] || 0), r.fc, r.adj, r.demand, r.stock, r.safety, r.need, r.batch, r.prop, st[r.status], r.short.join(', ')])];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), 'Tien ke hoach');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['Mã', 'Tên', 'Loại', 'ĐVT', 'Cần dùng', 'Tồn', 'Đang về', 'Chênh lệch'], ...mat.list.map((x) => [x.code, x.m?.name || '', x.m?.type || '', x.m?.unit || '', x.req, x.stock, x.inc, x.gap])]), 'Nhu cau NVL');
  XLSX.writeFile(wb, `Tien_ke_hoach_${per}.xlsx`);
}

/* ================= Màn hình: Kế hoạch sản xuất ================= */
async function genPlan() {
  const per = state.settings.period;
  const { rows } = computePre(per);
  const props = rows.filter((r) => r.prop > 0);
  if (!props.length) { toast('Không có mã nào cần sản xuất trong kỳ này.'); return; }
  const existing = state.plan.filter((r) => r.period === per);
  if (existing.length && !(await ask('Tạo lại kế hoạch?', `Kỳ ${periodLabel(per)} đã có ${existing.length} lệnh. Tạo lại sẽ thay toàn bộ lệnh của kỳ này bằng đề xuất mới.`, 'Tạo lại'))) return;
  const lMap = byCode(state.lines), cursor = {};
  const plan = props.map((r) => {
    const line = r.p?.line && lMap[r.p.line] ? r.p.line : state.lines[0]?.code || '';
    const start = cursor[line] || firstWorkday(per);
    const days = Math.max(1, Math.ceil(r.prop / (num(lMap[line]?.capacity) || r.prop)));
    const end = addWorkdays(start, days);
    const nx = new Date(`${end}T00:00:00`); nx.setDate(nx.getDate() + 1); cursor[line] = addWorkdays(dateStr(nx), 1);
    return { id: Math.random().toString(36).slice(2, 9), period: per, product: r.code, qty: r.prop, line, start, status: 'Nháp', note: r.short.length ? 'Thiếu NVL: ' + r.short.join(', ') : '' };
  });
  state.plan = state.plan.filter((r) => r.period !== per).concat(plan);
  save('plan'); ui.view = 'plan'; render();
  toast(`Đã tạo ${plan.length} lệnh sản xuất ở trạng thái Nháp`);
}
function planDays(r) { const l = state.lines.find((x) => x.code === r.line); const c = num(l?.capacity); return c > 0 ? Math.max(1, Math.ceil(num(r.qty) / c)) : 1; }
function viewPlan() {
  const per = state.settings.period, pMap = byCode(state.products);
  const idx = state.plan.map((r, i) => i).filter((i) => state.plan[i].period === per)
    .sort((a, b) => (state.plan[a].line || '').localeCompare(state.plan[b].line || '') || (state.plan[a].start || '').localeCompare(state.plan[b].start || ''));
  const rows = idx.map((i) => state.plan[i]);
  const active = rows.filter((r) => r.status !== 'Huỷ');
  const mat = materialNeeds(active.filter((r) => r.status !== 'Hoàn thành').map((r) => ({ product: r.product, qty: num(r.qty) })));
  const wd = workdaysInMonth(per);
  const loads = state.lines.map((l) => {
    const d = active.filter((r) => r.line === l.code).reduce((a, r) => a + planDays(r), 0);
    const pct = Math.round((d / wd) * 100);
    return `<div class="load"><span><b class="mono">${esc(l.code)}</b> <small class="muted">${esc(l.name)}</small></span>
      <div class="bar" role="img" aria-label="${pct}%"><i class="${pct > 100 ? 'over' : pct > 85 ? 'high' : ''}" style="width:${Math.min(100, pct)}%"></i></div>
      <span class="num"><b>${pct}%</b> <small class="muted">${d}/${wd} ngày</small></span></div>`;
  }).join('');
  const statusCls = { 'Nháp': 'neutral', 'Đã duyệt': 'ok', 'Đang SX': 'warn', 'Hoàn thành': 'ok', 'Huỷ': 'bad' };
  return `${sampleBanner()}
  <section class="panel">
    <div class="panel-head"><div><h2>Kế hoạch sản xuất · ${periodLabel(per)}</h2><p>Sửa số lượng, chọn line, ngày bắt đầu và trạng thái. Ngày kết thúc tính theo công suất line, bỏ Chủ nhật.</p></div>
    <div class="row">
      <button class="btn sm ghost" data-act="genPlan">Tạo lại từ tiền KH</button>
      ${rows.length ? `<button class="btn sm ghost" data-act="approveAll">Duyệt tất cả</button><button class="btn sm ghost" data-act="exportPlan">Xuất Excel</button>` : ''}
    </div></div>
    ${rows.length ? `<div class="table-wrap"><table><thead><tr><th>Mã SP</th><th class="num">Số lượng</th><th>Line</th><th>Bắt đầu</th><th class="num">Ngày</th><th>Kết thúc</th><th>Trạng thái</th><th>Ghi chú</th><th></th></tr></thead><tbody>
    ${idx.map((i) => { const r = state.plan[i]; const p = pMap[r.product]; const d = planDays(r);
      return `<tr>
      <td><input list="dl-product" class="mono" value="${esc(r.product)}" data-plan="${i}" data-f="product" aria-label="Mã SP"><span class="sub ${p ? '' : 'unknown'}">${p ? esc(p.name) : 'Chưa có trong danh mục'}</span></td>
      <td><input type="number" step="any" value="${esc(r.qty)}" data-plan="${i}" data-f="qty" aria-label="Số lượng"></td>
      <td>${selectHTML('', optionsFor('line'), r.line, `data-plan="${i}" data-f="line" aria-label="Line"`)}</td>
      <td><input type="date" value="${esc(r.start)}" data-plan="${i}" data-f="start" aria-label="Ngày bắt đầu"></td>
      <td class="num">${d}</td>
      <td class="mono">${fmtDate(addWorkdays(r.start, d))}</td>
      <td>${selectHTML('', optionsFor('status'), r.status, `data-plan="${i}" data-f="status" aria-label="Trạng thái"`)} <span class="chip ${statusCls[r.status] || 'neutral'}" style="margin-top:4px">${esc(r.status)}</span></td>
      <td><input value="${esc(r.note)}" data-plan="${i}" data-f="note" aria-label="Ghi chú" style="min-width:160px"></td>
      <td><button class="icon-btn" data-delplan="${i}" aria-label="Xoá lệnh">×</button></td></tr>`; }).join('')}
    <tr class="total"><td>Tổng</td><td class="num">${fmt(sum(active, (r) => r.qty))}</td><td colspan="7"></td></tr>
    </tbody></table></div>` : `<div class="empty"><b>Chưa có kế hoạch cho kỳ ${periodLabel(per)}</b><span>Bấm “Tạo lại từ tiền KH” để sinh lệnh SX từ đề xuất của bảng tiền kế hoạch.</span></div>`}
    <div class="row"><button class="btn sm ghost" data-act="addPlan">+ Thêm lệnh thủ công</button></div>
  </section>
  <div class="grid-2">
    <section class="panel"><div class="panel-head"><div><h2>Tải line sản xuất</h2><p>Số ngày chạy so với ${wd} ngày làm việc trong tháng.</p></div></div>
      <div class="loads">${loads || '<div class="empty">Chưa khai báo line. Thêm ở Danh mục.</div>'}</div></section>
    ${matPanel(mat, 'Nhu cầu NVL theo kế hoạch', 'Tính cho các lệnh chưa hoàn thành, không gồm lệnh huỷ.')}
  </div>`;
}
function exportPlan() {
  if (!window.XLSX) { toast('Chưa tải được thư viện Excel.'); return; }
  const per = state.settings.period, pMap = byCode(state.products), lMap = byCode(state.lines);
  const rows = state.plan.filter((r) => r.period === per);
  const aoa = [['Kỳ', 'Mã SP', 'Tên SP', 'ĐVT', 'Số lượng', 'Line', 'Tên line', 'Bắt đầu', 'Số ngày', 'Kết thúc', 'Trạng thái', 'Ghi chú'],
    ...rows.map((r) => { const d = planDays(r); return [periodLabel(per), r.product, pMap[r.product]?.name || '', pMap[r.product]?.unit || '', num(r.qty), r.line, lMap[r.line]?.name || '', r.start, d, addWorkdays(r.start, d), r.status, r.note]; })];
  const mat = materialNeeds(rows.filter((r) => r.status !== 'Huỷ' && r.status !== 'Hoàn thành').map((r) => ({ product: r.product, qty: num(r.qty) })));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), 'Ke hoach SX');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['Mã', 'Tên', 'Loại', 'ĐVT', 'Cần dùng', 'Tồn', 'Đang về', 'Chênh lệch'], ...mat.list.map((x) => [x.code, x.m?.name || '', x.m?.type || '', x.m?.unit || '', x.req, x.stock, x.inc, x.gap])]), 'Nhu cau NVL');
  XLSX.writeFile(wb, `Ke_hoach_SX_${per}.xlsx`);
}

/* ================= Màn hình: Danh mục ================= */
function viewMaster() {
  const tabs = [['products', `Sản phẩm · ${state.products.length}`], ['materials', `NVL - phụ gia · ${state.materials.length}`], ['bom', `BOM · ${state.bom.length}`], ['lists', 'Kênh · Kho · Line']];
  const k = ui.master;
  let body;
  if (k === 'lists') body = listsPanel();
  else body = importPanel(k) + gridPanel(k, SOURCES[k].label, '');
  return `<div class="notice info">Mã sản phẩm và mã NVL - phụ gia ở đây là <b>mã chuẩn</b> dùng chung cho mọi nguồn. Mã trong file nhập không khớp sẽ được đánh dấu đỏ.</div>
  <div class="seg" role="group" aria-label="Chọn danh mục">${tabs.map(([x, l]) => `<button data-master="${x}" aria-pressed="${x === k}">${l}</button>`).join('')}</div>${body}`;
}
function listsPanel() {
  const simple = (key, title, ph) => `<section class="panel"><div class="panel-head"><div><h2>${title}</h2><p>Dùng cho danh sách chọn nhanh.</p></div></div>
    <div class="row">${state[key].map((c, i) => `<span class="chip neutral" style="padding:4px 4px 4px 10px">${esc(c)}<button class="icon-btn" style="width:24px;height:24px;font-size:15px" data-dellist="${key}" data-i="${i}" aria-label="Xoá ${esc(c)}">×</button></span>`).join('')}</div>
    <form class="row" data-addlist="${key}"><input placeholder="${ph}" required style="flex:1;min-width:160px"><button class="btn sm" type="submit">Thêm</button></form></section>`;
  return `<div class="grid-2">${simple('channels', 'Kênh bán', 'vd Xuất khẩu')}${simple('warehouses', 'Kho', 'vd Kho TP Đà Nẵng')}</div>
  <section class="panel"><div class="panel-head"><div><h2>Line sản xuất</h2><p>Công suất dùng để tính số ngày chạy và tải line.</p></div><button class="btn sm ghost" data-act="addLine">+ Thêm line</button></div>
    <div class="table-wrap"><table><thead><tr><th>Mã line</th><th>Tên line</th><th class="num">Công suất / ngày (ĐVT SP)</th><th></th></tr></thead><tbody>
    ${state.lines.map((l, i) => `<tr><td><input class="mono" value="${esc(l.code)}" data-line="${i}" data-f="code" aria-label="Mã line"></td><td><input value="${esc(l.name)}" data-line="${i}" data-f="name" aria-label="Tên line"></td><td><input type="number" value="${esc(l.capacity)}" data-line="${i}" data-f="capacity" aria-label="Công suất"></td><td><button class="icon-btn" data-delline="${i}" aria-label="Xoá line">×</button></td></tr>`).join('')}
    </tbody></table></div></section>`;
}

/* ================= Màn hình: Cài đặt & đám mây ================= */
const SQL = `create table if not exists tkh_state (
  id text primary key,          -- mã nhà máy / workspace
  data jsonb not null,
  updated_at timestamptz default now()
);
alter table tkh_state enable row level security;
-- Bản đầu: cho phép người dùng đã đăng nhập đọc/ghi.
create policy "tkh rw" on tkh_state for all
  to anon, authenticated using (true) with check (true);`;
function viewSettings() {
  const c = state.settings.cloud;
  return `<section class="panel">
    <div class="panel-head"><div><h2>Đề xuất lưu trữ đám mây</h2><p>App luôn lưu trên máy (chạy được khi mất mạng). Đám mây dùng để nhiều người cùng xem một bộ số liệu.</p></div></div>
    <div class="table-wrap"><table class="stack"><thead><tr><th>Phương án</th><th>Phù hợp khi</th><th>Chi phí</th><th>Đánh giá</th></tr></thead><tbody>
      <tr><td data-l="Phương án"><b>Supabase</b> (PostgreSQL)</td><td data-l="Phù hợp khi">Dữ liệu dạng bảng, cần phân quyền đăng nhập, sau này làm báo cáo SQL</td><td data-l="Chi phí">Miễn phí 500MB, gói Pro ~25 USD/tháng</td><td data-l="Đánh giá"><span class="chip ok">Khuyến nghị</span></td></tr>
      <tr><td data-l="Phương án"><b>Firebase Firestore</b></td><td data-l="Phù hợp khi">Cần đồng bộ thời gian thực giữa nhiều máy</td><td data-l="Chi phí">Miễn phí hạn mức nhỏ, tính theo lượt đọc/ghi</td><td data-l="Đánh giá"><span class="chip neutral">Phương án 2</span></td></tr>
      <tr><td data-l="Phương án"><b>SharePoint / OneDrive</b> (Microsoft 365)</td><td data-l="Phù hợp khi">Công ty đã dùng Microsoft 365, muốn giữ file Excel làm gốc</td><td data-l="Chi phí">Có sẵn trong M365</td><td data-l="Đánh giá"><span class="chip neutral">Nếu đã có M365</span></td></tr>
      <tr><td data-l="Phương án"><b>Google Sheets</b> + Apps Script</td><td data-l="Phù hợp khi">Quy mô nhỏ, muốn xem dữ liệu ngay trên Sheet</td><td data-l="Chi phí">Miễn phí</td><td data-l="Đánh giá"><span class="chip warn">Chậm khi nhiều dòng</span></td></tr>
    </tbody></table></div>
  </section>
  <section class="panel">
    <div class="panel-head"><div><h2>Kết nối Supabase</h2><p>Tạo project tại supabase.com, chạy đoạn SQL bên dưới trong SQL Editor, rồi dán URL và anon key.</p></div>
      <span>${cloudReady() ? '<span class="chip ok">Đã cấu hình</span>' : '<span class="chip neutral">Chưa kết nối</span>'}</span></div>
    <form class="form-grid" id="cloudForm">
      <label class="field" style="grid-column:1/-1"><span>Project URL</span><input id="cUrl" value="${esc(c.url)}" placeholder="https://xxxx.supabase.co" inputmode="url" autocomplete="off"></label>
      <label class="field" style="grid-column:1/-1"><span>Anon public key</span><input id="cKey" value="${esc(c.key)}" placeholder="eyJhbGciOi…" autocomplete="off"></label>
      <label class="field"><span>Mã nhà máy / workspace</span><input id="cWs" value="${esc(c.workspace)}"></label>
      <label class="field"><span>Tự đồng bộ khi lưu</span>${selectHTML('cAuto', [['0', 'Tắt'], ['1', 'Bật']], c.auto ? '1' : '0')}</label>
      <button class="btn" type="submit">Lưu cấu hình</button>
    </form>
    <div class="row"><button class="btn ghost" data-act="cloudPush" ${cloudReady() ? '' : 'disabled'}>Đẩy lên đám mây</button><button class="btn ghost" data-act="cloudPull" ${cloudReady() ? '' : 'disabled'}>Tải từ đám mây về</button>
      <small class="muted">${state.meta.cloudAt ? 'Đồng bộ lần cuối: ' + relTime(state.meta.cloudAt) : ''}</small></div>
    <details><summary>SQL tạo bảng</summary><pre class="sql">${esc(SQL)}</pre></details>
  </section>
  <section class="panel">
    <div class="panel-head"><div><h2>Sao lưu & dữ liệu</h2><p>Sao lưu toàn bộ dữ liệu ra file JSON hoặc khôi phục.</p></div></div>
    <div class="row">
      <button class="btn ghost" data-act="backup">Tải bản sao lưu</button>
      <label class="btn ghost" for="restoreIn">Khôi phục từ file<input type="file" id="restoreIn" accept=".json" hidden></label>
      <button class="btn ghost" data-act="loadSample">Nạp dữ liệu mẫu</button>
      <button class="btn danger" data-act="wipe">Xoá toàn bộ dữ liệu</button>
    </div>
  </section>
  <section class="panel">
    <div class="panel-head"><div><h2>Cài lên điện thoại</h2></div></div>
    <p><b>iPhone (Safari):</b> bấm nút Chia sẻ → “Thêm vào MH chính”.<br><b>Android (Chrome):</b> menu ⋮ → “Cài đặt ứng dụng” hoặc “Thêm vào màn hình chính”.<br>
    <span class="muted">App mở toàn màn hình, không có thanh địa chỉ, và vẫn dùng được khi mất mạng.</span></p>
  </section>`;
}
const cloudReady = () => { const c = state.settings.cloud; return !!(c.url && c.key && c.workspace); };
function cloudHeaders() { const c = state.settings.cloud; return { apikey: c.key, Authorization: `Bearer ${c.key}`, 'Content-Type': 'application/json' }; }
function snapshot() { const s = JSON.parse(JSON.stringify(state)); s.settings.cloud = { ...s.settings.cloud, key: '' }; return s; }
async function cloudPush(silent) {
  const c = state.settings.cloud;
  try {
    const r = await fetch(`${c.url.replace(/\/$/, '')}/rest/v1/${c.table}`, {
      method: 'POST', headers: { ...cloudHeaders(), Prefer: 'resolution=merge-duplicates,return=minimal' },
      body: JSON.stringify({ id: c.workspace, data: snapshot(), updated_at: new Date().toISOString() }),
    });
    if (!r.ok) throw new Error(`${r.status} ${await r.text()}`);
    state.meta.cloudAt = new Date().toISOString(); DB.set('state', state);
    if (!silent) { toast('Đã đẩy dữ liệu lên đám mây'); render(); }
  } catch (e) { toast('Đồng bộ thất bại: ' + e.message.slice(0, 120)); }
}
async function cloudPull() {
  const c = state.settings.cloud;
  try {
    const r = await fetch(`${c.url.replace(/\/$/, '')}/rest/v1/${c.table}?id=eq.${encodeURIComponent(c.workspace)}&select=data,updated_at`, { headers: cloudHeaders() });
    if (!r.ok) throw new Error(`${r.status} ${await r.text()}`);
    const [row] = await r.json();
    if (!row) { toast('Chưa có dữ liệu trên đám mây cho workspace này.'); return; }
    if (!(await ask('Tải dữ liệu đám mây?', `Bản trên đám mây cập nhật ${relTime(row.updated_at)}. Dữ liệu trên máy này sẽ bị thay thế.`, 'Tải về'))) return;
    const keep = state.settings.cloud;
    state = migrate(row.data); state.settings.cloud = keep; state.meta.cloudAt = new Date().toISOString();
    save(); render(); toast('Đã tải dữ liệu từ đám mây');
  } catch (e) { toast('Tải thất bại: ' + e.message.slice(0, 120)); }
}
function migrate(s) {
  const base = emptyState();
  const out = { ...base, ...s, meta: { ...base.meta, ...(s?.meta || {}) }, settings: { ...base.settings, ...(s?.settings || {}) } };
  out.settings.cloud = { ...base.settings.cloud, ...(s?.settings?.cloud || {}) };
  return out;
}

/* ================= Toast & hộp xác nhận ================= */
let toastTimer;
function toast(msg) { const t = $('#toast'); t.textContent = msg; t.hidden = false; clearTimeout(toastTimer); toastTimer = setTimeout(() => (t.hidden = true), 3200); }
function ask(title, body, okLabel = 'Đồng ý', danger) {
  return new Promise((res) => {
    $('#modalTitle').textContent = title; $('#modalBody').textContent = body;
    const ok = $('#modalOk'); ok.textContent = okLabel; ok.className = danger ? 'btn danger' : 'btn';
    $('#modal').hidden = false; ok.focus();
    const done = (v) => { $('#modal').hidden = true; ok.onclick = $('#modalCancel').onclick = null; res(v); };
    ok.onclick = () => done(true); $('#modalCancel').onclick = () => done(false);
  });
}

/* ================= Sự kiện ================= */
function blankRow(key) {
  const per = state.settings.period;
  return {
    forecast: { period: per, product: '', channel: state.channels[0] || '', qty: 0 },
    fgStock: { product: '', warehouse: state.warehouses[0] || '', qty: 0 },
    rmStock: { material: '', warehouse: state.warehouses[state.warehouses.length - 1] || '', qty: 0, incoming: 0 },
    adjust: { period: per, product: '', channel: state.channels[0] || '', qty: 0, reason: REASONS[0], note: '' },
    products: { code: '', name: '', unit: 'hộp', group: '', batch: 0, safetyDays: 7, line: state.lines[0]?.code || '' },
    materials: { code: '', name: '', type: 'NVL', unit: 'kg' },
    bom: { product: '', material: '', qty: 0 },
  }[key];
}
document.addEventListener('click', async (e) => {
  const t = e.target.closest('button, [data-go]'); if (!t) return;
  const d = t.dataset;
  if (d.go) return go(d.go);
  if (d.src) { ui.src = d.src; ui.importer = null; return render(); }
  if (d.master) { ui.master = d.master; ui.importer = null; return render(); }
  if (d.del) { state[d.del].splice(+d.i, 1); save(d.del); return render(); }
  if (d.delplan) { state.plan.splice(+d.delplan, 1); save('plan'); return render(); }
  if (d.delline) { state.lines.splice(+d.delline, 1); save(); return render(); }
  if (d.dellist) { state[d.dellist].splice(+d.i, 1); save(); return render(); }
  switch (d.act) {
    case 'template': return downloadTemplate(d.key);
    case 'impCancel': ui.importer = null; return render();
    case 'impCommit': return commitImport();
    case 'addRow': state[d.key].unshift(blankRow(d.key)); save(d.key); return render();
    case 'clearKey': if (await ask('Xoá hết dữ liệu?', `Xoá toàn bộ ${state[d.key].length} dòng ${SOURCES[d.key].label}.`, 'Xoá hết', true)) { state[d.key] = []; save(d.key); render(); } return;
    case 'genPlan': return genPlan();
    case 'exportPre': return exportPre();
    case 'exportPlan': return exportPlan();
    case 'approveAll': state.plan.forEach((r) => { if (r.period === state.settings.period && r.status === 'Nháp') r.status = 'Đã duyệt'; }); save('plan'); render(); return toast('Đã duyệt các lệnh nháp');
    case 'addPlan': state.plan.push({ id: Math.random().toString(36).slice(2, 9), period: state.settings.period, product: '', qty: 0, line: state.lines[0]?.code || '', start: firstWorkday(state.settings.period), status: 'Nháp', note: '' }); save('plan'); return render();
    case 'addLine': state.lines.push({ code: 'L' + (state.lines.length + 1), name: 'Line mới', capacity: 1000 }); save(); return render();
    case 'clearSample': if (await ask('Xoá dữ liệu mẫu?', 'Toàn bộ dữ liệu mẫu (danh mục, forecast, tồn kho, kế hoạch) sẽ bị xoá. Kênh, kho và line được giữ lại để bạn sửa.', 'Xoá mẫu', true)) { const keep = state; state = emptyState(); state.channels = keep.channels; state.warehouses = keep.warehouses; state.lines = keep.lines; state.settings = keep.settings; save(); render(); toast('Đã xoá dữ liệu mẫu'); } return;
    case 'loadSample': if (await ask('Nạp dữ liệu mẫu?', 'Dữ liệu hiện tại sẽ bị thay bằng dữ liệu mẫu.', 'Nạp mẫu', true)) { const cloud = state.settings.cloud; state = sampleState(); state.settings.cloud = cloud; save(); render(); } return;
    case 'wipe': if (await ask('Xoá toàn bộ dữ liệu?', 'Mọi danh mục, nguồn dữ liệu và kế hoạch trên máy này sẽ bị xoá. Nên tải bản sao lưu trước.', 'Xoá toàn bộ', true)) { const cloud = state.settings.cloud; state = emptyState(); state.settings.cloud = cloud; save(); render(); } return;
    case 'backup': {
      const blob = new Blob([JSON.stringify(snapshot(), null, 1)], { type: 'application/json' });
      const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = `tien-ke-hoach-${dateStr(new Date())}.json`; a.click();
      return setTimeout(() => URL.revokeObjectURL(a.href), 1000);
    }
    case 'cloudPush': return cloudPush();
    case 'cloudPull': return cloudPull();
  }
});
document.addEventListener('change', (e) => {
  const t = e.target, d = t.dataset;
  if (t.id === 'periodSel') { state.settings.period = t.value; save(); return render(); }
  if (t.id === 'fileIn' && t.files[0]) return handleFile(t.files[0], d.key);
  if (t.id === 'restoreIn' && t.files[0]) {
    t.files[0].text().then(async (txt) => {
      try { const s = JSON.parse(txt); if (!(await ask('Khôi phục dữ liệu?', 'Dữ liệu hiện tại sẽ bị thay bằng bản sao lưu.', 'Khôi phục', true))) return; const cloud = state.settings.cloud; state = migrate(s); state.settings.cloud = cloud; save(); render(); toast('Đã khôi phục'); }
      catch { toast('File sao lưu không hợp lệ.'); }
    }); return;
  }
  if (t.id === 'preFilter') { ui.preFilter = t.value; return render(); }
  if (d.allp) { ui.filter[d.allp + ':all'] = t.checked; return render(); }
  // Trình nhập Excel
  const imp = ui.importer;
  if (imp) {
    if ('impMap' in d) { imp.map[d.impMap] = t.value; return render(); }
    if ('impDef' in d) { imp.defaults[d.impDef] = t.value; return render(); }
    if ('impSheet' in d) { imp.sheet = t.value; imp.headerRow = null; loadSheet(imp); return render(); }
    if ('impHdr' in d) { imp.headerRow = +t.value - 1; imp.headers = (imp.data[imp.headerRow] || []).map((h) => String(h).trim()); imp.map = guessMap(SOURCES[imp.key].fields, imp.headers); return render(); }
    if ('impMode' in d) { imp.mode = t.value; return; }
  }
  // Ô sửa trực tiếp
  if (d.k && d.f) {
    const f = SOURCES[d.k].fields.find((x) => x.k === d.f);
    state[d.k][+d.i][d.f] = f.type === 'num' ? num(t.value) : t.value.trim();
    save(d.k); if (f.type === 'product' || f.type === 'material') render(); return;
  }
  if (d.plan != null && d.f) { const r = state.plan[+d.plan]; r[d.f] = d.f === 'qty' ? num(t.value) : t.value; save('plan'); return render(); }
  if (d.line != null && d.f) {
    const l = state.lines[+d.line], old = l.code; l[d.f] = d.f === 'capacity' ? num(t.value) : t.value.trim();
    if (d.f === 'code' && old !== l.code) { state.products.forEach((p) => { if (p.line === old) p.line = l.code; }); state.plan.forEach((p) => { if (p.line === old) p.line = l.code; }); }
    save(); return;
  }
});
let filterTimer;
document.addEventListener('input', (e) => {
  const d = e.target.dataset;
  if (d.filter != null) {
    ui.filter[d.filter] = e.target.value; clearTimeout(filterTimer);
    filterTimer = setTimeout(() => { const pos = e.target.selectionStart; render(); const el = document.querySelector(`[data-filter="${d.filter}"]`); if (el) { el.focus(); try { el.setSelectionRange(pos, pos); } catch {} } }, 250);
  }
});
document.addEventListener('submit', (e) => {
  e.preventDefault();
  const f = e.target;
  if (f.id === 'adjForm') {
    state.adjust.unshift({ period: $('#adjPeriod').value, product: $('#adjProduct').value.trim(), channel: $('#adjChannel').value, qty: num($('#adjQty').value), reason: $('#adjReason').value, note: $('#adjNote').value.trim() });
    save('adjust'); render(); toast('Đã thêm cân chỉnh'); return;
  }
  if (f.id === 'cloudForm') {
    Object.assign(state.settings.cloud, { url: $('#cUrl').value.trim(), key: $('#cKey').value.trim(), workspace: $('#cWs').value.trim() || 'nha-may-1', auto: $('#cAuto').value === '1' });
    save(); render(); toast('Đã lưu cấu hình đám mây'); return;
  }
  if (f.dataset.addlist) {
    const k = f.dataset.addlist, v = f.querySelector('input').value.trim();
    if (v && !state[k].includes(v)) { state[k].push(v); save(); render(); }
  }
});
// Kéo thả file vào vùng nhập
document.addEventListener('dragover', (e) => { const z = e.target.closest?.('#drop'); if (z) { e.preventDefault(); z.classList.add('over'); } });
document.addEventListener('dragleave', (e) => { const z = e.target.closest?.('#drop'); if (z) z.classList.remove('over'); });
document.addEventListener('drop', (e) => {
  const z = e.target.closest?.('#drop'); if (!z) return;
  e.preventDefault(); z.classList.remove('over');
  const file = e.dataTransfer.files[0]; if (file) handleFile(file, $('#fileIn').dataset.key);
});

/* ================= Cài đặt PWA ================= */
let deferredPrompt;
window.addEventListener('beforeinstallprompt', (e) => { e.preventDefault(); deferredPrompt = e; $('#installBtn').hidden = false; });
$('#installBtn').addEventListener('click', async () => { if (!deferredPrompt) return; deferredPrompt.prompt(); await deferredPrompt.userChoice; deferredPrompt = null; $('#installBtn').hidden = true; });
if ('serviceWorker' in navigator && location.protocol !== 'file:') {
  window.addEventListener('load', () => navigator.serviceWorker.register('sw.js').catch(() => {}));
}

/* ================= Khởi động ================= */
(async function init() {
  await DB.open();
  const saved = await DB.get('state');
  state = saved ? migrate(saved) : sampleState();
  try { const v = localStorage.getItem('tkh:view'); if (v && VIEWS.some((x) => x[0] === v)) ui.view = v; } catch {}
  render();
  if (!saved) DB.set('state', state);
})();
