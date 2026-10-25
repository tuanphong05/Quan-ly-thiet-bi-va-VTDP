// === NHẬP DỮ LIỆU THEO MẪU "PHÂN HỆ ANTEN" VÀO WEB ===
// Dán vào Apps Script (cùng project với Code.gs, file riêng tên Import.gs).
// Cách dùng: dán sheet theo mẫu vào cùng file Google Sheet (thành một tab), sửa khối NHAP bên dưới,
// chọn hàm nhapTheoMau rồi bấm Chạy.
//
// Hàm nhận cột theo TÊN TIÊU ĐỀ (không phụ thuộc thứ tự cột) nên chỉ cần giữ đúng tên cột của mẫu.
// Hàm hiểu cấu trúc mẫu:
//   - Dòng chỉ có 1 ô chữ ở cột "Loại hệ thống"  -> tên loại hệ thống (VD "Radar PSR/SSR")
//   - Dòng chỉ có 1 ô chữ ở cột khác              -> tên phân hệ / tủ (VD "Phân hệ Anten"), hiện thành tab trên web
//   - Các dòng còn lại có tên                     -> vật tư (khối thiết bị hoặc vật tư dự phòng)
const NHAP = {
  NGUON: 'Phân hệ anten',                    // tên tab chứa bảng theo mẫu
  DICH: 'equip',                             // 'equip' = Khối thiết bị | 'spare' = Vật tư dự phòng
  NHOM: 'Hệ thống PSR/MSSR',   // phải trùng đúng 1 trong 7 hệ thống trên web
  TU: '',  // để trống = lấy tên phân hệ ghi trong sheet; điền tên để ép tất cả vào một tab
  TU_CU: 'Tủ ACU'                               // chỉ dùng cho hàm thayTheTu: tên tủ cũ cần thay thế
};

// tên tiêu đề (viết thường, bỏ dấu) -> trường dữ liệu
const BO_COT = {
  'loai he thong': 'stt', 'stt': 'stt', 'so thu tu': 'stt', 'loai vat tu': 'loai', 'ten ky thuat': 'tenkt', 'ten su dung': 'ten',
  'don vi tinh/so luong': 'sl', 'dvt/so luong': 'sl', 'so luong': 'sl',
  'serial': 'sn', 'serial number': 'sn', 's/n': 'sn', 'p/n': 'pn', 'part number': 'pn',
  'hang san xuat': 'hang', 'xuat xu': 'xuatxu', 'nam san xuat': 'nam', 'he thong': 'hethong',
  'tinh trang vat tu': 'tinhtrang', 'phan nhom thiet bi': 'nhomtb', 'dia diem': 'diadiem',
  'don vi quan ly': 'donvi', 'tai san vtdp': 'taisan', 'tai san vat tu du phong': 'taisan',
  'nguyen gia': 'nguyengia', 'thong tin bo sung': 'ghichu',
  'tinh trang hoat dong': 'tt', 'phan he': 'tu', 'tu thiet bi': 'tu'
};
const chuan = s => String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '')
  .replace(/đ/g, 'd').replace(/\s+/g, ' ').trim();

// Đọc sheet nguồn theo mẫu, trả về danh sách vật tư (chưa kiểm tra trùng)
function docMau() {
  const src = SS().getSheetByName(NHAP.NGUON);
  if (!src) throw Error('Không tìm thấy sheet "' + NHAP.NGUON + '"');
  if (NHAP.DICH !== 'equip' && NHAP.DICH !== 'spare') throw Error('DICH phải là "equip" hoặc "spare"');
  const rg = src.getDataRange(), raw = rg.getDisplayValues(), v = raw.map(r => r.slice());
  // Ô gộp chỉ giữ giá trị ở ô đầu: chép giá trị đó ra cả vùng gộp (dùng giá trị hiển thị nên giữ nguyên số 0 đầu của Serial)
  rg.getMergedRanges().forEach(m => {
    const r0 = m.getRow() - 1, c0 = m.getColumn() - 1, val = raw[r0][c0];
    for (let i = 0; i < m.getNumRows(); i++)
      for (let j = 0; j < m.getNumColumns(); j++) v[r0 + i][c0 + j] = val;
  });

  // Tìm dòng tiêu đề: dòng đầu tiên (trong 10 dòng đầu) có từ 3 tên cột trở lên khớp mẫu
  let hr = -1, col = {};
  for (let i = 0; i < Math.min(raw.length, 10) && hr < 0; i++) {
    const c = {}; let n = 0;
    raw[i].forEach((t, j) => { const k = BO_COT[chuan(t)]; if (k && c[k] === undefined) { c[k] = j; n++; } });
    if (n >= 3) { hr = i; col = c; }
  }
  if (hr < 0) throw Error('Không thấy dòng tiêu đề theo mẫu (Loại vật tư, Tên sử dụng, Serial Number...)');

  const get = (r, k) => (col[k] === undefined ? '' : String(v[r][col[k]] || '').trim());
  const out = []; let phanHe = '';
  for (let i = hr + 1; i < raw.length; i++) {
    const co = raw[i].map((t, j) => [String(t || '').trim(), j]).filter(x => x[0]);
    if (!co.length) continue;
    if (co.length === 1 && !SO.test(co[0][0])) {          // dòng nhóm
      if (co[0][1] === col.stt) phanHe = ''; else phanHe = co[0][0];   // ô chữ ở cột đầu = tên loại hệ thống (bỏ qua)
      continue;
    }
    const ten = get(i, 'ten') || get(i, 'tenkt') || get(i, 'loai');
    if (!ten) continue;
    const o = {
      stt: get(i, 'stt'), loai: get(i, 'loai'), tenkt: get(i, 'tenkt'), ten,
      sl: get(i, 'sl') || '1', sn: get(i, 'sn'), pn: get(i, 'pn'), hang: get(i, 'hang'), xuatxu: get(i, 'xuatxu'),
      nam: get(i, 'nam'), hethong: get(i, 'hethong'),
      tinhtrang: get(i, 'tinhtrang') || (NHAP.DICH === 'spare' ? 'Dự phòng' : 'Đang sử dụng'),
      nhomtb: get(i, 'nhomtb'), diadiem: get(i, 'diadiem'), donvi: get(i, 'donvi'), taisan: get(i, 'taisan'),
      nguyengia: get(i, 'nguyengia'), ghichu: get(i, 'ghichu'),
      nhom: NHAP.NHOM, tu: NHAP.TU || get(i, 'tu') || phanHe, tt: get(i, 'tt') || 'Tốt'
    };
    out.push(o);
  }
  return out;
}

// Khóa nhận diện một vật tư
const SO = /^[\d.,]+$/;   // số thứ tự dạng 1, 2, 5,1, 5.2
const kh = x => [x.nhom, x.tu, x.loai, x.tenkt, x.ten, x.pn, x.sn].map(t => String(t || '').trim()).join('|');

function nhapTheoMau() {
  capNhatSTT();   // điền luôn số thứ tự cho các dòng cũ còn trống
  const dst = SS().getSheetByName(NHAP.DICH), h = hdr(NHAP.DICH);
  const have = new Map();   // khóa -> số dòng giống hệt đã có: chạy lại không bị nhân đôi
  rd(NHAP.DICH).forEach(x => { const k = kh(x); have.set(k, (have.get(k) || 0) + 1); });
  const out = [];
  docMau().forEach(o => {
    const k = kh(o);
    if (have.get(k) > 0) { have.set(k, have.get(k) - 1); return; }
    out.push(Object.assign({ id: nid(), anh: '', anhtem: '', cauhinh: '', capnhat: today(), nguoi: 'Nhập từ sheet gốc' }, o));
  });
  if (!out.length) { Logger.log('Không có dòng mới để nhập.'); return; }
  const dau = dst.getLastRow() + 1, cuoi = dau + out.length - 1;
  if (cuoi > dst.getMaxRows()) dst.insertRowsAfter(dst.getMaxRows(), cuoi - dst.getMaxRows());
  dst.getRange(dau, 1, out.length, h.length).setValues(out.map(o => h.map(c => o[c] ?? '')));
  Logger.log('Đã nhập ' + out.length + ' dòng vào sheet ' + NHAP.DICH + ' (phân hệ: ' + (NHAP.TU || 'theo sheet') + ').');
}

// Điền Số thứ tự (cột stt) cho các dòng ĐÃ nhập trước đó mà còn trống, lấy từ sheet nguồn. Không xóa/không thêm dòng nào.
function capNhatSTT() {
  const items = docMau(), q = new Map();
  items.forEach(o => { const k = kh(o); if (!q.has(k)) q.set(k, []); q.get(k).push(o.stt); });
  const s = SS().getSheetByName(NHAP.DICH); hdr(NHAP.DICH);
  const v = s.getDataRange().getDisplayValues(), h = v[0], ci = h.indexOf('stt');
  const idx = {}; ['nhom', 'tu', 'loai', 'tenkt', 'ten', 'pn', 'sn'].forEach(c => idx[c] = h.indexOf(c));
  let n = 0;
  const col = v.slice(1).map(r => {
    if (r[ci] !== '' || r[0] === '') return [r[ci]];
    const k = kh(Object.fromEntries(Object.keys(idx).map(c => [c, r[idx[c]]]))), a = q.get(k);
    if (a && a.length) { n++; return [a.shift()]; }
    return [r[ci]];
  });
  if (col.length) s.getRange(2, ci + 1, col.length, 1).setNumberFormat('@').setValues(col);
  Logger.log('Đã điền số thứ tự cho ' + n + ' dòng.');
}

// THAY THẾ dữ liệu của một tủ: sao lưu sheet đích, xóa các dòng có phân hệ/tủ = NHAP.TU_CU, rồi nhập lại từ sheet nguồn.
// VD: đặt NGUON = 'Tủ điều khiển Anten', TU = 'Tủ điều khiển Anten', TU_CU = 'Tủ ACU' rồi chạy thayTheTu.
function thayTheTu() {
  if (!NHAP.TU_CU) throw Error('Chưa điền TU_CU');
  if (!SS().getSheetByName(NHAP.NGUON)) throw Error('Không tìm thấy sheet "' + NHAP.NGUON + '"');
  const dst = SS().getSheetByName(NHAP.DICH), h = hdr(NHAP.DICH), ci = h.indexOf('tu'), ni = h.indexOf('nhom');
  const v = dst.getDataRange().getDisplayValues();
  dst.copyTo(SS()).setName('saoluu_' + NHAP.DICH + '_' + Utilities.formatDate(new Date(), 'GMT+7', 'yyyyMMdd_HHmmss'));
  let n = 0;
  for (let i = v.length - 1; i >= 1; i--)
    if (v[i][ci] === NHAP.TU_CU && v[i][ni] === NHAP.NHOM) { dst.deleteRow(i + 1); n++; }
  Logger.log('Đã xóa ' + n + ' dòng của "' + NHAP.TU_CU + '" (đã sao lưu sheet trước khi xóa).');
  nhapTheoMau();
}

// XÓA toàn bộ dữ liệu của một tủ (cả Khối thiết bị lẫn Vật tư dự phòng) để tab của tủ đó biến mất khỏi web.
// Cách dùng: đặt NHAP.TU_CU = tên tủ cần xóa (VD 'Tủ Radar Distribution'), NHAP.NHOM = đúng nhóm hệ thống, rồi chạy xoaTu.
// Hàm tự sao lưu mỗi sheet thành sheet mới tên saoluu_... trước khi xóa.
function xoaTu() {
  if (!NHAP.TU_CU) throw Error('Chưa điền TU_CU');
  let n = 0;
  ['equip', 'spare'].forEach(k => {
    const dst = SS().getSheetByName(k), h = hdr(k), ci = h.indexOf('tu'), ni = h.indexOf('nhom');
    const v = dst.getDataRange().getDisplayValues();
    if (!v.slice(1).some(r => r[ci] === NHAP.TU_CU && r[ni] === NHAP.NHOM)) return;
    dst.copyTo(SS()).setName('saoluu_' + k + '_' + Utilities.formatDate(new Date(), 'GMT+7', 'yyyyMMdd_HHmmss'));
    for (let i = v.length - 1; i >= 1; i--)
      if (v[i][ci] === NHAP.TU_CU && v[i][ni] === NHAP.NHOM) { dst.deleteRow(i + 1); n++; }
  });
  Logger.log('Đã xóa ' + n + ' dòng của "' + NHAP.TU_CU + '" (có sao lưu trước khi xóa).');
}
