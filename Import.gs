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
  NHOM: 'Hệ thống PSR/MSSR Easat EA29190',   // phải trùng đúng 1 trong 7 hệ thống trên web
  TU: ''                                     // để trống = lấy tên phân hệ ghi trong sheet; điền tên để ép tất cả vào một tab
};

// tên tiêu đề (viết thường, bỏ dấu) -> trường dữ liệu
const BO_COT = {
  'loai he thong': 'loaihe', 'loai vat tu': 'loai', 'ten ky thuat': 'tenkt', 'ten su dung': 'ten',
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

function nhapTheoMau() {
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

  const dst = SS().getSheetByName(NHAP.DICH), h = hdr(NHAP.DICH);
  const kh = x => [x.nhom, x.tu, x.loai, x.tenkt, x.ten, x.pn, x.sn].map(t => String(t || '').trim()).join('|');
  const have = new Map();   // khóa -> số dòng giống hệt đã có: chạy lại không bị nhân đôi
  rd(NHAP.DICH).forEach(x => { const k = kh(x); have.set(k, (have.get(k) || 0) + 1); });

  const get = (r, k) => (col[k] === undefined ? '' : String(v[r][col[k]] || '').trim());
  const out = []; let loaiHe = '', phanHe = '';
  for (let i = hr + 1; i < raw.length; i++) {
    const co = raw[i].map((t, j) => [String(t || '').trim(), j]).filter(x => x[0]);
    if (!co.length) continue;
    if (co.length === 1 && !/^\d+$/.test(co[0][0])) {          // dòng nhóm
      if (co[0][1] === col.loaihe) { loaiHe = co[0][0]; phanHe = ''; } else phanHe = co[0][0];
      continue;
    }
    const ten = get(i, 'ten') || get(i, 'tenkt') || get(i, 'loai');
    if (!ten) continue;
    const a = get(i, 'loaihe');                                 // cột A của dòng vật tư thường chứa số thứ tự
    const o = {
      loaihe: a && !/^\d+$/.test(a) ? a : loaiHe, loai: get(i, 'loai'), tenkt: get(i, 'tenkt'), ten,
      sl: get(i, 'sl') || '1', sn: get(i, 'sn'), pn: get(i, 'pn'), hang: get(i, 'hang'), xuatxu: get(i, 'xuatxu'),
      nam: get(i, 'nam'), hethong: get(i, 'hethong'),
      tinhtrang: get(i, 'tinhtrang') || (NHAP.DICH === 'spare' ? 'Dự phòng' : 'Đang sử dụng'),
      nhomtb: get(i, 'nhomtb'), diadiem: get(i, 'diadiem'), donvi: get(i, 'donvi'), taisan: get(i, 'taisan'),
      nguyengia: get(i, 'nguyengia'), ghichu: get(i, 'ghichu'),
      nhom: NHAP.NHOM, tu: NHAP.TU || get(i, 'tu') || phanHe, tt: get(i, 'tt') || 'Tốt'
    };
    const k = kh(o);
    if (have.get(k) > 0) { have.set(k, have.get(k) - 1); continue; }
    out.push(Object.assign({ id: nid(), anh: '', anhtem: '', cauhinh: '', capnhat: today(), nguoi: 'Nhập từ sheet gốc' }, o));
  }
  if (!out.length) { Logger.log('Không có dòng mới để nhập.'); return; }
  const dau = dst.getLastRow() + 1, cuoi = dau + out.length - 1;
  if (cuoi > dst.getMaxRows()) dst.insertRowsAfter(dst.getMaxRows(), cuoi - dst.getMaxRows());
  dst.getRange(dau, 1, out.length, h.length).setValues(out.map(o => h.map(c => o[c] ?? '')));
  Logger.log('Đã nhập ' + out.length + ' dòng vào sheet ' + NHAP.DICH + ' (phân hệ: ' + (NHAP.TU || 'theo sheet') + ').');
}
