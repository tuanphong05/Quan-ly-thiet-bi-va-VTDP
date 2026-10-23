// === NHẬP DỮ LIỆU TỪ SHEET GỐC VÀO WEB ===
// Dán vào Apps Script (cùng project với Code.gs, tạo thêm file Import.gs),
// sửa 3 dòng cấu hình bên dưới, chọn hàm nhapThietBi rồi bấm Chạy.
const NGUON = 'TỦ ACU';  // tên sheet gốc đang chứa bảng
const DONG_TD = 2;       // số dòng chứa tiêu đề cột (dữ liệu bắt đầu ngay dưới dòng này)
const NHOM = 'Hệ thống PSR/MSSR Easat EA29190'; // phải trùng đúng tên 1 trong 7 hệ thống trên web:
// Hệ thống PSR/MSSR Easat EA29190 | Hệ thống VCCS | Hệ thống VHF | Hệ thống RAPS | Hệ thống KVM | Hệ thống Radar thời tiết WRS400 | Hệ thống ATM

const TU = 'Tủ ACU';     // tủ thiết bị: Tủ ACU | Tủ PTC | Tủ RSCC | Tủ MSSR | Tủ COMMS | Tủ Radar Distribution | Tủ Transponder

// Thứ tự cột sheet gốc: A Stt | B Tên vật tư | C PN | D SN | E Tình trạng | F Nơi đặt | G Ghi chú | H Mô tả
function nhapThietBi() {
  const src = SS().getSheetByName(NGUON);
  if (!src) throw Error('Không tìm thấy sheet "' + NGUON + '"');
  const rg = src.getDataRange(), v = rg.getValues();
  // Ô gộp (merge) chỉ giữ giá trị ở ô đầu, nên chép giá trị đó ra cả vùng gộp
  rg.getMergedRanges().forEach(m => {
    const r0 = m.getRow() - 1, c0 = m.getColumn() - 1, val = v[r0][c0];
    for (let i = 0; i < m.getNumRows(); i++)
      for (let j = 0; j < m.getNumColumns(); j++) v[r0 + i][c0 + j] = val;
  });
  const dst = SS().getSheetByName('equip'), h = H.equip.split(',');
  const bo = g => String(g || '').split(' | ').filter(x => x && !x.startsWith('Nơi đặt:')).join(' | ');
  const kh = (tu, ten, pn, sn, g) => [tu, ten, pn, sn, g].map(x => String(x || '').trim()).join('|');
  const have = new Map(); // khóa -> số dòng giống hệt đã có trong sheet equip
  rd('equip').forEach(x => { const k = kh(x.tu, x.ten, x.pn, x.sn, bo(x.ghichu)); have.set(k, (have.get(k) || 0) + 1); });
  const out = [];
  for (let i = DONG_TD; i < v.length; i++) {
    const [, ten, pn, sn, , noi, dp, mota] = v[i].map(x => String(x).trim());
    if (!ten) continue;
    const ghichu = [mota, dp].filter(Boolean).join(' | '); // cột Ghi chú: mô tả + "Không có dự phòng"; nơi đặt đã có ở cột Tủ
    const key = kh(TU, ten, pn, sn, ghichu);
    if (have.get(key) > 0) { have.set(key, have.get(key) - 1); continue; } // dòng đã nhập rồi: bỏ qua, chạy lại không bị trùng
    out.push({ id: nid(), ten, sl: 1, nhom: NHOM, tu: TU, tt: 'Tốt', sn, pn, anh: '', ghichu,
               capnhat: today(), nguoi: 'Nhập từ sheet gốc' });
  }
  if (!out.length) { Logger.log('Không có dòng mới để nhập.'); return; }
  dst.getRange(dst.getLastRow() + 1, 1, out.length, h.length)
     .setValues(out.map(o => h.map(c => o[c] ?? '')));
  Logger.log('Đã nhập ' + out.length + ' dòng vào sheet equip.');
}

// Chạy 1 lần nếu trước đó đã nhập dữ liệu với nhóm cũ (tên nhóm = tên tab nguồn):
// sửa NHOM cho đúng hệ thống rồi chạy hàm này để đổi nhóm các dòng cũ trong sheet equip.
function doiNhomCu() {
  const s = SS().getSheetByName('equip'), n = s.getLastRow();
  if (n < 2) return;
  const col = H.equip.split(',').indexOf('nhom') + 1, r = s.getRange(2, col, n - 1, 1);
  let k = 0;
  r.setValues(r.getValues().map(([x]) => { if (x === NGUON) { k++; return [NHOM]; } return [x]; }));
  Logger.log('Đã đổi nhóm ' + k + ' dòng.');
}

// Chạy 1 lần để làm sạch Ghi chú của các dòng đã nhập trước đó: bỏ đoạn "Nơi đặt: ..." (đã có ở cột Tủ).
function lamSachGhiChu() {
  const s = SS().getSheetByName('equip'), n = s.getLastRow();
  if (n < 2) return;
  const col = H.equip.split(',').indexOf('ghichu') + 1, r = s.getRange(2, col, n - 1, 1);
  let k = 0;
  r.setValues(r.getValues().map(([x]) => {
    const y = String(x).split(' | ').filter(t => t && !t.startsWith('Nơi đặt:')).join(' | ');
    if (y !== x) k++; return [y];
  }));
  Logger.log('Đã làm sạch ' + k + ' dòng.');
}

// =====================================================================
// NHẬP VẬT TƯ DỰ PHÒNG TỪ SHEET GỐC VÀO WEB (sheet "spare")
// Cách dùng: sửa khối cấu hình VT bên dưới cho khớp tab cần nhập rồi chạy hàm nhapVatTuDuPhong.
// COT: điền CHỮ CÁI CỘT trong sheet gốc (A, B, C...). Để '' nếu sheet gốc không có cột đó.
// =====================================================================
const VT = {
  NGUON: 'Spare - Tủ PTC',                     // tên tab chứa bảng vật tư dự phòng
  DONG_TD: 2,                                  // bỏ qua 2 dòng đầu (tiêu đề cột + dòng gộp "TỦ PTC"); dữ liệu bắt đầu từ dòng 3
  NHOM: 'Hệ thống PSR/MSSR Easat EA29190',     // phải trùng 1 trong 7 hệ thống trên web
  TU: 'Tủ PTC',                                // tủ thiết bị ('' nếu hệ thống không chia tủ)
  COT: { ten: 'B', pn: 'C', sn: 'D', sl: '', tt: '', cauhinh: '', ghichu: 'G' }
};

function nhapVatTuDuPhong() {
  const src = SS().getSheetByName(VT.NGUON);
  if (!src) throw Error('Không tìm thấy sheet "' + VT.NGUON + '"');
  const rg = src.getDataRange(), v = rg.getValues();
  rg.getMergedRanges().forEach(m => {  // ô gộp: chép giá trị ra cả vùng gộp
    const r0 = m.getRow() - 1, c0 = m.getColumn() - 1, val = v[r0][c0];
    for (let i = 0; i < m.getNumRows(); i++)
      for (let j = 0; j < m.getNumColumns(); j++) v[r0 + i][c0 + j] = val;
  });
  const ci = l => l ? l.toUpperCase().charCodeAt(0) - 65 : -1;
  const lay = (row, k) => { const i = ci(VT.COT[k]); return i < 0 ? '' : String(row[i] ?? '').trim(); };
  const TT = ['Tốt', 'Cảnh báo', 'Hư hỏng', 'Đang sửa chữa'];
  const kh = (tu, ten, pn, sn, g) => [tu, ten, pn, sn, g].map(x => String(x || '').trim()).join('|');
  const dst = SS().getSheetByName('spare'), h = H.spare.split(',');
  const have = new Map();
  rd('spare').forEach(x => { const k = kh(x.tu, x.ten, x.pn, x.sn, x.ghichu); have.set(k, (have.get(k) || 0) + 1); });
  const out = [];
  for (let i = VT.DONG_TD; i < v.length; i++) {
    const ten = lay(v[i], 'ten'); if (!ten) continue;
    const pn = lay(v[i], 'pn'), sn = lay(v[i], 'sn'), ghichu = lay(v[i], 'ghichu'), cauhinh = lay(v[i], 'cauhinh');
    const sl = parseInt(lay(v[i], 'sl'), 10), tt = lay(v[i], 'tt');
    const key = kh(VT.TU, ten, pn, sn, ghichu);
    if (have.get(key) > 0) { have.set(key, have.get(key) - 1); continue; }
    out.push({ id: nid(), ten, sl: isNaN(sl) ? 1 : sl, nhom: VT.NHOM, tu: VT.TU, tt: TT.indexOf(tt) >= 0 ? tt : 'Tốt',
               sn, pn, anh: '', anhtem: '', cauhinh, ghichu, capnhat: today(), nguoi: 'Nhập từ sheet gốc' });
  }
  if (!out.length) { Logger.log('Không có dòng mới để nhập.'); return; }
  dst.getRange(dst.getLastRow() + 1, 1, out.length, h.length).setValues(out.map(o => h.map(c => o[c] ?? '')));
  Logger.log('Đã nhập ' + out.length + ' dòng vào sheet spare.');
}
