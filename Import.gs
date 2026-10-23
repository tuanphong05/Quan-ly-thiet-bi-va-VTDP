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
  const have = new Set(rd('equip').map(x => [x.ten, x.pn, x.ghichu].join('|')));
  const out = [];
  for (let i = DONG_TD; i < v.length; i++) {
    const [, ten, pn, sn, , noi, dp, mota] = v[i].map(x => String(x).trim());
    if (!ten) continue;
    const ghichu = [mota, noi && 'Nơi đặt: ' + noi, dp].filter(Boolean).join(' | ');
    const key = [ten, pn, ghichu].join('|');
    if (have.has(key)) continue; // chạy lại nhiều lần cũng không bị trùng
    have.add(key);
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
