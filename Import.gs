// === NHẬP DỮ LIỆU TỪ SHEET GỐC VÀO WEB ===
// Dán vào Apps Script (cùng project với Code.gs, tạo thêm file Import.gs),
// sửa 3 dòng cấu hình bên dưới, chọn hàm nhapThietBi rồi bấm Chạy.
const NGUON = 'TỦ ACU';  // tên sheet gốc đang chứa bảng
const DONG_TD = 2;       // số dòng chứa tiêu đề cột (dữ liệu bắt đầu ngay dưới dòng này)
const NHOM = 'TỦ ACU';   // giá trị cột "Thuộc nhóm hệ thống" hiển thị trên web

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
    out.push({ id: nid(), ten, sl: 1, nhom: NHOM, tt: 'Tốt', sn, pn, anh: '', ghichu,
               capnhat: today(), nguoi: 'Nhập từ sheet gốc' });
  }
  if (!out.length) { Logger.log('Không có dòng mới để nhập.'); return; }
  dst.getRange(dst.getLastRow() + 1, 1, out.length, h.length)
     .setValues(out.map(o => h.map(c => o[c] ?? '')));
  Logger.log('Đã nhập ' + out.length + ' dòng vào sheet equip.');
}
