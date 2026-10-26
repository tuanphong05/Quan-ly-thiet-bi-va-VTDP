// === QUẢN LÝ THIẾT BỊ & VẬT TƯ DỰ PHÒNG - LTH ===
// Dán vào Apps Script của Google Sheets, chạy setup() 1 lần, rồi Deploy > Web app
// (Execute as: Me, Who has access: Anyone).
// LƯU Ý khi dán đè code mới: điền lại FOLDER_ID (thư mục Drive lưu ảnh) ở dòng dưới.
const FOLDER_ID = 'DAN_ID_THU_MUC_DRIVE_LUU_ANH';

// Các cột theo mẫu sheet "Phân hệ anten" (thêm vào sau các cột cũ để không ảnh hưởng dữ liệu đã có):
//  stt=Số thứ tự (cột đầu của sheet mẫu, tiêu đề trong sheet ghi "Loại hệ thống") | loai=Loại vật tư | tenkt=Tên kỹ thuật | ten=Tên sử dụng | sl=Đơn vị tính/Số lượng
//  sn=Serial Number | pn=Part Number | hang=Hãng sản xuất | xuatxu=Xuất xứ | nam=Năm sản xuất | hethong=Hệ thống
//  tinhtrang=Tình trạng vật tư | nhomtb=Phân nhóm thiết bị | diadiem=Địa điểm | donvi=Đơn vị quản lý
//  taisan=Tài sản vật tư dự phòng | nguyengia=Nguyên giá | ghichu=Thông tin bổ sung
// Các cột riêng của web: nhom=nhóm hệ thống (menu) | tu=phân hệ/tủ (tab) | tt=tình trạng hoạt động (Tốt/Thiết bị gặp sự cố/Đang sửa chữa)
const MAU = 'stt,loai,tenkt,hang,xuatxu,nam,hethong,tinhtrang,nhomtb,diadiem,donvi,taisan,nguyengia';
const H = {
  equip: 'id,ten,sl,nhom,tt,sn,pn,anh,ghichu,capnhat,nguoi,tu,' + MAU,
  spare: 'id,ten,sl,nhom,tt,sn,pn,anh,anhtem,cauhinh,capnhat,nguoi,tu,ghichu,' + MAU,
  fail: 'id,equipId,ngay,mota,nguyennhan,khacphuc,nguoi,ngaykp,anhhuong',
  log: 'id,equipId,ngay,loai,noidung,thuchien,ketqua,tiep,nguoi,capnhat,xacnhan,xnnguoi',   // nhật ký bảo trì, bảo dưỡng của từng thiết bị
  cfg: 'id,spareId,ngay,nguoi,noidung',
  swap: 'id,ngay,nguoi,equipId,spareId,sl,ghichu,thoigian,nguoitt,ttsau',
  users: 'user,pass,role,ten'
};
const today = () => Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd');
const nid = () => Utilities.getUuid().slice(0, 8);
const SS = () => SpreadsheetApp.getActive();

// Đọc dòng tiêu đề thật của sheet; tự thêm các cột còn thiếu (định dạng văn bản để giữ nguyên số 0 đầu của Serial).
function hdr(k) {
  const s = SS().getSheetByName(k) || SS().insertSheet(k);
  const h = s.getLastColumn() ? s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0].map(String) : [];
  H[k].split(',').forEach(c => {
    if (h.indexOf(c) >= 0) return;
    h.push(c);
    const col = h.length;
    if (col > s.getMaxColumns()) s.insertColumnsAfter(s.getMaxColumns(), col - s.getMaxColumns());
    s.getRange(1, col, s.getMaxRows(), 1).setNumberFormat('@');
    s.getRange(1, col).setValue(c).setFontWeight('bold');
  });
  return h;
}
function setup() {
  for (const k in H) {
    const s = SS().getSheetByName(k) || SS().insertSheet(k);
    if (!s.getLastRow()) s.getRange(1, 1, s.getMaxRows(), s.getMaxColumns()).setNumberFormat('@');
    hdr(k);
  }
  const u = SS().getSheetByName('users');
  if (u.getLastRow() < 2) u.appendRow(['admin', '123456', 'quantri', 'Quản trị']); // role: quantri | xem
}
function rd(k) {
  if (!SS().getSheetByName(k)) hdr(k);   // sheet mới (VD log) tự được tạo
  const v = SS().getSheetByName(k).getDataRange().getValues(), h = v.shift();
  return v.filter(r => r[0] !== '').map(r => Object.fromEntries(h.map((c, i) =>
    [c, r[i] instanceof Date ? Utilities.formatDate(r[i], 'GMT+7', 'yyyy-MM-dd') : r[i]])));
}
function wr(k, o) {
  const s = SS().getSheetByName(k), h = hdr(k), n = s.getLastRow();
  const ids = n > 1 ? s.getRange(2, 1, n - 1, 1).getValues().flat().map(String) : [];
  const i = ids.indexOf(String(o.id)), a = h.map(c => o[c] ?? '');
  i < 0 ? s.appendRow(a) : s.getRange(i + 2, 1, 1, h.length).setValues([a]);
}
function rm(k, id) {
  const s = SS().getSheetByName(k), n = s.getLastRow();
  const ids = n > 1 ? s.getRange(2, 1, n - 1, 1).getValues().flat().map(String) : [];
  const i = ids.indexOf(String(id));
  if (i >= 0) s.deleteRow(i + 2);
}
const all = () => ({ equip: rd('equip'), spare: rd('spare'), fail: rd('fail'), log: rd('log'), cfg: rd('cfg'), swap: rd('swap') });
const soLuong = v => parseFloat(String(v ?? '').replace(',', '.')) || 0;   // "2 bộ" -> 2

function doPost(e) {
  let r;
  try {
    const q = JSON.parse(e.postData.contents), c = CacheService.getScriptCache();
    if (q.a === 'login') {
      const u = rd('users').find(x => x.user == q.u && String(x.pass) == String(q.p));
      if (!u) throw Error('Sai tài khoản hoặc mật khẩu');
      const tk = Utilities.getUuid(), us = { u: u.user, role: u.role, ten: u.ten };
      c.put(tk, JSON.stringify(us), 21600);
      r = { tk, us, d: all() };
    } else {
      const j = c.get(q.tk);
      if (!j) throw Error('Hết phiên, vui lòng đăng nhập lại');
      const us = JSON.parse(j);
      if (q.a !== 'all' && us.role !== 'quantri') throw Error('Bạn chỉ có quyền xem');
      if (q.a === 'save' || q.a === 'del') if (!H[q.t] || q.t === 'users') throw Error('Bảng không hợp lệ');
      if (q.a === 'save') {
        const o = q.row; o.id = o.id || nid(); o.capnhat = today(); if (!(q.t === 'fail' && o.nguoi)) o.nguoi = us.ten;
        if (q.t === 'fail') o.ngaykp = String(o.khacphuc || '').trim() ? (o.ngaykp || today()) : '';
        if (q.t === 'spare' && o.cauhinh) {
          const old = rd('spare').find(x => x.id == o.id);
          if (!old || old.cauhinh != o.cauhinh)
            wr('cfg', { id: nid(), spareId: o.id, ngay: today(), nguoi: us.ten, noidung: o.cauhinh });
        }
        if (q.t === 'log') xuLyNhatKy(o, us);
        wr(q.t, o);
        if (q.t === 'fail' || q.t === 'log') dongBoSuCo(o.equipId, us);
        r = all();
      } else if (q.a === 'del') {
        const f = (q.t === 'fail' || q.t === 'log') ? rd(q.t).find(x => x.id == q.id) : null;
        rm(q.t, q.id); if (f) dongBoSuCo(f.equipId, us); r = all();
      }
      else if (q.a === 'swap') {
        const o = q.row, sl = Number(o.sl), sp = rd('spare').find(x => x.id == o.spareId), co = sp ? soLuong(sp.sl) : 0;
        if (!sp || !(sl >= 1) || sl > co) throw Error('Không đủ số lượng vật tư dự phòng');
        const dv = String(sp.sl).replace(/^[\s\d.,]+/, '').trim();   // giữ nguyên đơn vị tính nếu có
        sp.sl = (co - sl) + (dv ? ' ' + dv : ''); sp.capnhat = today(); sp.nguoi = us.ten; wr('spare', sp);
        wr('swap', { id: nid(), ngay: today(), nguoi: us.ten, equipId: o.equipId, spareId: o.spareId, sl, ghichu: o.ghichu,
          thoigian: o.thoigian, nguoitt: o.nguoitt, ttsau: o.ttsau });
        if (['Tốt', 'Thiết bị gặp sự cố', 'Đang sửa chữa'].indexOf(o.ttsau) >= 0) capNhatSauThay(o.equipId, o.ttsau, sp.ten, us);
        r = all();
      } else if (q.a === 'up') {
        const f = DriveApp.getFolderById(FOLDER_ID).createFile(
          Utilities.newBlob(Utilities.base64Decode(q.b64), 'image/jpeg', Date.now() + '.jpg'));
        f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        r = { id: f.getId() };
      } else r = all();
    }
  } catch (err) { r = { err: err.message }; }
  return ContentService.createTextOutput(JSON.stringify(r)).setMimeType(ContentService.MimeType.JSON);
}

// Chạy 1 LẦN sau khi dán code mới (chạy lại nhiều lần vẫn an toàn): thêm các cột của mẫu "Phân hệ anten" vào
// sheet equip và spare, rồi điền giá trị mặc định cho các dòng cũ còn trống:
//  - Đơn vị quản lý = "Trung tâm KSKL Long Thành"
//  - Tình trạng vật tư = "Đang sử dụng" (khối thiết bị) hoặc "Dự phòng" (vật tư dự phòng)
function chuyenMau() {
  const DV = 'Trung tâm KSKL Long Thành';
  ['equip', 'spare'].forEach(k => {
    const s = SS().getSheetByName(k), h = hdr(k), n = s.getLastRow();
    if (n < 2) return;
    const dien = (c, fn) => {
      const r = s.getRange(2, h.indexOf(c) + 1, n - 1, 1), v = r.getValues();
      r.setValues(v.map((x, i) => [x[0] === '' ? fn() : x[0]]));
    };
    dien('donvi', () => DV);
    dien('tinhtrang', () => (k === 'spare' ? 'Dự phòng' : 'Đang sử dụng'));
  });
  Logger.log('Đã chuyển sang mẫu mới.');
}

// Chạy 1 LẦN khi đổi tên một nhóm hệ thống: đổi giá trị cột nhom trong các sheet equip và spare.
// Sửa 2 dòng dưới rồi chạy. Mặc định: "Hệ thống PSR/MSSR Easat EA29190" -> "Hệ thống PSR/MSSR".
function doiTenNhom() {
  const CU = 'Hệ thống PSR/MSSR Easat EA29190', MOI = 'Hệ thống PSR/MSSR';
  let n = 0;
  ['equip', 'spare'].forEach(k => {
    const s = SS().getSheetByName(k), h = hdr(k), r = s.getLastRow();
    if (r < 2) return;
    const g = s.getRange(2, h.indexOf('nhom') + 1, r - 1, 1), v = g.getValues();
    g.setValues(v.map(x => { if (String(x[0]).trim() === CU) { n++; return [MOI]; } return x; }));
  });
  Logger.log('Đã đổi tên nhóm hệ thống cho ' + n + ' dòng.');
}

// Chạy 1 LẦN: đổi số thứ tự dạng "5,1" thành "5.1" trong sheet equip và spare.
function chuanHoaSTT() {
  let n = 0;
  ['equip', 'spare'].forEach(k => {
    const s = SS().getSheetByName(k), h = hdr(k), r = s.getLastRow();
    if (r < 2) return;
    const g = s.getRange(2, h.indexOf('stt') + 1, r - 1, 1), v = g.getDisplayValues();
    g.setNumberFormat('@').setValues(v.map(x => { const t = String(x[0]).replace(/,/g, '.'); if (t !== x[0]) n++; return [t]; }));
  });
  Logger.log('Đã đổi dạng số thứ tự cho ' + n + ' dòng.');
}

// Chạy 1 LẦN: đổi tình trạng hoạt động "Hư hỏng" / "Thiết bị đang sự cố" thành "Thiết bị gặp sự cố" trong sheet equip và spare.
function doiTinhTrangSuCo() {
  let n = 0;
  ['equip', 'spare'].forEach(k => {
    const s = SS().getSheetByName(k), h = hdr(k), r = s.getLastRow();
    if (r < 2) return;
    const g = s.getRange(2, h.indexOf('tt') + 1, r - 1, 1), v = g.getValues();
    g.setValues(v.map(x => { if (['Hư hỏng', 'Thiết bị đang sự cố'].indexOf(String(x[0]).trim()) >= 0) { n++; return ['Thiết bị gặp sự cố']; } return x; }));
  });
  Logger.log('Đã đổi tình trạng cho ' + n + ' dòng.');
}

// Đồng bộ tình trạng hoạt động của thiết bị theo nhật ký (ưu tiên từ trên xuống):
//  - có lần "Sửa chữa" CHƯA xác nhận khắc phục       -> "Đang sửa chữa"
//  - có sự cố CHƯA có "Cách khắc phục"               -> "Thiết bị gặp sự cố"
//  - không còn gì đang mở                             -> trở lại "Tốt" (nếu đang là sự cố/hư hỏng/đang sửa chữa)
function dongBoSuCo(eid, us) {
  let sh = 'equip', e = rd('equip').find(x => x.id == eid);
  if (!e) { sh = 'spare'; e = rd('spare').find(x => x.id == eid); }   // vật tư dự phòng cũng có nhật ký sự cố
  if (!e) return;
  const DANG = 'Thiết bị gặp sự cố';
  const mo = rd('fail').filter(f => f.equipId == eid && !String(f.khacphuc || '').trim()).length;
  const sua = rd('log').filter(l => l.equipId == eid && l.loai === 'Sửa chữa' && !String(l.xacnhan || '').trim()).length;
  let tt = e.tt;
  if (sua > 0) tt = 'Đang sửa chữa';      // đang có lần sửa chữa chưa xác nhận khắc phục
  else if (mo > 0) tt = DANG;
  else if ([DANG, 'Thiết bị đang sự cố', 'Hư hỏng', 'Đang sửa chữa', 'Cảnh báo'].indexOf(e.tt) >= 0) tt = 'Tốt';
  if (tt !== e.tt) { e.tt = tt; e.capnhat = today(); e.nguoi = us.ten; wr(sh, e); }
}

// Xử lý nhật ký loại "Sửa chữa": tick xác nhận khắc phục -> ghi ngày/người xác nhận và đóng các sự cố đang mở của thiết bị.
function xuLyNhatKy(o, us) {
  if (o.loai === 'Sửa chữa') {
    if (o.xong) { o.xacnhan = o.xacnhan || today(); o.xnnguoi = o.xnnguoi || us.ten; }
  } else { o.xacnhan = ''; o.xnnguoi = ''; }
  delete o.xong;
  if (o.loai !== 'Bảo dưỡng định kỳ') o.tiep = '';   // chỉ bảo dưỡng định kỳ mới có ngày bảo dưỡng tiếp theo
  if (o.xacnhan)
    rd('fail').filter(f => f.equipId == o.equipId && !String(f.khacphuc || '').trim())
      .forEach(f => { f.khacphuc = 'Đã sửa chữa: ' + (o.noidung || ''); f.ngaykp = today(); wr('fail', f); });
}

// Sau khi thay thế: đặt "Tình trạng hoạt động" của thiết bị theo lựa chọn của người dùng.
// Nếu chọn "Tốt": các sự cố đang mở của thiết bị tự đóng, các lần sửa chữa chưa xác nhận tự được xác nhận.
function capNhatSauThay(eid, ttsau, tenVT, us) {
  let sh = 'equip', e = rd('equip').find(x => x.id == eid);
  if (!e) { sh = 'spare'; e = rd('spare').find(x => x.id == eid); }
  if (!e) return;
  if (ttsau === 'Tốt') {
    rd('fail').filter(f => f.equipId == eid && !String(f.khacphuc || '').trim())
      .forEach(f => { f.khacphuc = 'Đã thay thế bằng vật tư dự phòng: ' + tenVT; f.ngaykp = today(); wr('fail', f); });
    rd('log').filter(l => l.equipId == eid && l.loai === 'Sửa chữa' && !String(l.xacnhan || '').trim())
      .forEach(l => { l.xacnhan = today(); l.xnnguoi = us.ten; wr('log', l); });
  }
  if (e.tt !== ttsau) { e.tt = ttsau; e.capnhat = today(); e.nguoi = us.ten; wr(sh, e); }
}
