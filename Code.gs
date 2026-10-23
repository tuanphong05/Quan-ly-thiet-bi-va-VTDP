// === QUẢN LÝ THIẾT BỊ & VẬT TƯ DỰ PHÒNG - LTH ===
// Dán vào Apps Script của Google Sheets, chạy setup() 1 lần, rồi Deploy > Web app
// (Execute as: Me, Who has access: Anyone).
const FOLDER_ID = 'DAN_ID_THU_MUC_DRIVE_LUU_ANH';
const H = {
  equip: 'id,ten,sl,nhom,tt,sn,pn,anh,ghichu,capnhat,nguoi,tu',
  spare: 'id,ten,sl,nhom,tt,sn,pn,anh,anhtem,cauhinh,capnhat,nguoi,tu,ghichu',
  fail: 'id,equipId,ngay,mota,nguyennhan,khacphuc,nguoi',
  cfg: 'id,spareId,ngay,nguoi,noidung',
  swap: 'id,ngay,nguoi,equipId,spareId,sl,ghichu',
  users: 'user,pass,role,ten'
};
const today = () => Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd');
const nid = () => Utilities.getUuid().slice(0, 8);
const SS = () => SpreadsheetApp.getActive();

function setup() {
  for (const k in H) {
    const s = SS().getSheetByName(k) || SS().insertSheet(k);
    if (s.getLastRow()) continue;
    const h = H[k].split(',');
    s.getRange('A:Z').setNumberFormat('@');
    s.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold');
  }
  const u = SS().getSheetByName('users');
  if (u.getLastRow() < 2) u.appendRow(['admin', '123456', 'quantri', 'Quản trị']); // role: quantri | xem
}
function rd(k) {
  const v = SS().getSheetByName(k).getDataRange().getValues(), h = v.shift();
  return v.filter(r => r[0] !== '').map(r => Object.fromEntries(h.map((c, i) =>
    [c, r[i] instanceof Date ? Utilities.formatDate(r[i], 'GMT+7', 'yyyy-MM-dd') : r[i]])));
}
function wr(k, o) {
  const s = SS().getSheetByName(k), h = H[k].split(','), n = s.getLastRow();
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
const all = () => ({ equip: rd('equip'), spare: rd('spare'), fail: rd('fail'), cfg: rd('cfg'), swap: rd('swap') });

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
        const o = q.row; o.id = o.id || nid(); o.capnhat = today(); o.nguoi = us.ten;
        if (q.t === 'spare' && o.cauhinh) {
          const old = rd('spare').find(x => x.id == o.id);
          if (!old || old.cauhinh != o.cauhinh)
            wr('cfg', { id: nid(), spareId: o.id, ngay: today(), nguoi: us.ten, noidung: o.cauhinh });
        }
        wr(q.t, o); r = all();
      } else if (q.a === 'del') { rm(q.t, q.id); r = all(); }
      else if (q.a === 'swap') {
        const o = q.row, sl = Number(o.sl), sp = rd('spare').find(x => x.id == o.spareId);
        if (!sp || sl < 1 || sl > Number(sp.sl)) throw Error('Không đủ số lượng vật tư dự phòng');
        sp.sl = Number(sp.sl) - sl; sp.capnhat = today(); sp.nguoi = us.ten; wr('spare', sp);
        wr('swap', { id: nid(), ngay: today(), nguoi: us.ten, equipId: o.equipId, spareId: o.spareId, sl, ghichu: o.ghichu });
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

// Chạy 1 LẦN sau khi dán code mới (chạy lại nhiều lần vẫn an toàn): thêm cột còn thiếu
// ("tu" cho equip và spare, "ghichu" cho spare) và gán các dòng cũ của hệ thống
// PSR/MSSR Easat EA29190 vào "Tủ ACU" nếu chưa có tủ.
function nangCap() {
  const HT = 'Hệ thống PSR/MSSR Easat EA29190', TU_CU = 'Tủ ACU';
  const can = { equip: ['tu'], spare: ['tu', 'ghichu'] };
  for (const k in can) {
    const s = SS().getSheetByName(k), h = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0];
    can[k].forEach(c => { if (h.indexOf(c) < 0) { s.getRange(1, h.length + 1).setValue(c).setFontWeight('bold'); h.push(c); } });
    const n = s.getLastRow(); if (n < 2) continue;
    const cN = h.indexOf('nhom') + 1, cT = h.indexOf('tu') + 1;
    const nh = s.getRange(2, cN, n - 1, 1).getValues(), r = s.getRange(2, cT, n - 1, 1), tu = r.getValues();
    r.setNumberFormat('@').setValues(tu.map((x, i) => [x[0] || (nh[i][0] === HT ? TU_CU : '')]));
  }
  Logger.log('Đã nâng cấp xong.');
}
