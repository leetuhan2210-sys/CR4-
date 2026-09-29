/**
*  =====================================================================
*  AEF26 - THÊM KM MỚI HÀNG LOẠT TỪ WEB APP  (bản 4: thêm Sub-sector)
*  File riêng, nằm CÙNG project Apps Script với WA_Code.gs, MAIL DATA.gs, mã KM.gs, gui mail.gs.
*  KHÔNG cần sửa gì trong các file cũ.
*
*  Form nhập đủ TOÀN BỘ cột của sheet TEST, trừ cột "MÃ KM" (luôn tự sinh).
*  Dò cột theo TIÊU ĐỀ (không dùng số cột cố định) nên chèn/xoá cột trên sheet không làm hỏng.
*
*  Dùng lại (nếu có trong project, không có thì bỏ qua bước kiểm tra tương ứng):
*    WA_Code.gs  : WA_CONFIG, WA_getCombinedHeaderLabels_, WA_buildHeaderMap_, WA_firstCol_,
*                  WA_findMergedGroupRange_, WA_findColInRange_, WA_timCotTheoNhan_,
*                  WA_getPicGroupsFromSetupSheet_, WA_findPicOfPerson_,
*                  WA_getHandlerOptionsFromSetupSheet_, WA_getOrCreateEditLog_, WA_colLetter_
*    MAIL DATA.gs: TB_chayCoKhoa_
*    mã KM.gs    : MK_PREFIX_MAP, mkLayDanhSachMaHienCo, mkSinhMaKeTiepAnToan, MK_COL_*
*    gui mail.gs : D_QUOC_GIA_TQ, D_ghepChucVuToChucTQ_, D_COL, D_LOAI_THU
*
*  LĨNH VỰC + SUB-SECTOR (bản 4):
*    - Danh sách chuẩn đọc từ sheet THIẾT LẬP FILE, 2 cột có tiêu đề dòng 1 là "Lĩnh vực" và "Sub-sector"
*      (dạng khối: tên Lĩnh vực ở dòng đầu khối, các Sub-sector xếp bên dưới tới khi gặp Lĩnh vực kế tiếp).
*    - Lưu vào 2 CỘT RIÊNG của TEST: "Lĩnh vực" (Q) và "Sub-sector" (cột mới, dò theo tiêu đề nên đặt ở đâu cũng được).
*    - Sub-sector phải thuộc đúng Lĩnh vực đã chọn; chọn Lĩnh vực có Sub-sector thì bắt buộc chọn Sub-sector.
*    - Gõ/dán dạng gộp "1. Nông nghiệp, Lâm nghiệp & Thủy sản - Trồng trọt" hoặc chỉ gõ "Trồng trọt"
*      vào ô Lĩnh vực: tự tách đúng vào 2 cột.
*    - Danh sách được nhớ tạm 5 phút. Vừa sửa THIẾT LẬP FILE mà muốn áp dụng ngay: chạy NM_xoaCacheLinhVuc().
*  =====================================================================
*/

var NM_MAX_ROWS = 200;       // tối đa số dòng mỗi lần bấm Lưu
var NM_CACHE_GIAY = 21600;   // 6 giờ: nhớ các dòng đã lưu để bấm Lưu lại không ghi trùng

/**
 * Khai báo trường nhập. Thứ tự hiển thị trên giao diện = thứ tự cột thật trên sheet
 * (sắp theo số cột dò được), không phụ thuộc thứ tự khai báo ở đây.
 *
 * Cách dò cột:
 *   "header"          -> dò thẳng theo tiêu đề. Dùng cho cột đứng riêng.
 *   "zone" + "sub"    -> dò TRONG nhóm gộp ở dòng 4.
 *   "zone" + "smart"  -> dò trong nhóm gộp theo nhiều mức (cột "Tiến độ" của 2 đợt có nhãn khác nhau).
 *   BẮT BUỘC dùng zone cho "Phụ trách"/"Tiến độ" của 2 đợt vì trùng chữ với cột "Phụ trách" chính.
 *
 * required: bắt buộc (luôn hiện, không ẩn được)   common: có ô "Giá trị chung cho cả đợt"
 * group   : core luôn hiện; internal/contact/dot1/dot2/gmail bật tắt bằng chip
 * suggest : gợi ý từ giá trị đang có trên sheet, tự sửa về đúng cách viết đang dùng
 * warnNew : cảnh báo khi gõ giá trị chưa từng có trên sheet (dễ là gõ sai chính tả)
 * multi   : ô nhiều giá trị, ngăn cách bằng dấu phẩy
 * log     : ghi theo định dạng nhật ký "dd.MM - nội dung"
 * email   : 'primary' | 'cc' | 'other'
 * raw     : ghi nguyên giá trị, không ép kiểu văn bản (cột ngày/giờ)
 */
function NM_fieldDefs_() {
  var C = WA_CONFIG;
  var LOC_TIENDO = {
    chuaTu: ['tien do'],
    khongChuaTu: ['btc', 'lam viec', 'phan hoi'],
    boQua: ['Phụ trách', 'Nhân sự hỗ trợ', 'Phản hồi của Khách', 'Tiến độ làm việc của BTC']
  };

  return [
    // ---------- Bắt buộc + hồ sơ chính (luôn hiện) ----------
    { key: 'pic',       header: C.COL_PHUTRACH,     label: 'Phụ trách',          group: 'core', required: true, common: true },
    { key: 'hoTro',     header: C.COL_NHANSU_HOTRO, label: 'Nhân sự hỗ trợ',     group: 'core', required: true, common: true, hint: 'Phải thuộc nhóm của Phụ trách (sheet THIẾT LẬP FILE)' },
    { key: 'doiTuong',  header: C.COL_DOITUONG,     label: 'Đối tượng',          group: 'core', required: true, common: true, hint: 'Quyết định tiền tố Mã KM. KM phụ tự lấy theo KM chính.' },
    { key: 'danhXung',  header: C.COL_DANHXUNG,     label: 'Danh xưng',          group: 'core' },
    { key: 'title',     header: C.COL_TITLE,        label: 'Title',              group: 'core' },
    { key: 'hoTen',     header: C.COL_HOTEN,        label: 'Họ và tên',          group: 'core', required: true },
    { key: 'fullName',  header: C.COL_FULLNAME,     label: 'Full name',          group: 'core', required: true },
    { key: 'chucVu',    header: C.COL_CHUCVU,       label: 'Chức vụ',            group: 'core', required: true, hint: 'Nhiều chức vụ tách bằng ";" và phải khớp số đoạn với Tên đơn vị' },
    { key: 'position',  header: C.COL_POSITION,     label: 'Position',           group: 'core', required: true, hint: 'Nhiều chức vụ tách bằng ";" và phải khớp số đoạn với Organizagtion' },
    { key: 'donVi',     header: C.COL_DONVI,        label: 'Tên đơn vị',         group: 'core', required: true },
    { key: 'org',       header: C.COL_ORGANIZATION, label: 'Organizagtion',      group: 'core', required: true },
    { key: 'linhVuc',   header: C.COL_LINHVUC,      label: 'Lĩnh vực',           group: 'core', common: true },
    { key: 'subSector', headers: [C.COL_SUBSECTOR || 'Sub-sector', 'Sub-sector', 'Sub sector', 'Subsector', 'Lĩnh vực con', 'Lĩnh vực phụ'],
                                                    label: 'Sub-sector',         group: 'core', common: true, afterKey: 'linhVuc',
                                                    hint: 'Chọn sau Lĩnh vực, chỉ hiện Sub-sector thuộc Lĩnh vực đó' },
    { key: 'quocGia',   header: C.COL_QUOCGIA,      label: 'Quốc gia',           group: 'core', required: true },
    { key: 'email',     header: C.COL_EMAIL,        label: 'Thư điện tử chính',  group: 'core', email: 'primary' },
    { key: 'emailCc',   header: C.COL_EMAIL_CC,     label: 'Thư điện tử CC phụ', group: 'core', email: 'cc', hint: 'Nhiều email cách nhau bằng dấu phẩy' },
    { key: 'ghiChu',    header: C.COL_GHICHU,       label: 'Ghi chú',            group: 'core', log: true },

    // ---------- Thông tin nội bộ (tuỳ chọn) ----------
    { key: 'kmChinh',    header: C.COL_KM_CHINH,     label: 'Thuộc nhóm KM chính của (nếu có)', group: 'internal', ref: true, hint: 'Chỉ điền nếu là KM phụ: gõ Full name hoặc Mã KM của KM chính' },
    { key: 'kenhMoi',    header: C.COL_KENH_MOI,     label: 'Kênh mời',             group: 'internal', common: true },
    { key: 'nguoiMoi',   header: C.COL_NGUOI_MOI,    label: 'Người mời trực tiếp',  group: 'internal', common: true },
    { key: 'loaiThuMoi', header: C.COL_LOAI_THU_MOI, label: 'Loại thư mời',         group: 'internal', common: true },
    { key: 'daiTho',     header: C.COL_DAITHO,       label: 'Đài thọ',              group: 'internal', common: true },
    { key: 'phanLoai',   header: C.COL_PHANLOAI,     label: 'Phân loại',            group: 'internal', common: true },

    // ---------- Liên hệ phụ / đại diện ----------
    { key: 'lienLacKhac', header: C.COL_LIENLAC_KHAC, label: 'Thông tin liên lạc khác', group: 'contact' },
    { key: 'sdt',         header: C.COL_SDT,          label: 'Số điện thoại',         group: 'contact', phone: true, hint: 'Định dạng (+mã quốc gia) số. Bấm vào ô để nhập nhiều số.' },    
    { key: 'diaChi',      header: C.COL_DIACHI,       label: 'Địa chỉ',               group: 'contact' },
    { key: 'emailThay',   header: C.COL_EMAIL_THAY,   label: 'Email liên hệ thay',    group: 'contact', email: 'other' },
    { key: 'tenThay',     header: C.COL_TEN_THAY,     label: 'Tên người liên hệ thay', group: 'contact' },
    { key: 'quanHe',      header: C.COL_QUANHE,       label: 'Quan hệ',               group: 'contact' },

    // ---------- Đợt 1 ----------
    { key: 'xacNhan1',  header: C.COL_XACNHAN_DOT1,                  label: 'Phản hồi tham dự Save the day',  group: 'dot1' },
    { key: 'd1Pic',     zone: 'dot1', sub: C.DOT1_COL_PHUTRACH,       label: 'Đợt 1 · Phụ trách',              group: 'dot1', common: true },
    { key: 'd1TienDo',  zone: 'dot1', smart: { candidates: [C.DOT1_COL_TIENDO, 'Tiến độ', 'Tiến độ liên hệ', 'Trạng thái'], chuaTu: LOC_TIENDO.chuaTu, khongChuaTu: LOC_TIENDO.khongChuaTu, boQua: LOC_TIENDO.boQua },
                                                                      label: 'Đợt 1 · Tiến độ liên hệ',        group: 'dot1', common: true },
    { key: 'd1PhanHoi', zone: 'dot1', sub: C.DOT1_COL_PHANHOI_KHACH, label: 'Đợt 1 · Phản hồi của Khách (1)', group: 'dot1', log: true },
    { key: 'd1Btc',     zone: 'dot1', sub: C.DOT1_COL_TIENDO_BTC,    label: 'Đợt 1 · Tiến độ làm việc của BTC (1)', group: 'dot1', log: true },

    // ---------- Đợt 2 ----------
    { key: 'xacNhan2',  header: C.COL_XACNHAN_DOT2,                  label: 'Phản hồi tham dự chính thức',    group: 'dot2' },
    { key: 'd2Pic',     zone: 'dot2', sub: C.DOT2_COL_PHUTRACH,       label: 'Đợt 2 · Phụ trách',              group: 'dot2', common: true },
    { key: 'd2TienDo',  zone: 'dot2', smart: { candidates: [C.DOT2_COL_TIENDO, 'Tiến độ liên hệ', 'Tiến độ', 'Trạng thái'], chuaTu: LOC_TIENDO.chuaTu, khongChuaTu: LOC_TIENDO.khongChuaTu, boQua: LOC_TIENDO.boQua },
                                                                      label: 'Đợt 2 · Tiến độ',                group: 'dot2', common: true },
    { key: 'd2PhanHoi', zone: 'dot2', sub: C.DOT2_COL_PHANHOI_KHACH, label: 'Đợt 2 · Phản hồi của Khách (2)', group: 'dot2', log: true },
    { key: 'd2Btc',     zone: 'dot2', sub: C.DOT2_COL_TIENDO_BTC,    label: 'Đợt 2 · Tiến độ làm việc của BTC (2)', group: 'dot2', log: true },

    // ---------- Gmail & phản hồi gần nhất (bình thường để trống, hệ thống tự ghi) ----------
    { key: 'maGmail', header: C.COL_GMAIL_THREAD,               label: 'Mã Gmail',   group: 'gmail', hint: 'Thường để trống. Script gửi mail tự ghi Thread ID vào đây.' },
    { key: 'ngayPh',  zone: 'latest', sub: C.COL_NGAY_PHANHOI,  label: 'Ngày (Day)', group: 'gmail', raw: true, hint: 'Thường để trống, tự cập nhật khi quét Gmail.' },
    { key: 'gioPh',   zone: 'latest', sub: C.COL_THANG_PHANHOI, label: 'Giờ (Time)', group: 'gmail', raw: true, hint: 'Thường để trống, tự cập nhật khi quét Gmail.' }
  ];
}

// ================== HELPER ==================
// Chuẩn hoá để SO KHỚP: bỏ dấu tiếng Việt, bỏ emoji/ký tự đặc biệt, hạ chữ thường.
function NM_norm_(s) {
  return (s === null || s === undefined ? '' : s).toString().toLowerCase()
    .normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/đ/g, 'd')
    .replace(/[^a-z0-9]+/g, ' ').trim();
}

// Làm sạch giá trị TRƯỚC KHI GHI: dữ liệu dán từ Word/web/email thường lẫn khoảng trắng không ngắt
// (U+00A0) và ký tự vô hình (U+200B...). Chúng làm hỏng mọi phép so khớp chính xác phía sau
// (TextFinder của mã KM.gs, dò email của MAIL DATA.gs) mà mắt thường không thấy.
function NM_sach_(v) {
  if (v === null || v === undefined) return '';
  return v.toString()
    .replace(/[​-‍⁠﻿]/g, '')
    .replace(/[   ]/g, ' ')
    .replace(/\r\n?/g, '\n')
    .split('\n').map(function (l) { return l.replace(/[ \t]+/g, ' ').trim(); }).join('\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim()
    .normalize('NFC');
}

// Khoá so trùng tên: bỏ danh xưng/học hàm, không phân biệt thứ tự từ.
// "Dr. Klaus Schwab" = "Schwab Klaus" = "klaus  SCHWAB".
var NM_TU_BO_QUA_TEN = ['mr', 'mrs', 'ms', 'miss', 'dr', 'prof', 'assoc', 'he', 'phd', 'sir',
  'ong', 'ba', 'ts', 'pgs', 'gs', 'ths', 'ngai', 'anh', 'chi'];
function NM_khoaTen_(s) {
  return NM_norm_(s).split(' ')
    .filter(function (t) { return t && t.length > 1 && NM_TU_BO_QUA_TEN.indexOf(t) === -1; })
    .sort().join(' ');
}

// Tìm lựa chọn khớp: đúng nguyên văn → đúng sau chuẩn hoá → chứa chuỗi (chỉ khi duy nhất 1 khớp).
function NM_khopLuaChon_(value, options) {
  var v = (value || '').toString().trim();
  if (!v || !options || !options.length) return v;
  if (options.indexOf(v) !== -1) return v;
  var n = NM_norm_(v);
  if (!n) return null;
  var exact = null, partial = [];
  options.forEach(function (o) {
    var on = NM_norm_(o);
    if (on === n) { if (!exact) exact = o; }
    else if (on.indexOf(n) !== -1) partial.push(o);
  });
  if (exact) return exact;
  return partial.length === 1 ? partial[0] : null;
}

var NM_EMAIL_RE = /^[^\s@,;<>]+@[^\s@,;<>]+\.[^\s@,;<>]{2,}$/;

// Tách email từ chuỗi dán vào, chấp nhận cả dạng "Tên <a@b.com>", "mailto:a@b.com", "a@b.com; c@d.com".
// Hạ chữ thường (MAIL DATA.gs so khớp email ở dạng chữ thường) và bỏ trùng.
function NM_layEmail_(s) {
  var t = NM_sach_(s).replace(/mailto:/gi, ' ');
  var out = [];
  t.split(/[\s,;<>()"'\[\]]+/).forEach(function (x) {
    x = x.replace(/^[.:]+|[.:]+$/g, '').toLowerCase();
    if (x && x.indexOf('@') !== -1 && out.indexOf(x) === -1) out.push(x);
  });
  return out;
}

function NM_tachDoan_(s) {
  return String(s || '').split(/[;；]/).map(function (x) { return x.trim(); }).filter(Boolean);
}

// Tách ô nhiều lựa chọn "A, B, C" nhưng KHÔNG cắt đôi tên có dấu phẩy
// (VD "1. Nông nghiệp, Lâm nghiệp & Thủy sản"): ghép các mảnh liền nhau cho tới khi khớp 1 tên trong options.
// Mảnh không khớp tên nào được gộp lại thành 1 mục tới trước mục khớp kế tiếp.
function NM_tachDaGiaTri_(v, options) {
  var p = NM_sach_(v).split(',').map(function (x) { return x.trim(); }).filter(Boolean);
  var biet = {};
  (options || []).forEach(function (o) { var k = NM_norm_(o); if (k) biet[k] = true; });
  if (!Object.keys(biet).length) return p;
  var dai = function (i) {
    for (var j = p.length; j > i; j--) if (biet[NM_norm_(p.slice(i, j).join(', '))]) return j;
    return -1;
  };
  var out = [], i = 0;
  while (i < p.length) {
    var j = dai(i);
    if (j !== -1) { out.push(p.slice(i, j).join(', ')); i = j; continue; }
    var k = i + 1;
    while (k < p.length && dai(k) === -1) k++;
    out.push(p.slice(i, k).join(', '));
    i = k;
  }
  return out;
}
function NM_gop_(arr) {
  var seen = {}, out = [];
  (arr || []).forEach(function (x) { var k = NM_norm_(x); if (x && k && !seen[k]) { seen[k] = true; out.push(x); } });
  return out.join(', ');
}

function NM_giaTriRong_(v) {
  return v === '' || v === null || v === undefined || v === false || !v.toString().trim();
}

// Khớp Đối tượng với tiền tố Mã KM theo PHẦN CHỮ, bỏ qua số thứ tự đầu dòng, vì danh sách
// Đối tượng trên sheet được đánh số lại theo thời gian ("4. Tỷ phú" -> "5. Tỷ phú").
var NM_TIENTO_CACHE_ = null;
function NM_boSoThuTu_(s) {
  return String(s || '').replace(/^[\s]*\d+[\s]*[.)-][\s]*/, '').trim();
}
function NM_tienTo_(doiTuong) {
  if (typeof mkTienTo_ === 'function') return mkTienTo_(doiTuong);   // dùng chung với mã KM.gs nếu có
  if (typeof MK_PREFIX_MAP === 'undefined') return null;
  var raw = String(doiTuong || '').trim().normalize('NFC');
  if (!raw) return null;
  if (MK_PREFIX_MAP[raw]) return MK_PREFIX_MAP[raw];
  if (!NM_TIENTO_CACHE_) {
    NM_TIENTO_CACHE_ = {};
    Object.keys(MK_PREFIX_MAP).forEach(function (k) {
      var key = NM_norm_(NM_boSoThuTu_(k));
      if (key && !NM_TIENTO_CACHE_[key]) NM_TIENTO_CACHE_[key] = MK_PREFIX_MAP[k];
    });
  }
  return NM_TIENTO_CACHE_[NM_norm_(NM_boSoThuTu_(raw))] || null;
}

// Giá trị văn bản trông giống số / ngày / công thức sẽ bị Google Sheets tự đổi kiểu khi ghi:
//   "0903 123 456" -> 903123456 (mất số 0),  "+84 90..." -> số,  "1/2" -> ngày 1 tháng 2,
//   "=IMPORTRANGE(...)" -> chạy công thức.
// Thêm dấu ' ở đầu để Sheets giữ nguyên là văn bản (dấu ' không hiện trong ô).
function NM_giaTriGhi_(m, v) {
  if (typeof v !== 'string' || !v) return v;
  if (m.raw || m.kind === 'list' || m.kind === 'checkbox') return v;
  if (/^[=+\-@]/.test(v) || /^[\d\s.,()\/:+\-]+$/.test(v)) return "'" + v;
  return v;
}

// Dòng dữ liệu cuối THẬT của TEST: quét ngược các cột Mã KM / Full name / Họ và tên / Email.
// Không tin getLastRow() vì sheet có giá trị rác ở dòng rất xa (nguồn dropdown).
function NM_timDongCuoi_(sheet, map) {
  var start = WA_CONFIG.DATA_START_ROW;
  var last = sheet.getLastRow();
  if (last < start) return start - 1;
  var cols = [WA_CONFIG.COL_MAKM, WA_CONFIG.COL_FULLNAME, WA_CONFIG.COL_HOTEN, WA_CONFIG.COL_EMAIL]
    .map(function (h) { return WA_firstCol_(map, h); }).filter(Boolean);
  if (!cols.length) throw new Error('Không tìm thấy cột Mã KM / Full name / Họ và tên / Thư điện tử chính trong sheet TEST.');
  var minC = Math.min.apply(null, cols), maxC = Math.max.apply(null, cols);
  var values = sheet.getRange(start, minC, last - start + 1, maxC - minC + 1).getValues();
  for (var i = values.length - 1; i >= 0; i--) {
    for (var k = 0; k < cols.length; k++) {
      if (!NM_giaTriRong_(values[i][cols[k] - minC])) return start + i;
    }
  }
  return start - 1;
}

// ================== LĨNH VỰC / SUB-SECTOR ==================
var NM_SETUP_SHEET = 'THIẾT LẬP FILE';
var NM_SETUP_HDR_LINHVUC = ['Lĩnh vực'];
var NM_SETUP_HDR_SUB = ['Sub-sector', 'Sub sector', 'Subsector', 'Lĩnh vực con', 'Lĩnh vực phụ'];
var NM_SECTOR_CACHE_KEY = 'NM_SECTOR_MAP_v1';
var NM_SECTOR_CACHE_GIAY = 60;   // web gọi WA_getSectorMap luôn đọc mới; các chỗ khác nhớ 1 phút
var NM_BANG_LV_REQ_ = null;   // nhớ trong 1 lần gọi máy chủ

// Cột đầu tiên có tiêu đề khớp 1 trong các tên (so sau chuẩn hoá: bỏ dấu, bỏ ngoặc, không phân biệt hoa thường).
function NM_timCotNhieuTen_(labels, names) {
  var want = names.map(NM_norm_).filter(Boolean);
  for (var k = 0; k < want.length; k++) {
    for (var i = 0; i < labels.length; i++) {
      if (NM_norm_(labels[i]) === want[k]) return i + 1;
    }
  }
  return null;
}

// Đọc bảng Lĩnh vực → [Sub-sector] ở THIẾT LẬP FILE (2 cột, đọc 1 lần, nhớ tạm 5 phút).
// Trả về { ok, sectors:[...], subs:{ sector:[sub,...] }, allSubs:[...], subToSectors:{ normSub:[sector,...] }, note }
function NM_docBangLinhVuc_(boQuaCache) {
  if (NM_BANG_LV_REQ_ && !boQuaCache) return NM_BANG_LV_REQ_;
  var cache = null;
  try { cache = CacheService.getScriptCache(); } catch (e) {}
  if (cache && !boQuaCache) {
    try { var c = cache.get(NM_SECTOR_CACHE_KEY); if (c) { NM_BANG_LV_REQ_ = JSON.parse(c); return NM_BANG_LV_REQ_; } } catch (e) {}
  }

  var out = { ok: false, sectors: [], subs: {}, allSubs: [], subToSectors: {}, note: '' };
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NM_SETUP_SHEET);
  if (!sh) { out.note = 'Không thấy sheet "' + NM_SETUP_SHEET + '".'; return out; }
  var lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) { out.note = 'Sheet "' + NM_SETUP_SHEET + '" chưa có dữ liệu.'; return out; }

  var head = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var cSec = NM_timCotNhieuTen_(head, NM_SETUP_HDR_LINHVUC);
  var cSub = NM_timCotNhieuTen_(head, NM_SETUP_HDR_SUB);
  if (!cSec || !cSub) {
    out.note = 'Dòng 1 của "' + NM_SETUP_SHEET + '" chưa có đủ 2 tiêu đề "Lĩnh vực" và "Sub-sector".';
    return out;
  }
  var minC = Math.min(cSec, cSub), maxC = Math.max(cSec, cSub);
  var values = sh.getRange(2, minC, lastRow - 1, maxC - minC + 1).getValues();

  var cur = '', seenSub = {};
  values.forEach(function (row) {
    var sec = NM_sach_(row[cSec - minC]);
    var sub = NM_sach_(row[cSub - minC]);
    if (sec) {
      cur = sec;
      if (!out.subs[cur]) { out.subs[cur] = []; out.sectors.push(cur); }
    }
    if (!sub || !cur) return;
    if (out.subs[cur].indexOf(sub) === -1) out.subs[cur].push(sub);
    var k = NM_norm_(sub);
    (out.subToSectors[k] = out.subToSectors[k] || []);
    if (out.subToSectors[k].indexOf(cur) === -1) out.subToSectors[k].push(cur);
    if (!seenSub[sub]) { seenSub[sub] = true; out.allSubs.push(sub); }
  });
  out.ok = out.sectors.length > 0;
  if (!out.ok) out.note = 'Cột "Lĩnh vực" của "' + NM_SETUP_SHEET + '" đang trống.';

  NM_BANG_LV_REQ_ = out;
  if (cache && out.ok) { try { cache.put(NM_SECTOR_CACHE_KEY, JSON.stringify(out), NM_SECTOR_CACHE_GIAY); } catch (e) {} }
  return out;
}

// Chạy tay sau khi sửa danh sách Lĩnh vực / Sub-sector để web app dùng ngay bản mới.
function NM_xoaCacheLinhVuc() {
  try { CacheService.getScriptCache().remove(NM_SECTOR_CACHE_KEY); } catch (e) {}
  NM_BANG_LV_REQ_ = null;
  var b = NM_docBangLinhVuc_(true);
  Logger.log(b.ok ? ('Đã nạp lại ' + b.sectors.length + ' Lĩnh vực, ' + b.allSubs.length + ' Sub-sector.') : b.note);
}

// Cho giao diện hồ sơ khách (dropdown Sub-sector phụ thuộc Lĩnh vực).
function WA_getSectorMap() {
  // Lần gọi đầu trong mỗi lượt chạy: đọc thẳng sheet (và làm mới cache). Gọi lại trong cùng lượt: dùng bản vừa đọc.
  var b = NM_docBangLinhVuc_(!NM_BANG_LV_REQ_);  
  return b.ok ? { sectors: b.sectors, subs: b.subs } : null;
}

// Tách giá trị gõ/dán vào ô Lĩnh vực thành { sec, sub }. Nhận các dạng:
//   "1. Nông nghiệp, Lâm nghiệp & Thủy sản - Trồng trọt"  (dấu nối - – — | › > /)
//   "1. Nông nghiệp, Lâm nghiệp & Thủy sản (Trồng trọt)"
//   "Trồng trọt"  (chỉ Sub-sector, tự suy ra Lĩnh vực nếu Sub-sector chỉ thuộc 1 Lĩnh vực)
// Không tách được → null. Chỉ nhận khớp tuyệt đối sau chuẩn hoá để không đoán sai.
function NM_tachLinhVucGop_(s, bang) {
  var t = NM_sach_(s);
  if (!t || !bang || !bang.ok) return null;
  var secIdx = {};
  bang.sectors.forEach(function (x) { secIdx[NM_norm_(x)] = x; });
  function subTrong(sec, sub) {
    var n = NM_norm_(sub), ds = bang.subs[sec] || [];
    for (var i = 0; i < ds.length; i++) if (NM_norm_(ds[i]) === n) return ds[i];
    return null;
  }
  var m = t.match(/^(.*\S)\s*\(([^()]+)\)\s*$/);
  if (m) {
    var s1 = secIdx[NM_norm_(m[1])], b1 = s1 && subTrong(s1, m[2]);
    if (b1) return { sec: s1, sub: b1 };
  }
  var re = /\s+[-–—|›>\/]\s+|\s*[|›>]\s*/g, hit;
  while ((hit = re.exec(t)) !== null) {
    var left = t.slice(0, hit.index), right = t.slice(hit.index + hit[0].length);
    var s2 = secIdx[NM_norm_(left)], b2 = s2 && subTrong(s2, right);
    if (b2) return { sec: s2, sub: b2 };
  }
  var ds = bang.subToSectors[NM_norm_(t)] || [];
  if (ds.length === 1) return { sec: ds[0], sub: subTrong(ds[0], t) };
  return null;
}

function NM_context_() {
  var C = WA_CONFIG;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(C.SHEET_DATA);
  if (!sheet) throw new Error('Không tìm thấy sheet "' + C.SHEET_DATA + '".');

  var labels = WA_getCombinedHeaderLabels_(sheet, C.HEADER_ROW_DATA_1, C.HEADER_ROW_DATA_2);
  var map = WA_buildHeaderMap_(labels);
  var width = 0;
  labels.forEach(function (l, i) { if (l) width = i + 1; }); // độ rộng thật theo header, không dùng getLastColumn()

  var zones = {
    latest: WA_findMergedGroupRange_(sheet, C.HEADER_ROW_DATA_1, C.GROUP_LATEST_RESPONSE),
    dot1:   WA_findMergedGroupRange_(sheet, C.HEADER_ROW_DATA_1, C.GROUP_DOT1),
    dot2:   WA_findMergedGroupRange_(sheet, C.HEADER_ROW_DATA_1, C.GROUP_DOT2)
  };

  var lastDataRow = NM_timDongCuoi_(sheet, map);
  var tplRow = lastDataRow >= C.DATA_START_ROW ? lastDataRow : C.DATA_START_ROW;
  var tplFormulas = sheet.getRange(tplRow, 1, 1, width).getFormulas()[0];

  var fields = [], missing = [], skipped = [], daDung = {};
  NM_fieldDefs_().forEach(function (f) {
    if (f.zone) {
      var rg = zones[f.zone];
      if (!rg) { missing.push(f.label + ' (không thấy nhóm gộp ở dòng 4)'); return; }
      f.col = f.smart
        ? WA_timCotTheoNhan_(labels, rg, f.smart.candidates, f.smart.chuaTu, f.smart.khongChuaTu, f.smart.boQua)
        : WA_findColInRange_(labels, rg, f.sub);
      if (!f.col) { missing.push(f.label + ' (không thấy cột "' + (f.sub || f.smart.candidates[0]) + '" trong nhóm)'); return; }
    } else if (f.headers) {
      // Nhiều cách đặt tiêu đề được chấp nhận (VD "Sub-sector" / "Subsector"), so sau khi chuẩn hoá.
      f.col = NM_timCotNhieuTen_(labels, f.headers);
      if (!f.col) { missing.push(f.label + ' (chưa có cột tiêu đề "' + f.headers[0] + '" trên TEST)'); return; }
      f.header = labels[f.col - 1];
    } else {
      f.col = WA_firstCol_(map, f.header);
      if (!f.col) { missing.push(f.label + ' (tiêu đề "' + f.header + '")'); return; }
    }
    // Chốt chặn: không bao giờ để 2 trường trỏ cùng 1 cột (sẽ ghi đè lẫn nhau).
    if (daDung[f.col]) {
      missing.push(f.label + ' (dò ra trùng cột ' + WA_colLetter_(f.col) + ' của "' + daDung[f.col] + '")');
      return;
    }
    // Cột đang là công thức ở dòng mẫu → sheet tự tính, không cho nhập để khỏi ghi đè công thức.
    if (tplFormulas[f.col - 1]) { skipped.push(f.label); return; }
    daDung[f.col] = f.label;
    fields.push(f);
  });

  // Trường bắt buộc mà không dò được cột → dừng hẳn, tránh ghi dữ liệu thiếu.
  var thieuBatBuoc = NM_fieldDefs_().filter(function (d) {
    return d.required && !fields.some(function (f) { return f.key === d.key; });
  });
  if (thieuBatBuoc.length) {
    throw new Error('Không dò được cột bắt buộc: ' + thieuBatBuoc.map(function (d) { return d.label; }).join(', ') +
      '. Kiểm tra tiêu đề dòng 4/5 của sheet TEST (chạy WA_debugPrintHeaders để xem).');
  }

  return {
    ss: ss, sheet: sheet, labels: labels, map: map, width: width, zones: zones,
    lastDataRow: lastDataRow, tplRow: tplRow, tplFormulas: tplFormulas,
    fields: fields, missing: missing, skipped: skipped
  };
}

// So số cột CỐ ĐỊNH đang khai báo trong mã KM.gs / gui mail.gs với vị trí THẬT của tiêu đề trên sheet.
// Lệch = các file đó đang đọc/ghi nhầm cột (thường do có người chèn hoặc xoá cột trên TEST).
function NM_kiemTraLechCot_(ctx) {
  var colOf = {};
  ctx.fields.forEach(function (f) { colOf[f.key] = f.col; });
  colOf.maKM = WA_firstCol_(ctx.map, WA_CONFIG.COL_MAKM);

  var out = [];
  function so(tenFile, tenHang, giaTri, key, nhan) {
    if (giaTri === undefined || giaTri === null || !colOf[key]) return;
    if (Number(giaTri) !== colOf[key]) {
      out.push(tenFile + ': ' + tenHang + ' = ' + giaTri + ' (cột ' + WA_colLetter_(Number(giaTri)) +
        ') nhưng "' + nhan + '" trên sheet đang ở cột ' + WA_colLetter_(colOf[key]) + ' (' + colOf[key] + ')');
    }
  }
  try {
    if (typeof mkDongBoCotTheoTieuDe_ === 'function') throw 'mã KM.gs đã tự dò cột theo tiêu đề';
    if (typeof MK_COL_MA_KM !== 'undefined')     so('mã KM.gs', 'MK_COL_MA_KM', MK_COL_MA_KM, 'maKM', 'MÃ KM');
    if (typeof MK_COL_LA_KM_PHU !== 'undefined') so('mã KM.gs', 'MK_COL_LA_KM_PHU', MK_COL_LA_KM_PHU, 'kmChinh', 'Thuộc nhóm KM chính của');
    if (typeof MK_COL_DOI_TUONG !== 'undefined') so('mã KM.gs', 'MK_COL_DOI_TUONG', MK_COL_DOI_TUONG, 'doiTuong', 'Đối tượng');
    if (typeof MK_COL_HO_VA_TEN !== 'undefined') so('mã KM.gs', 'MK_COL_HO_VA_TEN', MK_COL_HO_VA_TEN, 'fullName', 'Full name');
    if (typeof MK_COL_MA_GMAIL !== 'undefined')  so('mã KM.gs', 'MK_COL_MA_GMAIL', MK_COL_MA_GMAIL, 'maGmail', 'Mã Gmail');
  } catch (e) { /* bỏ qua */ }
  try {
    if (typeof D_COL !== 'undefined') {
      var D = D_COL;
      so('gui mail.gs', 'D_COL.MA_KM', D.MA_KM, 'maKM', 'MÃ KM');
      so('gui mail.gs', 'D_COL.DANH_XUNG', D.DANH_XUNG, 'danhXung', 'Danh xưng');
      so('gui mail.gs', 'D_COL.TITLE', D.TITLE, 'title', 'Title');
      so('gui mail.gs', 'D_COL.HO_TEN_VN', D.HO_TEN_VN, 'hoTen', 'Họ và tên');
      so('gui mail.gs', 'D_COL.FULLNAME', D.FULLNAME, 'fullName', 'Full name');
      so('gui mail.gs', 'D_COL.CHUC_VU_VN', D.CHUC_VU_VN, 'chucVu', 'Chức vụ');
      so('gui mail.gs', 'D_COL.POSITION', D.POSITION, 'position', 'Position');
      so('gui mail.gs', 'D_COL.DON_VI_VN', D.DON_VI_VN, 'donVi', 'Tên đơn vị');
      so('gui mail.gs', 'D_COL.ORGANIZATION', D.ORGANIZATION, 'org', 'Organizagtion');
      so('gui mail.gs', 'D_COL.QUOC_GIA', D.QUOC_GIA, 'quocGia', 'Quốc gia');
      so('gui mail.gs', 'D_COL.EMAIL', D.EMAIL, 'email', 'Thư điện tử chính');
      so('gui mail.gs', 'D_COL.CC_EMAIL', D.CC_EMAIL, 'emailCc', 'Thư điện tử CC phụ');
      so('gui mail.gs', 'D_COL.THREAD_ID', D.THREAD_ID, 'maGmail', 'Mã Gmail');
      so('gui mail.gs', 'D_COL.PHAN_HOI_STD', D.PHAN_HOI_STD, 'xacNhan1', 'Phản hồi tham dự Save the day');
    }
    if (typeof D_LOAI_THU !== 'undefined') {
      if (D_LOAI_THU.D1) so('gui mail.gs', 'D_LOAI_THU.D1.statusCol', D_LOAI_THU.D1.statusCol, 'd1TienDo', 'Đợt 1 · Tiến độ liên hệ');
      if (D_LOAI_THU.D2) so('gui mail.gs', 'D_LOAI_THU.D2.statusCol', D_LOAI_THU.D2.statusCol, 'd2TienDo', 'Đợt 2 · Tiến độ');
    }
  } catch (e2) { /* bỏ qua */ }
  return out;
}

// Đọc danh sách lựa chọn từ 1 rule Data Validation (list gõ cứng hoặc tham chiếu vùng).
function NM_optionsFromRule_(rule) {
  if (!rule) return null;
  var t = rule.getCriteriaType();
  var vals = rule.getCriteriaValues();
  if (t === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) return vals[0];
  if (t === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) {
    var seen = {}, out = [];
    vals[0].getValues().forEach(function (r) {
      r.forEach(function (v) {
        var s = NM_sach_(v);
        if (s && !seen[s]) { seen[s] = true; out.push(s); }
      });
    });
    return out;
  }
  return null;
}

// Giá trị riêng biệt đang có trên sheet của vài cột (1 lần đọc). multiKeys: tách theo dấu phẩy.
function NM_docGiaTriRieng_(ctx, fields) {
  var out = {};
  var cols = fields.map(function (f) { return f.col; });
  if (!cols.length || ctx.lastDataRow < WA_CONFIG.DATA_START_ROW) return out;
  var minC = Math.min.apply(null, cols), maxC = Math.max.apply(null, cols);
  var values = ctx.sheet.getRange(WA_CONFIG.DATA_START_ROW, minC,
    ctx.lastDataRow - WA_CONFIG.DATA_START_ROW + 1, maxC - minC + 1).getValues();
  fields.forEach(function (f) {
    var seen = {}, list = [];
    values.forEach(function (row) {
      var s = NM_sach_(row[f.col - minC]);
      if (!s) return;
      (f.multi ? s.split(',') : [s]).forEach(function (p) {
        p = p.trim();
        var k = NM_norm_(p);
        if (p && k && !seen[k]) { seen[k] = true; list.push(p); }
      });
    });
    out[f.key] = list.sort(function (a, b) { return a.localeCompare(b, 'vi'); });
  });
  return out;
}

// Kiểu nhập của từng trường.
function NM_fieldMeta_(ctx) {
  var groups = WA_getPicGroupsFromSetupSheet_();
  var rules = ctx.sheet.getRange(ctx.tplRow, 1, 1, ctx.width).getDataValidations()[0];
  var distinct = NM_docGiaTriRieng_(ctx, ctx.fields.filter(function (f) { return f.suggest; }));
  var byKey = {};
  var doiTuongKhongMa = [];
  var bangLV = NM_docBangLinhVuc_();

  var list = ctx.fields.map(function (f) {
    var kind = 'text', options = f.options || null, strict = false, goiY = null;
    var rule = null;
    try { rule = rules[f.col - 1]; } catch (e) {}

    // Nguồn sự thật: bảng WA_DROPDOWN_MAP trong WA_Code.gs (đọc danh sách ở sheet THIẾT LẬP FILE).
    var dd = (typeof WA_dropdownTheoTieuDe_ === 'function')
      ? WA_dropdownTheoTieuDe_(f.zone ? (f.sub || ctx.labels[f.col - 1]) : f.header) : null;
    if (dd && !f.ref) {
      kind = 'list'; options = dd.options; strict = true; f.multi = dd.multi;
    } else if (f.ref) {
      // Cột KM chính: dropdown trên sheet lấy từ danh sách tên → dùng làm gợi ý, KHÔNG kiểm tra
      // theo danh sách (KM chính có thể đang nằm ngay trong đợt nhập này, chưa có trên sheet).
      kind = 'ref';
      options = (dd && dd.options.length) ? dd.options : (distinct.fullName || []);
    } else if (f.suggest) {
      kind = 'suggest';
      options = distinct[f.key] || [];
    } else {
      try {
        if (rule) {
          if (rule.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.CHECKBOX) {
            kind = 'checkbox'; options = ['Có', 'Không']; strict = true;
          } else {
            var dyn = NM_optionsFromRule_(rule);
            if (dyn && dyn.length) { kind = 'list'; options = dyn; strict = !rule.getAllowInvalid(); }
          }
        }
      } catch (e) { /* đọc rule lỗi thì dùng danh sách dự phòng trong WA_CONFIG */ }
      if (kind === 'text' && options && options.length) kind = 'list';
    }
    // Lĩnh vực / Sub-sector: danh sách chuẩn lấy thẳng từ THIẾT LẬP FILE (không phụ thuộc Data validation
    // của TEST). NHIỀU LỰA CHỌN: "A, B", từng mục phải nằm trong danh sách (kiểm ở NM_kiemTraRows_).
    if (bangLV.ok && (f.key === 'linhVuc' || f.key === 'subSector')) {
      kind = 'list'; f.multi = true;
      options = f.key === 'linhVuc' ? bangLV.sectors : bangLV.allSubs;
    }    

    // Ô nhiều lựa chọn: giá trị là chuỗi ghép nhiều mục nên không so "chặt" cả ô được,
    // từng mục vẫn phải nằm trong danh sách (kiểm ở NM_kiemTraRows_).
    if (f.multi) strict = false;

    // Đối tượng: chỉ gợi ý những giá trị mã KM.gs sinh được mã.
    if (f.key === 'doiTuong' && options && typeof MK_PREFIX_MAP !== 'undefined') {
      var ok = options.filter(function (o) { return NM_tienTo_(o); });
      doiTuongKhongMa = options.filter(function (o) { return !NM_tienTo_(o); });
      if (ok.length) goiY = ok;
    }

    var m = {
      key: f.key, label: f.label, header: f.header || '', group: f.group, col: f.col,
      colLetter: WA_colLetter_(f.col),
      required: !!f.required, common: !!f.common, free: !!f.free, log: !!f.log,
      multi: !!f.multi, warnNew: !!f.warnNew, email: f.email || '', raw: !!f.raw, hint: f.hint || '', phone: !!f.phone,      
      kind: kind, options: options, goiY: goiY, strict: strict,
      afterKey: f.afterKey || '', sortCol: f.col
    };
    byKey[f.key] = m;
    return m;
  });
  // Trường "đi kèm" (Sub-sector) luôn đứng ngay sau trường chính trên giao diện,
  // dù cột thật trên sheet nằm ở đâu.
  list.forEach(function (m) {
    if (m.afterKey && byKey[m.afterKey]) m.sortCol = byKey[m.afterKey].col + 0.5;
  });
  return { list: list, byKey: byKey, picGroups: groups, doiTuongKhongMa: doiTuongKhongMa, bangLV: bangLV };
}

// ================== API CHO GIAO DIỆN ==================
function WA_getNewGuestFormMeta() {
  var ctx = NM_context_();
  var meta = NM_fieldMeta_(ctx);
  var user = '';
  try { user = Session.getActiveUser().getEmail() || ''; } catch (e) {}
  return {
    fields: meta.list,
    picGroups: meta.picGroups,
    missing: ctx.missing,
    skipped: ctx.skipped,
    doiTuongKhongMa: meta.doiTuongKhongMa,
    linhVucMap: meta.bangLV.ok ? { sectors: meta.bangLV.sectors, subs: meta.bangLV.subs } : null,
    linhVucNote: meta.bangLV.ok ? '' : meta.bangLV.note,
    lechCot: NM_kiemTraLechCot_(ctx),
    user: user,
    maxRows: NM_MAX_ROWS,
    nextRow: Math.max(ctx.lastDataRow + 1, WA_CONFIG.DATA_START_ROW)
  };
}

function WA_checkNewGuests(rows, common) {
  rows = rows || [];
  var ctx = NM_context_();
  var meta = NM_fieldMeta_(ctx);
  var tach = NM_tachDongDaLuu_(rows);
  var results = NM_kiemTraRows_(ctx, meta, tach.conLai, common || {});
  tach.daLuu.forEach(function (d) {
    results.push({ id: d.id, vals: {}, norm: {}, errors: [], warnings: [{ key: 'fullName',
      msg: 'Dòng này ĐÃ được lưu ở lần bấm trước (dòng ' + d.info.row + (d.info.maKM ? ', mã ' + d.info.maKM : '') + '). Bấm Lưu sẽ không ghi lại.' }] });
  });
  return { results: results, nextRow: Math.max(ctx.lastDataRow + 1, WA_CONFIG.DATA_START_ROW) };
}

function WA_addNewGuests(rows, common, force) {
  return TB_chayCoKhoa_('Thêm KM mới', 60000, function () {
    return NM_themMoi_(rows, common || {}, !!force);
  });
}

// ================== CHỐNG GHI TRÙNG KHI BẤM LƯU LẠI ==================
// Mỗi dòng trên giao diện có 1 mã _uid cố định. Lưu xong, mã đó được nhớ 6 giờ.
// Trường hợp mạng chập chờn: máy chủ đã ghi xong nhưng trình duyệt không nhận được phản hồi,
// người dùng bấm Lưu lần nữa → các dòng đã ghi được nhận ra và bỏ qua, không sinh thêm KM trùng.
function NM_tachDongDaLuu_(rows) {
  var out = { daLuu: [], conLai: [] };
  var keys = rows.map(function (r) { return r && r._uid ? 'NM_' + String(r._uid).slice(0, 60) : null; });
  var hit = {};
  try { var k = keys.filter(Boolean); if (k.length) hit = CacheService.getScriptCache().getAll(k); } catch (e) {}
  rows.forEach(function (r, i) {
    var h = keys[i] && hit[keys[i]];
    if (h) { try { out.daLuu.push({ id: r._id, info: JSON.parse(h) }); return; } catch (e) {} }
    out.conLai.push(r);
  });
  return out;
}

function NM_nhoDongDaLuu_(rows, created) {
  var obj = {};
  rows.forEach(function (r, i) {
    if (!r._uid || !created[i]) return;
    obj['NM_' + String(r._uid).slice(0, 60)] = JSON.stringify({
      row: created[i].row, maKM: created[i].maKM || '', fullName: created[i].fullName || '', org: created[i].org || ''
    });
  });
  try { if (Object.keys(obj).length) CacheService.getScriptCache().putAll(obj, NM_CACHE_GIAY); } catch (e) {}
}

// ================== ĐỌC DỮ LIỆU ĐANG CÓ (1 LẦN) ==================
function NM_docDuLieuHienCo_(ctx, meta) {
  var C = WA_CONFIG;
  var ex = { byMa: {}, byName: {}, nameKey: {}, emailOwner: {}, nameToMa: {} };
  var col = function (k, h) { return meta.byKey[k] ? meta.byKey[k].col : (h ? WA_firstCol_(ctx.map, h) : null); };

  var cols = {
    maKM: WA_firstCol_(ctx.map, C.COL_MAKM),
    fullName: col('fullName', C.COL_FULLNAME),
    hoTen: col('hoTen', C.COL_HOTEN),
    doiTuong: col('doiTuong', C.COL_DOITUONG),
    email: col('email', C.COL_EMAIL),
    emailCc: col('emailCc', C.COL_EMAIL_CC),
    emailThay: col('emailThay', C.COL_EMAIL_THAY)
  };
  var used = Object.keys(cols).map(function (k) { return cols[k]; }).filter(Boolean);
  if (ctx.lastDataRow < C.DATA_START_ROW || !used.length) return ex;

  var minC = Math.min.apply(null, used), maxC = Math.max.apply(null, used);
  var values = ctx.sheet.getRange(C.DATA_START_ROW, minC, ctx.lastDataRow - C.DATA_START_ROW + 1, maxC - minC + 1).getValues();
  var LOAI = { email: 'email chính', emailCc: 'CC', emailThay: 'email liên hệ thay' };

  values.forEach(function (row) {
    var g = function (c) { return c ? NM_sach_(row[c - minC]) : ''; };
    var ma = g(cols.maKM), ten = g(cols.fullName), tenVn = g(cols.hoTen);
    if (!ma && !ten && !tenVn) return;
    var label = (ma || '(chưa có mã)') + (ten ? ' · ' + ten : (tenVn ? ' · ' + tenVn : ''));
    var info = { ma: ma, fullName: ten, hoTen: tenVn, doiTuong: g(cols.doiTuong), label: label };

    if (ma) ex.byMa[ma.toLowerCase()] = info;
    // Định danh khách CHỈ theo Full name (cột tiếng Anh). Cột tiếng Việt chỉ dùng cho thư khách TQ.
    var kTen = NM_norm_(ten);
    if (kTen) {
      (ex.byName[kTen] = ex.byName[kTen] || []).push(info);
      if (ma && !ex.nameToMa[kTen]) ex.nameToMa[kTen] = ma;
    }
    var kk = NM_khoaTen_(ten);
    if (kk && !ex.nameKey[kk]) ex.nameKey[kk] = label;
    ['email', 'emailCc', 'emailThay'].forEach(function (k) {
      NM_layEmail_(g(cols[k])).forEach(function (e) {
        if (!ex.emailOwner[e]) ex.emailOwner[e] = label + ' (' + LOAI[k] + ')';
      });
    });
  });
  return ex;
}

// ================== KIỂM TRA DỮ LIỆU ==================
// Trả về cho từng dòng:
//   { id, vals (giá trị sẽ ghi), norm (giá trị đã chuẩn hoá để hiện lại trên bảng),
//     errors: [{key,msg}] (chặn lưu), warnings: [{key,msg}] (hỏi xác nhận rồi vẫn lưu được) }
function NM_kiemTraRows_(ctx, meta, rows, common, exPre) {
  var ex = exPre || NM_docDuLieuHienCo_(ctx, meta);
  var coBangTienTo = (typeof MK_PREFIX_MAP !== 'undefined');
  var picGroups = meta.picGroups || {};
  var coNhomPic = Object.keys(picGroups).length > 0;
  var YES = ['co', 'x', 'yes', 'y', 'true', '1', 'v', 'dai tho'];
  var NO = ['khong', 'no', 'n', 'false', '0'];
  var tqList = (typeof D_QUOC_GIA_TQ !== 'undefined' && D_QUOC_GIA_TQ && D_QUOC_GIA_TQ.length) ? D_QUOC_GIA_TQ : ['trung quốc'];
  var TQ_ALIAS = ['china', 'prc', 'cn', 'trung hoa', 'peoples republic of china', 'people s republic of china', 'zhongguo', 'trung quoc'];
  var bangLV = meta.bangLV || { ok: false, sectors: [], subs: {}, subToSectors: {} };
  var coCotSub = !!meta.byKey.subSector;

  // Chỉ mục chuẩn hoá cho các trường gợi ý (tự sửa về đúng cách viết đang có trên sheet)
  var idx = {};
  meta.list.forEach(function (m) {
    if (m.kind !== 'suggest' && m.kind !== 'ref') return;
    idx[m.key] = {};
    (m.options || []).forEach(function (o) { var k = NM_norm_(o); if (k && !idx[m.key][k]) idx[m.key][k] = o; });
  });
  var tenTrungQuoc = (idx.quocGia && idx.quocGia['trung quoc']) || 'Trung Quốc';

  // ---------- VÒNG 1: chuẩn hoá từng ô ----------
  var pre = rows.map(function (raw) {
    var res = { id: raw._id, vals: {}, norm: {}, errors: [], warnings: [], _common: {}, _typed: {} };
    var err = function (k, m) { res.errors.push({ key: k, msg: m }); };
    var warn = function (k, m) { res.warnings.push({ key: k, msg: m }); };

    meta.list.forEach(function (m) {
      var rawTyped = (raw[m.key] === undefined || raw[m.key] === null) ? '' : String(raw[m.key]);
      var v = NM_sach_(rawTyped);
      // Sub-sector đã được tách từ ô Lĩnh vực của CHÍNH dòng này → không để giá trị chung đè lên.
      if (m.key === 'subSector' && !v && res._tachSub) return;
      var fromCommon = false;
      if (!v && m.common) { v = NM_sach_(common[m.key]); fromCommon = !!v; }
      // Sub-sector chung chỉ áp cho dòng dùng đúng Lĩnh vực chung (dòng tự gõ Lĩnh vực khác thì bỏ qua).
      if (m.key === 'subSector' && fromCommon && NM_sach_(raw.linhVuc) &&
          NM_norm_(raw.linhVuc) !== NM_norm_(common.linhVuc)) { v = ''; fromCommon = false; }
      if (!v) return;
      res._typed[m.key] = rawTyped;
      if (fromCommon) res._common[m.key] = true;
      if (m.phone) {
        var kqP = WA_chuanHoaSdt_(v);
        kqP.loi.forEach(function (x) { err(m.key, m.label + ': ' + x); });
        if (kqP.giaDinh.length) warn(m.key, m.label + ': "' + kqP.giaDinh.join('", "') + '" không có mã quốc gia nên đã hiểu là số Việt Nam (+84). Sửa lại nếu là số nước ngoài.');
        res.vals[m.key] = kqP.giaTri;
        var hienThi = kqP.vals.join('; ');
        if (!kqP.loi.length && !fromCommon && rawTyped !== hienThi) res.norm[m.key] = hienThi;
        return;
      }

      // Lĩnh vực gõ/dán dạng gộp "Lĩnh vực - Sub-sector" hoặc chỉ gõ Sub-sector → tách ra 2 cột.
      if (m.key === 'linhVuc' && bangLV.ok) {
        var secOut = [], subThem = [];
        NM_tachDaGiaTri_(v, bangLV.sectors).forEach(function (p) {
          var nP = NM_norm_(p), dung = null;
          bangLV.sectors.forEach(function (x) { if (!dung && NM_norm_(x) === nP) dung = x; });
          var tach = dung ? null : NM_tachLinhVucGop_(p, bangLV);
          if (tach) { secOut.push(tach.sec); subThem.push(tach.sub); }
          else secOut.push(dung || p);
        });
        v = NM_gop_(secOut);
        if (subThem.length && coCotSub && !NM_sach_(raw.subSector)) {
          res.vals.subSector = NM_gop_(subThem); res.norm.subSector = res.vals.subSector; res._tachSub = true;
        }
      }      

      if (m.email || m.kind === 'ref') { res.vals[m.key] = v; return; } // xử lý ở dưới / vòng 2

      if (m.kind === 'checkbox') {
        var n = NM_norm_(v);
        if (YES.indexOf(n) !== -1) { res.vals[m.key] = true; if (!fromCommon) res.norm[m.key] = 'Có'; }
        else if (NO.indexOf(n) !== -1) { res.vals[m.key] = false; if (!fromCommon) res.norm[m.key] = 'Không'; }
        else err(m.key, m.label + ': chỉ nhận "Có" hoặc "Không" (đang là "' + v + '")');
        return;
      }

      if (m.kind === 'list' && m.options && m.options.length) {
        if (m.multi) {
          var ok = [], bad = [];
          NM_tachDaGiaTri_(v, m.options).forEach(function (p) {
            p = p.trim(); if (!p) return;
            var mt = NM_khopLuaChon_(p, m.options);            
            if (mt) { if (ok.indexOf(mt) === -1) ok.push(mt); } else bad.push(p);
          });
          if (bad.length && !m.free) err(m.key, m.label + ': "' + bad.join('", "') + '" không có trong danh sách của sheet');
          v = ok.concat(bad).join(', ');
        } else {
          var matched = NM_khopLuaChon_(v, m.options);
          if (matched) v = matched;
          else if (m.strict) err(m.key, m.label + ': "' + v + '" không có trong danh sách của sheet');
          else if (!m.free) warn(m.key, m.label + ': "' + v + '" không có trong danh sách gợi ý');
        }
      } else if (m.kind === 'suggest') {
        var moi = [];
        var parts = (m.multi ? v.split(',') : [v]).map(function (p) {
          p = p.trim(); if (!p) return '';
          var k = NM_norm_(p);
          if (m.key === 'quocGia' && TQ_ALIAS.indexOf(k) !== -1) return tenTrungQuoc;
          if (idx[m.key] && idx[m.key][k]) return idx[m.key][k];
          moi.push(p);
          return p;
        }).filter(Boolean);
        v = parts.join(', ');
        if (moi.length && m.warnNew) {
          warn(m.key, m.label + ': "' + moi.join('", "') + '" chưa từng có trên sheet. Kiểm tra chính tả để Bộ lọc không bị tách thành 2 nhóm.');
        }
      }

      res.vals[m.key] = v;
      if (!fromCommon && rawTyped !== v) res.norm[m.key] = v;
    });

    // --- Email: tách, hạ chữ thường, bỏ trùng; CC không lặp lại email chính ---
    var prim = [];
    if (meta.byKey.email && res.vals.email !== undefined) {
      prim = NM_layEmail_(res.vals.email);
      if (!prim.length) err('email', 'Thư điện tử chính: không thấy địa chỉ email nào trong "' + res.vals.email + '"');
      prim.forEach(function (e) { if (!NM_EMAIL_RE.test(e)) err('email', 'Email chính không hợp lệ: ' + e); });
      if (prim.length > 1) warn('email', 'Email chính đang có ' + prim.length + ' địa chỉ. Nên để 1 địa chỉ, chuyển phần còn lại sang cột CC.');
      res.vals.email = prim.join(', ');
    }
    [['emailCc', 'Email CC'], ['emailThay', 'Email liên hệ thay']].forEach(function (p) {
      var k = p[0];
      if (!meta.byKey[k] || res.vals[k] === undefined) return;
      var list = NM_layEmail_(res.vals[k]).filter(function (e) { return prim.indexOf(e) === -1 || k !== 'emailCc'; });
      if (!list.length && k === 'emailCc' && NM_layEmail_(res.vals[k]).length) { res.vals[k] = ''; res.norm[k] = ''; return; }
      if (!list.length) { warn(k, p[1] + ': không thấy địa chỉ email nào trong "' + res.vals[k] + '"'); return; }
      list.forEach(function (e) { if (!NM_EMAIL_RE.test(e)) warn(k, p[1] + ' không hợp lệ, sẽ bị bỏ qua khi gửi/đối soát: ' + e); });
      res.vals[k] = list.join(', ');
    });
    ['email', 'emailCc', 'emailThay'].forEach(function (k) {
      if (res.vals[k] !== undefined && res._typed[k] !== undefined && res._typed[k] !== res.vals[k]) res.norm[k] = res.vals[k];
    });

    return res;
  });

  // ---------- VÒNG 2: kiểm tra chéo giữa các ô / các dòng / dữ liệu đang có ----------
  var batchByName = {};
  pre.forEach(function (r, i) {
    var k = NM_norm_(r.vals.fullName); if (!k) return;
    (batchByName[k] = batchByName[k] || []).push(i);
  });
  var kmOptIdx = idx.kmChinh || {};
  function tenThamChieu(p) {
    // Cột KM chính luôn chứa FULL NAME của KM chính (đúng chuỗi mà dropdown trên sheet đang dùng).
    var a = NM_norm_(p.fullName);
    if (a && kmOptIdx[a]) return kmOptIdx[a];
    return p.fullName || '';
  }

  var seenName = {}, seenEmail = {};

  pre.forEach(function (res, i) {
    var err = function (k, m) { res.errors.push({ key: k, msg: m }); };
    var warn = function (k, m) { res.warnings.push({ key: k, msg: m }); };
    var v = res.vals;
    var soDong = i + 1;

    // --- 1) KM chính (KM phụ) ---
    var doiTuongCha = '';
    if (v.kmChinh) {
      var ref = v.kmChinh, canon = '';
      var cha = ex.byMa[ref.toLowerCase()] || null;
      var dsTen = cha ? [] : (ex.byName[NM_norm_(ref)] || []);
      var trongDot = batchByName[NM_norm_(ref)] || [];

      if (!cha && dsTen.length > 1) {
        err('kmChinh', 'Có ' + dsTen.length + ' khách cùng tên "' + ref + '" (' +
          dsTen.map(function (x) { return x.ma || '?'; }).join(', ') + '). Gõ Mã KM của KM chính thay cho tên.');
      } else if (!cha && dsTen.length === 1) {
        cha = dsTen[0];
      }

      if (cha) {
        if (cha.ma && cha.ma.indexOf('.') !== -1) {
          err('kmChinh', '"' + ref + '" là KM phụ (mã ' + cha.ma + '), không làm KM chính được. Chọn KM chính của người đó.');
        } else if (!cha.fullName) {
          err('kmChinh', 'KM chính ' + (cha.ma || ref) + ' chưa có Full name trên sheet. Bổ sung Full name cho KM chính trước.');
        } else {
          canon = tenThamChieu(cha);
          doiTuongCha = cha.doiTuong;
        }
      } else if (!dsTen.length) {
        var khac = trongDot.filter(function (j) { return j !== i; });
        if (trongDot.length === 1 && trongDot[0] === i) {
          err('kmChinh', 'Một khách không thể là KM chính của chính mình.');
        } else if (khac.length > 1) {
          err('kmChinh', 'Có ' + khac.length + ' dòng trong bảng cùng tên "' + ref + '". Không xác định được KM chính.');
        } else if (khac.length === 1) {
          var p = pre[khac[0]];
          if (p.vals.kmChinh) {
            err('kmChinh', 'Dòng ' + (khac[0] + 1) + ' ("' + ref + '") cũng là KM phụ, không làm KM chính được.');
          } else {
            canon = p.vals.fullName || ref;
            doiTuongCha = p.vals.doiTuong || '';
          }
        } else {
          err('kmChinh', 'Không tìm thấy KM chính "' + ref + '" theo Full name hoặc Mã KM, trên sheet TEST lẫn trong bảng đang nhập.');
        }
      }
      if (canon && canon !== ref) { v.kmChinh = canon; res.norm.kmChinh = canon; }
    }

    // --- 2) KM phụ luôn cùng Đối tượng với KM chính (mã KM.gs cũng sẽ ép như vậy khi điều chỉnh mã) ---
    if (v.kmChinh && doiTuongCha) {
      if (!v.doiTuong || res._common.doiTuong) {
        if (v.doiTuong !== doiTuongCha) res.norm.doiTuong = doiTuongCha;
        v.doiTuong = doiTuongCha;
      } else if (v.doiTuong !== doiTuongCha) {
        warn('doiTuong', 'KM phụ phải cùng Đối tượng với KM chính. Đã đổi "' + v.doiTuong + '" thành "' + doiTuongCha + '".');
        v.doiTuong = doiTuongCha; res.norm.doiTuong = doiTuongCha;
      }
    }

    // --- 3) Bắt buộc ---
    meta.list.forEach(function (m) {
      if (m.required && NM_giaTriRong_(v[m.key])) err(m.key, 'Bắt buộc nhập: ' + m.label);
    });

    // --- 4) KM chính phải sinh được Mã KM ---
    if (!v.kmChinh && v.doiTuong && coBangTienTo && !NM_tienTo_(v.doiTuong)) {
      err('doiTuong', 'Đối tượng "' + v.doiTuong + '" chưa có tiền tố Mã KM trong mã KM.gs (MK_PREFIX_MAP) nên không sinh được mã. Chọn Đối tượng khác hoặc báo quản lý bổ sung.');
    }

    // --- 5) Phụ trách / Nhân sự hỗ trợ khớp THIẾT LẬP FILE ---
    if (coNhomPic && v.pic && !picGroups[v.pic]) {
      err('pic', '"' + v.pic + '" không phải PIC trong sheet THIẾT LẬP FILE (cột B).');
    }
    if (coNhomPic && v.pic && v.hoTro && picGroups[v.pic] && picGroups[v.pic].indexOf(v.hoTro) === -1) {
      var nhomThat = (typeof WA_findPicOfPerson_ === 'function') ? WA_findPicOfPerson_(v.hoTro, picGroups) : '';
      err('hoTro', v.hoTro + ' không thuộc nhóm PIC ' + v.pic +
        (nhomThat ? ' (đang thuộc nhóm ' + nhomThat + ')' : ' (không có trong THIẾT LẬP FILE)') +
        '. Tab "Mail đang xử lý" chia mail theo nhóm PIC nên lệch nhóm sẽ làm mail bị xếp sai người.');
    }

    // --- 6) Position ↔ Organizagtion phải khớp số đoạn ";" (gui mail.gs sẽ BỎ QUA khách này nếu lệch) ---
    var pos = NM_tachDoan_(v.position), org = NM_tachDoan_(v.org);
    if (Math.max(pos.length, org.length) > 1 && pos.length !== org.length) {
      err('position', 'Position có ' + pos.length + ' đoạn, Organizagtion có ' + org.length + ' đoạn (tách bằng ";"). ' +
        'Hai cột phải khớp từng cặp, nếu không hệ thống gửi thư sẽ báo lỗi và không gửi cho khách này.');
    }
    var cv = NM_tachDoan_(v.chucVu), dv = NM_tachDoan_(v.donVi);
    if (Math.max(cv.length, dv.length) > 1 && cv.length !== dv.length) {
      warn('chucVu', 'Chức vụ có ' + cv.length + ' đoạn, Tên đơn vị có ' + dv.length + ' đoạn (tách bằng ";"). Nên khớp từng cặp như bản tiếng Anh.');
    }

    // --- 6b) Lĩnh vực ↔ Sub-sector (NHIỀU GIÁ TRỊ): mỗi Sub-sector phải thuộc 1 trong các Lĩnh vực đã chọn;
    //         đã chọn Lĩnh vực có Sub-sector thì phải chọn ít nhất 1 Sub-sector.
    if (bangLV.ok && coCotSub) {
      var secs = NM_tachDaGiaTri_(v.linhVuc || '', bangLV.sectors);
      var subs = NM_tachDaGiaTri_(v.subSector || '', bangLV.allSubs);
      if (subs.length && !secs.length && meta.byKey.linhVuc) {          // chưa có Lĩnh vực: suy ra từ Sub-sector
        subs.forEach(function (b) {
          var dsSec = bangLV.subToSectors[NM_norm_(b)] || [];
          if (dsSec.length === 1) { if (secs.indexOf(dsSec[0]) === -1) secs.push(dsSec[0]); }
          else if (dsSec.length > 1) err('linhVuc', 'Sub-sector "' + b + '" có ở nhiều Lĩnh vực (' + dsSec.join(' / ') + '). Chọn Lĩnh vực.');
          else err('subSector', 'Sub-sector "' + b + '" không có trong danh sách của THIẾT LẬP FILE.');
        });
        if (secs.length) { v.linhVuc = NM_gop_(secs); res.norm.linhVuc = v.linhVuc; }
      }
      var secBiet = secs.filter(function (s) { return bangLV.subs[s]; });
      if (secBiet.length) {
        var pool = [];
        secBiet.forEach(function (s) { pool = pool.concat(bangLV.subs[s]); });
        if (subs.length && !pool.length) {
          err('subSector', 'Các Lĩnh vực đã chọn không có Sub-sector. Để trống ô Sub-sector.');
        } else if (subs.length) {
          var dung = [], sai = [];
          subs.forEach(function (b) {
            var nB = NM_norm_(b), k = null;
            pool.forEach(function (x) { if (!k && NM_norm_(x) === nB) k = x; });
            if (k) { if (dung.indexOf(k) === -1) dung.push(k); } else sai.push(b);
          });
          if (sai.length) {
            err('subSector', 'Sub-sector "' + sai.join('", "') + '" không thuộc Lĩnh vực đã chọn (' + secBiet.join(' / ') + ').');
          } else {
            var gop = NM_gop_(dung);
            if (gop !== v.subSector) { v.subSector = gop; res.norm.subSector = gop; }
          }
        } else if (pool.length) {
          err('subSector', 'Bắt buộc chọn Sub-sector cho Lĩnh vực "' + secBiet.join('", "') + '"');
        }
      }
    }
    // --- 7) Khách Trung Quốc: thư tiếng Hoa cần phần chữ Hán trong ngoặc ở Chức vụ / Tên đơn vị ---
    var qg = NM_sach_(v.quocGia).toLowerCase();
    if (qg && tqList.indexOf(qg) !== -1 && typeof D_ghepChucVuToChucTQ_ === 'function') {
      try { D_ghepChucVuToChucTQ_(v.chucVu || '', v.donVi || ''); }
      catch (e) {
        warn('chucVu', 'Khách Trung Quốc: ' + String(e && e.message ? e.message : e)
          .replace(/\s*Sửa ở sheet TEST rồi chạy lại\.?/i, '').replace(/\s*\(cột [A-Z]+\)/g, '') +
          ' Nếu chưa bổ sung, bảng điều khiển sẽ báo lỗi khi gửi thư tiếng Hoa.');
      }
    }

    // --- 8) Trùng tên (bỏ qua danh xưng, không phân biệt thứ tự từ) ---
    var baoTen = {};
    [v.fullName].forEach(function (n) {
      var k = NM_khoaTen_(n); if (!k) return;
      if (ex.nameKey[k] && !baoTen[ex.nameKey[k]]) {
        baoTen[ex.nameKey[k]] = true;
        warn('fullName', 'Có thể trùng khách đã có: ' + ex.nameKey[k]);
      } else if (seenName[k] !== undefined && seenName[k] !== i && !baoTen['#' + seenName[k]]) {
        baoTen['#' + seenName[k]] = true;
        warn('fullName', 'Có thể trùng với dòng ' + (seenName[k] + 1) + ' trong bảng');
      } else if (seenName[k] === undefined) seenName[k] = i;
    });

    // --- 9) Email trùng với khách khác (kể cả CC / liên hệ thay) → MAIL DATA.gs có thể gán mail nhầm Mã KM ---
    var lop = [['email', 'email chính'], ['emailCc', 'CC'], ['emailThay', 'email liên hệ thay']];
    var cuaDong = {};
    lop.forEach(function (p) {
      NM_layEmail_(v[p[0]]).forEach(function (e) {
        if (cuaDong[e]) return; cuaDong[e] = true;
        if (ex.emailOwner[e]) {
          warn(p[0], 'Email ' + e + ' đã có ở ' + ex.emailOwner[e] + '. Mail phản hồi từ địa chỉ này có thể bị gán nhầm Mã KM.');
        } else if (seenEmail[e] !== undefined && seenEmail[e] !== i) {
          warn(p[0], 'Email ' + e + ' trùng với dòng ' + (seenEmail[e] + 1) + ' trong bảng.');
        } else seenEmail[e] = i;
      });
    });
    if (meta.byKey.email && NM_giaTriRong_(v.email)) warn('email', 'Chưa có email chính, sẽ không gửi thư mời được.');
  });

  return pre.map(function (r) {
    return { id: r.id, vals: r.vals, norm: r.norm, errors: r.errors, warnings: r.warnings };
  });
}

// ================== SINH MÃ KM CHO CÁC DÒNG MỚI ==================
// KHÔNG gọi mkDieuChinhMa() / mkDienMaConThieu(): hai hàm đó quét lại TOÀN BỘ sheet bằng
// getRange() từng ô (và mkDieuChinhMa còn gọi SpreadsheetApp.getUi() — web app sẽ lỗi).
// Chỉ sinh mã cho đúng các dòng vừa thêm, dùng chung bộ đếm + hàm chống trùng của mã KM.gs.

function NM_phanTichMa_(tapMaHienCo) {
  var main = {}, sub = {};
  tapMaHienCo.forEach(function (ma) {
    var m = String(ma).match(/^([A-ZÀ-Ỹ]+)(\d+)(?:\.(\d+))?$/);
    if (!m) return;
    var pre = m[1], so = parseInt(m[2], 10);
    if (m[3] === undefined) {
      if (!(pre in main) || so > main[pre]) main[pre] = so;
    } else {
      var cha = pre + so, s = parseInt(m[3], 10);
      if (!(cha in sub) || s > sub[cha]) sub[cha] = s;
    }
  });
  return { main: main, sub: sub };
}

// Đẩy bộ đếm lên ít nhất bằng số đang có trên sheet, để mkSinhMaKeTiepAnToan() không phải quay vòng
// hàng trăm lần ghi PropertiesService khi bộ đếm đang thấp hơn dữ liệu thật.
function NM_dayBoDem_(khoa, soToiThieu) {
  if (!soToiThieu) return;
  var props = PropertiesService.getScriptProperties();
  var key = 'MK_CTR_' + khoa;
  var cur = parseInt(props.getProperty(key) || '0', 10);
  if (soToiThieu > cur) props.setProperty(key, String(soToiThieu));
}

function NM_sinhMaKM_(ctx, results, ex) {
  var notes = [];
  var codes = results.map(function () { return ''; });

  if (typeof MK_PREFIX_MAP === 'undefined' ||
      typeof mkLayDanhSachMaHienCo !== 'function' ||
      typeof mkSinhMaKeTiepAnToan !== 'function') {
    notes.push('Không tìm thấy file "mã KM.gs" trong project nên các dòng vừa thêm chưa có Mã KM. Mở sheet chạy "Điều chỉnh mã" để cấp mã.');
    return { codes: codes, notes: notes, changed: false };
  }

  if (typeof mkDongBoCotTheoTieuDe_ === 'function') mkDongBoCotTheoTieuDe_();
  var tap = mkLayDanhSachMaHienCo(ctx.sheet);
  var stats = NM_phanTichMa_(tap);
  var nameToMa = {};
  Object.keys(ex.nameToMa).forEach(function (k) { nameToMa[k] = ex.nameToMa[k]; });

  // Vòng 1: KM CHÍNH
  results.forEach(function (r, i) {
    if (r.vals.kmChinh) return;
    var prefix = NM_tienTo_(r.vals.doiTuong);
    if (!prefix) { notes.push('Dòng ' + (i + 1) + ': Đối tượng "' + (r.vals.doiTuong || '(trống)') + '" chưa có tiền tố mã nên Mã KM để trống.'); return; }
    NM_dayBoDem_('MAIN_' + prefix, stats.main[prefix] || 0);
    var ma = mkSinhMaKeTiepAnToan('MAIN_' + prefix, tap, function (so) { return prefix + so; });
    codes[i] = ma;
    var kFn = NM_norm_(r.vals.fullName); if (kFn) nameToMa[kFn] = ma;
  });

  // Vòng 2: KM PHỤ (sau vòng 1 để tra được mã cha vừa sinh)
  results.forEach(function (r, i) {
    if (codes[i]) return;
    var ref = (r.vals.kmChinh || '').toString().trim();
    if (!ref) return;
    var cha = ex.byMa[ref.toLowerCase()];
    var maCha = cha ? cha.ma : (nameToMa[NM_norm_(ref)] || '');
    if (!maCha) { notes.push('Dòng ' + (i + 1) + ': không tìm thấy mã của KM chính "' + ref + '" nên Mã KM để trống.'); return; }
    NM_dayBoDem_('SUB_' + maCha, stats.sub[maCha.toUpperCase()] || 0);
    codes[i] = mkSinhMaKeTiepAnToan('SUB_' + maCha, tap, function (so) { return maCha + '.' + so; });
  });

  return { codes: codes, notes: notes, changed: true };
}

// ================== GHI VÀO SHEET TEST ==================
function NM_themMoi_(rowsAll, common, force) {
  if (!rowsAll || !rowsAll.length) throw new Error('Chưa có dòng nào để lưu.');
  if (rowsAll.length > NM_MAX_ROWS) throw new Error('Mỗi lần lưu tối đa ' + NM_MAX_ROWS + ' dòng.');

  var C = WA_CONFIG;
  var tach = NM_tachDongDaLuu_(rowsAll);
  var createdCu = tach.daLuu.map(function (d) {
    return { id: d.id, row: d.info.row, maKM: d.info.maKM || '', fullName: d.info.fullName || '', org: d.info.org || '', daLuuTruoc: true };
  });
  var rows = tach.conLai;
  var notes = [];
  if (createdCu.length) notes.push(createdCu.length + ' dòng đã được lưu ở lần bấm trước nên không ghi lại lần nữa.');
  if (!rows.length) return { ok: true, created: createdCu, notes: notes, results: [] };

  var ctx = NM_context_();
  var meta = NM_fieldMeta_(ctx);
  var ex = NM_docDuLieuHienCo_(ctx, meta);
  var results = NM_kiemTraRows_(ctx, meta, rows, common, ex);

  var hasErr = results.some(function (r) { return r.errors.length; });
  var hasWarn = results.some(function (r) { return r.warnings.length; });
  if (hasErr) return { ok: false, results: results };
  if (hasWarn && !force) return { ok: false, needConfirm: true, results: results };

  var sheet = ctx.sheet;
  var n = results.length;
  var startRow = Math.max(ctx.lastDataRow + 1, C.DATA_START_ROW);
  var endRow = startRow + n - 1;
  if (endRow > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), endRow - sheet.getMaxRows());

  // --- 1) Gom giá trị theo cột ---
  var tz = Session.getScriptTimeZone();
  var now = new Date();
  var dateStr = Utilities.formatDate(now, tz, 'dd.MM');
  var writeCols = {};
  meta.list.forEach(function (m) {
    var any = results.some(function (r) { return r.vals[m.key] !== undefined && r.vals[m.key] !== ''; });
    if (!any) return;
    writeCols[m.col] = results.map(function (r) {
      var v = r.vals[m.key];
      if (v === undefined || v === null || v === '') return m.kind === 'checkbox' ? false : '';
      if (m.log) v = dateStr + ' - ' + v;
      return NM_giaTriGhi_(m, v);
    });
  });
  var colList = Object.keys(writeCols).map(Number).sort(function (a, b) { return a - b; });
  if (!colList.length) throw new Error('Không có dữ liệu nào để ghi.');

  // --- 2) An toàn: các ô đích phải đang trống ---
  var colMa = WA_firstCol_(ctx.map, C.COL_MAKM);
  var kiemCols = colList.concat(colMa ? [colMa] : []);
  var minC = Math.min.apply(null, kiemCols), maxC = Math.max.apply(null, kiemCols);
  var cur = sheet.getRange(startRow, minC, n, maxC - minC + 1).getValues();
  for (var r = 0; r < n; r++) {
    for (var k = 0; k < kiemCols.length; k++) {
      var cv = cur[r][kiemCols[k] - minC];
      if (!NM_giaTriRong_(cv)) {
        throw new Error('Ô ' + WA_colLetter_(kiemCols[k]) + (startRow + r) + ' trong sheet TEST đang có dữ liệu ("' +
          cv.toString().substring(0, 40) + '"). CHƯA GHI GÌ CẢ. Báo quản lý kiểm tra vùng cuối của sheet TEST.');
      }
    }
  }

  // --- 3) Sao định dạng, dropdown và công thức từng dòng từ dòng dữ liệu cuối ---
  if (ctx.lastDataRow >= C.DATA_START_ROW) {
    var tpl = sheet.getRange(ctx.tplRow, 1, 1, ctx.width);
    var dest = sheet.getRange(startRow, 1, n, ctx.width);
    tpl.copyTo(dest, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    tpl.copyTo(dest, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
    ctx.tplFormulas.forEach(function (fx, i) {
      if (!fx || writeCols[i + 1]) return;
      if (/ARRAYFORMULA/i.test(fx)) return;
      sheet.getRange(ctx.tplRow, i + 1).copyTo(
        sheet.getRange(startRow, i + 1, n, 1), SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);
    });
  }

  // --- 4) Ghi theo từng cụm cột liền nhau ---
  var a = 0;
  while (a < colList.length) {
    var b = a;
    while (b + 1 < colList.length && colList[b + 1] === colList[b] + 1) b++;
    var block = [];
    for (var rr = 0; rr < n; rr++) {
      var line = [];
      for (var cc = a; cc <= b; cc++) line.push(writeCols[colList[cc]][rr]);
      block.push(line);
    }
    sheet.getRange(startRow, colList[a], n, b - a + 1).setValues(block);
    a = b + 1;
  }
  SpreadsheetApp.flush();

  var created = results.map(function (x, i) {
    return { id: x.id, row: startRow + i, maKM: '',
      fullName: (x.vals.fullName || x.vals.hoTen || '').toString(), org: (x.vals.org || x.vals.donVi || '').toString() };
  });
  // Dữ liệu đã nằm trên sheet: nhớ ngay, để lỡ bước sau lỗi thì bấm Lưu lại cũng không ghi trùng.
  NM_nhoDongDaLuu_(rows, created);

  // --- 5) Sinh Mã KM (lỗi ở bước này KHÔNG làm mất dữ liệu đã ghi) ---
  var gen;
  try { gen = NM_sinhMaKM_(ctx, results, ex); }
  catch (e) {
    gen = { codes: results.map(function () { return ''; }), changed: false,
      notes: ['Lỗi khi sinh Mã KM: ' + (e && e.message ? e.message : e) + '. Dữ liệu ĐÃ được lưu. Báo quản lý chạy "Điều chỉnh mã" trên sheet để cấp mã.'] };
  }
  if (gen.changed && colMa) {
    sheet.getRange(startRow, colMa, n, 1).setValues(gen.codes.map(function (c) { return [c || '']; }));
  }
  created.forEach(function (c, i) { c.maKM = gen.codes[i] || ''; });
  NM_nhoDongDaLuu_(rows, created);
  notes = notes.concat(gen.notes || []);

  // --- 6) Dấu vết cho quản lý ---
  var user = '';
  try { user = Session.getActiveUser().getEmail() || ''; } catch (e) {}
  user = user || 'không rõ';
  var stamp = Utilities.formatDate(now, tz, 'dd/MM/yyyy HH:mm');
  //  a) Ghi chú (note) trên ô Full name: rê chuột vào ô trên sheet là thấy ai nhập, lúc nào.
  try {
    var colFn = meta.byKey.fullName.col;
    sheet.getRange(startRow, colFn, n, 1).setNotes(created.map(function () { return ['Thêm qua web app\n' + user + '\n' + stamp]; }));
  } catch (e) {}
  //  b) Nhật ký WA_Log_Chinh_Sua: 1 dòng / KM, kèm các cảnh báo người nhập đã chọn bỏ qua.
  try {
    var logSheet = WA_getOrCreateEditLog_(ctx.ss);
    logSheet.getRange(logSheet.getLastRow() + 1, 1, n, 6).setValues(created.map(function (c, i) {
      var w = results[i].warnings.map(function (x) { return x.msg; });
      var tomTat = [c.fullName, c.org, results[i].vals.pic ? 'PIC: ' + results[i].vals.pic : ''].filter(Boolean).join(' | ');
      if (w.length) tomTat += ' || Cảnh báo đã bỏ qua: ' + w.join(' ; ');
      return [now, user, c.maKM,
        'Thêm KM mới qua web (dòng ' + c.row + ')' + (w.length ? ' · bỏ qua ' + w.length + ' cảnh báo' : ''),
        '', tomTat.substring(0, 45000)];
    }));
  } catch (e) {}
  SpreadsheetApp.flush();

  return {
    ok: true, created: createdCu.concat(created), notes: notes, results: results,
    nextRow: endRow + 1
  };
}
