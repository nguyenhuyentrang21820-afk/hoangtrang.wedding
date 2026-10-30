/**
 * Nhận xác nhận tham dự + lời chúc từ thiệp cưới và lưu vào Google Sheet.
 *
 * Cài đặt:
 * 1. Tạo một Google Sheet mới → Tiện ích mở rộng → Apps Script.
 * 2. Xoá hết code mẫu, dán toàn bộ file này vào, bấm Lưu.
 * 3. Triển khai → Tùy chọn triển khai mới → chọn loại "Ứng dụng web":
 *      - Thực thi dưới dạng: Tôi
 *      - Người có quyền truy cập: Bất kỳ ai
 *    Bấm Triển khai, cấp quyền, rồi sao chép "URL ứng dụng web" (kết thúc bằng /exec).
 * 4. Mở index.html, dán URL đó vào dòng: const SHEET_URL='...';
 *
 * Mỗi lượt gửi là một dòng trong trang tính "Phản hồi". Bỏ tích cột "Hiển thị"
 * nếu không muốn một lời chúc hiện trên thiệp.
 * Khi sửa code này, phải Triển khai → Quản lý các bản triển khai → sửa → Phiên bản mới
 * thì thay đổi mới có hiệu lực (URL giữ nguyên).
 */

const SHEET_NAME = 'Phản hồi';
const HEADERS = ['Thời gian', 'Tên', 'Tham dự', 'Số người', 'Khách của', 'Lời chúc', 'Hiển thị'];

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// Cắt độ dài và chặn nội dung bị Sheet hiểu thành công thức (=, +, -, @)
function clean_(v, max) {
  const s = String(v == null ? '' : v).trim().slice(0, max);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function doPost(e) {
  let data;
  try {
    data = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: 'bad_request' });
  }
  const name = clean_(data.name, 80);
  if (!name) return json_({ ok: false, error: 'missing_name' });
  const attend = data.attend === 'yes' ? 'Có' : data.attend === 'no' ? 'Không' : '';

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = getSheet_();
    sh.appendRow([new Date(), name, attend, clean_(data.count, 20), clean_(data.side, 20), clean_(data.wish, 1000), true]);
    sh.getRange(sh.getLastRow(), HEADERS.length).insertCheckboxes().check();
  } finally {
    lock.releaseLock();
  }
  return json_({ ok: true });
}

// Trả về danh sách lời chúc (mới nhất trước) để hiện trong mục "Your wishes"
function doGet() {
  const sh = getSheet_();
  const rows = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, HEADERS.length).getValues() : [];
  const wishes = rows
    .filter(r => r[5] && r[6] !== false)
    .reverse()
    .slice(0, 300)
    .map(r => ({ name: String(r[1]).replace(/^'/, ''), text: String(r[5]).replace(/^'/, '') }));
  return json_({ wishes: wishes });
}
