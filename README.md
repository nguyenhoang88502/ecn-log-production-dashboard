# ECN Shopfloor Dashboard (Power BI PBIP) - v10

Hai bản cùng dữ liệu (đọc trực tiếp từ SharePoint):
- `ECN_Shopfloor.pbip`: 3 tab (Log / Lịch sử / Performance).
- `ECN_Assembly.pbip`: **chỉ tab Log**, cho assembly xem.

Mở `ECN_Shopfloor.pbip` trong thư mục này (KHÔNG mở từ trong file zip) -> Refresh.

## Nguồn dữ liệu: SharePoint
- File: `ECN Log.xlsx`, sheet `ECN`, site `NPIInformation`, thư mục
  `Shared Documents / General / Sharing to Production, Planner and Buyer`.
- Đổi nguồn: Transform data -> parameter `pSiteUrl`, `pFolderPath`, `pFileName`, `pSheetName`.
- Lần Refresh đầu Power BI hỏi đăng nhập: chọn **Microsoft account / Organizational account**, đăng nhập tài khoản Wahl,
  bấm Connect. Không chọn Anonymous (lỗi 403 và bị nhớ; xóa ở File > Options > Data source settings).
- Người Refresh cần quyền đọc site NPIInformation, không chỉ link thư mục.
- Muốn đọc file trên máy: sửa query `ECN Raw` thành `Excel.Workbook(File.Contents("C:\...\ECN Log.xlsx"), null, true)` rồi lấy sheet `ECN`.

## Tuần chạy và dời lịch (thứ tự ưu tiên)
1. **Cột `New item (est MP week)` (R) + `New item (est MP year)` (S)**: tuần MP của từng dòng item; ECN lấy tuần sớm nhất
   (so sánh theo năm rồi tuần, nên `W02` năm 2027 đứng sau `W44` năm 2026). Giá trị gõ kiểu `W44 --> W43` lấy số cuối.
2. Nếu R trống: **dòng Note có tuần (`W51`, `week 51`) và chữ `change`** (hoặc `dời`, `đổi`) = dời lịch sang tuần đó.
3. R trống và S có năm lớn hơn năm của Date (ví dụ 2027): hiện **năm** đó, nhóm "Sắp tới".
4. Còn lại: **tuần ISO của cột `Date`**.
- Dòng Note không thỏa điều kiện = sự kiện trong lịch sử. Tuần/ngày đứng đầu dòng (`W40:`, `9/23)`) là lúc ghi, không phải tuần dời.
- "Dời N lần" = số dòng Note là dời lịch. Cột "Lịch sử dời tuần": `W23 → W42` (tuần của Date -> các lần dời -> tuần MP hiện tại).
- Điền R/S cho dòng cũ: dùng `Prompt_Claude_Excel.md` (Claude in Excel). Thử trước bằng `fill_est.py`
  (kết quả mẫu: `preview/DRY_RUN_dien_R_S_chi_de_xem.xlsx`, KHÔNG dùng thay file thật).
- Nhóm tuần: Tuần này / Tuần sau / Sắp tới / Tuần cũ chưa xong (ECN còn mở, tuần đã qua) / Chưa có lịch / Đã xong.
  Tab Log luôn hiện Tuần này + Tuần sau; tab Performance có chip chọn nhóm tuần.

## 3 tab
| Tab | Nội dung |
|---|---|
| **Log - hôm nay chạy gì** | Mở lên hiện sẵn **Tuần này (bên trái)** và **Tuần sau (bên phải)**, mỗi bên có tiêu đề `TUẦN NÀY · W42 · n ECN`. Mỗi bên 4 cột chữ lớn: ECN number (+ `dời từ Wxx` nếu có dời), FG, Description of Change (2 thay đổi đầu + `+n khác`), cỡ chữ 13-17pt (1pt = 1.33px, đã chỉnh để không vỡ chữ), EC Instruction (Immediately tô đỏ nhạt). Không hiện Note. Lọc: Mã FG (gõ để tìm), EC Instruction, XÓA LỌC. |
| **Lịch sử** | Mỗi dòng Note = một sự kiện. Chip lọc **Dời lịch / Khác**, lọc Phân loại, FG, Giấy ECN. 4 ô: Sự kiện, Dòng dời lịch, ECN có dời lịch, Dòng chưa nhận dạng. |
| **Performance** | 4 ô: Tổng ECN, FG đang mở, Giấy ECN đã gửi, Ngày từ ban hành đến hoàn tất (trung vị). 6 hình: ECN theo Model (top 10) / Ban hành -> hoàn tất / ECN đang mở theo giấy / **ECN mở mới mỗi tuần (line chart 26 tuần)** / Mở mới vs hoàn tất (13 tuần) / **Top 5 FG (treemap)**. |

## Số kỳ vọng ngày 2026-10-13 (W42) với file `ECN Log.xlsx` mới, Python tính
| Nhóm tuần | Số ECN |
|---|---|
| Tuần này | 1 (2026/QA/087: W23 -> W42) |
| Tuần sau | 0 |
| Sắp tới | 35 |
| Tuần cũ chưa xong | 71 |
| Chưa có lịch | 1 |
| Đã xong / hủy | 160 |

268 ECN, 628 dòng. Cột R đang có 100 dòng điền sẵn (99 W.. + 1 gõ tay), cột S chưa có dòng nào, nên con số "Tuần cũ chưa xong"
sẽ giảm sau khi chạy prompt điền R/S (các dòng còn #N/A sẽ lấy từ Date, đa số vẫn là tuần cũ).
Đặt parameter `pTodayOverride` = `2026-10-07` để xem lại ngày khác.

## Lưu ý
- File chưa được Power BI Desktop kiểm tra ở môi trường này. `preview/*.png` chỉ là ảnh mô phỏng bố cục bằng dữ liệu thật, KHÔNG phải ảnh chụp Power BI.
- Slicer "Giấy ECN" đã bỏ ở mọi tab, thay bằng ô Mã FG lớn hơn (gõ mã FG là ra ECN liên quan).
- `2026/QA/008` ("W40: Lanh will inform - W44") không có chữ change nên theo Date (W05, tuần cũ chưa xong). Muốn tính `- W44` là dời lịch thì thêm chữ change vào dòng đó.
- Nút XÓA LỌC dùng hành động "Clear all slicers"; nếu không có tác dụng thì dùng biểu tượng cục tẩy trên từng bộ lọc.
- Nhiều ECN cũ được đóng hàng loạt cùng một ngày nên nhóm "> 30 ngày" lớn.
- Quy tắc nhập Excel: `ECN_Entry_Rules.xlsx`. Bản cũ trong `archive/`.
- Power BI Import không tự refresh mỗi 5 phút; không có carousel tự chuyển tab.
