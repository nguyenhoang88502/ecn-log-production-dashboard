# Quy tắc nhập ECN Log (để dashboard tự đọc)

Mỗi ô văn bản tự do được dashboard tách thành thông tin có cấu trúc. Nhập đúng quy tắc thì không phải sửa dashboard; dòng không đọc được sẽ hiện ở trang 3 mục "Chưa nhận dạng".

## 1. Quy tắc theo cột

| Cột | Quy tắc | Đúng | Sai |
|---|---|---|---|
| ECN Number | Dạng YYYY/QA/NNN, ví dụ 2026/QA/192. Dòng con của cùng ECN lặp lại ĐÚNG số này (không để trống, không thừa dấu cách). | `2026/QA/192` | Dòng con để trống số ECN; 2026/QA/ 192 |
| Date | Ngày ban hành / kế hoạch của ECN, nhập là ngày thật của Excel. Tuần chạy của ECN = tuần của ngày này (trừ khi Note dời lịch). Nếu trống, dashboard dùng cột Date Issued. | `07/10/2026` | '7-10' dạng chữ; để trống |
| FG | Mã thành phẩm, mỗi mã 7 số, ngăn cách bằng dấu phẩy. Không xuống dòng để ngăn mã. | `1004150, 1004330, 1009196` | 1004150 1004330 / xuống dòng |
| Model /Customer No. | Mỗi model một mục, ngăn bằng dấu phẩy. Viết cụ thể từng model, không viết 'all 9953'. | `9865-1301, 9865-1301VR` | all 9953 |
| Old item | 1 mã mỗi ô, 7 số bắt đầu bằng 1 hoặc 3. Thêm mã mới thì để trống ô này. | `1001733` | 1001733, 1001734 trong 1 ô |
| New item | 1 mã mỗi ô, 7 số bắt đầu bằng 1 hoặc 3. Bỏ mã (không thay) thì để trống ô này. | `3034045` | 'hộp mới' bằng chữ |
| Buyer | Chỉ 3 giá trị: Local / Import / In-house. | `Local` | FGS, SUBCONT, ASSM, Local-VV |
| Status | Ongoing / Done / Cancel. Để trống = chưa bắt đầu. Viết đúng chữ hoa/thường. | `Ongoing` | ongoing, Closed, OK |
| Description of Change | Mỗi dòng một việc, mở đầu bằng động từ Thay / Thêm / Bỏ / Cập nhật (hoặc Replace / Add / Remove / Update), nêu tên vật tư rồi mới tới mã. Mã: 7 số hoặc #số (2-5 chữ số). Nhiều mã ngăn bằng dấu phẩy. | `Thay hộp màu 3029521 bằng hộp màu mới 3031550` | Đổi hộp màu mới (không có mã) |
| EC Instruction | Chỉ 2 giá trị: Running Change hoặc Immediately (không dư khoảng trắng cuối). | `Immediately` | Immediately␠ |
| Date Issued / Date Completed | Ngày thật của Excel. | `07/10/2026` | đang làm |
| ECN form sent date / received date | Ngày thật, hoặc chữ 'Đã gửi giấy' / 'Đã đưa giấy'. Có giá trị = đã gửi / đã nhận. | `Đã gửi giấy` | x |
| New item (est MP week) | Chữ W + số tuần ISO, không dấu cách. Tuần 1-12 trong khi đang ở tuần 40+ được hiểu là năm sau. | `W44` | 44 / Wk44 / W 44 |
| Note | Mỗi dòng một sự kiện, dòng mới nhất nằm TRÊN CÙNG (Alt+Enter để xuống dòng). Muốn DỜI LỊCH: viết dòng có tuần (Wxx hoặc week xx) VÀ chữ 'change'. Dòng khác chỉ là lịch sử sự kiện. | `Est week 51 will change` | Dời lịch mà không ghi tuần; ghi tuần nhưng thiếu chữ change |

## 2. Cú pháp cột Note

Một dòng = một sự kiện. Dòng mới nhất nằm trên cùng. Mẫu: `[tuần hoặc ngày]: [người] [việc] [- Wxx]`.

| Thông tin | Viết như | Cách đọc |
|---|---|---|
| DỜI LỊCH (quan trọng) | `Est week 51 will change  /  est W51 change  /  Diep: Est change at week 44` | Dòng có Wxx hoặc 'week xx' VÀ có chữ change = dời lịch sang tuần đó. 'change' có thể thay bằng 'dời' hoặc 'đổi'. Dòng dời nằm trên cùng là lịch hiện tại; các dòng dời cũ bên dưới là lịch sử dời tuần. |
| Không dời lịch | `W40: Lanh follow up  /  9/23 update) est W45 clear stock  /  W32: inform to change BOM` | Không có tuần cần dời + change: chỉ là sự kiện trong tab Lịch sử, tuần chạy vẫn theo Date. Tuần/ngày đứng đầu dòng (W40: / 9/23)) là lúc ghi, không phải tuần dời. |
| Chỉ có năm | `Est 2027 will change` | Tính là dời lịch (hiện 'năm 2027') nhưng không đổi tuần chạy. |
| Người phụ trách | `Lanh, Diep, Hien, Thao, Nga, Tra, Cuong` | Viết đúng tên trong danh mục, chữ cái đầu viết hoa. Thêm người mới: sửa parameter pPeople trong Power BI. |
| Dời lịch | `W44: dời do chờ NCC | W40: Lanh will inform | W34: chờ clear tồn` | Thêm dòng MỚI Ở TRÊN CÙNG với tuần mới. Giữ lại các dòng cũ bên dưới: dashboard đếm số lần dời = số tuần khác nhau - 1. |
| Tuần dự kiến | `est W45 clear stock  /  Est week 51 will change  /  est end 2026` | Từ 'est' + W45 hoặc 'week 45' (trong 5 chữ kế tiếp). Muốn nói năm: est 2027 / est end 2026. |
| Hạn hoàn tất | `Finished by 31-Oct-26` | Dạng dd-Mon-yy, tiếng Anh: 31-Oct-26. |
| Đơn hàng cuối | `last order Nov-2026` | Sau 'last order ' ghi tháng-năm. |
| Tồn kho cần xả | `clear stock 3031595 (local)  /  stock 398pcs` | Trong dòng có chữ 'stock' thì các mã 7 số sẽ được lấy làm mã tồn; số lượng viết liền pcs: 398pcs. |
| Đã xong | `Done  /  OK  /  Updated` | Dòng chỉ 1-3 từ, từ đầu là Done / OK / Updated. |
| Hủy | `Cancel` | Dòng chỉ có chữ Cancel. |

### Từ khóa phân loại dòng Note (khớp từ trên xuống, dòng đầu tiên khớp thắng)

| Phân loại | Dòng chứa một trong các cụm |
|---|---|
| Đã xong | Done / OK / Updated (dòng ≤ 3 từ) |
| Hủy | Cancel |
| Chưa có nhu cầu / đơn | no more demand, không thấy có demand, chưa có kh, chưa tạo mo, no order, không cần clear |
| Chờ Buyer / giấy tờ | not yet email, buyer not action, wrote mail, đã làm giấy, buyer informed, buyer noted, buyer updated, buyer est |
| Chờ khách / thiết kế | customer, cus., design, khách |
| Chờ NCC / hàng về | vendor, delivery, arrives, last order, quick connect, viet may, deliver, viet may, quick connect dev, vinh vien |
| Xử lý tồn kho | stock, tồn, scrap, clear, wh , inventory, keep old, available |
| Setup / BOM | setup, bom, mo , level |
| Hạn hoàn tất | finished by |
| Theo dõi | follow, recheck, check, inform, update, est, wait, will change |
| Chưa nhận dạng | không khớp cụm nào ở trên |

### Ví dụ đúng

- `W40: Lanh will inform - W44`  ->  Tuần 40, Lanh, dời MP sang W44
- `9/23 update) est W45 clear stock`  ->  Ngày 23/09, dự kiến W45, xử lý tồn kho
- `W34: waiting for Quick Connect Dev, last order Nov-2026`  ->  Tuần 34, chờ NCC, đơn cuối Nov-2026
- `Finished by 31-Oct-26`  ->  Hạn hoàn tất 31/10/2026
- `8/26) no more demand in 2026, stock 398pcs`  ->  Chưa có nhu cầu, tồn 398 pcs
- `Sep 28 2026: wrote mail to a Hien change BOM`  ->  Ngày 28/09, Hien, chờ giấy tờ

### Ví dụ chưa đọc được (lấy từ file thật)

- `w40: Lanh will infrom`  ->  Gõ sai 'infrom' -> Chưa nhận dạng. Viết 'inform'.
- `FGS / SUBCONT / ASSM`  ->  Đây là loại Buyer, ghi vào cột Buyer, không ghi vào Note.
- `Đang làm`  ->  Không có tuần, ngày, người, từ khóa. Viết: W41: Lanh follow up.
- `W34 ~ 35 still having stock`  ->  Dùng dấu : sau tuần: W34-35: still having stock.

## 3. Từ khóa tạo Tag và Vật tư (từ Description of Change)

| Tag | Từ khóa |
|---|---|
| Bao bì | hộp, thùng, carton, túi, blister, mút, packaging, color box, graphic card, card, quai, bvld, mailer, vỉ, bao bì, pe bag, foam, khay, lót, chèn |
| Cơ khí | thân, trimmer, lưỡi, blade, vít, screw, nắp, housing, cơ khí, lò xo, spring, gá, khuôn, linh kiện nhựa, motor, mô tơ, đầu cắt, comb, lược, dao, ốc, phôi, jack, đầu cắm, cụm |
| Mạch | mạch, pcb, bo mạch, board, chip, dây điện, wire, pin, battery, adapter, sạc, charger, led, switch, công tắc, mach, usb, type c, bóng đèn, 4v, 5v |
| Nhãn | nhãn, label, tem, laser, sticker, sách, manual, hướng dẫn, in ấn, bản vẽ, barcode, mã vạch, decal, logo, print, nội dung in, phiên bản, version, warning, warranty, instruction, ib #, fill |

| Vật tư | Từ khóa (cụm ≤ 3 ký tự phải đứng riêng thành một từ) |
|---|---|
| Hộp màu | hộp màu, color box, clor box, hộp, box |
| Thùng carton | thùng, carton |
| Túi / pouch | túi, pouch, zip |
| Nhãn | nhãn, label, tem |
| Sách HD | sách, book, manual, instruction |
| Graphic card | graphic card, card |
| Blister / khay | blister, khay |
| Mút / chèn | mút, chèn, foam |
| Bản mạch | mạch, pcb, mach |
| Thân / vỏ | thân, vỏ, housing |
| Dao / lưỡi | dao, blade, lưỡi |
| Trimmer | trimmer, đầu tỉa |
| Pin / sạc | pin, sạc, battery, 4v, 5v, usb |
| Laser / in | laser, in xoay, logo, print |
| Chống ẩm | desiccant, chống ẩm |
| Mô tơ | mô tơ, motor |

## 4. Động từ trong Description of Change

- Bỏ: bỏ, remove, romove, hủy, xóa, scrap
- Thêm: thêm, add, bổ
- Thay: thay, replace, rplace, đổi, change
- Cập nhật: update, updete, cập
- Nối cũ -> mới: bằng, by, thành, to, sang, ->, →, ➔

Mã đứng sau động từ Bỏ/Thay là mã cũ; sau Thêm hoặc sau từ nối (bằng, by, to...) là mã mới.

## 5. Thêm người, NCC, từ khóa mới

- Người và NCC: Power BI Desktop -> Transform data -> parameter `pPeople`, `pVendors` (cách nhau bằng dấu phẩy).
- Từ khóa Tag / Vật tư / Phân loại: sửa danh sách ở đầu `ecn_parse.py` rồi chạy lại `build_ecn_pbip.py`, hoặc sửa trực tiếp hàm `fnTags`, `fnComps`, `fnNoteLine` trong Power Query.
