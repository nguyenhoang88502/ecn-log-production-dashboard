Bạn đang làm việc trên sheet "ECN" của workbook ECN Log. Hãy điền hai cột, TỪNG DÒNG MỘT:

- Cột R: "New item (est MP week)" - chữ `W` + 2 số, ví dụ `W07`, `W44`.
- Cột S: "New item (est MP year)" - số 4 chữ số, ví dụ `2026`.

DỮ LIỆU
- Dòng 4 là tiêu đề, dữ liệu từ dòng 5 đến dòng cuối có "ECN Number" ở cột A.
- Mỗi dòng là một mã vật tư (item) và có thể có tuần MP riêng. Một ECN có thể có các dòng với tuần khác nhau (ví dụ ECN 2026/QA/111 có W51 và W52). KHÔNG ép các dòng của một ECN về cùng một tuần.
- Cột B = Date. Cột N = Date Issued. Cột V = Note (nhiều dòng trong một ô, dòng mới nhất nằm TRÊN CÙNG).
- `#N/A` hoặc ô trống ở cột R nghĩa là chưa có.

TUẦN (cột R) - làm theo thứ tự, dừng ở bước đầu tiên có kết quả
1. R đã có `Wxx` hợp lệ: GIỮ NGUYÊN, không ghi đè.
2. Ô Note CỦA CHÍNH DÒNG ĐÓ (rồi đến các ô Note của các dòng khác cùng ECN), đọc từ trên xuống. Một dòng là DỜI LỊCH khi có CẢ HAI:
   - một số tuần: `W51`, `w46`, `W 51`, `week 51`, `wk 51`, `tuần 51`, và
   - một trong các từ `change`, `dời`, `đổi`.
   Ví dụ: `Est week 51 will change`, `est W51 change`, `Diep: Est change at week 44`, `10/04) est W43 change`.
   Phần `Wxx:` hoặc ngày `9/23)` đứng ĐẦU dòng là lúc ghi, KHÔNG phải tuần dời; bỏ nó trước khi tìm tuần.
   Lấy dòng dời lịch trên cùng. Nếu có nhiều số tuần trong dòng, lấy số cuối.
3. Note của dòng nói một NĂM lớn hơn năm của Date mà không có tuần (ví dụ `est in 04.2027`, `no FC until mid 2027`, `Lanh will inform - 2027`): để trống R, chỉ điền S = năm đó.
4. Còn lại: R = tuần ISO của Date: `="W"&TEXT(WEEKNUM(B5,21),"00")`. Nếu Date trống thì dùng Date Issued (cột N).
5. Không có gì để dựa vào: để trống R và S.

NĂM (cột S) - chỉ điền khi S đang trống, không ghi đè
- Năm gốc = năm ISO của Date: `=YEAR(B5+4-WEEKDAY(B5,2))` (ví dụ 29/12/2025 là W01 năm 2026).
- Nếu tuần ở R nhỏ hơn tuần ISO của Date quá 26 tuần (Date ở W50, R = W03) thì năm = năm gốc + 1.
- Nếu ô Note của dòng có năm 20xx lớn hơn năm gốc (ví dụ `- 2027`, `Est 2027 will change`) thì dùng năm đó.

CÁCH GHI
- Ghi GIÁ TRỊ, không để công thức. Cột R định dạng Text.
- Không sửa cột nào khác, không chèn/xóa cột, không đổi thứ tự dòng.
- Tô nền vàng nhạt các ô R/S do bạn điền.
- Tạo sheet mới "ECN fill log": Dòng | ECN Number | R cũ | R mới | S mới | Nguồn (Giữ nguyên / Note / Note (chỉ năm) / Date) | Dòng Note đã dùng.

BÁO CÁO CUỐI
- Số dòng theo từng nguồn: Giữ nguyên, Note, Note (chỉ năm), Date, Để trống.
- Liệt kê các dòng có năm S lớn hơn năm của Date để tôi xem lại.
- Liệt kê các dòng có tuần R nhỏ hơn tuần ISO của Date (kế hoạch nằm trước ngày ban hành).
