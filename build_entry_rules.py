# -*- coding: utf-8 -*-
"""Writes the data-entry rules (markdown + xlsx) from the same dictionaries the parser uses."""
import os, sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import ecn_parse as P

ROOT = os.path.dirname(os.path.abspath(__file__))

COLUMN_RULES = [
    ("ECN Number", "Dạng YYYY/QA/NNN, ví dụ 2026/QA/192. Dòng con của cùng ECN lặp lại ĐÚNG số này (không để trống, không thừa dấu cách).",
     "2026/QA/192", "Dòng con để trống số ECN; 2026/QA/ 192"),
    ("Date", "Ngày ban hành / kế hoạch của ECN, nhập là ngày thật của Excel. Tuần chạy của ECN = tuần của ngày này "
     "(trừ khi Note dời lịch). Nếu trống, dashboard dùng cột Date Issued.",
     "07/10/2026", "'7-10' dạng chữ; để trống"),
    ("FG", "Mã thành phẩm, mỗi mã 7 số, ngăn cách bằng dấu phẩy. Không xuống dòng để ngăn mã.",
     "1004150, 1004330, 1009196", "1004150 1004330 / xuống dòng"),
    ("Model /Customer No.", "Mỗi model một mục, ngăn bằng dấu phẩy. Viết cụ thể từng model, không viết 'all 9953'.",
     "9865-1301, 9865-1301VR", "all 9953"),
    ("Old item", "1 mã mỗi ô, 7 số bắt đầu bằng 1 hoặc 3. Thêm mã mới thì để trống ô này.",
     "1001733", "1001733, 1001734 trong 1 ô"),
    ("New item", "1 mã mỗi ô, 7 số bắt đầu bằng 1 hoặc 3. Bỏ mã (không thay) thì để trống ô này.",
     "3034045", "'hộp mới' bằng chữ"),
    ("Buyer", "Chỉ 3 giá trị: Local / Import / In-house.", "Local", "FGS, SUBCONT, ASSM, Local-VV"),
    ("Status", "Ongoing / Done / Cancel. Để trống = chưa bắt đầu. Viết đúng chữ hoa/thường.",
     "Ongoing", "ongoing, Closed, OK"),
    ("Description of Change", "Mỗi dòng một việc, mở đầu bằng động từ Thay / Thêm / Bỏ / Cập nhật (hoặc Replace / Add / Remove / Update), "
     "nêu tên vật tư rồi mới tới mã. Mã: 7 số hoặc #số (2-5 chữ số). Nhiều mã ngăn bằng dấu phẩy.",
     "Thay hộp màu 3029521 bằng hộp màu mới 3031550", "Đổi hộp màu mới (không có mã)"),
    ("EC Instruction", "Chỉ 2 giá trị: Running Change hoặc Immediately (không dư khoảng trắng cuối).",
     "Immediately", "Immediately␠"),
    ("Date Issued / Date Completed", "Ngày thật của Excel.", "07/10/2026", "đang làm"),
    ("ECN form sent date / received date", "Ngày thật, hoặc chữ 'Đã gửi giấy' / 'Đã đưa giấy'. Có giá trị = đã gửi / đã nhận.",
     "Đã gửi giấy", "x"),
    ("New item (est MP week)", "Chữ W + số tuần ISO, không dấu cách. Tuần 1-12 trong khi đang ở tuần 40+ được hiểu là năm sau.",
     "W44", "44 / Wk44 / W 44"),
    ("Note", "Mỗi dòng một sự kiện, dòng mới nhất nằm TRÊN CÙNG (Alt+Enter để xuống dòng). Muốn DỜI LỊCH: viết dòng có "
     "tuần (Wxx hoặc week xx) VÀ chữ 'change'. Dòng khác chỉ là lịch sử sự kiện.",
     "Est week 51 will change", "Dời lịch mà không ghi tuần; ghi tuần nhưng thiếu chữ change"),
]

NOTE_RULES = [
    ("DỜI LỊCH (quan trọng)", "Est week 51 will change  /  est W51 change  /  Diep: Est change at week 44",
     "Dòng có Wxx hoặc 'week xx' VÀ có chữ change = dời lịch sang tuần đó. 'change' có thể thay bằng 'dời' hoặc 'đổi'. "
     "Dòng dời nằm trên cùng là lịch hiện tại; các dòng dời cũ bên dưới là lịch sử dời tuần."),
    ("Không dời lịch", "W40: Lanh follow up  /  9/23 update) est W45 clear stock  /  W32: inform to change BOM",
     "Không có tuần cần dời + change: chỉ là sự kiện trong tab Lịch sử, tuần chạy vẫn theo Date. "
     "Tuần/ngày đứng đầu dòng (W40: / 9/23)) là lúc ghi, không phải tuần dời."),
    ("Chỉ có năm", "Est 2027 will change", "Tính là dời lịch (hiện 'năm 2027') nhưng không đổi tuần chạy."),
    ("Người phụ trách", "Lanh, Diep, Hien, Thao, Nga, Tra, Cuong",
     "Viết đúng tên trong danh mục, chữ cái đầu viết hoa. Thêm người mới: sửa parameter pPeople trong Power BI."),
    ("Dời lịch", "W44: dời do chờ NCC | W40: Lanh will inform | W34: chờ clear tồn",
     "Thêm dòng MỚI Ở TRÊN CÙNG với tuần mới. Giữ lại các dòng cũ bên dưới: dashboard đếm số lần dời = số tuần khác nhau - 1."),
    ("Tuần dự kiến", "est W45 clear stock  /  Est week 51 will change  /  est end 2026",
     "Từ 'est' + W45 hoặc 'week 45' (trong 5 chữ kế tiếp). Muốn nói năm: est 2027 / est end 2026."),
    ("Hạn hoàn tất", "Finished by 31-Oct-26", "Dạng dd-Mon-yy, tiếng Anh: 31-Oct-26."),
    ("Đơn hàng cuối", "last order Nov-2026", "Sau 'last order ' ghi tháng-năm."),
    ("Tồn kho cần xả", "clear stock 3031595 (local)  /  stock 398pcs",
     "Trong dòng có chữ 'stock' thì các mã 7 số sẽ được lấy làm mã tồn; số lượng viết liền pcs: 398pcs."),
    ("Đã xong", "Done  /  OK  /  Updated", "Dòng chỉ 1-3 từ, từ đầu là Done / OK / Updated."),
    ("Hủy", "Cancel", "Dòng chỉ có chữ Cancel."),
]

EXAMPLES_OK = [
    ("W40: Lanh will inform - W44", "Tuần 40, Lanh, dời MP sang W44"),
    ("9/23 update) est W45 clear stock", "Ngày 23/09, dự kiến W45, xử lý tồn kho"),
    ("W34: waiting for Quick Connect Dev, last order Nov-2026", "Tuần 34, chờ NCC, đơn cuối Nov-2026"),
    ("Finished by 31-Oct-26", "Hạn hoàn tất 31/10/2026"),
    ("8/26) no more demand in 2026, stock 398pcs", "Chưa có nhu cầu, tồn 398 pcs"),
    ("Sep 28 2026: wrote mail to a Hien change BOM", "Ngày 28/09, Hien, chờ giấy tờ"),
]
EXAMPLES_BAD = [
    ("w40: Lanh will infrom", "Gõ sai 'infrom' -> Chưa nhận dạng. Viết 'inform'."),
    ("FGS / SUBCONT / ASSM", "Đây là loại Buyer, ghi vào cột Buyer, không ghi vào Note."),
    ("Đang làm", "Không có tuần, ngày, người, từ khóa. Viết: W41: Lanh follow up."),
    ("W34 ~ 35 still having stock", "Dùng dấu : sau tuần: W34-35: still having stock."),
]


def build_md(path):
    L = ["# Quy tắc nhập ECN Log (để dashboard tự đọc)", "",
         "Mỗi ô văn bản tự do được dashboard tách thành thông tin có cấu trúc. "
         "Nhập đúng quy tắc thì không phải sửa dashboard; dòng không đọc được sẽ hiện ở trang 3 mục \"Chưa nhận dạng\".", "",
         "## 1. Quy tắc theo cột", "",
         "| Cột | Quy tắc | Đúng | Sai |", "|---|---|---|---|"]
    for c, r, ok, bad in COLUMN_RULES:
        L.append("| %s | %s | `%s` | %s |" % (c, r, ok, bad))
    L += ["", "## 2. Cú pháp cột Note", "",
          "Một dòng = một sự kiện. Dòng mới nhất nằm trên cùng. Mẫu: `[tuần hoặc ngày]: [người] [việc] [- Wxx]`.", "",
          "| Thông tin | Viết như | Cách đọc |", "|---|---|---|"]
    for a, b, c in NOTE_RULES:
        L.append("| %s | `%s` | %s |" % (a, b, c))
    L += ["", "### Từ khóa phân loại dòng Note (khớp từ trên xuống, dòng đầu tiên khớp thắng)", "",
          "| Phân loại | Dòng chứa một trong các cụm |", "|---|---|"]
    L.append("| Đã xong | Done / OK / Updated (dòng ≤ 3 từ) |")
    L.append("| Hủy | Cancel |")
    for name, kws in P.CATS:
        k = kws + ([v.lower() for v in P.VENDORS] if name == P.CAT_VENDOR else [])
        L.append("| %s | %s |" % (name, ", ".join(k)))
    L.append("| Chưa nhận dạng | không khớp cụm nào ở trên |")
    L += ["", "### Ví dụ đúng", ""]
    L += ["- `%s`  ->  %s" % e for e in EXAMPLES_OK]
    L += ["", "### Ví dụ chưa đọc được (lấy từ file thật)", ""]
    L += ["- `%s`  ->  %s" % e for e in EXAMPLES_BAD]
    L += ["", "## 3. Từ khóa tạo Tag và Vật tư (từ Description of Change)", "",
          "| Tag | Từ khóa |", "|---|---|"]
    for n, ws in P.TAGS:
        L.append("| %s | %s |" % (n, ", ".join(ws)))
    L += ["", "| Vật tư | Từ khóa (cụm ≤ 3 ký tự phải đứng riêng thành một từ) |", "|---|---|"]
    for n, ws in P.COMPONENTS:
        L.append("| %s | %s |" % (n, ", ".join(ws)))
    L += ["", "## 4. Động từ trong Description of Change", "",
          "- Bỏ: %s" % ", ".join(P.V_REMOVE), "- Thêm: %s" % ", ".join(P.V_ADD),
          "- Thay: %s" % ", ".join(P.V_REPLACE), "- Cập nhật: %s" % ", ".join(P.V_UPDATE),
          "- Nối cũ -> mới: %s" % ", ".join(P.V_CONNECT), "",
          "Mã đứng sau động từ Bỏ/Thay là mã cũ; sau Thêm hoặc sau từ nối (bằng, by, to...) là mã mới.", "",
          "## 5. Thêm người, NCC, từ khóa mới", "",
          "- Người và NCC: Power BI Desktop -> Transform data -> parameter `pPeople`, `pVendors` (cách nhau bằng dấu phẩy).",
          "- Từ khóa Tag / Vật tư / Phân loại: sửa danh sách ở đầu `ecn_parse.py` rồi chạy lại `build_ecn_pbip.py`, "
          "hoặc sửa trực tiếp hàm `fnTags`, `fnComps`, `fnNoteLine` trong Power Query.", ""]
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(L))


def build_xlsx(path):
    wb = openpyxl.Workbook()
    head = PatternFill("solid", fgColor="1E293B")
    thin = Side(style="thin", color="CBD5E1")

    def sheet(title, headers, rows, widths):
        ws = wb.create_sheet(title)
        ws.append(headers)
        for r in rows:
            ws.append(list(r))
        for c in ws[1]:
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = head
        for i, wd in enumerate(widths, 1):
            ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = wd
        for row in ws.iter_rows(min_row=1):
            for c in row:
                c.alignment = Alignment(wrap_text=True, vertical="top")
                c.border = Border(top=thin, bottom=thin, left=thin, right=thin)
        ws.freeze_panes = "A2"
        return ws

    wb.remove(wb.active)
    sheet("Quy tắc cột", ["Cột", "Quy tắc", "Ví dụ đúng", "Ví dụ sai"], COLUMN_RULES, [28, 80, 36, 36])
    sheet("Cú pháp Note", ["Thông tin", "Viết như", "Cách đọc"], NOTE_RULES, [22, 60, 80])
    cats = [("Đã xong", "Done / OK / Updated (dòng ≤ 3 từ)"), ("Hủy", "Cancel")]
    for name, kws in P.CATS:
        k = kws + ([v.lower() for v in P.VENDORS] if name == P.CAT_VENDOR else [])
        cats.append((name, ", ".join(k)))
    cats.append(("Chưa nhận dạng", "không khớp cụm nào ở trên"))
    sheet("Phân loại Note", ["Phân loại (từ trên xuống)", "Dòng chứa một trong các cụm"], cats, [30, 100])
    sheet("Tag & Vật tư", ["Nhóm", "Tên", "Từ khóa"],
          [("Tag", n, ", ".join(ws)) for n, ws in P.TAGS]
          + [("Vật tư", n, ", ".join(ws)) for n, ws in P.COMPONENTS], [12, 22, 110])
    sheet("Ví dụ", ["Loại", "Dòng Note", "Giải thích"],
          [("ĐÚNG",) + e for e in EXAMPLES_OK] + [("CHƯA ĐỌC ĐƯỢC",) + e for e in EXAMPLES_BAD],
          [18, 70, 70])
    ws = sheet("Danh mục", ["Người (pPeople)", "NCC (pVendors)", "Buyer", "Status", "EC Instruction"],
               [], [20, 26, 14, 14, 18])
    lists = [P.PEOPLE, P.VENDORS, ["Local", "Import", "In-house"],
             ["Ongoing", "Done", "Cancel"], ["Running Change", "Immediately"]]
    for i in range(max(len(x) for x in lists)):
        ws.append([x[i] if i < len(x) else None for x in lists])
    # a tiny entry template showing the dropdowns to copy into the real sheet
    t = wb.create_sheet("Mẫu nhập")
    t.append(["ECN Number", "Date", "FG", "Model /Customer No.", "Old item", "New item",
              "Buyer", "Status", "Description of Change", "EC Instruction", "Note"])
    t.append(["2026/QA/999", None, "1004150, 1004330", "9865-1301, 9865-1301VR", 1006982, 1011819,
              "Local", "Ongoing", "Thay graphic card 1006982 bằng 1011819", "Running Change",
              "W41: Lanh will inform - W44"])
    for c in t[1]:
        c.font = Font(bold=True, color="FFFFFF"); c.fill = head
    for i, wd in enumerate([16, 12, 26, 28, 12, 12, 12, 12, 48, 18, 44], 1):
        t.column_dimensions[openpyxl.utils.get_column_letter(i)].width = wd
    t["B2"].number_format = "dd/mm/yyyy"
    import datetime as dt
    t["B2"] = dt.date(2026, 10, 7)
    for rng, fml in [("G2:G500", '"Local,Import,In-house"'), ("H2:H500", '"Ongoing,Done,Cancel"'),
                     ("J2:J500", '"Running Change,Immediately"')]:
        dv = DataValidation(type="list", formula1=fml, allow_blank=True)
        t.add_data_validation(dv)
        dv.add(rng)
    wb.save(path)


build_md(os.path.join(ROOT, "ECN_Entry_Rules.md"))
build_xlsx(os.path.join(ROOT, "ECN_Entry_Rules.xlsx"))
print("rules written")
