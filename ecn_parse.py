# -*- coding: utf-8 -*-
"""
Cell parsers for the ECN Log, written so the Power Query (M) version in
build_ecn_pbip.py can follow the same steps one-to-one.

This module is BOTH the verification mirror and the single source of the
keyword dictionaries: the generator renders these lists into the M code.
"""
import datetime as dt
import collections
import warnings
import openpyxl

warnings.filterwarnings("ignore")

# ---------------------------------------------------------------- dictionaries
PEOPLE = ["Lanh", "Diep", "Hien", "Thao", "Nga", "Tra", "Cuong"]
VENDORS = ["Viet May", "Quick Connect Dev", "Vinh Vien"]

PKG = ["hộp", "thùng", "carton", "túi", "blister", "mút", "packaging", "color box",
       "graphic card", "card", "quai", "bvld", "mailer", "vỉ", "bao bì", "pe bag", "foam",
       "khay", "lót", "chèn"]
MECH = ["thân", "trimmer", "lưỡi", "blade", "vít", "screw", "nắp", "housing", "cơ khí",
        "lò xo", "spring", "gá", "khuôn", "linh kiện nhựa", "motor", "mô tơ", "đầu cắt",
        "comb", "lược", "dao", "ốc", "phôi", "jack", "đầu cắm", "cụm"]
PCB = ["mạch", "pcb", "bo mạch", "board", "chip", "dây điện", "wire", "pin", "battery",
       "adapter", "sạc", "charger", "led", "switch", "công tắc", "mach", "usb", "type c",
       "bóng đèn", "4v", "5v"]
LABEL = ["nhãn", "label", "tem", "laser", "sticker", "sách", "manual", "hướng dẫn",
         "in ấn", "bản vẽ", "barcode", "mã vạch", "decal", "logo", "print", "nội dung in",
         "phiên bản", "version", "warning", "warranty", "instruction", "ib #", "fill"]

# component = what physically changes (keywords <= 3 chars match a whole token)
COMPONENTS = [
    ("Hộp màu", ["hộp màu", "color box", "clor box", "hộp", "box"]),
    ("Thùng carton", ["thùng", "carton"]),
    ("Túi / pouch", ["túi", "pouch", "zip"]),
    ("Nhãn", ["nhãn", "label", "tem"]),
    ("Sách HD", ["sách", "book", "manual", "instruction"]),
    ("Graphic card", ["graphic card", "card"]),
    ("Blister / khay", ["blister", "khay"]),
    ("Mút / chèn", ["mút", "chèn", "foam"]),
    ("Bản mạch", ["mạch", "pcb", "mach"]),
    ("Thân / vỏ", ["thân", "vỏ", "housing"]),
    ("Dao / lưỡi", ["dao", "blade", "lưỡi"]),
    ("Trimmer", ["trimmer", "đầu tỉa"]),
    ("Pin / sạc", ["pin", "sạc", "battery", "4v", "5v", "usb"]),
    ("Laser / in", ["laser", "in xoay", "logo", "print"]),
    ("Chống ẩm", ["desiccant", "chống ẩm"]),
    ("Mô tơ", ["mô tơ", "motor"]),
]

# description verbs (matched as whole lower-case tokens)
V_REMOVE = ["bỏ", "remove", "romove", "hủy", "xóa", "scrap"]
V_ADD = ["thêm", "add", "bổ"]
V_REPLACE = ["thay", "replace", "rplace", "đổi", "change"]
V_UPDATE = ["update", "updete", "cập"]
V_CONNECT = ["bằng", "by", "thành", "to", "sang", "->", "→", "➔"]

# note-line categories, first match wins (see rules in ENTRY RULES doc)
DONE_WORDS = ["done", "ok", "updated"]
CATS = [
    ("Chưa có nhu cầu / đơn", ["no more demand", "không thấy có demand", "chưa có kh",
                               "chưa tạo mo", "no order", "không cần clear"]),
    ("Chờ Buyer / giấy tờ", ["not yet email", "buyer not action", "wrote mail", "đã làm giấy",
                             "buyer informed", "buyer noted", "buyer updated", "buyer est"]),
    ("Chờ khách / thiết kế", ["customer", "cus.", "design", "khách"]),
    ("Chờ NCC / hàng về", ["vendor", "delivery", "arrives", "last order", "quick connect",
                           "viet may", "deliver"]),
    ("Xử lý tồn kho", ["stock", "tồn", "scrap", "clear", "wh ", "inventory", "keep old",
                       "available"]),
    ("Setup / BOM", ["setup", "bom", "mo ", "level"]),
    ("Hạn hoàn tất", ["finished by"]),
    ("Theo dõi", ["follow", "recheck", "check", "inform", "update", "est", "wait", "will change"]),
]
CAT_DONE = "Đã xong"
CAT_CANCEL = "Hủy"
CAT_UNKNOWN = "Chưa nhận dạng"
CAT_VENDOR = "Chờ NCC / hàng về"

TAGS = [("Bao bì", PKG), ("Cơ khí", MECH), ("Mạch", PCB), ("Nhãn", LABEL)]
DELIMS = [" ", ",", ";", ":", "(", ")", "[", "]", "/", "&", "\n", "\r", "\t", "!", "?", "="]


# ---------------------------------------------------------------- tokens
def tokens(s):
    if not s:
        return []
    for d in DELIMS:
        s = s.replace(d, "\x00")
    out = []
    for p in s.split("\x00"):
        p = p.strip().rstrip(".-")
        if p:
            out.append(p)
    return out


def is_digits(x):
    return x != "" and all(c in "0123456789" for c in x)


def code_kind(tok):
    if tok.startswith("#"):
        b = tok[1:]
        return "short" if is_digits(b) and 2 <= len(b) <= 5 else None
    if is_digits(tok):
        return "item" if len(tok) == 7 and tok[0] in "13" else None
    if "-" in tok:
        p = tok.split("-")
        if len(p) == 2 and is_digits(p[0]) and is_digits(p[1]) and 4 <= len(p[0]) <= 5:
            return "part"
    return None


def tags_of(desc):
    d = (desc or "").lower()
    return [n for n, ws in TAGS if any(w in d for w in ws)]


def components_of(desc):
    d = (desc or "").lower()
    toks = [t.lower() for t in tokens(desc)]
    out = []
    for label, kws in COMPONENTS:
        for k in kws:
            hit = (k in toks) if len(k) <= 3 else (k in d)
            if hit:
                out.append(label)
                break
    return out


def parse_desc(desc):
    """old / new / ref codes + verbs from the free-text Description of Change."""
    mode, old, new, ref, act = "ref", [], [], [], []

    def add(a):
        if a not in act:
            act.append(a)

    for tok in tokens(desc):
        l = tok.lower()
        if l in V_REMOVE:
            mode = "old"; add("Bỏ")
        elif l in V_ADD:
            mode = "new"; add("Thêm")
        elif l in V_REPLACE:
            mode = "old"; add("Thay")
        elif l in V_CONNECT:
            if not old and ref:
                old, ref = ref, []
            mode = "new"
        elif l in V_UPDATE:
            mode = "ref"; add("Cập nhật")
        elif code_kind(tok):
            lst = old if mode == "old" else new if mode == "new" else ref
            if tok not in lst:
                lst.append(tok)
    return dict(old=old, new=new, ref=ref, act=act)


# ---------------------------------------------------------------- notes
def iso_week(d):
    return d.isocalendar()[1]


def plan_wk(line):
    marker = " - W"
    if marker not in line:
        return None
    tail = line.rsplit(marker, 1)[1].strip()
    digits = "".join(c for c in tail if c in "0123456789")
    if tail and digits == tail and len(tail) <= 2:
        return int(digits)
    return None


def est_wk(ltoks):
    idx = [i for i, t in enumerate(ltoks) if t.startswith("est")]
    if idx:
        start = idx[0] + 1
    elif "change" in ltoks:          # "Week 34 will change" has no "est"
        start = 0
    else:
        return None, None
    win = ltoks[start: start + 5]
    wk = None
    for j, t in enumerate(win):
        if len(t) >= 2 and t[0] == "w" and is_digits(t[1:]):
            wk = int(t[1:]); break
        if t == "week" and j + 1 < len(win) and is_digits(win[j + 1]):
            wk = int(win[j + 1]); break
    if wk is not None:
        return wk, None
    yr = next((t for t in win if is_digits(t) and len(t) == 4), None)
    if yr is None:
        return None, None
    return None, ("cuối " if "end" in win else "") + yr


def parse_note_line(line, today):
    s0 = line.strip()
    s1 = s0[4:].strip() if s0.upper().startswith("OBS:") else s0
    colon = s1.find(":")
    head = s1[:colon].strip().upper() if 0 < colon <= 12 else ""
    wk_from = None
    if head.startswith("W") and head[1:] and all(c in "0123456789-~ " for c in head[1:]) \
            and any(c in "0123456789" for c in head):
        nums = [int(x) for x in head[1:].replace("~", "-").replace(" ", "").split("-") if x]
        wk_from = nums[-1] if nums else None
    s2 = s1[colon + 1:].strip() if wk_from is not None else s1
    p1, p2 = s2.find(")"), s2.find(":")
    end_p = p1 if 0 <= p1 <= 14 else p2 if 0 <= p2 <= 14 else -1
    dtok = s2[:end_p].strip() if end_p > 0 else ""
    dfirst = dtok.split(" ")[0] if dtok else ""
    dparts = dfirst.split("/")
    date_md = None
    if dfirst and len(dparts) == 2 and is_digits(dparts[0]) and is_digits(dparts[1]) \
            and 1 <= int(dparts[0]) <= 12 and 1 <= int(dparts[1]) <= 31:
        try:
            date_md = dt.date(today.year, int(dparts[0]), int(dparts[1]))
        except ValueError:
            date_md = None
    date_nm = None
    if date_md is None and dtok:
        for f in ("%b %d %Y", "%b %d, %Y"):
            try:
                date_nm = dt.datetime.strptime(dtok, f).date(); break
            except ValueError:
                pass
    date_from = date_md or date_nm
    s3 = s2[end_p + 1:].strip() if date_from is not None else s2
    lt = s3.lower()
    toks = tokens(s3)
    ltoks = [t.lower() for t in toks]
    who = [p for p in PEOPLE if p in toks]
    vend = [v for v in VENDORS if v.lower() in lt]
    ewk, etxt = est_wk(ltoks)
    finish = None
    i = lt.find("finished by ")
    if i >= 0:
        try:
            finish = dt.datetime.strptime(s3[i + 12:].strip(), "%d-%b-%y").date()
        except ValueError:
            finish = None
    last_order = None
    i = lt.find("last order ")
    if i >= 0:
        last_order = (s3[i + 11:] + ",").split(",")[0].strip()
    stock_codes = [t for t in toks if code_kind(t) == "item"] if "stock" in lt else []
    qty = next((t[:-3] for t in ltoks if t.endswith("pcs") and is_digits(t[:-3])), None)
    # category
    if ltoks and len(ltoks) <= 3 and ltoks[0] in DONE_WORDS:
        cat = CAT_DONE
    elif lt == "cancel":
        cat = CAT_CANCEL
    else:
        cat = CAT_UNKNOWN
        vl = [v.lower() for v in VENDORS]
        for label, kws in CATS:
            kk = kws + vl if label == CAT_VENDOR else kws
            if any(k in lt for k in kk):
                cat = label
                break
    if date_from is not None:
        key = date_from.toordinal()
    elif wk_from is not None:
        key = (dt.date(today.year, 1, 1) + dt.timedelta(days=(wk_from - 1) * 7)).toordinal()
    else:
        key = -1
    return dict(raw=s0, text=s3, wk_from=wk_from, date_from=date_from,
                who=", ".join(who), vendor=", ".join(vend), target=plan_wk(s3),
                est_wk=ewk, est_txt=etxt, finish=finish, last_order=last_order,
                stock=", ".join(dict.fromkeys(stock_codes)), qty=qty, cat=cat, key=key)


# ---------------------------------------------------------------- sheet
def txt(v):
    if v is None:
        return None
    s = str(v).strip()
    return s or None


def to_date(v):
    return v.date() if isinstance(v, dt.datetime) else None


def week_no(t):
    if t is None:
        return None
    s = str(t).strip().upper()
    return int(s[1:]) if s.startswith("W") and s[1:].isdigit() else None


def eff_wk(wk, twk):
    if wk is None:
        return None
    return wk + 53 if (wk <= 12 and twk >= 40) else wk


def _norm(h):
    return " ".join(str(h).replace("\n", " ").split()) if h is not None else ""


def load(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["ECN"]
    hdr_i, hdr = None, None
    for i, r in enumerate(ws.iter_rows(min_row=1, max_row=10, values_only=True), 1):
        if r and r[0] == "ECN Number":
            hdr_i, hdr = i, r
            break
    idx = {}
    for i, v in enumerate(hdr):
        idx.setdefault(_norm(v), i)

    def g(r, name):
        i = idx.get(name)
        return r[i] if i is not None and i < len(r) else None

    lines = []
    for r in ws.iter_rows(min_row=hdr_i + 1, max_col=max(idx.values()) + 1, values_only=True):
        if not txt(g(r, "ECN Number")):
            continue
        est = g(r, "New item (est MP week)")
        est = None if (est is None or str(est).startswith("#")) else txt(est)
        yr = g(r, "New item (est MP year)")
        try:
            yr = int(str(yr).strip()[:4]) if yr not in (None, "") and not str(yr).startswith("#") else None
        except ValueError:
            yr = None
        lines.append(dict(ecn=txt(g(r, "ECN Number")),
                          date=to_date(g(r, "Date")) or to_date(g(r, "Date Issued")),
                          fg=txt(g(r, "FG")), model=txt(g(r, "Model /Customer No.")),
                          old=txt(g(r, "Old item")), new=txt(g(r, "New item")),
                          buyer=txt(g(r, "Buyer (Local/ Pur/ Ser/ Inhouse)")),
                          status=txt(g(r, "Status")) or "", desc=txt(g(r, "Description of Change")),
                          instr=txt(g(r, "EC Instruction")), done=to_date(g(r, "Date Completed")),
                          sent=txt(g(r, "ECN form sent date")), recv=txt(g(r, "ECN form received date")),
                          est=est, est_year=yr, note=txt(g(r, "Note"))))
    return lines


def split_list(s):
    if not s:
        return []
    for d in [";", "\n", "\r"]:
        s = s.replace(d, ",")
    out = []
    for p in s.split(","):
        p = p.strip()
        if p and p not in out:
            out.append(p)
    return out


def build(lines, today):
    twk = iso_week(today)
    g = collections.OrderedDict()
    for l in lines:
        g.setdefault(l["ecn"], []).append(l)
    notes_all, ecn_rows, line_rows = [], [], []
    for k, rs in g.items():
        # ---- note lines (all distinct cells, all distinct lines)
        cells = list(dict.fromkeys(r["note"] for r in rs if r["note"]))
        seen, nl = set(), []
        for c in cells:
            for x in c.replace("\r", "").split("\n"):
                x = x.strip()
                if x and x not in seen:
                    seen.add(x); nl.append(x)
        notes = []
        for i, x in enumerate(nl, 1):
            n = parse_note_line(x, today); n["ecn"] = k; n["seq"] = i
            notes.append(n)
        notes_all += notes
        ns = sorted(notes, key=lambda n: (-n["key"], n["seq"]))
        active = [n for n in ns if n["cat"] != CAT_DONE]
        top = active[0] if active else (ns[0] if ns else None)

        # ---- status
        stats = [r["status"] for r in rs]
        act = [s for s in stats if s != "Cancel"]
        if not act:
            st = "Cancel"
        elif "Ongoing" in act or ("Done" in act and "" in act):
            st = "Ongoing"
        elif "" in act:
            st = "Chưa bắt đầu"
        else:
            st = "Done"
        openx = st in ("Ongoing", "Chưa bắt đầu")
        dates = [r["date"] for r in rs if r["date"]]
        issue = max(dates) if dates else None

        # ---- schedule
        ests = [week_no(r["est"]) for r in rs if week_no(r["est"])]
        est = min(ests) if ests else None
        targets = [n["target"] if n["target"] is not None else n["est_wk"] for n in ns]
        targets = [t for t in targets if t is not None]
        orig = targets[-1] if targets else None
        cur = est if est else (targets[0] if targets else None)
        moved = orig is not None and cur is not None and orig != cur
        reason = None
        tl = next((n for n in ns if n["target"] is not None), None)
        if tl:
            reason = tl["text"].rsplit(" - W", 1)[0][:70]
        ecur = eff_wk(cur, twk)
        if st == "Done":
            state = "ĐÃ HOÀN TẤT"
        elif st == "Cancel":
            state = "ĐÃ HỦY"
        elif moved:
            state = "DỜI LỊCH"
        elif ecur is not None and ecur < twk:
            state = "QUÁ HẠN"
        elif cur is not None:
            state = "ĐÚNG LỊCH"
        else:
            state = "CHƯA CÓ LỊCH"
        recent = issue is not None and (today - issue).days <= 6
        near = ecur is not None and twk <= ecur <= twk + 1
        if not openx:
            scope = "Đã xong / Hủy"
        elif recent or near or state in ("DỜI LỊCH", "QUÁ HẠN"):
            scope = "Áp dụng hôm nay"
        else:
            scope = "Đang mở khác"

        # ---- structured line mix
        nrep = sum(1 for r in rs if r["old"] and r["new"])
        nadd = sum(1 for r in rs if r["new"] and not r["old"])
        nrem = sum(1 for r in rs if r["old"] and not r["new"])
        # ---- description parsing
        comps, tg, actions = [], [], []
        old_c, new_c = [], []
        for r in rs:
            pd = parse_desc(r["desc"])
            for c in components_of(r["desc"]):
                if c not in comps: comps.append(c)
            for t in tags_of(r["desc"]):
                if t not in tg: tg.append(t)
            for a in pd["act"]:
                if a not in actions: actions.append(a)
            for c in pd["old"]:
                if c not in old_c: old_c.append(c)
            for c in pd["new"]:
                if c not in new_c: new_c.append(c)
        models = []
        fgs = []
        for r in rs:
            for m in split_list(r["model"]):
                if m not in models: models.append(m)
            for f in split_list(r["fg"]):
                if f not in fgs: fgs.append(f)
        buyers = collections.Counter(r["buyer"] for r in rs if r["buyer"])
        owners = list(dict.fromkeys(p for n in notes for p in n["who"].split(", ") if p))
        stock = list(dict.fromkeys(c for n in notes for c in n["stock"].split(", ") if c))
        finish = max([n["finish"] for n in notes if n["finish"]], default=None)
        last_order = next((n["last_order"] for n in ns if n["last_order"]), None)
        qty = next((n["qty"] for n in ns if n["qty"]), None)
        sent = any(r["sent"] for r in rs)
        recvd = any(r["recv"] for r in rs)
        comp_dates = [r["done"] for r in rs if r["done"]]
        instrs = [r["instr"] for r in rs if r["instr"]]
        instr = "Immediately" if "Immediately" in instrs else (instrs[0] if instrs else "—")
        ecn_rows.append(dict(
            ecn=k, issue=issue, status=st, state=state, scope=scope, est=est, cur=cur, orig=orig,
            moved=moved, reason=reason, instr=instr, open=openx,
            first_run=openx and ecur == twk, n=len(rs), tags=tg, comps=comps, actions=actions,
            mix=(nrep, nadd, nrem), models=models, fgs=fgs, buyers=dict(buyers),
            blocker=top["cat"] if top else None, blocker_text=top["text"][:110] if top else None,
            updated=(top["date_from"].strftime("%d/%m") if top and top["date_from"]
                     else ("W%d" % top["wk_from"] if top and top["wk_from"] else None)),
            owners=owners, stock=stock, qty=qty, finish=finish, last_order=last_order,
            est_note=next((n["est_wk"] or n["est_txt"] for n in ns if n["est_wk"] or n["est_txt"]), None),
            paper="Đã nhận" if recvd else ("Đã gửi" if sent else "Chưa gửi"),
            completed=max(comp_dates) if comp_dates else None,
            old_desc=old_c, new_desc=new_c))
    return ecn_rows, notes_all


if __name__ == "__main__":
    today = dt.date(2026, 10, 7)
    ecn, notes = build(load("ECN Log.xlsx"), today)
    print("ECN", len(ecn), "note lines", len(notes))
    print(collections.Counter(e["scope"] for e in ecn))
    print(collections.Counter(e["state"] for e in ecn))
    print("moved:", [(e["ecn"], e["orig"], e["cur"], e["reason"], e["scope"]) for e in ecn if e["moved"]])
    print("note categories:", collections.Counter(n["cat"] for n in notes))
    print("unrecognised:", [n["raw"][:70] for n in notes if n["cat"] == CAT_UNKNOWN])
    print("owners:", collections.Counter(p for e in ecn for p in e["owners"]))
    print("finish:", [(e["ecn"], str(e["finish"])) for e in ecn if e["finish"]][:5])
    print("est from note only:", sum(1 for e in ecn if e["est_note"] and not e["est"]))
    print("components:", collections.Counter(c for e in ecn for c in e["comps"]).most_common())
    print("blockers:", collections.Counter(e["blocker"] for e in ecn if e["open"]))
    print("paper:", collections.Counter(e["paper"] for e in ecn))


# ---------------------------------------------------------------- weekly qty
# Columns V..BM of the ECN sheet ("Old item (qty) + OPOs", "Old item (qty)2"...)
# hold the projected balance of the OLD item at the end of each week; row 3 has
# the week number and row 2 the Monday of that week. A drop in the balance is
# what production used that week; a balance at or below 0 means the old item is
# finished and the line must run the new item.
def to_week_start(v):
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, (int, float)) and 30000 < v < 80000:
        return (dt.datetime(1899, 12, 30) + dt.timedelta(days=int(v))).date()
    return None


def load_weeks(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["ECN"]
    grid = list(ws.iter_rows(min_row=1, max_row=10, values_only=True))
    h = next(i for i, r in enumerate(grid) if r[0] == "ECN Number")
    wk_row, dt_row = grid[h - 1], grid[h - 2]
    cols = []
    for i in range(len(grid[h])):
        w, d = wk_row[i], to_week_start(dt_row[i])
        if grid[h][i] is not None and isinstance(w, (int, float)) and 1 <= w <= 53 and w == int(w) and d:
            cols.append((i, int(w), d))
    out = []
    line_no = 0
    for r in ws.iter_rows(min_row=h + 2, values_only=True):
        if not txt(r[0]):
            continue
        line_no += 1
        for i, w, d in cols:
            v = r[i] if i < len(r) else None
            if isinstance(v, bool) or not isinstance(v, (int, float)):
                continue
            out.append(dict(line=line_no, ecn=txt(r[0]), old=txt(r[5]), new=txt(r[6]),
                            week=w, start=d, bal=float(v)))
    return out


def week_rel(start, today):
    monday = today - dt.timedelta(days=today.weekday())
    d = round((start - monday).days / 7)
    if d < -1:
        return "Đã qua", d
    return {-1: "Tuần trước", 0: "Tuần này", 1: "Tuần sau"}.get(d, "+%d tuần" % d), d


def calc_weeks(rows, today):
    by = collections.OrderedDict()
    for r in rows:
        by.setdefault(r["line"], []).append(r)
    out = []
    for line, rs in by.items():
        rs.sort(key=lambda r: r["start"])
        if all(r["bal"] == 0 for r in rs):     # item not in the MRP sheet: SUMIF gives 0
            continue
        for i, r in enumerate(rs):
            begin = rs[i - 1]["bal"] if i else r["bal"]
            demand = max(begin - r["bal"], 0.0)
            old_used = min(demand, max(begin, 0.0))
            new_used = demand - old_used
            if demand == 0:
                mode = None
            elif new_used == 0:
                mode = "Dùng mã cũ"
            elif r["new"] is None:
                mode = "Hết tồn (bỏ mã)"
            elif r["new"] == r["old"]:
                mode = "Thiếu tồn (giữ mã)"
            elif old_used == 0:
                mode = "Chạy mã mới"
            else:
                mode = "Chuyển mã trong tuần"
            rel, d = week_rel(r["start"], today)
            out.append(dict(r, begin=begin, demand=demand, old_used=old_used,
                            new_used=new_used, mode=mode, rel=rel, rel_n=d))
    return out


def runout(weeks):
    """first week whose balance is <= 0, per line."""
    first = {}
    for r in sorted(weeks, key=lambda r: (r["line"], r["start"])):
        if r["bal"] <= 0 and r["line"] not in first:
            first[r["line"]] = r
    return first


def ecn_week_summary(weeks, today):
    """per ECN: demand this week (deduped by old item) and the old-item run-out label."""
    monday = today - dt.timedelta(days=today.weekday())
    by = collections.defaultdict(list)
    for r in weeks:
        if r["old"]:
            by[r["ecn"]].append(r)
    out = {}
    for ecn, rs in by.items():
        now = {}
        for r in rs:
            if r["rel_n"] == 0:
                now[(r["old"], r["demand"])] = r["demand"]
        qty = sum(now.values()) if now else None
        lines = collections.defaultdict(list)
        for r in rs:
            lines[r["line"]].append(r)
        fut, since, keeps, last = [], [], [], max(r["start"] for r in rs)
        for ln, lr in lines.items():
            st = line_stock(sorted(lr, key=lambda r: r["start"]), monday)
            if st["fut"]:
                fut.append(st["fut"])
            elif st["since"]:
                since.append(st["since"])
            else:
                keeps.append(st["last"])
        if fut:
            d = min(fut)
            n = (d - monday).days // 7
            label = "Hết mã cũ tuần này" if n == 0 else "Hết mã cũ W%d · còn %d tuần" % (iso_week(d), n)
        elif keeps and since:
            label = "Đã hết %d/%d mã cũ · %d mã còn tồn" % (len(since), len(since) + len(keeps), len(keeps))
        elif keeps:
            label = "Còn tồn mã cũ sau W%d" % iso_week(max(keeps))
        else:
            label = "Đã hết mã cũ từ W%d" % iso_week(max(since))
        out[ecn] = dict(qty_now=qty, runout=label)
    return out


def line_stock(recs, monday):
    """recs sorted by week. fut = first week from now with balance <= 0 while the
    old item is still in stock now; since = start of the current <= 0 streak."""
    pos = [i for i, r in enumerate(recs) if r["start"] <= monday]
    now = pos[-1] if pos else 0
    now_bal = recs[now]["bal"]
    if now_bal > 0:
        f = [r for r in recs if r["start"] >= monday and r["bal"] <= 0]
        return dict(fut=f[0]["start"] if f else None, since=None, last=recs[-1]["start"])
    up = [i for i in range(now + 1) if recs[i]["bal"] > 0]
    return dict(fut=None, since=recs[(up[-1] + 1) if up else 0]["start"], last=recs[-1]["start"])


# ---------------------------------------------------------------- v5: schedule week from Note
# Rule from the user: every line of the Note cell is a new schedule, newest on TOP,
# the first characters "Wnn" are the week. A line without W-prefix falls back to
# est/target week in the text, then to the ISO week of its date.
def line_week(n):
    if n["wk_from"] is not None:
        return n["wk_from"]
    if n["est_wk"] is not None:
        return n["est_wk"]
    if n["target"] is not None:
        return n["target"]
    if n["date_from"] is not None:
        return iso_week(n["date_from"])
    return None


REL_ORDER = {"Tuần này": 1, "Tuần sau": 2, "Sắp tới": 3, "Quá hạn": 4, "Chưa có lịch": 5, "Đã xong": 6}


def week_view(lines, today):
    twk = iso_week(today)
    g = collections.OrderedDict()
    for l in lines:
        g.setdefault(l["ecn"], []).append(l)
    out = {}
    for k, rs in g.items():
        cells = list(dict.fromkeys(r["note"] for r in rs if r["note"]))
        seen, rows = set(), []
        for ci, c in enumerate(cells, 1):
            for pi, x in enumerate([y.strip() for y in c.replace("\r", "").split("\n") if y.strip()], 1):
                if x in seen:
                    continue
                seen.add(x)
                n = parse_note_line(x, today)
                n["cell"], n["pos"], n["line"] = ci, pi, x
                n["lwk"] = line_week(n)
                rows.append(n)
        top = {}
        for n in sorted(rows, key=lambda n: (n["cell"], n["pos"])):
            if n["lwk"] is not None and n["cell"] not in top:
                top[n["cell"]] = n["lwk"]
        eff = [eff_wk(w, twk) for w in top.values()]
        est = [week_no(r["est"]) for r in rs if week_no(r["est"])]
        est = min(est) if est else None
        cur = max(eff) if eff else eff_wk(est, twk)
        pfx = {eff_wk(n["wk_from"], twk) for n in rows if n["wk_from"] is not None}
        moves = len(pfx) - 1 if len(pfx) > 1 else 0
        runno = None if cur is None else (cur - 53 if cur > 53 else cur)
        d = None if cur is None else cur - twk
        stats = [r["status"] for r in rs]
        act = [x for x in stats if x != "Cancel"]
        is_open = bool(act) and not all(x == "Done" for x in act)
        rel = ("Đã xong" if not is_open else "Chưa có lịch" if d is None else "Quá hạn" if d < 0
               else "Tuần này" if d == 0 else "Tuần sau" if d == 1 else "Sắp tới")
        ordr = sorted(rows, key=lambda n: (n["cell"], n["pos"]))
        out[k] = dict(run="W%02d" % runno if runno else "—", rel=rel, moves=moves,
                      note_raw="  |  ".join(n["line"] for n in ordr) or "—")
    return out


# ---------------------------------------------------------------- v6: week from Date + moves from Note
# User rule: the week an ECN runs = ISO week of its Date. A Note line that has a week
# ("W51", "week 51") AND "change" is a schedule move (dời lịch) to that week; every
# other line is just an event in the history.
CHANGE_WORDS = ["change", "dời", "đổi"]


def week_tokens(ltoks):
    out = []
    n = len(ltoks)
    for i, t in enumerate(ltoks):
        body = t[1:].replace("~", "-")
        if len(t) >= 2 and t[0] == "w" and body and all(c in "0123456789-" for c in body) \
                and any(c in "0123456789" for c in body):
            parts = [x for x in body.split("-") if x]
            if parts and is_digits(parts[-1]):
                out.append(int(parts[-1]))
                continue
        if t in ("week", "wk", "w", "tuần", "tuan") and i + 1 < n and is_digits(ltoks[i + 1]) \
                and len(ltoks[i + 1]) <= 2:
            out.append(int(ltoks[i + 1]))
    return out


def move_info(text):
    """-> (is_move, move_week or None, year or None). text = note line without its W../date prefix."""
    lt = text.lower()
    ltoks = [t.lower() for t in tokens(text)]
    chg = any(w in lt for w in CHANGE_WORDS)
    wk = week_tokens(ltoks)
    yr = [t for t in ltoks if is_digits(t) and len(t) == 4 and t.startswith("20")]
    is_move = chg and (bool(wk) or bool(yr))
    return is_move, (wk[-1] if chg and wk else None), (yr[-1] if chg and yr else None)


def iso_year(d):
    return d.isocalendar()[0]


def week_start_of(wk, today):
    y = iso_year(today)
    tw = iso_week(today)
    if wk <= 12 and tw >= 40:
        y += 1
    elif wk >= 40 and tw <= 12:
        y -= 1
    jan4 = dt.date(y, 1, 4)
    return jan4 - dt.timedelta(days=jan4.weekday()) + dt.timedelta(days=(wk - 1) * 7)


REL6 = {"Tuần này": 1, "Tuần sau": 2, "Sắp tới": 3, "Tuần cũ chưa xong": 4, "Chưa có lịch": 5, "Đã xong": 6}


def week_view6(lines, today):
    monday = today - dt.timedelta(days=today.weekday())
    g = collections.OrderedDict()
    for l in lines:
        g.setdefault(l["ecn"], []).append(l)
    out = {}
    for k, rs in g.items():
        dates = [r["date"] for r in rs if r["date"]]
        issue = max(dates) if dates else None
        plan = issue - dt.timedelta(days=issue.weekday()) if issue else None
        cells = list(dict.fromkeys(r["note"] for r in rs if r["note"]))
        seen, ev = set(), []
        for ci, c in enumerate(cells, 1):
            for pi, x in enumerate([y.strip() for y in c.replace("\r", "").split("\n") if y.strip()], 1):
                if x in seen:
                    continue
                seen.add(x)
                n = parse_note_line(x, today)
                im, mw, my = move_info(n["text"])
                ev.append(dict(cell=ci, pos=pi, line=x, is_move=im, move_wk=mw, move_yr=my, note=n))
        ev.sort(key=lambda e: (e["cell"], e["pos"]))
        mv = [e for e in ev if e["is_move"]]
        mwk = [e["move_wk"] for e in mv if e["move_wk"] is not None]
        cur = week_start_of(mwk[0], today) if mwk else plan
        stats = [r["status"] for r in rs]
        act = [x for x in stats if x != "Cancel"]
        is_open = bool(act) and not all(x == "Done" for x in act)
        d = None if cur is None else (cur - monday).days // 7
        rel = ("Đã xong" if not is_open else "Chưa có lịch" if d is None else "Tuần cũ chưa xong" if d < 0
               else "Tuần này" if d == 0 else "Tuần sau" if d == 1 else "Sắp tới")
        chain = ["W%02d" % iso_week(plan)] if plan else []
        chain += ["W%02d" % w for w in reversed(mwk)]
        out[k] = dict(run="W%02d" % iso_week(cur) if cur else "—", rel=rel, moves=len(mv),
                      hist=" → ".join(chain) if chain else "—", plan=plan, cur=cur, open=is_open,
                      note_raw="  |  ".join(e["line"] for e in ev) or "—", events=ev)
    return out


# ---------------------------------------------------------------- v10: est MP week + year (columns R, S) first
def week_start_y(wk, yr):
    jan4 = dt.date(yr, 1, 4)
    return jan4 - dt.timedelta(days=jan4.weekday()) + dt.timedelta(days=(wk - 1) * 7)


def week_view7(lines, today):
    """est MP week + year (R, S) > Note move (Wxx/week xx + change) > Date; year only => future year."""
    base = week_view6(lines, today)
    monday = today - dt.timedelta(days=today.weekday())
    g = collections.OrderedDict()
    for l in lines:
        g.setdefault(l["ecn"], []).append(l)
    for k, rs in g.items():
        v = base[k]
        starts, years = [], []
        for r in rs:
            wk = None
            if r.get("est"):
                toks = week_tokens([t.lower() for t in tokens(r["est"])])
                wk = toks[-1] if toks else None
            yr = r.get("est_year")
            if wk is not None:
                starts.append(week_start_y(wk, yr) if yr else week_start_of(wk, today))
            elif yr:
                years.append(yr)
        plan = v["plan"]
        note_cur = v["cur"] if v["moves"] and v["cur"] != plan else None
        est_start = min(starts) if starts else None
        year_start = dt.date(min(years), 1, 1) if years else None
        year_only = (est_start is None and note_cur is None and year_start is not None
                     and (plan is None or year_start.year > plan.year))
        cur = est_start or note_cur or (year_start if year_only else plan)
        d = None if cur is None else (cur - monday).days // 7
        rel = ("Đã xong" if not v["open"] else "Chưa có lịch" if d is None else "Tuần cũ chưa xong" if d < 0
               else "Tuần này" if d == 0 else "Tuần sau" if d == 1 else "Sắp tới")
        nums = ([iso_week(plan)] if plan else [])
        nums += [e["move_wk"] for e in reversed(v["events"]) if e["is_move"] and e["move_wk"] is not None]
        if cur and not year_only:
            nums.append(iso_week(cur))
        chain = []
        for n in nums:
            if not chain or chain[-1] != n:
                chain.append(n)
        txt_chain = ["W%02d" % n for n in chain] + ([str(cur.year)] if year_only else [])
        v.update(cur=cur, rel=rel,
                 run=(str(cur.year) if year_only else "W%02d" % iso_week(cur)) if cur else "—",
                 hist=" → ".join(txt_chain) if len(txt_chain) > 1 else ("Đúng kế hoạch " + txt_chain[0] if txt_chain else "—"),
                 year_only=year_only)
    return base
