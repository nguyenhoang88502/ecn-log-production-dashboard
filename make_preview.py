# -*- coding: utf-8 -*-
"""Layout preview (NOT a Power BI screenshot): same positions / sizes / data as the report."""
import datetime as dt, collections, statistics, html, os, sys, subprocess, base64
import ecn_parse as p

ROOT = os.path.dirname(os.path.abspath(__file__))
TODAY = dt.date.fromisoformat(os.environ.get("ECN_PREVIEW_TODAY", "2026-10-13"))
ONLY_LOG = os.environ.get("ECN_PREVIEW_ONLY_LOG") == "1"
lines = p.load(os.path.join(ROOT, "ECN Log.xlsx"))
ecn, _ = p.build(lines, TODAY)
E = {e["ecn"]: e for e in ecn}
WV = p.week_view7(lines, TODAY)
by = collections.defaultdict(list)
for l in lines:
    by[l["ecn"]].append(l)
logo = base64.b64encode(open(os.path.join(ROOT, "assets/img/wahl-logo.png"), "rb").read()).decode()
esc = html.escape


def change_text(k):
    seen, out = [], []
    for r in by[k]:
        ch = ("%s → %s" % (r["old"], r["new"]) if r["old"] and r["new"] else
              "%s → BỎ" % r["old"] if r["old"] else "THÊM → %s" % r["new"] if r["new"] else "—")
        t = (r["desc"][:70] + " (" + ch + ")") if r["desc"] and ch != "—" else (r["desc"][:70] if r["desc"] else ch)
        if t not in seen:
            seen.append(t)
    a, b, c = E[k]["mix"]
    mix = [("%d thay" % a) if a else "", ("%d thêm" % b) if b else "", ("%d bỏ" % c) if c else ""]
    mix = " · ".join(x for x in mix if x)
    n = len(by[k])
    txt = ("[" + mix + "]  " if n > 1 and mix else "") + "  |  ".join(seen[:3])
    return txt + ("  |  +%d thay đổi khác" % (len(seen) - 3) if len(seen) > 3 else "")


def short(l):
    return ", ".join(l[:2]) + (" +%d" % (len(l) - 2) if len(l) > 2 else "") if l else "—"


CSS = """
body{margin:0;width:1920px;height:1080px;background:#F3F5F9;font-family:'Segoe UI',Arial,sans-serif;color:#1C2430;position:relative;overflow:hidden}
.a{position:absolute;box-sizing:border-box}
.box{background:#fff;border:1px solid #DDE2EA;border-radius:10px}
.band{left:0;top:0;width:1920px;height:96px;background:#3F4246}
.strip{left:0;top:96px;width:1920px;height:6px;background:#F2C230}
h1{margin:0;font-size:35px;color:#fff;font-weight:700}
.sub{font-size:16px;color:#D5D8DD}
.lbl{font-size:15px;font-weight:700;padding:6px 12px}
.chip{display:inline-block;border:1px solid #bbb;padding:7px 12px;margin:2px;font-size:16px;background:#fff}
.chip.on{background:#3F4246;color:#fff}
.kpi{text-align:center}.kpi b{display:block;font-size:46px;margin-top:14px}.kpi span{font-size:17px;color:#5B6575}
table{border-collapse:collapse;width:100%;table-layout:fixed}
th{background:#3F4246;color:#fff;font-size:15px;text-align:left;padding:10px 8px}
td{font-size:inherit;padding:9px 8px;border-bottom:1px solid #DDE2EA;vertical-align:top;word-wrap:break-word}
tr:nth-child(even) td{background:#F6F8FB}
td.k{color:#3F4246;font-size:19px;font-weight:700}
.imm{background:#FBD5D1 !important}
.t{font-size:16px;font-weight:700;padding:10px 14px}
"""


def head(title, sub):
    return ('<div class="a band"></div><div class="a strip"></div>'
            '<div class="a" style="left:24px;top:10px"><h1>%s</h1><div class="sub">%s</div></div>'
            '<div class="a" style="left:1170px;top:30px;font-size:22px;font-weight:700;color:#F2C230">%s  ·  W%d</div>'
            '<img class="a" style="left:1660px;top:14px;height:68px" src="data:image/png;base64,%s">'
            % (title, sub, TODAY.strftime("%A %d/%m/%Y"), p.iso_week(TODAY), logo))


def slicers(week_on, note):
    chips = "".join('<span class="chip %s">%s</span>' % ("on" if c in week_on else "", c)
                    for c in ["Tuần này", "Tuần sau", "Sắp tới", "Tuần cũ chưa xong", "Chưa có lịch", "Đã xong"])
    s = '<div class="a box" style="left:40px;top:112px;width:800px;height:76px"><div class="lbl">Tuần chạy (lấy từ Note)</div>%s</div>' % chips
    for x, w, t in [(852, 520, "Mã FG (gõ để tìm)"), (1384, 260, "EC Instruction")]:
        s += ('<div class="a box" style="left:%dpx;top:112px;width:%dpx;height:76px"><div class="lbl">%s</div>'
              '<div style="padding:0 12px;font-size:16px">All ▾</div></div>' % (x, w, t))
    s += ('<div class="a" style="left:1678px;top:112px;width:202px;height:76px;background:#3F4246;color:#fff;'
          'font-size:18px;font-weight:700;text-align:center;line-height:76px">XÓA LỌC</div>')
    return s


def kpis(vals):
    s = ""
    for i, (v, l, c, sz) in enumerate(vals):
        s += ('<div class="a box kpi" style="left:%dpx;top:198px;width:445px;height:122px"><b style="color:%s;font-size:%dpx">%s</b>'
              '<span>%s</span></div>' % (40 + i * 465, c, sz, esc(str(v)), esc(l)))
    return s


def page1(week_on, name):
    rows = [k for k, v in WV.items() if v["rel"] in week_on]
    rows.sort(key=lambda k: (p.REL6[WV[k]["rel"]], -(E[k]["issue"].toordinal() if E[k]["issue"] else 0), k))
    ks = rows
    fg_open = {f for k in ks if E[k]["open"] for f in E[k]["fgs"]}
    models = sorted({m for k in ks for m in E[k]["models"]})
    mtxt = ", ".join(models[:6]) + ("  … +%d" % (len(models) - 6) if len(models) > 6 else "") if models else "—"
    body = ""
    for k in ks:
        instr = E[k]["instr"]
        fgs = short(E[k]["fgs"]) if E[k]["fgs"] else ("Model " + short(E[k]["models"]) if E[k]["models"] else "—")
        hist = WV[k]["hist"] if WV[k]["moves"] else ("Đúng kế hoạch " + WV[k]["hist"] if WV[k]["hist"] != "—" else "—")
        body += ("<tr><td class='k'>%s</td><td class='k'>%s</td><td>%s</td><td>%s</td><td>%s</td><td>%s</td><td>%s</td><td class='%s'>%s</td></tr>"
                 % (esc(WV[k]["run"] + ("  ·  dời %d lần" % WV[k]["moves"] if WV[k]["moves"] else "")), esc(k),
                    esc(fgs), esc(change_text(k)), esc(hist), esc(WV[k]["note_raw"]),
                    E[k]["issue"].strftime("%d/%m/%Y") if E[k]["issue"] else "", "imm" if instr == "Immediately" else "", esc(instr)))
    cols = [150, 150, 190, 400, 190, 400, 120, 140]
    table = ("<table style='font-size:16px'><colgroup>%s</colgroup><tr><th>Tuần</th><th>ECN number</th><th>FG</th><th>Description of Change</th>"
             "<th>Lịch sử dời tuần</th><th>Note (mới nhất ở trên)</th><th>Date Issued</th><th>EC Instruction</th></tr>%s</table>"
             % ("".join('<col style="width:%dpx">' % c for c in cols), body))
    h = (head("ECN HÔM NAY CHẠY GÌ", "Log  ·  tuần chạy = tuần của Date; Note có Wxx/week xx + change = dời lịch") + slicers(week_on, "")
         + kpis([(len(ks), "ECN", "#1C2430", 46), (len(fg_open), "FG đang mở", "#1F5FBF", 46),
                 (sum(1 for k in ks if E[k]["paper"] != "Chưa gửi"), "Giấy ECN đã gửi", "#C77700", 46),
                 (mtxt, "Model", "#1C2430", 20)])
         + '<div class="a box" style="left:40px;top:334px;width:1840px;height:722px;overflow:hidden">%s</div>' % table)
    open(os.path.join(ROOT, name), "w", encoding="utf-8").write(
        "<!doctype html><meta charset=utf-8><style>%s</style><body>%s</body>" % (CSS, h))
    return len(ks)


def page3(name, kinds):
    rows = []
    for k in sorted(WV, reverse=True):
        for e in WV[k]["events"]:
            kind = "Dời lịch" if e["is_move"] else "Khác"
            if kind in kinds:
                n = e["note"]
                when = n["date_from"].strftime("%d/%m") if n["date_from"] else ("W%d" % n["wk_from"] if n["wk_from"] else "")
                mv = ("W%02d" % e["move_wk"]) if e["move_wk"] else ("năm %s" % e["move_yr"] if e["move_yr"] else "")
                rows.append((k, (short(E[k]["models"]) if E[k]["models"] else short(E[k]["fgs"])), when, kind, mv,
                             n["who"], n["cat"], e["line"]))
    cols = [160, 240, 100, 130, 120, 130, 200, 720]
    body = "".join("<tr><td class='k'>%s</td><td class='k'>%s</td><td>%s</td><td style='%s'>%s</td><td>%s</td><td>%s</td><td>%s</td><td>%s</td></tr>"
                   % (esc(a), esc(b), esc(c), "background:#FFE3A3" if d == "Dời lịch" else "", d, esc(m), esc(w), esc(cat), esc(t))
                   for a, b, c, d, m, w, cat, t in rows)
    table = ("<table style='font-size:16px'><colgroup>%s</colgroup><tr><th>ECN number</th><th>Model / FG</th><th>Tuần/ngày ghi</th>"
             "<th>Loại</th><th>Dời sang</th><th>Người</th><th>Phân loại</th><th>Nội dung</th></tr>%s</table>"
             % ("".join('<col style="width:%dpx">' % c for c in cols), body))
    allev = [e for k in WV for e in WV[k]["events"]]
    chips = "".join('<span class="chip %s">%s</span>' % ("on" if c in kinds and len(kinds) == 1 else "", c) for c in ["Dời lịch", "Khác"])
    out = (head("LỊCH SỬ ECN", "Mỗi dòng Note là một sự kiện  ·  Dời lịch = có Wxx/week xx + change")
           + '<div class="a box" style="left:40px;top:112px;width:420px;height:76px"><div class="lbl">Loại sự kiện</div>%s</div>' % chips)
    for x, w, t in [(472, 330, "Phân loại ghi chú"), (814, 520, "Mã FG (gõ để tìm)"), (1346, 260, "EC Instruction")]:
        out += ('<div class="a box" style="left:%dpx;top:112px;width:%dpx;height:76px"><div class="lbl">%s</div>'
                '<div style="padding:0 12px;font-size:16px">All ▾</div></div>' % (x, w, t))
    out += ('<div class="a" style="left:1678px;top:112px;width:202px;height:76px;background:#3F4246;color:#fff;'
            'font-size:18px;font-weight:700;text-align:center;line-height:76px">XÓA LỌC</div>')
    out += kpis([(len(allev), "Sự kiện (dòng Note)", "#1C2430", 46),
                 (sum(1 for e in allev if e["is_move"]), "Dòng dời lịch", "#C77700", 46),
                 (sum(1 for k in WV if WV[k]["moves"]), "ECN có dời lịch", "#1F5FBF", 46),
                 (sum(1 for e in allev if e["note"]["cat"] == p.CAT_UNKNOWN), "Dòng chưa nhận dạng", "#D1352B", 46)])
    out += '<div class="a box" style="left:40px;top:334px;width:1840px;height:722px;overflow:hidden">%s</div>' % table
    open(os.path.join(ROOT, name), "w", encoding="utf-8").write(
        "<!doctype html><meta charset=utf-8><style>%s</style><body>%s</body>" % (CSS, out))
    return len(rows)


def brief_text(k):
    seen = []
    for r in by[k]:
        ch = ("%s → %s" % (r["old"], r["new"]) if r["old"] and r["new"] else
              "%s → BỎ" % r["old"] if r["old"] else "THÊM → %s" % r["new"] if r["new"] else "—")
        t = (r["desc"][:70] + " (" + ch + ")") if r["desc"] and ch != "—" else (r["desc"][:70] if r["desc"] else ch)
        if t not in seen:
            seen.append(t)
    return "  |  ".join(x[:60] for x in seen[:2]) + ("  |  +%d khác" % (len(seen) - 2) if len(seen) > 2 else "")


def page_log(name):
    # Power BI font sizes are in points: px = pt x 1.333  (17pt = 22.7px, 15pt = 20px, 13pt = 17.3px)
    def table(ks):
        ks = sorted(ks, key=lambda k: (-(E[k]["issue"].toordinal() if E[k]["issue"] else 0), k))
        body = ""
        for k in ks:
            fgs = short(E[k]["fgs"]) if E[k]["fgs"] else ("Model " + short(E[k]["models"]) if E[k]["models"] else "—")
            ecn_txt = k + ("  ·  dời từ " + WV[k]["hist"].split(" → ")[0] if WV[k]["moves"] else "")
            instr = E[k]["instr"]
            body += ("<tr><td class='k' style='font-size:22.7px'>%s</td><td class='k' style='font-size:22.7px'>%s</td>"
                     "<td>%s</td><td class='%s'>%s</td></tr>"
                     % (esc(ecn_txt), esc(fgs), esc(brief_text(k)), "imm" if instr == "Immediately" else "", esc(instr)))
        return ("<table style='font-size:20px;width:880px'><colgroup><col style='width:240px'><col style='width:150px'>"
                "<col style='width:335px'><col style='width:150px'></colgroup>"
                "<tr><th style='font-size:17.3px'>ECN number</th><th style='font-size:17.3px'>FG</th>"
                "<th style='font-size:17.3px'>Description of Change</th><th style='font-size:17.3px'>EC Instruction</th></tr>%s</table>" % body)
    nxt = TODAY + dt.timedelta(days=7)
    out = head("ECN HÔM NAY CHẠY GÌ", "Tuần này (bên trái)  ·  Tuần sau (bên phải)")
    out += ('<div class="a box" style="left:40px;top:112px;width:700px;height:76px"><div class="lbl">Mã FG (gõ để tìm)</div>'
            '<div style="padding:0 12px;font-size:16px">All ▾</div></div>'
            '<div class="a box" style="left:752px;top:112px;width:260px;height:76px"><div class="lbl">EC Instruction</div>'
            '<div style="padding:0 12px;font-size:16px">All ▾</div></div>'
            '<div class="a" style="left:1024px;top:112px;width:202px;height:76px;background:#3F4246;color:#fff;'
            'font-size:18px;font-weight:700;text-align:center;line-height:76px">XÓA LỌC</div>')
    n = 0
    for x, rel, title, d, col in [(40, "Tuần này", "TUẦN NÀY", TODAY, "#3F4246"), (975, "Tuần sau", "TUẦN SAU", nxt, "#1F5FBF")]:
        ks = [k for k, v in WV.items() if v["rel"] == rel]
        n += len(ks)
        out += ('<div class="a" style="left:%dpx;top:198px;width:905px;height:66px;background:%s;color:#fff;border-radius:10px;'
                'font-size:30px;font-weight:700;text-align:center;line-height:66px">%s  ·  W%02d  ·  %d ECN</div>'
                % (x, col, title, p.iso_week(d), len(ks)))
        out += ('<div class="a box" style="left:%dpx;top:274px;width:905px;height:782px;overflow:hidden">%s</div>'
                % (x, table(ks)))
    open(os.path.join(ROOT, name), "w", encoding="utf-8").write(
        "<!doctype html><meta charset=utf-8><style>%s</style><body>%s</body>" % (CSS, out))
    return n


def bars(items, color, w=558, h=290, horizontal=True):
    if not items:
        return ""
    mx = max(v for _, v in items) or 1
    s = '<svg width="%d" height="%d" style="margin:0 14px">' % (w, h)
    if horizontal:
        rh = h / len(items)
        for i, (lab, v) in enumerate(items):
            bw = (w - 190) * v / mx
            s += ('<text x="0" y="%d" font-size="14" fill="#1C2430">%s</text><rect x="150" y="%d" width="%d" height="%d" fill="%s"/>'
                  '<text x="%d" y="%d" font-size="14">%s</text>'
                  % (i * rh + rh * .62, esc(lab[:20]), i * rh + 5, bw, rh - 10, color, 156 + bw, i * rh + rh * .62, v))
    else:
        cw = (w - 20) / len(items)
        for i, (lab, v) in enumerate(items):
            bh = (h - 50) * v / mx
            s += ('<rect x="%d" y="%d" width="%d" height="%d" fill="%s"/><text x="%d" y="%d" font-size="14" text-anchor="middle">%s</text>'
                  '<text x="%d" y="%d" font-size="13" text-anchor="middle">%s</text>'
                  % (10 + i * cw + cw * .15, h - 30 - bh, cw * .7, bh, color, 10 + i * cw + cw / 2, h - 34 - bh, v,
                     10 + i * cw + cw / 2, h - 10, esc(lab)))
    return s + "</svg>"


def page2(name):
    mon = TODAY - dt.timedelta(days=TODAY.weekday())
    weeks = [mon + dt.timedelta(days=7 * i) for i in range(-12, 1)]
    mc = collections.Counter(m for e in ecn for m in e["models"])
    fc = collections.Counter(f for e in ecn for f in e["fgs"])
    days = {}
    for e in ecn:
        if e["status"] == "Done" and e["completed"] and e["issue"]:
            days[e["ecn"]] = (e["completed"] - e["issue"]).days
    bk = collections.Counter()
    for e in ecn:
        d = days.get(e["ecn"])
        bk["Đã hủy" if e["status"] == "Cancel" else "Chưa hoàn tất" if d is None else "≤ 7 ngày" if d <= 7 else
           "8–14 ngày" if d <= 14 else "15–30 ngày" if d <= 30 else "> 30 ngày"] += 1
    order = ["≤ 7 ngày", "8–14 ngày", "15–30 ngày", "> 30 ngày", "Chưa hoàn tất", "Đã hủy"]
    form = collections.Counter(e["paper"] for e in ecn if e["open"])
    iss = [sum(1 for e in ecn if e["issue"] and e["issue"] - dt.timedelta(days=e["issue"].weekday()) == w) for w in weeks]
    com = [sum(1 for e in ecn if e["ecn"] in days and e["completed"] - dt.timedelta(days=e["completed"].weekday()) == w) for w in weeks]
    med = []
    mon26 = [mon + dt.timedelta(days=7 * i) for i in range(-25, 1)]
    for w in mon26:
        med.append(sum(1 for e in ecn if e["issue"] and e["issue"] - dt.timedelta(days=e["issue"].weekday()) == w))
    lab = ["W%02d" % p.iso_week(w) for w in weeks]
    fg_open = {f for e in ecn if e["open"] for f in e["fgs"]}
    def tree(items):
        colors = ["#1F5FBF", "#16803C", "#C77700", "#7A4BB8", "#0F8B8D"]
        cells = "".join('<div style="flex:%d 1 0;background:%s;color:#fff;display:flex;align-items:flex-start;padding:10px;'
                        'font-size:18px;font-weight:700;margin:2px;min-width:90px">%s<br>%d</div>' % (v, colors[i], esc(k), v)
                        for i, (k, v) in enumerate(items))
        return '<div style="display:flex;flex-wrap:wrap;height:290px;margin:0 14px">%s</div>' % cells
    def panel(i, title, inner):
        x = [40, 667, 1294][i % 3]; y = [334, 700][i // 3]
        return ('<div class="a box" style="left:%dpx;top:%dpx;width:586px;height:356px"><div class="t">%s</div>%s</div>'
                % (x, y, esc(title), inner))
    def cmp(a, b):
        mx = max(a + b) or 1; w, h = 558, 290; cw = (w - 20) / len(a)
        s = '<svg width="%d" height="%d" style="margin:0 14px">' % (w, h)
        for i, (x, y) in enumerate(zip(a, b)):
            for j, (v, c) in enumerate([(x, "#1F5FBF"), (y, "#16803C")]):
                bh = (h - 50) * v / mx
                s += '<rect x="%d" y="%d" width="%d" height="%d" fill="%s"/>' % (10 + i * cw + j * cw * .4, h - 30 - bh, cw * .38, bh, c)
            s += '<text x="%d" y="%d" font-size="12" text-anchor="middle">%s</text>' % (10 + i * cw + cw / 2, h - 10, lab[i])
        return s + "</svg>"
    def line(v):
        mx = max(v) or 1; w, h = 558, 290; cw = (w - 20) / (len(v) - 1); labl = ["W%02d" % p.iso_week(x) for x in mon26]
        pts = " ".join("%d,%d" % (10 + i * cw, h - 40 - (h - 70) * x / mx) for i, x in enumerate(v))
        s = '<svg width="%d" height="%d" style="margin:0 14px"><polyline fill="none" stroke="#16803C" stroke-width="3" points="%s"/>' % (w, h, pts)
        for i, x in enumerate(v):
            s += '<text x="%d" y="%d" font-size="12" text-anchor="middle">%d</text><text x="%d" y="%d" font-size="12" text-anchor="middle">%s</text>' % (10 + i * cw, h - 46 - (h - 70) * x / mx, x, 10 + i * cw, h - 10, labl[i] if i % 3 == 0 else "")
        return s + "</svg>"
    h = (head("PERFORMANCE ECN", "Hiệu suất xử lý ECN  ·  theo model, tuần, giấy ECN") + slicers([], "")
         + kpis([(len(ecn), "Tổng ECN", "#1C2430", 46), (len(fg_open), "FG đang mở", "#1F5FBF", 46),
                 (sum(1 for e in ecn if e["paper"] != "Chưa gửi"), "Giấy ECN đã gửi", "#C77700", 46),
                 (round(statistics.median(days.values())), "Ngày từ ban hành đến hoàn tất", "#16803C", 46)])
         + panel(0, "ECN theo Model (top 10)", bars(mc.most_common(10), "#1F5FBF"))
         + panel(1, "Từ ngày ban hành đến ngày hoàn tất", bars([(o, bk[o]) for o in order], "#16803C", horizontal=False))
         + panel(2, "ECN đang mở: giấy đã gửi / chưa gửi", bars([(k, form[k]) for k in ["Chưa gửi", "Đã gửi", "Đã nhận"]], "#C77700", horizontal=False))
         + panel(3, "ECN mở mới mỗi tuần (xu hướng 26 tuần)", line(med))
         + panel(4, "ECN mỗi tuần: mở mới vs hoàn tất (13 tuần)", cmp(iss, com))
         + panel(5, "Top 5 FG đổi nhiều nhất", tree(sorted(fc.items(), key=lambda kv: (-kv[1], kv[0]))[:5])))
    open(os.path.join(ROOT, name), "w", encoding="utf-8").write(
        "<!doctype html><meta charset=utf-8><style>%s</style><body>%s</body>" % (CSS, h))


tag = os.environ.get("ECN_PREVIEW_TAG", "")
jobs = [("_pvlog.html", "1_Log%s.png" % tag)]
print("log rows:", page_log("_pvlog.html"))
if not ONLY_LOG:
    print("history rows:", page3("_pv3.html", ["Dời lịch", "Khác"]))
    page2("_pv2.html")
    jobs += [("_pv3.html", "2_Lich_su.png"), ("_pv2.html", "3_Performance.png")]
chrome = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
os.makedirs(os.path.join(ROOT, "preview"), exist_ok=True)
for src, dst in jobs:
    subprocess.run([chrome, "--headless=new", "--disable-gpu", "--hide-scrollbars", "--window-size=1920,1080",
                    "--screenshot=" + os.path.join(ROOT, "preview", dst),
                    "file:///" + os.path.join(ROOT, src).replace("\\", "/")], check=False,
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    os.remove(os.path.join(ROOT, src))
print(sorted(os.listdir(os.path.join(ROOT, "preview"))))
