# -*- coding: utf-8 -*-
"""
ECN Shopfloor dashboard - PBIP generator (v2: parses free-text cells).
Source: sheet "ECN" of ECN Log.xlsx (real data, no mock).
The M code follows ecn_parse.py step by step; ecn_parse.py is the verified mirror.
"""
import json, os, uuid, zipfile, shutil, re, sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import ecn_parse as P

ROOT = os.path.dirname(os.path.abspath(__file__))
VARIANT = os.environ.get("ECN_VARIANT", "full")      # full | assembly
NAME = "ECN_Assembly" if VARIANT == "assembly" else "ECN_Shopfloor"
SM = os.path.join(ROOT, NAME + ".SemanticModel")
RP = os.path.join(ROOT, NAME + ".Report")
XLSX = os.path.join(ROOT, "ECN Log.xlsx")
ECN, LIN, NOT, MOD, FGT, WKS, CAL = ("ECN", "ECN Lines", "ECN Notes", "ECN Models", "ECN FG",
                                      "ECN Weeks", "ECN Calendar")


def w(path, obj, raw=False):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        f.write(obj if raw else json.dumps(obj, indent=2, ensure_ascii=False))


def L(s): return s.strip("\n").split("\n")
def tag(): return str(uuid.uuid4())
def mlist(ws): return "{" + ", ".join('"%s"' % x for x in ws) + "}"
def mnested(defs): return "{" + ", ".join('{"%s", %s}' % (a, mlist(b)) for a, b in defs) + "}"


# ==========================================================================
# M : small helpers
# ==========================================================================
M_TODAY = r'''
let
    // "Today" is evaluated at refresh time in Vietnam time (UTC+7).
    // Type a date as yyyy-MM-dd into pTodayOverride to test another day.
    Result =
        if pTodayOverride <> null and Text.Trim(pTodayOverride) <> ""
        then Date.FromText(Text.Trim(pTodayOverride), [Format = "yyyy-MM-dd", Culture = "en-US"])
        else Date.From(DateTimeZone.SwitchZone(DateTimeZone.UtcNow(), 7))
in
    Result
'''

M_FN_ISOWEEK = r'''
(d as date) as number =>
let
    dow = Date.DayOfWeek(d, Day.Monday),
    thu = Date.AddDays(d, 3 - dow)
in
    Number.IntegerDivide(Date.DayOfYear(thu) - 1, 7) + 1
'''

M_FN_WEEKNO = r'''
(t as nullable text) as nullable number =>
let
    s = if t = null then "" else Text.Upper(Text.Trim(t)),
    n = if Text.StartsWith(s, "W") then try Number.FromText(Text.Middle(s, 1)) otherwise null else null
in
    n
'''

M_FN_PLANWK = r'''
(line as text) as nullable number =>
let
    marker = " - W",
    tail   = if Text.Contains(line, marker)
             then Text.Trim(Text.AfterDelimiter(line, marker, {0, RelativePosition.FromEnd}))
             else null,
    digits = if tail = null then null else Text.Select(tail, {"0".."9"}),
    ok     = tail <> null and digits <> "" and digits = tail and Text.Length(tail) <= 2
in
    if ok then Number.FromText(digits) else null
'''

M_FN_EFFWK = r'''
(wk as nullable number, todayWk as number) as nullable number =>
    if wk = null then null
    else if wk <= 12 and todayWk >= 40 then wk + 53
    else wk
'''

M_FN_ISDIGITS = r'''
(x as text) as logical =>
    x <> "" and Text.Select(x, {"0".."9"}) = x
'''

M_FN_TOKENS = r'''
(s as nullable text) as list =>
let
    // Split on spaces, commas, colons, brackets, slashes... then drop a trailing "." or "-".
    t     = if s = null then "" else s,
    parts = Splitter.SplitTextByAnyDelimiter(
                {" ", ",", ";", ":", "(", ")", "[", "]", "/", "&",
                 "#(lf)", "#(cr)", "#(tab)", "!", "?", "="}
            )(t),
    clean = List.Transform(parts, each Text.TrimEnd(Text.Trim(_), {".", "-"}))
in
    List.Select(clean, each _ <> "")
'''

M_FN_SPLITLIST = r'''
(s as nullable text) as list =>
    if s = null then {}
    else
        let
            x = Text.Replace(Text.Replace(Text.Replace(s, ";", ","), "#(lf)", ","), "#(cr)", ","),
            p = List.Transform(Text.Split(x, ","), Text.Trim)
        in
            List.Distinct(List.Select(p, each _ <> ""))
'''

M_FN_CODEKIND = r'''
(tok as text) as nullable text =>
let
    // item  = 7 digits starting 1 or 3      (1011819, 3034045)
    // short = #2..5 digits                   (#7410, #11476)
    // part  = 4-5 digits - digits            (97149-002, 91585-1201)
    body  = if Text.StartsWith(tok, "#") then Text.Middle(tok, 1) else null,
    parts = Text.Split(tok, "-")
in
    if body <> null then
        (if fnIsDigits(body) and Text.Length(body) >= 2 and Text.Length(body) <= 5 then "short" else null)
    else if fnIsDigits(tok) then
        (if Text.Length(tok) = 7 and (Text.StartsWith(tok, "1") or Text.StartsWith(tok, "3")) then "item" else null)
    else if List.Count(parts) = 2 and fnIsDigits(parts{0}) and fnIsDigits(parts{1})
            and Text.Length(parts{0}) >= 4 and Text.Length(parts{0}) <= 5 then "part"
    else null
'''

M_FN_TAGS = (r'''
(desc as nullable text) as text =>
let
    d   = if desc = null then "" else Text.Lower(desc),
    Pkg = %s,
    Mec = %s,
    Pcb = %s,
    Lbl = %s,
    hit = (words as list) as logical =>
              List.AnyTrue(List.Transform(words, each Text.Contains(d, _))),
    tags = List.Select(
               {
                   if hit(Pkg) then "Bao bì" else "",
                   if hit(Mec) then "Cơ khí" else "",
                   if hit(Pcb) then "Mạch"   else "",
                   if hit(Lbl) then "Nhãn"   else ""
               },
               each _ <> ""
           )
in
    Text.Combine(tags, " · ")
''' % (mlist(P.PKG), mlist(P.MECH), mlist(P.PCB), mlist(P.LABEL)))

M_FN_COMPS = (r'''
(desc as nullable text) as text =>
let
    // Which physical part changes. Keywords of 3 characters or less must equal a whole word.
    d     = if desc = null then "" else Text.Lower(desc),
    ltoks = List.Transform(fnTokens(desc), Text.Lower),
    Defs  = %s,
    hit   = (kws as list) as logical =>
                List.AnyTrue(
                    List.Transform(
                        kws,
                        each if Text.Length(_) <= 3 then List.Contains(ltoks, _) else Text.Contains(d, _)
                    )
                ),
    out   = List.Select(List.Transform(Defs, each if hit(_{1}) then _{0} else null), each _ <> null)
in
    Text.Combine(out, ", ")
''' % mnested(P.COMPONENTS))

M_FN_PARSEDESC = (r'''
(desc as nullable text) as record =>
let
    // Reads "Replace 1006982 by 1011819", "Remove 1003942 ... Add item 1011829",
    // "Thay hộp 3029144 bằng hộp 3031758, Bỏ nhãn 3031690" word by word.
    Rem = %s,
    Add = %s,
    Rep = %s,
    Upd = %s,
    Con = %s,
    start = [mode = "ref", old = {}, new = {}, ref = {}, act = {}],
    step = (s as record, tok as text) as record =>
        let
            l    = Text.Lower(tok),
            kind = fnCodeKind(tok)
        in
            if List.Contains(Rem, l) then
                Record.Combine({s, [mode = "old", act = List.Distinct(s[act] & {"Bỏ"})]})
            else if List.Contains(Add, l) then
                Record.Combine({s, [mode = "new", act = List.Distinct(s[act] & {"Thêm"})]})
            else if List.Contains(Rep, l) then
                Record.Combine({s, [mode = "old", act = List.Distinct(s[act] & {"Thay"})]})
            else if List.Contains(Con, l) then
                (if List.IsEmpty(s[old]) and not List.IsEmpty(s[ref])
                 then Record.Combine({s, [mode = "new", old = s[ref], ref = {}]})
                 else Record.Combine({s, [mode = "new"]}))
            else if List.Contains(Upd, l) then
                Record.Combine({s, [mode = "ref", act = List.Distinct(s[act] & {"Cập nhật"})]})
            else if kind <> null then
                (if s[mode] = "old" then Record.Combine({s, [old = List.Distinct(s[old] & {tok})]})
                 else if s[mode] = "new" then Record.Combine({s, [new = List.Distinct(s[new] & {tok})]})
                 else Record.Combine({s, [ref = List.Distinct(s[ref] & {tok})]}))
            else s
in
    List.Accumulate(fnTokens(desc), start, step)
''' % (mlist(P.V_REMOVE), mlist(P.V_ADD), mlist(P.V_REPLACE), mlist(P.V_UPDATE), mlist(P.V_CONNECT)))

M_FN_ESTWK = r'''
(ltoks as list) as record =>
let
    // "est W45", "Est week 51 will change", "Estimate week 31", "est end 2026", "Week 34 will change"
    idxs = List.Select(List.Positions(ltoks), each Text.StartsWith(ltoks{_}, "est")),
    i0   = if not List.IsEmpty(idxs) then idxs{0}
           else if List.Contains(ltoks, "change") then -1
           else null,
    win  = if i0 = null then {} else List.FirstN(List.Skip(ltoks, i0 + 1), 5),
    wkAt = (j as number) as nullable number =>
        let t = win{j} in
            if Text.Length(t) >= 2 and Text.StartsWith(t, "w") and fnIsDigits(Text.Middle(t, 1))
            then Number.FromText(Text.Middle(t, 1))
            else if t = "week" and j + 1 < List.Count(win) and fnIsDigits(win{j + 1})
            then Number.FromText(win{j + 1})
            else null,
    wk   = List.First(List.RemoveNulls(List.Transform(List.Positions(win), wkAt)), null),
    yr   = List.First(List.Select(win, each fnIsDigits(_) and Text.Length(_) = 4), null)
in
    if wk <> null then [Wk = wk, Txt = null]
    else if yr <> null then [Wk = null, Txt = (if List.Contains(win, "end") then "cuối " else "") & yr]
    else [Wk = null, Txt = null]
'''

M_FN_NOTELINE = (r'''
(line as text, T as date, people as list, vendors as list) as record =>
let
    s0 = Text.Trim(line),
    s1 = if Text.StartsWith(Text.Upper(s0), "OBS:") then Text.Trim(Text.Middle(s0, 4)) else s0,

    // ---- "W40:" / "W34-40:" week prefix -------------------------------------
    colon    = Text.PositionOf(s1, ":"),
    head     = if colon > 0 and colon <= 12 then Text.Upper(Text.Trim(Text.Start(s1, colon))) else "",
    headRest = if head = "" then "" else Text.Middle(head, 1),
    isWk     = Text.StartsWith(head, "W") and headRest <> ""
               and Text.Select(headRest, {"0".."9", "-", "~", " "}) = headRest
               and Text.Select(head, {"0".."9"}) <> "",
    nums     = if isWk
               then List.Transform(
                        List.Select(
                            Text.Split(Text.Replace(Text.Replace(headRest, "~", "-"), " ", ""), "-"),
                            each _ <> ""),
                        Number.FromText)
               else {},
    wkFrom   = if List.IsEmpty(nums) then null else List.Last(nums),
    s2       = if wkFrom <> null then Text.Trim(Text.Middle(s1, colon + 1)) else s1,

    // ---- "9/23 update)" / "8/6)" / "9/18:" / "Sep 28 2026:" date prefix ----
    p1     = Text.PositionOf(s2, ")"),
    p2     = Text.PositionOf(s2, ":"),
    endP   = if p1 >= 0 and p1 <= 14 then p1 else if p2 >= 0 and p2 <= 14 then p2 else -1,
    dtok   = if endP > 0 then Text.Trim(Text.Start(s2, endP)) else "",
    dfirst = if dtok = "" then "" else List.First(Text.Split(dtok, " ")),
    dparts = Text.Split(dfirst, "/"),
    isMD   = dfirst <> "" and List.Count(dparts) = 2
             and fnIsDigits(dparts{0}) and fnIsDigits(dparts{1})
             and Number.FromText(dparts{0}) >= 1 and Number.FromText(dparts{0}) <= 12
             and Number.FromText(dparts{1}) >= 1 and Number.FromText(dparts{1}) <= 31,
    dateMD = if isMD then try #date(Date.Year(T), Number.FromText(dparts{0}), Number.FromText(dparts{1})) otherwise null else null,
    dateNm = if dateMD = null and dtok <> ""
             then try Date.FromText(dtok, [Culture = "en-US"]) otherwise null
             else null,
    dateFrom = if dateMD <> null then dateMD else dateNm,
    s3     = if dateFrom <> null then Text.Trim(Text.Middle(s2, endP + 1)) else s2,

    lt    = Text.Lower(s3),
    toks  = fnTokens(s3),
    ltoks = List.Transform(toks, Text.Lower),

    who   = List.Select(people, each List.Contains(toks, _)),
    vend  = List.Select(vendors, each Text.Contains(lt, Text.Lower(_))),
    est   = fnEstWk(ltoks),

    iFin  = Text.PositionOf(lt, "finished by "),
    finish= if iFin >= 0
            then try Date.FromText(Text.Trim(Text.Middle(s3, iFin + 12)), [Format = "dd-MMM-yy", Culture = "en-US"]) otherwise null
            else null,
    iLast = Text.PositionOf(lt, "last order "),
    lastO = if iLast >= 0 then Text.Trim(Text.BeforeDelimiter(Text.Middle(s3, iLast + 11) & ",", ",")) else null,

    stockC = if Text.Contains(lt, "stock")
             then List.Distinct(List.Select(toks, each fnCodeKind(_) = "item"))
             else {},
    qtys   = List.Transform(
                 List.Select(ltoks, each Text.EndsWith(_, "pcs") and fnIsDigits(Text.Start(_, Text.Length(_) - 3))),
                 each Text.Start(_, Text.Length(_) - 3)),

    // ---- category: first rule that matches wins ----------------------------
    CatDefs = %s,
    vl      = List.Transform(vendors, Text.Lower),
    Defs    = List.Transform(CatDefs, each if _{0} = "%s" then {_{0}, _{1} & vl} else _),
    found   = List.First(
                  List.Select(Defs, each List.AnyTrue(List.Transform(_{1}, (k) => Text.Contains(lt, k)))),
                  null),
    isDone  = List.Count(ltoks) > 0 and List.Count(ltoks) <= 3 and List.Contains(%s, ltoks{0}),
    cat     = if isDone then "%s"
              else if lt = "cancel" then "%s"
              else if found <> null then found{0}
              else "%s",

    key     = if dateFrom <> null then Number.From(dateFrom)
              else if wkFrom <> null then Number.From(Date.AddDays(#date(Date.Year(T), 1, 1), (wkFrom - 1) * 7))
              else -1,

    // ---- schedule move: the line has a week (Wxx / week xx) AND the word "change" ----
    chg     = List.AnyTrue(List.Transform(%s, each Text.Contains(lt, _))),
    wkTok   = fnWeekTokens(ltoks),
    yrTok   = List.Select(ltoks, each fnIsDigits(_) and Text.Length(_) = 4 and Text.StartsWith(_, "20")),
    isMove  = chg and (not List.IsEmpty(wkTok) or not List.IsEmpty(yrTok)),
    moveWk  = if chg and not List.IsEmpty(wkTok) then List.Last(wkTok) else null
in
    [
        Text      = s3,
        WkFrom    = wkFrom,
        DateFrom  = dateFrom,
        Who       = Text.Combine(who, ", "),
        Vendor    = Text.Combine(vend, ", "),
        Target    = fnPlanWk(s3),
        EstWk     = est[Wk],
        EstTxt    = est[Txt],
        FinishBy  = finish,
        LastOrder = lastO,
        Stock     = Text.Combine(stockC, ", "),
        Qty       = List.First(qtys, null),
        Category  = cat,
        Key       = key,
        // The user's rule: the W.. at the start of the line is its week. Fallbacks:
        // est/target week written in the text, then the ISO week of the line's date.
        LineWk    = if wkFrom <> null then wkFrom
                    else if est[Wk] <> null then est[Wk]
                    else if fnPlanWk(s3) <> null then fnPlanWk(s3)
                    else if dateFrom <> null then fnIsoWeek(dateFrom)
                    else null,
        IsMove    = isMove,
        MoveWk    = moveWk,
        MoveTxt   = if moveWk <> null then "W" & Text.PadStart(Text.From(moveWk), 2, "0")
                    else if isMove then "năm " & List.Last(yrTok)
                    else null
    ]
''' % (mnested(P.CATS), P.CAT_VENDOR, mlist(P.DONE_WORDS), P.CAT_DONE, P.CAT_CANCEL, P.CAT_UNKNOWN,
                                                       mlist(P.CHANGE_WORDS)))

# ==========================================================================
# M : tables
# ==========================================================================
M_FN_WEEKTOKENS = r'''
(ltoks as list) as list =>
let
    // Week numbers written in a line: "w51", "w34-40" (-> 40), "week 38", "wk 12".
    n   = List.Count(ltoks),
    one = (i as number) as nullable number =>
        let
            t     = ltoks{i},
            body  = Text.Replace(Text.Middle(t, 1), "~", "-"),
            nxt   = if i + 1 < n then ltoks{i + 1} else "",
            isW   = Text.Length(t) >= 2 and Text.StartsWith(t, "w")
                    and Text.Select(body, {"0".."9", "-"}) = body
                    and Text.Select(body, {"0".."9"}) <> "",
            parts = if isW then List.Select(Text.Split(body, "-"), each _ <> "") else {},
            isWk  = (t = "week" or t = "wk") and fnIsDigits(nxt)
        in
            if isW and not List.IsEmpty(parts) and fnIsDigits(List.Last(parts)) then Number.FromText(List.Last(parts))
            else if isWk then Number.FromText(nxt)
            else null
in
    List.RemoveNulls(List.Transform(List.Positions(ltoks), one))
'''

M_FN_WEEKSTART = r'''
(wk as number, T as date) as date =>
let
    // Monday of ISO week wk, in the year that is closest to today's week
    // (W01..W12 seen from W40+ = next year, W40+ seen from W01..W12 = last year).
    isoYear = Date.Year(Date.AddDays(Date.StartOfWeek(T, Day.Monday), 3)),
    tw      = fnIsoWeek(T),
    yy      = if wk <= 12 and tw >= 40 then isoYear + 1
              else if wk >= 40 and tw <= 12 then isoYear - 1
              else isoYear,
    w1      = Date.StartOfWeek(#date(yy, 1, 4), Day.Monday)
in
    Date.AddDays(w1, (wk - 1) * 7)
'''

M_BASE = r'''
let
    Raw       = #"ECN Raw",

    // Header row = the row whose first cell is "ECN Number" (row 4 today).
    FirstCol  = List.First(Table.ColumnNames(Raw)),
    HdrIdx    = List.PositionOf(Table.Column(Raw, FirstCol), "ECN Number"),
    Promoted  = Table.PromoteHeaders(Table.Skip(Raw, HdrIdx), [PromoteAllScalars = true]),

    Clean =
        Table.TransformColumnNames(
            Promoted,
            each Text.Trim(
                     Text.Replace(
                         Text.Replace(
                             Text.Replace(_, "#(lf)", " "),
                         "#(cr)", " "),
                     "  ", " ")
                 )
        ),

    // Only what the shopfloor needs. Price, scrap, OPO qty columns fall away.
    Keep =
        Table.SelectColumns(
            Clean,
            {
                "ECN Number", "Date", "FG", "Model /Customer No.",
                "Old item", "New item", "Buyer (Local/ Pur/ Ser/ Inhouse)", "Status",
                "Description of Change", "EC Instruction", "Date Issued", "Date Completed",
                "ECN form sent date", "ECN form received date",
                "New item (est MP week)", "Note"
            } & #"Week Map"[Col],      // + the weekly old-item balance columns
            MissingField.UseNull
        ),

    // The est-MP-week column is a formula that returns #N/A when empty.
    NoErrors =
        Table.ReplaceErrorValues(
            Keep,
            List.Transform(Table.ColumnNames(Keep), each {_, null})
        ),

    ToText = (v as any) as nullable text =>
        if v = null then null
        else let t = Text.Trim(Text.From(v)) in if t = "" then null else t,

    ToDate = (v as any) as nullable date =>
        try Date.From(v) otherwise null,

    Cleaned =
        Table.TransformColumns(
            NoErrors,
            {
                {"ECN Number",             each ToText(_), type nullable text},
                {"FG",                     each ToText(_), type nullable text},
                {"Model /Customer No.",    each ToText(_), type nullable text},
                {"Old item",               each ToText(_), type nullable text},
                {"New item",               each ToText(_), type nullable text},
                {"Buyer (Local/ Pur/ Ser/ Inhouse)", each ToText(_), type nullable text},
                {"Status",                 each ToText(_), type nullable text},
                {"Description of Change",  each ToText(_), type nullable text},
                {"EC Instruction",         each ToText(_), type nullable text},
                {"ECN form sent date",     each ToText(_), type nullable text},
                {"ECN form received date", each ToText(_), type nullable text},
                {"New item (est MP week)", each ToText(_), type nullable text},
                {"Note",                   each ToText(_), type nullable text},
                {"Date",                   each ToDate(_), type nullable date},
                {"Date Issued",            each ToDate(_), type nullable date},
                {"Date Completed",         each ToDate(_), type nullable date}
            }
        ),

    Renamed =
        Table.RenameColumns(
            Cleaned,
            {
                {"Model /Customer No.",    "Model / Customer No."},
                {"New item (est MP week)", "Est MP Text"},
                {"Buyer (Local/ Pur/ Ser/ Inhouse)", "Buyer"}
            }
        ),

    HasEcn = Table.SelectRows(Renamed, each [ECN Number] <> null),

    AddIssue =
        Table.AddColumn(
            HasEcn, "Issue Date",
            each if [Date] <> null then [Date] else [Date Issued],
            type nullable date
        ),

    AddEst =
        Table.AddColumn(AddIssue, "Est Wk", each fnWeekNo([Est MP Text]), type nullable number),

    AddChange =
        Table.AddColumn(
            AddEst, "Change",
            each if [Old item] <> null and [New item] <> null then [Old item] & " → " & [New item]
                 else if [Old item] <> null then [Old item] & " → BỎ"
                 else if [New item] <> null then "THÊM → " & [New item]
                 else "—",
            type text
        ),

    AddRank =
        Table.AddColumn(
            AddChange, "Change Rank",
            each if [Old item] <> null and [New item] <> null then 0
                 else if [New item] <> null then 1
                 else 2,
            Int64.Type
        ),

    AddTags =
        Table.AddColumn(AddRank, "Tags", each fnTags([Description of Change]), type text),

    AddComps =
        Table.AddColumn(AddTags, "Components", each fnComps([Description of Change]), type text),

    // Codes written inside Description of Change (also catches #short codes).
    AddParsed =
        Table.AddColumn(AddComps, "Parsed", each fnParseDesc([Description of Change]), type record),

    AddDescOld =
        Table.AddColumn(AddParsed, "Desc Old", each Text.Combine([Parsed][old], ", "), type text),
    AddDescNew =
        Table.AddColumn(AddDescOld, "Desc New", each Text.Combine([Parsed][new], ", "), type text),
    AddDescRef =
        Table.AddColumn(AddDescNew, "Desc Ref", each Text.Combine([Parsed][ref], ", "), type text),

    // Structure first (old/new columns), then the verbs found in the text.
    AddAction =
        Table.AddColumn(
            AddDescRef, "Action",
            each if [Old item] <> null and [New item] <> null then "Thay"
                 else if [New item] <> null then "Thêm"
                 else if [Old item] <> null then "Bỏ"
                 else if not List.IsEmpty([Parsed][act]) then Text.Combine([Parsed][act], " + ")
                 else "—",
            type text
        ),

    NoParsed = Table.RemoveColumns(AddAction, {"Parsed"}),

    AddStatus =
        Table.AddColumn(
            NoParsed, "Line Status",
            each if [Status] = null then "Chưa bắt đầu" else [Status],
            type text
        ),

    AddIndex = Table.AddIndexColumn(AddStatus, "Line No", 1, 1, Int64.Type)
in
    AddIndex
'''

M_NOTES_BASE = r'''
let
    People  = List.Select(List.Transform(Text.Split(pPeople, ","), Text.Trim), each _ <> ""),
    Vendors = List.Select(List.Transform(Text.Split(pVendors, ","), Text.Trim), each _ <> ""),
    T       = TodayDate,
    Grouped = Table.Group(#"ECN Base", {"ECN Number"}, {{"Rows", each _, type table}}),

    // Every distinct cell of Note of the ECN, every distinct line of each cell.
    PerEcn = (ecn as text, rs as table) as list =>
        let
            cells    = List.Distinct(List.RemoveNulls(Table.Column(rs, "Note"))),
            cellLines = List.Transform(
                            cells,
                            (c) => List.Select(
                                       List.Transform(Text.Split(Text.Replace(c, "#(cr)", ""), "#(lf)"), Text.Trim),
                                       each _ <> "")),
            flat     = List.Combine(
                           List.Transform(
                               List.Positions(cellLines),
                               (ci) => List.Transform(
                                           List.Positions(cellLines{ci}),
                                           (pi) => [Cell = ci + 1, Pos = pi + 1, Line = cellLines{ci}{pi}]))),
            texts    = List.Distinct(List.Transform(flat, each [Line])),
            uniq     = List.Transform(texts, (t) => List.First(List.Select(flat, each [Line] = t))),
            recs     = List.Transform(
                           List.Positions(uniq),
                           (i) => Record.Combine({
                                      [#"ECN Number" = ecn, Seq = i + 1, Cell = uniq{i}[Cell],
                                       Pos = uniq{i}[Pos], Line = uniq{i}[Line]],
                                      fnNoteLine(uniq{i}[Line], T, People, Vendors)
                                  })
                       )
        in
            recs,

    AllRecs = List.Combine(List.Transform(Table.ToRecords(Grouped), each PerEcn(_[ECN Number], _[Rows]))),

    Tbl =
        Table.FromRecords(
            AllRecs,
            type table [
                #"ECN Number" = nullable text, Seq = nullable Int64.Type, Line = nullable text,
                Cell = nullable Int64.Type, Pos = nullable Int64.Type, LineWk = nullable Int64.Type,
                IsMove = nullable logical, MoveWk = nullable Int64.Type, MoveTxt = nullable text,
                Text = nullable text,
                WkFrom = nullable Int64.Type, DateFrom = nullable date,
                Who = nullable text, Vendor = nullable text, Target = nullable Int64.Type,
                EstWk = nullable Int64.Type, EstTxt = nullable text,
                FinishBy = nullable date, LastOrder = nullable text,
                Stock = nullable text, Qty = nullable text, Category = nullable text,
                Key = nullable number
            ],
            MissingField.UseNull
        )
in
    Tbl
'''

M_NOTES = r'''
let
    Source = #"ECN Notes Base",
    AddWhen =
        Table.AddColumn(
            Source, "Cập nhật",
            each if [DateFrom] <> null then Date.ToText([DateFrom], "dd/MM")
                 else if [WkFrom] <> null then "W" & Text.From([WkFrom])
                 else null,
            type text
        ),
    AddDue =
        Table.AddColumn(
            AddWhen, "Hạn / Est",
            each Text.Combine(
                     List.Select(
                         {
                             if [FinishBy] <> null then "Xong trước " & Date.ToText([FinishBy], "dd/MM/yyyy") else null,
                             if [EstWk] <> null then "Est W" & Text.From([EstWk])
                             else if [EstTxt] <> null then "Est " & [EstTxt] else null,
                             if [Target] <> null then "Kế hoạch W" & Text.From([Target]) else null,
                             if [LastOrder] <> null then "Đơn cuối " & [LastOrder] else null
                         },
                         each _ <> null
                     ),
                     " · "
                 ),
            type text
        ),
    AddStock =
        Table.AddColumn(
            AddDue, "Tồn cần xả",
            each if [Stock] = "" then null
                 else "Xả tồn " & [Stock] & (if [Qty] <> null then " · " & [Qty] & " pcs" else ""),
            type text
        ),
    AddKind =
        Table.AddColumn(
            AddStock, "Loại sự kiện",
            each if [IsMove] = true then "Dời lịch" else "Khác",
            type text
        ),
    Pick =
        Table.SelectColumns(
            AddKind,
            {"ECN Number", "Seq", "Line", "Text", "Cập nhật", "Who", "Vendor", "Category",
             "Hạn / Est", "Tồn cần xả", "Key", "Loại sự kiện", "MoveTxt"}
        ),
    Renamed = Table.RenameColumns(Pick, {{"Who", "Người phụ trách"}, {"Vendor", "NCC"},
                                         {"Category", "Phân loại"}, {"Text", "Nội dung"},
                                         {"MoveTxt", "Dời sang"}})
in
    Renamed
'''

M_LINES = r'''
let
    Source = #"ECN Base",
    Pick =
        Table.SelectColumns(
            Source,
            {"Line No", "ECN Number", "Action", "Old item", "New item", "Desc Old", "Desc New",
             "Desc Ref", "Description of Change", "Tags", "Components", "Buyer", "Line Status"}
        ),
    Renamed =
        Table.RenameColumns(
            Pick,
            {{"Tags", "Loại"}, {"Components", "Vật tư"},
             {"Desc Old", "Mã cũ (mô tả)"}, {"Desc New", "Mã mới (mô tả)"},
             {"Desc Ref", "Mã liên quan"}}
        ),
    Typed =
        Table.TransformColumnTypes(
            Renamed,
            {
                {"Line No", Int64.Type}, {"ECN Number", type text}, {"Action", type text},
                {"Old item", type text}, {"New item", type text},
                {"Mã cũ (mô tả)", type text}, {"Mã mới (mô tả)", type text},
                {"Mã liên quan", type text}, {"Description of Change", type text},
                {"Loại", type text}, {"Vật tư", type text}, {"Buyer", type text},
                {"Line Status", type text}
            }
        )
in
    Typed
'''

M_MODELS = r'''
let
    Rows  = Table.SelectColumns(#"ECN Base", {"ECN Number", "Model / Customer No."}),
    Split = Table.AddColumn(Rows, "Model", each fnSplitList([#"Model / Customer No."])),
    Flat  = Table.ExpandListColumn(Split, "Model"),
    Pick  = Table.SelectRows(Table.SelectColumns(Flat, {"ECN Number", "Model"}), each [Model] <> null),
    Out   = Table.TransformColumnTypes(Table.Distinct(Pick), {{"ECN Number", type text}, {"Model", type text}})
in
    Out
'''

M_FG = r'''
let
    Rows  = Table.SelectColumns(#"ECN Base", {"ECN Number", "FG"}),
    Split = Table.AddColumn(Rows, "FG Code", each fnSplitList([FG])),
    Flat  = Table.ExpandListColumn(Split, "FG Code"),
    Pick  = Table.SelectRows(Table.SelectColumns(Flat, {"ECN Number", "FG Code"}), each [FG Code] <> null),
    // ECNs with no FG still get one row so they can be found
    NoFg  = List.Difference(List.Distinct(Rows[ECN Number]), List.Distinct(Pick[ECN Number])),
    All   = Table.Combine({Pick, Table.FromRows(List.Transform(NoFg, each {_, "—"}), {"ECN Number", "FG Code"})}),
    Find  = Table.AddColumn(Table.Distinct(All), "FG / ECN", each [FG Code] & "  ·  " & [ECN Number], type text),
    Out   = Table.TransformColumnTypes(Find, {{"ECN Number", type text}, {"FG Code", type text}})
in
    Out
'''


M_RAW = r'''
let
    // SharePoint.Contents authenticates at SITE level - it cannot be pointed at a folder.
    // Everything below the site is plain navigation: library -> folders -> file.
    Root = SharePoint.Contents(pSiteUrl, [ApiVersion = 15]),

    // The library is "Shared Documents" on most tenants, "Documents" on others.
    Documents =
        try Root{[Name = "Shared Documents"]}[Content]
        otherwise Root{[Name = "Documents"]}[Content],

    // pFolderPath = "General/Sharing to Production, Planner and Buyer"
    Segments =
        List.Select(
            List.Transform(Text.Split(pFolderPath, "/"), Text.Trim),
            each _ <> ""
        ),
    Folder =
        List.Accumulate(
            Segments,
            Documents,
            (state, segment) => state{[Name = segment]}[Content]
        ),

    Workbook = Folder{[Name = pFileName]}[Content],
    Source   = Excel.Workbook(Workbook, null, true),
    Raw      = Source{[Item = pSheetName, Kind = "Sheet"]}[Data]
in
    Raw
'''

M_WEEKMAP = r'''
let
    // Columns V..BM: header "Old item (qty)N", row above = ISO week, two rows above = Monday.
    Raw      = #"ECN Raw",
    FirstCol = List.First(Table.ColumnNames(Raw)),
    HdrIdx   = List.PositionOf(Table.Column(Raw, FirstCol), "ECN Number"),
    Hdr      = Record.FieldValues(Raw{HdrIdx}),
    Wk       = if HdrIdx >= 1 then Record.FieldValues(Raw{HdrIdx - 1}) else {},
    Dt       = if HdrIdx >= 2 then Record.FieldValues(Raw{HdrIdx - 2}) else {},
    CleanName = (h as any) as nullable text =>
        if h = null then null
        else Text.Trim(
                 Text.Replace(
                     Text.Replace(
                         Text.Replace(Text.From(h), "#(lf)", " "),
                     "#(cr)", " "),
                 "  ", " ")),
    // Some Mondays are real dates, some are bare serial numbers (46181).
    ToStart = (v as any) as nullable date =>
        if v = null then null
        else if Value.Is(v, type datetime) or Value.Is(v, type date) then Date.From(v)
        else if Value.Is(v, type number) and v > 30000 and v < 80000 then Date.From(v)
        else null,
    Rows = List.Transform(
               List.Positions(Hdr),
               (i) => [
                   Col       = CleanName(Hdr{i}),
                   WeekNo    = if i < List.Count(Wk) and Value.Is(Wk{i}, type number) then Wk{i} else null,
                   WeekStart = if i < List.Count(Dt) then ToStart(Dt{i}) else null
               ]),
    Keep = List.Select(
               Rows,
               each [Col] <> null and [WeekNo] <> null and [WeekStart] <> null
                    and [WeekNo] >= 1 and [WeekNo] <= 53 and [WeekNo] = Number.RoundDown([WeekNo])),
    Tbl  = Table.FromRecords(
               Keep,
               type table [Col = nullable text, WeekNo = nullable number, WeekStart = nullable date],
               MissingField.UseNull)
in
    Tbl
'''

M_WEEKS_BASE = r'''
let
    // The weekly cells are the projected balance of the OLD item at the end of
    // each week (SUMIF from the MRP sheet). A drop = what was used that week;
    // balance <= 0 = old item finished, the line runs the new item.
    T       = TodayDate,
    Monday  = Date.StartOfWeek(T, Day.Monday),
    WM      = #"Week Map",
    Base    = #"ECN Base",
    Ids     = {"Line No", "ECN Number", "Old item", "New item", "Change"},
    Cols    = List.Intersect({WM[Col], Table.ColumnNames(Base)}),
    Pick    = Table.SelectRows(Table.SelectColumns(Base, Ids & Cols), each [Old item] <> null),
    Unp     = Table.UnpivotOtherColumns(Pick, Ids, "Col", "Raw"),
    Num     = Table.SelectRows(Unp, each Value.Is([Raw], type number)),
    J       = Table.NestedJoin(Num, {"Col"}, WM, {"Col"}, "W", JoinKind.Inner),
    X       = Table.ExpandTableColumn(J, "W", {"WeekNo", "WeekStart"}, {"Week No", "Week Start"}),
    Slim    = Table.RenameColumns(Table.RemoveColumns(X, {"Col"}), {{"Raw", "Balance"}}),
    Grouped = Table.Group(Slim, {"Line No"},
                  {{"Rows", each Table.Sort(_, {{"Week Start", Order.Ascending}}), type table}}),

    // A line whose balance is 0 in every week is an item the MRP sheet does not
    // know (SUMIF returns 0) - no information, drop it.
    HasData = Table.SelectRows(Grouped, each List.AnyTrue(List.Transform([Rows][Balance], (b) => b <> 0))),

    Calc = (t as table) as table =>
        let
            recs = Table.ToRecords(t),
            out  = List.Transform(
                       List.Positions(recs),
                       (i) =>
                           let
                               r      = recs{i},
                               begin  = if i = 0 then r[Balance] else recs{i - 1}[Balance],
                               demand = List.Max({begin - r[Balance], 0}),
                               oldU   = List.Min({demand, List.Max({begin, 0})}),
                               newU   = demand - oldU,
                               mode   = if demand = 0 then "Không dùng"
                                        else if newU = 0 then "Dùng mã cũ"
                                        else if r[New item] = null then "Hết tồn (bỏ mã)"
                                        else if r[New item] = r[Old item] then "Thiếu tồn (giữ mã)"
                                        else if oldU = 0 then "Chạy mã mới"
                                        else "Chuyển mã trong tuần",
                               d      = Number.Round(Duration.Days(r[Week Start] - Monday) / 7),
                               rel    = if d < -1 then "Đã qua"
                                        else if d = -1 then "Tuần trước"
                                        else if d = 0 then "Tuần này"
                                        else if d = 1 then "Tuần sau"
                                        else "+" & Text.From(d) & " tuần"
                           in
                               Record.Combine({r, [
                                   Begin = begin, Demand = demand,
                                   #"Old Used" = oldU, #"New Used" = newU, Mode = mode,
                                   #"Rel N" = d, #"Rel Order" = (if d < -1 then -2 else d),
                                   #"Week Rel" = rel]})
                   )
        in
            Table.FromRecords(out),

    Out      = Table.Combine(List.Transform(HasData[Rows], Calc)),
    AddLabel = Table.AddColumn(
                   Out, "Week Label",
                   each "W" & Text.PadStart(Text.From([Week No]), 2, "0") & " · " & Date.ToText([Week Start], "dd/MM"),
                   type text),
    AddKey   = Table.AddColumn(
                   AddLabel, "Week Key",
                   each Date.Year([Week Start]) * 10000 + Date.Month([Week Start]) * 100 + Date.Day([Week Start]),
                   Int64.Type)
in
    AddKey
'''

M_LINESTOCK = r'''
let
    // Per ECN line: when the old item runs out (future), since when it is out (past).
    Monday = Date.StartOfWeek(TodayDate, Day.Monday),
    G = Table.Group(#"ECN Weeks Base", {"Line No", "ECN Number"},
            {{"Rows", each Table.Sort(_, {{"Week Start", Order.Ascending}}), type table}}),
    F = (t as table) as record =>
        let
            recs   = Table.ToRecords(t),
            pos    = List.Select(List.Positions(recs), each recs{_}[Week Start] <= Monday),
            now    = if List.IsEmpty(pos) then 0 else List.Last(pos),
            nowBal = recs{now}[Balance],
            fl     = List.Select(recs, each [Week Start] >= Monday and [Balance] <= 0),
            up     = List.Select(List.Positions(recs), each _ <= now and recs{_}[Balance] > 0),
            sIdx   = if List.IsEmpty(up) then 0 else List.Last(up) + 1,
            since  = if nowBal > 0 then null else recs{sIdx}[Week Start]
        in
            [
                Fut   = if nowBal > 0 and not List.IsEmpty(fl) then fl{0}[Week Start] else null,
                Since = since,
                Last  = List.Last(recs)[Week Start]
            ],
    A = Table.AddColumn(G, "S", each F([Rows])),
    E = Table.ExpandRecordColumn(Table.RemoveColumns(A, {"Rows"}), "S", {"Fut", "Since", "Last"})
in
    E
'''

M_WEEKS = r'''
let
    Source = #"ECN Weeks Base",
    Pick =
        Table.SelectColumns(
            Source,
            {"Line No", "ECN Number", "Old item", "New item", "Change", "Week No", "Week Start",
             "Week Label", "Week Key", "Week Rel", "Rel Order", "Rel N", "Balance", "Begin",
             "Demand", "Old Used", "New Used", "Mode"}
        ),
    Typed =
        Table.TransformColumnTypes(
            Pick,
            {
                {"Line No", Int64.Type}, {"ECN Number", type text}, {"Old item", type text},
                {"New item", type text}, {"Change", type text}, {"Week No", Int64.Type},
                {"Week Start", type date}, {"Week Label", type text}, {"Week Key", Int64.Type},
                {"Week Rel", type text}, {"Rel Order", Int64.Type}, {"Rel N", Int64.Type},
                {"Balance", type number}, {"Begin", type number}, {"Demand", type number},
                {"Old Used", type number}, {"New Used", type number}, {"Mode", type text}
            }
        )
in
    Typed
'''

M_CAL = r'''
let
    // Last 25 weeks + this week, Monday first. "Gần đây" marks the last 13 weeks.
    Monday = Date.StartOfWeek(TodayDate, Day.Monday),
    Starts = List.Transform(List.Numbers(-25, 26), each Date.AddDays(Monday, _ * 7)),
    T0     = Table.FromColumns({Starts}, {"Week Start"}),
    Typed  = Table.TransformColumnTypes(T0, {{"Week Start", type date}}),
    Lbl    = Table.AddColumn(
                 Typed, "Tuần",
                 each "W" & Text.PadStart(Text.From(fnIsoWeek([Week Start])), 2, "0"),
                 type text),
    Recent = Table.AddColumn(
                 Lbl, "Gần đây",
                 each if Duration.Days([Week Start] - Monday) >= -84 then "Yes" else "No",
                 type text)
in
    Recent
'''

M_ECN = r'''
let
    T   = TodayDate,
    TWk = fnIsoWeek(T),
    NB  = Table.Buffer(#"ECN Notes Base"),
    WB  = Table.Buffer(#"ECN Weeks Base"),
    LS  = Table.Buffer(#"Line Stock"),
    Monday = Date.StartOfWeek(T, Day.Monday),
    Grouped = Table.Group(#"ECN Base", {"ECN Number"}, {{"Rows", each _, type table}}),

    Build = (rs as table, ecn as text) as record =>
        let
            // ---- identity ---------------------------------------------------
            dates  = List.RemoveNulls(Table.Column(rs, "Issue Date")),
            issue  = if List.IsEmpty(dates) then null else List.Max(dates),
            models = List.Distinct(List.Combine(List.Transform(List.RemoveNulls(Table.Column(rs, "Model / Customer No.")), fnSplitList))),
            fgs    = List.Distinct(List.Combine(List.Transform(List.RemoveNulls(Table.Column(rs, "FG")), fnSplitList))),
            short  = (l as list) as text =>
                         Text.Combine(List.FirstN(l, 2), ", ")
                         & (if List.Count(l) > 2 then " +" & Text.From(List.Count(l) - 2) else ""),
            modelFg =
                if List.IsEmpty(models) and List.IsEmpty(fgs) then "—"
                else if List.IsEmpty(models) then "FG " & short(fgs)
                else if List.IsEmpty(fgs) or fgs = models then short(models)
                else short(models) & "  (FG " & short(fgs) & ")",

            // ---- status from its lines ("Cancel" lines are ignored) ----------
            stats    = List.Transform(Table.Column(rs, "Status"), each if _ = null then "" else _),
            act      = List.Select(stats, each _ <> "Cancel"),
            hasOn    = List.Contains(act, "Ongoing"),
            hasDone  = List.Contains(act, "Done"),
            hasBlank = List.Contains(act, ""),
            status   =
                if List.IsEmpty(act) then "Cancel"
                else if hasOn or (hasDone and hasBlank) then "Ongoing"
                else if hasBlank then "Chưa bắt đầu"
                else "Done",
            isOpen   = status = "Ongoing" or status = "Chưa bắt đầu",

            // ---- parsed Note lines, newest first ------------------------------
            ns       = Table.SelectRows(NB, each [ECN Number] = ecn),
            nsSorted = Table.Sort(ns, {{"Key", Order.Descending}, {"Seq", Order.Ascending}}),
            nsActive = Table.SelectRows(nsSorted, each [Category] <> "Đã xong"),
            top      = if Table.RowCount(nsActive) > 0 then nsActive{0}
                       else if Table.RowCount(nsSorted) > 0 then nsSorted{0}
                       else null,
            targets  = List.RemoveNulls(
                           List.Transform(Table.ToRecords(nsSorted),
                                          each if [Target] <> null then [Target] else [EstWk])),
            tgtRows  = Table.SelectRows(nsSorted, each [Target] <> null),
            owners   = List.Distinct(List.Combine(List.Transform(Table.Column(ns, "Who"), each List.Select(Text.Split(_, ", "), (x) => x <> "")))),
            stockL   = List.Distinct(List.Combine(List.Transform(Table.Column(ns, "Stock"), each List.Select(Text.Split(_, ", "), (x) => x <> "")))),
            finishL  = List.RemoveNulls(Table.Column(ns, "FinishBy")),
            finish   = if List.IsEmpty(finishL) then null else List.Max(finishL),
            lastOrd  = List.First(List.RemoveNulls(Table.Column(nsSorted, "LastOrder")), null),
            qty      = List.First(List.RemoveNulls(Table.Column(nsSorted, "Qty")), null),
            estNote  = List.First(
                           List.Transform(
                               Table.ToRecords(Table.SelectRows(nsSorted, each [EstWk] <> null or [EstTxt] <> null)),
                               each if [EstWk] <> null then "W" & Text.From([EstWk]) else [EstTxt]),
                           null),
            blocker  = if top = null then "Chưa có ghi chú" else top[Category],
            btext    = if top = null then null else Text.Start(top[Text], 110),
            updated  = if top = null then null
                       else if top[DateFrom] <> null then Date.ToText(top[DateFrom], "dd/MM")
                       else if top[WkFrom] <> null then "W" & Text.From(top[WkFrom])
                       else null,

            // ---- schedule: column est week + history found in Note ----------
            ests     = List.RemoveNulls(Table.Column(rs, "Est Wk")),
            est      = if List.IsEmpty(ests) then null else List.Min(ests),
            orig     = if List.IsEmpty(targets) then null else List.Last(targets),
            cur      = if est <> null then est else if List.IsEmpty(targets) then null else List.First(targets),
            moved    = orig <> null and cur <> null and orig <> cur,
            reason   = if Table.RowCount(tgtRows) = 0 then null
                       else Text.Start(Text.BeforeDelimiter(tgtRows{0}[Text], " - W", {0, RelativePosition.FromEnd}), 70),
            ecur     = fnEffWk(cur, TWk),

            state    =
                if status = "Done" then "ĐÃ HOÀN TẤT"
                else if status = "Cancel" then "ĐÃ HỦY"
                else if moved then "DỜI LỊCH"
                else if ecur <> null and ecur < TWk then "QUÁ HẠN"
                else if cur <> null then "ĐÚNG LỊCH"
                else "CHƯA CÓ LỊCH",

            badge    =
                if state = "DỜI LỊCH" then
                    "DỜI LỊCH: chuyển sang W" & Text.From(cur) & " do "
                    & (if reason = null or reason = "" then "(chưa ghi lý do)" else reason)
                else if state = "QUÁ HẠN" then "QUÁ HẠN: dự kiến W" & Text.From(cur) & ", chưa xong"
                else if state = "ĐÚNG LỊCH" then "ĐÚNG LỊCH · MP W" & Text.From(cur)
                else if state = "CHƯA CÓ LỊCH" then "CHƯA CÓ LỊCH MP"
                else if state = "ĐÃ HOÀN TẤT" then "ĐÃ HOÀN TẤT"
                else "ĐÃ HỦY",

            color    =
                if state = "DỜI LỊCH" then "#FBBF24"
                else if state = "QUÁ HẠN" then "#F87171"
                else if state = "ĐÚNG LỊCH" then "#4ADE80"
                else if state = "CHƯA CÓ LỊCH" then "#94A3B8"
                else if state = "ĐÃ HOÀN TẤT" then "#7DD3FC"
                else "#CBD5E1",

            prio     =
                if state = "DỜI LỊCH" then 1
                else if state = "QUÁ HẠN" then 2
                else if state = "CHƯA CÓ LỊCH" then 3
                else if state = "ĐÚNG LỊCH" then 4
                else 5,

            planText =
                if moved then "W" & Text.From(orig) & " → W" & Text.From(cur)
                else if cur <> null then "MP W" & Text.From(cur)
                else "MP: chưa có",

            // ---- scope: what the line must know about today ----------------
            recent   = issue <> null and Duration.Days(T - issue) <= 6,
            near     = ecur <> null and ecur >= TWk and ecur <= TWk + 1,
            scope    =
                if not isOpen then "Đã xong / Hủy"
                else if recent or near or state = "DỜI LỊCH" or state = "QUÁ HẠN" then "Áp dụng hôm nay"
                else "Đang mở khác",
            firstRun = if isOpen and ecur = TWk then "Yes" else "No",

            instrs   = List.RemoveNulls(Table.Column(rs, "EC Instruction")),
            instr    =
                if List.Contains(instrs, "Immediately") then "Immediately"
                else if List.IsEmpty(instrs) then "—"
                else instrs{0},

            // ---- what changes -------------------------------------------------
            nRep = Table.RowCount(Table.SelectRows(rs, each [Old item] <> null and [New item] <> null)),
            nAdd = Table.RowCount(Table.SelectRows(rs, each [New item] <> null and [Old item] = null)),
            nRem = Table.RowCount(Table.SelectRows(rs, each [Old item] <> null and [New item] = null)),
            mixL = List.Select(
                       {
                           if nRep > 0 then Text.From(nRep) & " thay" else null,
                           if nAdd > 0 then Text.From(nAdd) & " thêm" else null,
                           if nRem > 0 then Text.From(nRem) & " bỏ" else null
                       },
                       each _ <> null),
            mix      = Text.Combine(mixL, " · "),

            sorted   = Table.Sort(rs, {{"Change Rank", Order.Ascending}, {"Line No", Order.Ascending}}),
            items    =
                List.Distinct(
                    List.Transform(
                        Table.ToRecords(sorted),
                        each if [Description of Change] <> null and [Change] <> "—"
                             then Text.Start([Description of Change], 70) & " (" & [Change] & ")"
                             else if [Description of Change] <> null
                             then Text.Start([Description of Change], 70)
                             else [Change]
                    )
                ),
            nItems   = List.Count(items),
            summary  =
                (if Table.RowCount(rs) > 1 and mix <> "" then "[" & mix & "]  " else "")
                & Text.Combine(List.FirstN(items, 3), "  |  ")
                & (if nItems > 3 then "  |  +" & Text.From(nItems - 3) & " thay đổi khác" else ""),

            splitNN  = (col as text) =>
                List.Select(
                    List.Distinct(List.Combine(List.Transform(List.RemoveNulls(Table.Column(rs, col)),
                                                              each Text.Split(_, (if col = "Tags" then " · " else ", "))))),
                    each _ <> ""),
            tagAll   = splitNN("Tags"),
            compAll  = splitNN("Components"),
            tagText  = if List.IsEmpty(tagAll) then "Khác" else Text.Combine(tagAll, " · "),
            compText = Text.Combine(compAll, ", "),

            buyers   = List.RemoveNulls(Table.Column(rs, "Buyer")),
            buyerTxt = Text.Combine(
                           List.Transform(List.Distinct(buyers),
                                          (b) => b & " ×" & Text.From(List.Count(List.Select(buyers, (x) => x = b)))),
                           " · "),

            // ---- deadline / stock / paperwork --------------------------------
            dueTxt   = Text.Combine(
                           List.Select(
                               {
                                   if finish <> null then "Xong trước " & Date.ToText(finish, "dd/MM/yyyy") else null,
                                   if lastOrd <> null then "Đơn cuối " & lastOrd else null,
                                   if estNote <> null and est = null then "Dự kiến " & estNote else null
                               },
                               each _ <> null),
                           " · "),
            stockTxt = if List.IsEmpty(stockL) then ""
                       else "Xả tồn " & Text.Combine(stockL, ", ") & (if qty <> null then " · " & qty & " pcs" else ""),
            sideTxt  = Text.Combine(List.Select({dueTxt, stockTxt}, each _ <> ""), " · "),
            waitTxt  = if top = null then "—"
                       else blocker
                            & (if List.IsEmpty(owners) then "" else " · " & Text.Combine(owners, ", "))
                            & (if updated = null then "" else " · " & updated),
            noteCol  = if sideTxt = "" and btext = null then "—"
                       else if sideTxt = "" then btext
                       else if btext = null then sideTxt
                       else sideTxt & "  |  " & btext,

            sent     = not List.IsEmpty(List.RemoveNulls(Table.Column(rs, "ECN form sent date"))),
            recvd    = not List.IsEmpty(List.RemoveNulls(Table.Column(rs, "ECN form received date"))),
            doneL    = List.RemoveNulls(Table.Column(rs, "Date Completed")),

            // ---- weekly old-item stock -------------------------------------
            ls     = Table.SelectRows(LS, each [ECN Number] = ecn),
            futL   = List.RemoveNulls(ls[Fut]),
            sinceL = List.RemoveNulls(Table.SelectRows(ls, each [Fut] = null)[Since]),
            keepL  = Table.SelectRows(ls, each [Fut] = null and [Since] = null)[Last],
            wkN    = (d as date) as text => Text.From(fnIsoWeek(d)),
            nFut   = if List.IsEmpty(futL) then null
                     else Number.IntegerDivide(Duration.Days(List.Min(futL) - Monday), 7),
            runout =
                if Table.RowCount(ls) = 0 then "Chưa có dữ liệu tồn"
                else if not List.IsEmpty(futL) then
                    (if nFut = 0 then "Hết mã cũ tuần này"
                     else "Hết mã cũ W" & wkN(List.Min(futL)) & " · còn " & Text.From(nFut) & " tuần")
                else if not List.IsEmpty(keepL) and not List.IsEmpty(sinceL) then
                    "Đã hết " & Text.From(List.Count(sinceL)) & "/"
                    & Text.From(List.Count(sinceL) + List.Count(keepL)) & " mã cũ · "
                    & Text.From(List.Count(keepL)) & " mã còn tồn"
                else if not List.IsEmpty(keepL) then "Còn tồn mã cũ sau W" & wkN(List.Max(keepL))
                else "Đã hết mã cũ từ W" & wkN(List.Max(sinceL)),
            // same old item in two lines = same balance: count it once
            nowRows = Table.Distinct(
                          Table.SelectColumns(
                              Table.SelectRows(WB, each [ECN Number] = ecn and [Rel N] = 0),
                              {"Old item", "Demand"})),
            qtyNow  = if Table.RowCount(nowRows) = 0 then null else List.Sum(nowRows[Demand]),
            nowTxt  = if qtyNow = null then "Tuần này: chưa có số liệu"
                      else if qtyNow = 0 then "Tuần này: không dùng mã cũ"
                      else "Tuần này: " & Number.ToText(qtyNow, "#,0", "en-US") & " pcs",
            title   = ecn & "  ·  " & modelFg & (if instr = "Immediately" then "  ·  IMMEDIATELY" else ""),

            // ---- week the ECN runs: ISO week of Date; a Note line with Wxx / week xx + change moves it --
            planStart = if issue = null then null else Date.StartOfWeek(issue, Day.Monday),
            mv        = Table.Sort(
                            Table.SelectRows(ns, each [IsMove] = true),
                            {{"Cell", Order.Ascending}, {"Pos", Order.Ascending}}),
            mvWk      = List.RemoveNulls(mv[MoveWk]),
            moves     = Table.RowCount(mv),
            curStart  = if not List.IsEmpty(mvWk) then fnWeekStart(List.First(mvWk), T) else planStart,
            runNo     = if curStart = null then null else fnIsoWeek(curStart),
            runTxt    = if runNo = null then "—" else "W" & Text.PadStart(Text.From(runNo), 2, "0"),
            relD      = if curStart = null then null
                        else Number.IntegerDivide(Duration.Days(curStart - Monday), 7),
            rel       = if not isOpen then "Đã xong"
                        else if relD = null then "Chưa có lịch"
                        else if relD < 0 then "Tuần cũ chưa xong"
                        else if relD = 0 then "Tuần này"
                        else if relD = 1 then "Tuần sau"
                        else "Sắp tới",
            relOrd    = if rel = "Tuần này" then 1 else if rel = "Tuần sau" then 2 else if rel = "Sắp tới" then 3
                        else if rel = "Tuần cũ chưa xong" then 4 else if rel = "Chưa có lịch" then 5 else 6,
            runShow   = runTxt & (if moves > 0 then "  ·  dời " & Text.From(moves) & " lần" else ""),
            pad2      = (n as number) as text => "W" & Text.PadStart(Text.From(n), 2, "0"),
            histTxt   = if planStart = null then "—"
                        else if List.IsEmpty(mvWk) then "Đúng kế hoạch " & pad2(fnIsoWeek(planStart))
                        else pad2(fnIsoWeek(planStart))
                             & Text.Combine(List.Transform(List.Reverse(mvWk), each " → " & pad2(_)), ""),
            rawSort  = Table.Sort(ns, {{"Cell", Order.Ascending}, {"Pos", Order.Ascending}}),
            noteRaw  = if Table.RowCount(rawSort) = 0 then "—" else Text.Combine(rawSort[Line], "  |  "),
            fgShow   = if not List.IsEmpty(fgs) then short(fgs)
                       else if not List.IsEmpty(models) then "Model " & short(models)
                       else "—",

            // ---- lead time: date issued -> date completed ----------------------------
            compDate = if status = "Done" and not List.IsEmpty(doneL) then List.Max(doneL) else null,
            procDays = if compDate <> null and issue <> null then Duration.Days(compDate - issue) else null,
            bucketNo = if status = "Cancel" then 6
                       else if procDays = null then 5
                       else if procDays <= 7 then 1
                       else if procDays <= 14 then 2
                       else if procDays <= 30 then 3
                       else 4,
            bucket   = {"≤ 7 ngày", "8–14 ngày", "15–30 ngày", "> 30 ngày", "Chưa hoàn tất", "Đã hủy"}{bucketNo - 1},
            issueWk  = if issue = null then null else Date.StartOfWeek(issue, Day.Monday),
            doneWk   = if compDate = null then null else Date.StartOfWeek(compDate, Day.Monday),

            // ---- compact texts for the Log tab ------------------------------------------
            ecnShow  = ecn & (if moves > 0
                              then "  ·  dời từ " & (if planStart = null then "?" else pad2(fnIsoWeek(planStart)))
                              else ""),
            briefTxt = Text.Combine(List.Transform(List.FirstN(items, 2), each Text.Start(_, 60)), "  |  ")
                       & (if nItems > 2 then "  |  +" & Text.From(nItems - 2) & " khác" else "")
        in
            [
                #"Ngày ban hành"      = issue,
                #"Model / Customer No." = if List.IsEmpty(models) then "—" else Text.Combine(models, ", "),
                FG                    = if List.IsEmpty(fgs) then "—" else Text.Combine(fgs, ", "),
                #"FG Count"           = List.Count(fgs),
                #"Model / FG"         = modelFg,
                #"EC Instruction"     = instr,
                #"Thay đổi"           = summary,
                #"Line Mix"           = mix,
                #"Loại"               = tagText,
                #"Vật tư"             = compText,
                #"Loại / Vật tư"      = tagText & (if compText = "" then "" else "  |  " & compText),
                #"Nguồn VT"           = buyerTxt,
                #"Est Week No"        = cur,
                #"Original Week No"   = orig,
                #"Kế hoạch MP"        = planText,
                #"Schedule State"     = state,
                #"Trạng thái lịch"    = badge,
                #"Badge Color"        = color,
                Priority              = prio,
                Status                = status,
                Scope                 = scope,
                #"First Run"          = firstRun,
                #"Line Count"         = Table.RowCount(rs),
                #"Chờ gì"             = blocker,
                #"Đang chờ"           = waitTxt,
                #"Chi tiết chờ"       = btext,
                #"Hạn / tồn / ghi chú" = noteCol,
                #"Người phụ trách"    = Text.Combine(owners, ", "),
                #"Cập nhật cuối"      = updated,
                #"Hạn dự kiến"        = dueTxt,
                #"Tồn cần xả"         = stockTxt,
                #"Giấy ECN"           = if recvd then "Đã nhận" else if sent then "Đã gửi" else "Chưa gửi",
                #"Hoàn tất ngày"      = if List.IsEmpty(doneL) then null else List.Max(doneL),
                #"As Of Date"         = T,
                #"As Of Week"         = TWk,
                #"Card Title"         = title,
                #"Hết mã cũ"          = runout,
                #"SL tuần này"        = qtyNow,
                #"Tuần này"           = nowTxt,
                #"Kế hoạch / tồn mã cũ" = planText & "  |  " & runout & "  |  " & nowTxt,
                #"Đang chờ / ghi chú"   = (if noteCol = "—" then waitTxt else waitTxt & "  |  " & noteCol),
                #"Tuần chạy"          = runTxt,
                #"Tuần chạy No"       = runNo,
                #"Lịch tuần"          = runShow,
                #"Tuần so với nay"    = rel,
                #"Tuần Order"         = relOrd,
                #"Note Raw"           = noteRaw,
                #"FG hiển thị"        = fgShow,
                #"Ngày xử lý"         = procDays,
                #"Nhóm xử lý"         = bucket,
                #"Nhóm Order"         = bucketNo,
                #"Lịch sử dời"       = histTxt,
                #"Số lần dời"        = moves,
                #"ECN hiển thị"      = ecnShow,
                #"Thay đổi gọn"      = briefTxt,
                #"Ngày đóng"         = compDate,
                #"Tuần ban hành"      = issueWk,
                #"Tuần hoàn tất"      = doneWk
            ],

    Built    = Table.AddColumn(Grouped, "R", each Build([Rows], [ECN Number])),
    Dropped  = Table.RemoveColumns(Built, {"Rows"}),
    Fields   = {"Ngày ban hành", "Model / Customer No.", "FG", "FG Count", "Model / FG", "EC Instruction",
                "Thay đổi", "Line Mix", "Loại", "Vật tư", "Loại / Vật tư", "Nguồn VT",
                "Est Week No", "Original Week No", "Kế hoạch MP", "Schedule State", "Trạng thái lịch",
                "Badge Color", "Priority", "Status", "Scope", "First Run", "Line Count",
                "Chờ gì", "Đang chờ", "Chi tiết chờ", "Hạn / tồn / ghi chú", "Người phụ trách",
                "Cập nhật cuối", "Hạn dự kiến", "Tồn cần xả", "Giấy ECN", "Hoàn tất ngày",
                "As Of Date", "As Of Week", "Card Title", "Hết mã cũ", "SL tuần này", "Tuần này",
                "Kế hoạch / tồn mã cũ", "Đang chờ / ghi chú",
                "Tuần chạy", "Tuần chạy No", "Lịch tuần", "Tuần so với nay", "Tuần Order", "Note Raw",
                "FG hiển thị", "Ngày xử lý", "Nhóm xử lý", "Nhóm Order", "Lịch sử dời", "Số lần dời", "ECN hiển thị", "Thay đổi gọn", "Ngày đóng", "Tuần ban hành", "Tuần hoàn tất"},
    Expanded = Table.ExpandRecordColumn(Dropped, "R", Fields),
    Typed =
        Table.TransformColumnTypes(
            Expanded,
            {
                {"ECN Number", type text}, {"Ngày ban hành", type date},
                {"Model / Customer No.", type text}, {"FG", type text}, {"FG Count", Int64.Type},
                {"Model / FG", type text}, {"EC Instruction", type text},
                {"Thay đổi", type text}, {"Line Mix", type text}, {"Loại", type text},
                {"Vật tư", type text}, {"Loại / Vật tư", type text}, {"Nguồn VT", type text},
                {"Est Week No", Int64.Type}, {"Original Week No", Int64.Type},
                {"Kế hoạch MP", type text}, {"Schedule State", type text},
                {"Trạng thái lịch", type text}, {"Badge Color", type text},
                {"Priority", Int64.Type}, {"Status", type text}, {"Scope", type text},
                {"First Run", type text}, {"Line Count", Int64.Type},
                {"Chờ gì", type text}, {"Đang chờ", type text}, {"Chi tiết chờ", type text},
                {"Hạn / tồn / ghi chú", type text}, {"Người phụ trách", type text},
                {"Cập nhật cuối", type text}, {"Hạn dự kiến", type text}, {"Tồn cần xả", type text},
                {"Giấy ECN", type text}, {"Hoàn tất ngày", type date},
                {"As Of Date", type date}, {"As Of Week", Int64.Type},
                {"Card Title", type text}, {"Hết mã cũ", type text},
                {"SL tuần này", type number}, {"Tuần này", type text},
                {"Kế hoạch / tồn mã cũ", type text}, {"Đang chờ / ghi chú", type text},
                {"Tuần chạy", type text}, {"Tuần chạy No", Int64.Type}, {"Lịch tuần", type text},
                {"Tuần so với nay", type text}, {"Tuần Order", Int64.Type}, {"Note Raw", type text},
                {"FG hiển thị", type text}, {"Ngày xử lý", Int64.Type}, {"Nhóm xử lý", type text},
                {"Nhóm Order", Int64.Type}, {"Lịch sử dời", type text}, {"Số lần dời", Int64.Type},
                {"ECN hiển thị", type text}, {"Thay đổi gọn", type text},
                {"Ngày đóng", type date}, {"Tuần ban hành", type date},
                {"Tuần hoàn tất", type date}
            }
        )
in
    Typed
'''

# ==========================================================================
# MEASURES
# ==========================================================================
INT = '#,0'
MEASURES = [
    ("ECN áp dụng hôm nay", INT, '''
CALCULATE ( COUNTROWS ( 'ECN' ), 'ECN'[Scope] = "Áp dụng hôm nay" ) + 0
''', "Open ECNs the line must know about today: issued in the last 7 days, MP planned "
     "this or next week, or schedule moved / overdue."),
    ("ECN chạy lần đầu tuần này", INT, '''
CALCULATE ( COUNTROWS ( 'ECN' ), 'ECN'[First Run] = "Yes" ) + 0
''', "Open ECNs whose MP week is the current week."),
    ("Cảnh báo dời lịch", INT, '''
CALCULATE (
    COUNTROWS ( 'ECN' ),
    'ECN'[Schedule State] IN { "DỜI LỊCH", "QUÁ HẠN" }
) + 0
''', "Open ECNs with a moved plan (history in Note) or an MP week already passed."),
    ("ECN đang mở", INT, '''
CALCULATE (
    COUNTROWS ( 'ECN' ),
    'ECN'[Status] IN { "Ongoing", "Chưa bắt đầu" }
) + 0
''', "All ECNs not yet Done or Cancelled."),
    ("Badge Fill", None, '''
SELECTEDVALUE ( 'ECN'[Badge Color], "#94A3B8" )
''', "Field-value conditional format for the schedule badge cell."),
    ("Today Label", None, '''
VAR D = CALCULATE ( MAX ( 'ECN'[As Of Date] ), ALL ( 'ECN' ) )
VAR W = CALCULATE ( MAX ( 'ECN'[As Of Week] ), ALL ( 'ECN' ) )
RETURN
    FORMAT ( D, "dddd dd/MM/yyyy" ) & "   ·   W" & W
''', "Date of the last refresh in Vietnam time and its ISO week."),
]

NEW_MEASURES = [
    ("Số ECN", INT, '''
COUNTROWS ( 'ECN' )
''', "ECN count in the current filter."),
    ("FG đang mở", INT, '''
CALCULATE (
    DISTINCTCOUNT ( 'ECN FG'[FG Code] ),
    'ECN'[Status] IN { "Ongoing", "Chưa bắt đầu" },
    'ECN FG'[FG Code] <> "—"
)
''', "Distinct FG codes of the open ECNs in the filter."),
    ("Giấy ECN đã gửi", INT, '''
CALCULATE ( COUNTROWS ( 'ECN' ), 'ECN'[Giấy ECN] IN { "Đã gửi", "Đã nhận" } ) + 0
''', "ECNs whose ECN form is sent or received."),
    ("Danh sách Model", None, '''
VAR m = VALUES ( 'ECN Models'[Model] )
VAR n = COUNTROWS ( m )
VAR t = TOPN ( 6, m, 'ECN Models'[Model], ASC )
RETURN
    IF (
        n = 0,
        "—",
        CONCATENATEX ( t, 'ECN Models'[Model], ", ", 'ECN Models'[Model], ASC )
            & IF ( n > 6, "  … +" & ( n - 6 ), "" )
    )
''', "First 6 models of the ECNs in the filter."),
    ("Ngày xử lý (trung vị)", "0", '''
MEDIAN ( 'ECN'[Ngày xử lý] )
''', "Median days from date issued to date completed (Done ECNs with both dates)."),
    ("Top 10 Model", INT, '''
VAR curN = [Số ECN]
VAR cur = SELECTEDVALUE ( 'ECN Models'[Model] )
VAR tbl = ADDCOLUMNS ( ALLSELECTED ( 'ECN Models'[Model] ), "@n", [Số ECN] )
VAR rk =
    COUNTROWS (
        FILTER ( tbl, [@n] > curN || ( [@n] = curN && 'ECN Models'[Model] <= cur ) )
    )
RETURN
    IF ( NOT ISBLANK ( curN ) && rk <= 10, curN )
''', "ECN count, only for the 10 models with most ECNs (ties broken by name)."),
    ("Top 5 FG", INT, '''
VAR curN = [Số ECN]
VAR cur = SELECTEDVALUE ( 'ECN FG'[FG Code] )
VAR tbl =
    ADDCOLUMNS (
        FILTER ( ALLSELECTED ( 'ECN FG'[FG Code] ), 'ECN FG'[FG Code] <> "—" ),
        "@n", [Số ECN]
    )
VAR rk =
    COUNTROWS (
        FILTER ( tbl, [@n] > curN || ( [@n] = curN && 'ECN FG'[FG Code] <= cur ) )
    )
RETURN
    IF ( cur <> "—" && NOT ISBLANK ( curN ) && rk <= 5, curN )
''', "ECN count, only for the 5 FG codes with most ECNs (ties broken by code)."),
    ("Tiêu đề tuần này", None, '''
VAR d = CALCULATE ( MAX ( 'ECN'[As Of Date] ), ALL ( 'ECN' ) )
VAR n = CALCULATE ( COUNTROWS ( 'ECN' ), 'ECN'[Tuần so với nay] = "Tuần này" ) + 0
RETURN
    "TUẦN NÀY  ·  W" & FORMAT ( WEEKNUM ( d, 21 ), "00" ) & "  ·  " & n & " ECN"
''', "Header of the left list of the Log tab."),
    ("Tiêu đề tuần sau", None, '''
VAR d = CALCULATE ( MAX ( 'ECN'[As Of Date] ), ALL ( 'ECN' ) )
VAR n = CALCULATE ( COUNTROWS ( 'ECN' ), 'ECN'[Tuần so với nay] = "Tuần sau" ) + 0
RETURN
    "TUẦN SAU  ·  W" & FORMAT ( WEEKNUM ( d + 7, 21 ), "00" ) & "  ·  " & n & " ECN"
''', "Header of the right list of the Log tab."),
    ("ECN ban hành", INT, '''
CALCULATE (
    COUNTROWS ( 'ECN' ),
    TREATAS ( VALUES ( 'ECN Calendar'[Week Start] ), 'ECN'[Tuần ban hành] )
) + 0
''', "ECNs issued in the week."),
    ("ECN hoàn tất", INT, '''
CALCULATE (
    COUNTROWS ( 'ECN' ),
    TREATAS ( VALUES ( 'ECN Calendar'[Week Start] ), 'ECN'[Tuần hoàn tất] )
) + 0
''', "ECNs completed in the week."),
    ("Ngày xử lý theo tuần", "0", '''
CALCULATE (
    MEDIAN ( 'ECN'[Ngày xử lý] ),
    TREATAS ( VALUES ( 'ECN Calendar'[Week Start] ), 'ECN'[Tuần hoàn tất] )
)
''', "Median days to complete, by week of completion."),
    ("Instr Fill", None, '''
IF ( SELECTEDVALUE ( 'ECN'[EC Instruction] ) = "Immediately", "#FBD5D1", "#FFFFFF" )
''', "Cell colour for EC Instruction = Immediately."),
]

WEEK_MEASURES = [
    ("SL nhu cầu", INT, '''
SUMX (
    SUMMARIZE ( 'ECN Weeks', 'ECN Weeks'[Old item], 'ECN Weeks'[Week Start] ),
    CALCULATE ( MAX ( 'ECN Weeks'[Demand] ) )
)
''', "Old-item quantity used in the week. The same old item in several ECN lines is counted once."),
    ("SL dùng mã cũ", INT, '''
SUMX (
    SUMMARIZE ( 'ECN Weeks', 'ECN Weeks'[Old item], 'ECN Weeks'[Week Start] ),
    CALCULATE ( MAX ( 'ECN Weeks'[Old Used] ) )
)
''', "Part of the usage still covered by old-item stock."),
    ("SL mã mới / thiếu", INT, '''
SUMX (
    SUMMARIZE ( 'ECN Weeks', 'ECN Weeks'[Old item], 'ECN Weeks'[Week Start] ),
    CALCULATE ( MAX ( 'ECN Weeks'[New Used] ) )
)
''', "Part of the usage beyond old-item stock: runs on the new item (or is short if there is none)."),
    ("ECN có nhu cầu", INT, '''
CALCULATE ( DISTINCTCOUNT ( 'ECN Weeks'[ECN Number] ), 'ECN Weeks'[Demand] > 0 ) + 0
''', "ECNs whose old item is used in the selected week(s)."),
    ("SL tuần (pcs)", INT, '''
VAR x = [SL nhu cầu]
RETURN IF ( x > 0, x )
''', "Blank when 0 so cards only show lines that run."),
    ("Tồn mã cũ cuối tuần", INT, '''
IF (
    [SL nhu cầu] > 0,
    SUMX (
        SUMMARIZE ( 'ECN Weeks', 'ECN Weeks'[Old item], 'ECN Weeks'[Week Start] ),
        CALCULATE ( MAX ( 'ECN Weeks'[Balance] ) )
    )
)
''', "Old-item balance at the end of the week; negative = short / running the new item."),
]

NOTE_MEASURES = [
    ("Dòng ghi chú", INT, '''
COUNTROWS ( 'ECN Notes' )
''', "Parsed note lines."),
    ("Dòng chưa nhận dạng", INT, '''
CALCULATE ( COUNTROWS ( 'ECN Notes' ), 'ECN Notes'[Phân loại] = "Chưa nhận dạng" ) + 0
''', "Note lines no rule recognised - rewrite them following the entry rules."),
    ("Dòng dời lịch", INT, '''
CALCULATE ( COUNTROWS ( 'ECN Notes' ), 'ECN Notes'[Loại sự kiện] = "Dời lịch" ) + 0
''', "Note lines that move the schedule."),
    ("ECN có dời lịch", INT, '''
CALCULATE (
    DISTINCTCOUNT ( 'ECN Notes'[ECN Number] ),
    'ECN Notes'[Loại sự kiện] = "Dời lịch"
) + 0
''', "ECNs with at least one schedule move."),
    ("Event Fill", None, '''
IF ( SELECTEDVALUE ( 'ECN Notes'[Loại sự kiện] ) = "Dời lịch", "#FFE3A3", "#FFFFFF" )
''', "Cell colour for schedule-move lines."),
    ("Dòng đã nhận dạng %", '0%', '''
DIVIDE ( [Dòng ghi chú] - [Dòng chưa nhận dạng], [Dòng ghi chú] )
''', "Share of note lines the parser understood."),
]

# ==========================================================================
# MODEL.BIM
# ==========================================================================
def col(name, dtype, fmt=None, hidden=False, sort=None, desc=None):
    c = {"name": name, "dataType": dtype, "lineageTag": tag(),
         "summarizeBy": "none", "sourceColumn": name,
         "annotations": [{"name": "SummarizationSetBy", "value": "Automatic"}]}
    if fmt: c["formatString"] = fmt
    if hidden: c["isHidden"] = True
    if sort: c["sortByColumn"] = sort
    if desc: c["description"] = desc
    if dtype == "dateTime":
        c["annotations"].append({"name": "UnderlyingDateTimeDataType", "value": "Date"})
    return c


S, I, D = "string", "int64", "dateTime"
ecn_cols = [
    col("ECN Number", S), col("Ngày ban hành", D, "dd/MM/yyyy"),
    col("Model / Customer No.", S), col("FG", S), col("FG Count", I, INT),
    col("Model / FG", S, desc="Card title."), col("EC Instruction", S),
    col("Thay đổi", S, desc="Top 3 changes + line mix."), col("Line Mix", S),
    col("Loại", S, desc="Bao bì / Cơ khí / Mạch / Nhãn."), col("Vật tư", S, desc="Part that changes."),
    col("Loại / Vật tư", S), col("Nguồn VT", S, desc="Buyer column: Local / Import / In-house."),
    col("Est Week No", I, "0", hidden=True), col("Original Week No", I, "0", hidden=True),
    col("Kế hoạch MP", S), col("Schedule State", S),
    col("Trạng thái lịch", S, sort="Priority", desc="The badge text."),
    col("Badge Color", S, hidden=True), col("Priority", I, "0", hidden=True),
    col("Status", S), col("Scope", S), col("First Run", S), col("Line Count", I, INT),
    col("Chờ gì", S, desc="Category of the newest open note line."),
    col("Đang chờ", S), col("Chi tiết chờ", S), col("Hạn / tồn / ghi chú", S),
    col("Người phụ trách", S), col("Cập nhật cuối", S), col("Hạn dự kiến", S),
    col("Tồn cần xả", S), col("Giấy ECN", S), col("Hoàn tất ngày", D, "dd/MM/yyyy"),
    col("As Of Date", D, "yyyy-mm-dd", hidden=True), col("As Of Week", I, "0", hidden=True),
    col("Card Title", S, desc="ECN · model (FG) · IMMEDIATELY flag - first field of the cards."),
    col("Hết mã cũ", S, desc="When the old item runs out, from the weekly balance columns."),
    col("SL tuần này", "double", INT, desc="Old-item usage this week, each old item counted once."),
    col("Tuần này", S),
    col("Kế hoạch / tồn mã cũ", S), col("Đang chờ / ghi chú", S),
    col("Tuần chạy", S, desc="ISO week of Date, or the week of the newest Note line with Wxx/week xx + change."),
    col("Tuần chạy No", I, "0", hidden=True),
    col("Lịch tuần", S, desc="Running week + how many Note lines moved the schedule."),
    col("Tuần so với nay", S, sort="Tuần Order",
        desc="Tuần này / Tuần sau / Sắp tới / Tuần cũ chưa xong / Chưa có lịch / Đã xong."),
    col("Tuần Order", I, "0", hidden=True),
    col("Note Raw", S, desc="All Note lines, newest on top, joined with |."),
    col("FG hiển thị", S), col("Ngày xử lý", I, INT, desc="Date completed - date issued (Done ECNs)."),
    col("Nhóm xử lý", S, sort="Nhóm Order"), col("Nhóm Order", I, "0", hidden=True),
    col("Lịch sử dời", S, desc="Plan week (from Date) -> each move found in the Note, oldest first."),
    col("Số lần dời", I, INT),
    col("ECN hiển thị", S, desc="ECN number + 'dời từ Wxx' when the schedule moved."),
    col("Thay đổi gọn", S, desc="First 2 changes (60 chars each) + '+n khác'."),
    col("Ngày đóng", D, "dd/MM/yyyy", desc="Date completed of a Done ECN."),
    col("Tuần ban hành", D, "dd/MM/yyyy", hidden=True), col("Tuần hoàn tất", D, "dd/MM/yyyy", hidden=True),
]
cal_cols = [col("Week Start", D, "dd/MM/yyyy"), col("Tuần", S, sort="Week Start"),
            col("Gần đây", S, desc="Yes = last 13 weeks.")]
wks_cols = [
    col("Line No", I, "0", hidden=True), col("ECN Number", S), col("Old item", S),
    col("New item", S), col("Change", S), col("Week No", I, "0"),
    col("Week Start", D, "dd/MM/yyyy"), col("Week Label", S, sort="Week Key"),
    col("Week Key", I, "0", hidden=True), col("Week Rel", S, sort="Rel Order"),
    col("Rel Order", I, "0", hidden=True), col("Rel N", I, "0", hidden=True),
    col("Balance", "double", INT, desc="Old-item balance at the end of the week (sheet columns V..BM)."),
    col("Begin", "double", INT, hidden=True),
    col("Demand", "double", INT, desc="Drop of the balance = old-item quantity used that week."),
    col("Old Used", "double", INT), col("New Used", "double", INT), col("Mode", S),
]
lin_cols = [col("Line No", I, "0", hidden=True), col("ECN Number", S), col("Action", S),
            col("Old item", S), col("New item", S), col("Mã cũ (mô tả)", S),
            col("Mã mới (mô tả)", S), col("Mã liên quan", S), col("Description of Change", S),
            col("Loại", S), col("Vật tư", S), col("Buyer", S), col("Line Status", S)]
not_cols = [col("ECN Number", S), col("Seq", I, "0", hidden=True), col("Line", S),
            col("Nội dung", S), col("Cập nhật", S), col("Người phụ trách", S), col("NCC", S),
            col("Phân loại", S), col("Hạn / Est", S), col("Tồn cần xả", S),
            col("Key", "double", "0", hidden=True),
            col("Loại sự kiện", S, desc="Dời lịch = line has Wxx/week xx AND change; otherwise Khác."),
            col("Dời sang", S, desc="Week the line moves the schedule to.")]
mod_cols = [col("ECN Number", S), col("Model", S)]
fg_cols = [col("ECN Number", S), col("FG Code", S),
           col("FG / ECN", S, desc="Search box: type an FG code, pick the ECN.")]


def mk_measures(lst):
    out = []
    for mname, fmt, expr, desc in lst:
        m = {"name": mname, "expression": L(expr), "lineageTag": tag(), "description": desc}
        if fmt: m["formatString"] = fmt
        out.append(m)
    return out


def table(name, cols, m, desc, measures=None):
    t = {"name": name, "lineageTag": tag(), "description": desc, "columns": cols,
         "partitions": [{"name": name, "mode": "import",
                         "source": {"type": "m", "expression": L(m)}}],
         "annotations": [{"name": "PBI_ResultType", "value": "Table"}]}
    if measures: t["measures"] = measures
    return t


tables_ = [
    table(ECN, ecn_cols, M_ECN, "One row per ECN Number.", mk_measures(MEASURES + NEW_MEASURES)),
    table(LIN, lin_cols, M_LINES, "One row per line of the ECN sheet."),
    table(NOT, not_cols, M_NOTES, "One row per parsed line of the Note column.",
          mk_measures(NOTE_MEASURES)),
    table(MOD, mod_cols, M_MODELS, "Bridge: ECN Number x model token of Model / Customer No."),
    table(FGT, fg_cols, M_FG, "Bridge: ECN Number x FG code token."),
    table(WKS, wks_cols, M_WEEKS, "One row per ECN line x week: old-item balance and usage.",
          mk_measures(WEEK_MEASURES)),
    table(CAL, cal_cols, M_CAL, "Last 12 weeks + this week (weekly charts)."),
]


def param(name, value, required=True):
    v = "null" if value is None else '"%s"' % value
    return {"name": name, "kind": "m",
            "expression": ['%s meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=%s]'
                           % (v, "true" if required else "false")],
            "lineageTag": tag(), "queryGroup": "Parameters",
            "annotations": [{"name": "PBI_ResultType", "value": "Text"}]}


def fn(name, expr, rtype="Function"):
    return {"name": name, "kind": "m", "expression": L(expr), "lineageTag": tag(),
            "queryGroup": "Functions",
            "annotations": [{"name": "PBI_ResultType", "value": rtype}]}


FUNCS = [("fnIsoWeek", M_FN_ISOWEEK), ("fnWeekNo", M_FN_WEEKNO), ("fnPlanWk", M_FN_PLANWK),
         ("fnEffWk", M_FN_EFFWK), ("fnIsDigits", M_FN_ISDIGITS), ("fnTokens", M_FN_TOKENS),
         ("fnSplitList", M_FN_SPLITLIST), ("fnCodeKind", M_FN_CODEKIND), ("fnTags", M_FN_TAGS),
         ("fnComps", M_FN_COMPS), ("fnParseDesc", M_FN_PARSEDESC), ("fnEstWk", M_FN_ESTWK),
         ("fnWeekTokens", M_FN_WEEKTOKENS), ("fnWeekStart", M_FN_WEEKSTART),
         ("fnNoteLine", M_FN_NOTELINE)]

expressions = [
    param("pSiteUrl", "https://wahlclippercorporation.sharepoint.com/sites/NPIInformation"),
    param("pFolderPath", "General/Sharing to Production, Planner and Buyer"),
    param("pFileName", "ECN Log.xlsx"), param("pSheetName", "ECN"),
    param("pTodayOverride", None, required=False),
    param("pPeople", ", ".join(P.PEOPLE)), param("pVendors", ", ".join(P.VENDORS)),
    {"name": "TodayDate", "kind": "m", "expression": L(M_TODAY), "lineageTag": tag(),
     "queryGroup": "Functions", "annotations": [{"name": "PBI_ResultType", "value": "Date"}]},
] + [fn(n, m) for n, m in FUNCS] + [
    fn("ECN Raw", M_RAW, "Table"), fn("Week Map", M_WEEKMAP, "Table"),
    fn("ECN Base", M_BASE, "Table"), fn("ECN Notes Base", M_NOTES_BASE, "Table"),
    fn("ECN Weeks Base", M_WEEKS_BASE, "Table"), fn("Line Stock", M_LINESTOCK, "Table"),
]

rels = []
for child, extra in [(LIN, None), (NOT, None), (WKS, None), (MOD, "bothDirections"), (FGT, "bothDirections")]:
    r = {"name": tag(), "fromTable": child, "fromColumn": "ECN Number",
         "toTable": ECN, "toColumn": "ECN Number"}
    if extra: r["crossFilteringBehavior"] = extra
    rels.append(r)

model_bim = {
    "name": "SemanticModel", "compatibilityLevel": 1600,
    "model": {
        "culture": "en-US",
        "dataAccessOptions": {"legacyRedirects": True, "returnErrorValuesAsNull": True},
        "defaultPowerBIDataSourceVersion": "powerBI_V3",
        "sourceQueryCulture": "en-US",
        "tables": tables_, "relationships": rels, "expressions": expressions,
        "queryGroups": [{"folder": "Parameters", "annotations": []},
                        {"folder": "Functions", "annotations": []}],
        "annotations": [
            {"name": "PBI_QueryOrder",
             "value": json.dumps([e["name"] for e in expressions] + [ECN, LIN, NOT, MOD, FGT, WKS, CAL])},
            {"name": "__PBI_TimeIntelligenceEnabled", "value": "0"},
        ],
    },
}

w(os.path.join(SM, "model.bim"), model_bim)
# Desktop keeps the last loaded model in .pbi/cache.abf; a stale cache built by an
# older run (or at another compatibility level) blocks opening. It is rebuilt on refresh.
_cache = os.path.join(SM, ".pbi", "cache.abf")
if os.path.exists(_cache):
    os.remove(_cache)
w(os.path.join(SM, "definition.pbism"), {
    "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/semanticModel/definitionProperties/1.0.0/schema.json",
    "version": "4.2", "settings": {}})
w(os.path.join(SM, ".platform"), {
    "$schema": "https://developer.microsoft.com/json-schemas/fabric/gitIntegration/platformProperties/2.0.0/schema.json",
    "metadata": {"type": "SemanticModel", "displayName": NAME},
    "config": {"version": "2.0", "logicalId": tag()}})

# ==========================================================================
# REPORT  -  2 tabs as in the slide: Log (hôm nay chạy gì) + Performance
# ==========================================================================
VS = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/2.9.0/schema.json"
BG, PANEL, EDGE, TXT, MUTE = "#F3F5F9", "#FFFFFF", "#DDE2EA", "#1C2430", "#5B6575"
BRAND, YEL, BLUE, AMBER, GREEN, RED = "#3F4246", "#F2C230", "#1F5FBF", "#C77700", "#16803C", "#D1352B"
LOGO_SRC = os.path.join(ROOT, "assets", "img", "wahl-logo.png")
LOGO = "wahl_logo.png"


def lit(v): return {"expr": {"Literal": {"Value": v}}}
def solid(h): return {"solid": {"color": lit("'%s'" % h)}}


def meas(name, entity=ECN, display=None):
    p = {"field": {"Measure": {"Expression": {"SourceRef": {"Entity": entity}}, "Property": name}},
         "queryRef": "%s.%s" % (entity, name), "nativeQueryRef": name}
    if display: p["displayName"] = display
    return p


def column_(name, entity=ECN, active=False, display=None):
    p = {"field": {"Column": {"Expression": {"SourceRef": {"Entity": entity}}, "Property": name}},
         "queryRef": "%s.%s" % (entity, name), "nativeQueryRef": name}
    if active: p["active"] = True
    if display: p["displayName"] = display
    return p


def colref(entity, prop):
    return {"Column": {"Expression": {"SourceRef": {"Entity": entity}}, "Property": prop}}


def measref(entity, prop):
    return {"Measure": {"Expression": {"SourceRef": {"Entity": entity}}, "Property": prop}}


def container(title=None, fill=PANEL, border=True):
    base = {"background": [{"properties": {"show": lit("true"), "color": solid(fill),
                                           "transparency": lit("0L")}}],
            "border": [{"properties": {"show": lit("true" if border else "false"),
                                       "color": solid(EDGE), "radius": lit("10L")}}]}
    if title:
        base["title"] = [{"properties": {"show": lit("true"), "text": lit("'%s'" % title),
                                         "fontSize": lit("16D"), "bold": lit("true"),
                                         "fontColor": solid(TXT)}}]
    return base


NO_BG = {"background": [{"properties": {"show": lit("false")}}],
         "border": [{"properties": {"show": lit("false")}}]}

VISUALS = []


def visual(page, vid, vtype, x, y, wd, ht, z, query=None, objects=None, vco=None, sort=None,
           filter_config=None):
    v = {"$schema": VS, "name": vid,
         "position": {"x": x, "y": y, "z": z, "height": ht, "width": wd, "tabOrder": z},
         "visual": {"visualType": vtype}}
    if query:
        v["visual"]["query"] = {"queryState": query}
        if sort:
            v["visual"]["query"]["sortDefinition"] = {"sort": sort}
    if objects: v["visual"]["objects"] = objects
    if vco: v["visual"]["visualContainerObjects"] = vco
    v["visual"]["drillFilterOtherVisuals"] = True
    if filter_config: v["filterConfig"] = filter_config
    VISUALS.append((page, vid, v))


def only_filter(entity, prop, value, name):
    return {"filters": [{
        "name": name, "field": colref(entity, prop), "type": "Categorical",
        "filter": {"Version": 2, "From": [{"Name": "t", "Entity": entity, "Type": 0}],
                   "Where": [{"Condition": {"In": {
                       "Expressions": [{"Column": {"Expression": {"SourceRef": {"Source": "t"}},
                                                   "Property": prop}}],
                       "Values": [[{"Literal": {"Value": "'%s'" % value}}]]}}}]}}]}


def textbox(page, vid, x, y, wd, ht, z, runs, fill=None):
    v = {"$schema": VS, "name": vid,
         "position": {"x": x, "y": y, "z": z, "height": ht, "width": wd, "tabOrder": z},
         "visual": {"visualType": "textbox",
                    "objects": {"general": [{"properties": {"paragraphs": runs}}]},
                    "drillFilterOtherVisuals": True}}
    if fill == "transparent":
        v["visual"]["visualContainerObjects"] = NO_BG
    elif fill:
        v["visual"]["visualContainerObjects"] = {
            "background": [{"properties": {"show": lit("true"), "color": solid(fill),
                                           "transparency": lit("0L")}}]}
    VISUALS.append((page, vid, v))


def para(text, size="14pt", color=TXT, bold=False):
    ts = {"fontSize": size, "color": color}
    if bold: ts["fontWeight"] = "bold"
    return {"textRuns": [{"value": text, "textStyle": ts}]}


def band(page, p, title, sub):
    """dark Wahl band + yellow strip + logo + date/week."""
    textbox(page, p + "_band", 0, 0, 1920, 96, 1, [para(" ", "10pt", BRAND)], fill=BRAND)
    textbox(page, p + "_strip", 0, 96, 1920, 6, 2, [para(" ", "4pt", YEL)], fill=YEL)
    textbox(page, p + "_title", 24, 4, 1130, 90, 10, [
        para(title, "26pt", "#FFFFFF", True), para(sub, "12pt", "#D5D8DD")], fill="transparent")
    visual(page, p + "_today", "card", 1170, 18, 460, 62, 11,
        query={"Values": {"projections": [meas("Today Label")]}},
        objects={"labels": [{"properties": {"color": solid(YEL), "fontSize": lit("20L"),
                                            "bold": lit("true")}}],
                 "categoryLabels": [{"properties": {"show": lit("false")}}]},
        vco=NO_BG)
    VISUALS.append((page, p + "_logo", {
        "$schema": VS, "name": p + "_logo",
        "position": {"x": 1660, "y": 14, "z": 12, "height": 68, "width": 228, "tabOrder": 12},
        "visual": {"visualType": "image",
                   "objects": {
                       "general": [{"properties": {"imageUrl": {"expr": {"ResourcePackageItem": {
                           "PackageName": "RegisteredResources", "PackageType": 1,
                           "ItemName": LOGO}}}}}],
                       "imageScaling": [{"properties": {"imageScalingType": lit("'Fit'")}}]},
                   "drillFilterOtherVisuals": True}}))


def kpi(page, vid, measure_name, x, y, wd, ht, z, color, entity=ECN, size="46L", label=None):
    visual(page, vid, "card", x, y, wd, ht, z,
        query={"Values": {"projections": [meas(measure_name, entity, label)]}},
        objects={
            "labels": [{"properties": {"color": solid(color), "fontSize": lit(size),
                                       "bold": lit("true"), "labelDisplayUnits": lit("1D")}}],
            "categoryLabels": [{"properties": {"show": lit("true"), "color": solid(MUTE),
                                               "fontSize": lit("17L")}}],
        },
        vco=container())


def slicer(page, vid, entity, prop, x, y, wd, ht, z, title, dropdown=False,
           horizontal=False, default=None, search=False):
    objs = {"header": [{"properties": {"show": lit("false")}}],
            "items": [{"properties": {"textSize": lit("16D"), "fontColor": solid(TXT),
                                      "background": solid(PANEL)}}]}
    if dropdown: objs["data"] = [{"properties": {"mode": lit("'Dropdown'")}}]
    general = {}
    if horizontal:
        general["orientation"] = lit("1L")
    if search:
        general["selfFilterEnabled"] = lit("true")
    if default:
        vals = default if isinstance(default, list) else [default]
        general["filter"] = {"filter": {
            "Version": 2, "From": [{"Name": "t", "Entity": entity, "Type": 0}],
            "Where": [{"Condition": {"In": {
                "Expressions": [{"Column": {"Expression": {"SourceRef": {"Source": "t"}},
                                            "Property": prop}}],
                "Values": [[{"Literal": {"Value": "'%s'" % v}}] for v in vals]}}}]}}
    if general:
        objs["general"] = [{"properties": general}]
    visual(page, vid, "slicer", x, y, wd, ht, z,
           query={"Values": {"projections": [column_(prop, entity, active=True)]}},
           objects=objs, vco=container(title))


def clear_button(page, vid, x, y, wd, ht, z):
    VISUALS.append((page, vid, {
        "$schema": VS, "name": vid,
        "position": {"x": x, "y": y, "z": z, "height": ht, "width": wd, "tabOrder": z},
        "visual": {
            "visualType": "actionButton",
            "objects": {
                "icon": [{"properties": {"shapeType": lit("'blank'")}}],
                "text": [{"properties": {"show": lit("true")}},
                         {"properties": {"text": lit("'XÓA LỌC'"), "fontColor": solid("#FFFFFF"),
                                         "fontSize": lit("16D"), "bold": lit("true"),
                                         "horizontalAlignment": lit("'Center'")},
                          "selector": {"id": "default"}}],
                "fill": [{"properties": {"show": lit("true")}},
                         {"properties": {"fillColor": solid(BRAND), "transparency": lit("0D")},
                          "selector": {"id": "default"}}],
                "outline": [{"properties": {"show": lit("false")}}],
            },
            "visualContainerObjects": {
                "visualLink": [{"properties": {"show": lit("true"), "type": lit("'ClearAllSlicers'")}}],
            },
            "drillFilterOtherVisuals": True}}))


def loglist(page, vid, x, y, wd, ht, z, cols, sort=None, fsize="17D", fill_col=None,
            big="19D", head="15D", pad="9D", filter_config=None):
    """table styled as a list: wraps long text, no vertical grid, instruction cell coloured."""
    widths = [("%s.%s" % (e, c), wdt) for e, c, wdt, _ in cols]
    objs = {
        "columnHeaders": [{"properties": {"bold": lit("true"), "fontSize": lit(head),
                                          "fontColor": solid("#FFFFFF"), "backColor": solid(BRAND),
                                          "wordWrap": lit("true"), "outline": lit("'None'")}}],
        "values": [{"properties": {"fontSize": lit(fsize), "fontColor": solid(TXT),
                                   "backColor": solid(PANEL), "backColorSecondary": solid("#F6F8FB"),
                                   "wordWrap": lit("true")}}],
        "grid": [{"properties": {"gridHorizontal": lit("true"), "gridHorizontalColor": solid(EDGE),
                                 "gridHorizontalWeight": lit("1D"), "gridVertical": lit("false"),
                                 "rowPadding": lit(pad), "outlineColor": solid(EDGE)}}],
        "total": [{"properties": {"totals": lit("false")}}],
        "columnWidth": [{"properties": {"value": lit("%dD" % wdt)}, "selector": {"metadata": ref}}
                        for ref, wdt in widths],
    }
    objs["values"].append({"properties": {"fontColor": solid(BRAND), "fontSize": lit(big)},
                           "selector": {"metadata": "%s.%s" % (cols[0][0], cols[0][1])}})
    objs["values"].append({"properties": {"fontColor": solid(BRAND), "fontSize": lit(big)},
                           "selector": {"metadata": "%s.%s" % (cols[1][0], cols[1][1])}})
    fc = fill_col or (ECN, "EC Instruction", ECN, "Instr Fill")
    objs["values"].append({
        "properties": {"backColor": {"solid": {"color": {"expr": {"Measure": {
                           "Expression": {"SourceRef": {"Entity": fc[2]}}, "Property": fc[3]}}}}}},
        "selector": {"metadata": "%s.%s" % (fc[0], fc[1])}})
    proj = [column_(c, e, display=d) for e, c, _, d in cols]
    visual(page, vid, "tableEx", x, y, wd, ht, z, query={"Values": {"projections": proj}},
           sort=sort, objects=objs, vco=container(fill=PANEL), filter_config=filter_config)


def chart(page, vid, vtype, x, y, wd, ht, z, cat, ys, colors, title, sort=None, legend=False,
          filter_config=None, label_size="14D"):
    objs = {
        "dataPoint": [{"properties": {"fill": solid(c)},
                       "selector": {"metadata": "%s.%s" % (e, m)}} for (e, m), c in colors],
        "labels": [{"properties": {"show": lit("true"), "fontSize": lit(label_size),
                                   "color": solid(TXT)}}],
        "categoryAxis": [{"properties": {"fontSize": lit("14D"), "labelColor": solid(TXT),
                                         "showAxisTitle": lit("false")}}],
        "valueAxis": [{"properties": {"show": lit("false"), "showAxisTitle": lit("false"),
                                      "gridlineShow": lit("false")}}],
        "legend": [{"properties": {"show": lit("true" if legend else "false"),
                                   "position": lit("'Top'"), "fontSize": lit("14D"),
                                   "labelColor": solid(TXT)}}],
    }
    visual(page, vid, vtype, x, y, wd, ht, z,
           query={"Category": {"projections": [cat]}, "Y": {"projections": ys}},
           objects=objs, vco=container(title), sort=sort, filter_config=filter_config)


def asc(entity, prop): return {"field": colref(entity, prop), "direction": "Ascending"}
def desc_(entity, prop): return {"field": colref(entity, prop), "direction": "Descending"}
def mdesc(entity, prop): return {"field": measref(entity, prop), "direction": "Descending"}


P1, P3, P2 = "Hom_nay", "Lich_su", "Performance"
SY, SH = 112, 76           # slicer row
KY, KH = 198, 122          # KPI row
BY, BH = 334, 722          # body
KX = [40, 505, 970, 1435]


def clear_and_form(page, p, x_fg=852, fg_w=520, clear_x=1678):
    # FG code replaces the old "Giấy ECN" slicer: type the FG code, pick it.
    slicer(page, p + "_sl_fg", FGT, "FG Code", x_fg, SY, fg_w, SH, 210, "Mã FG (gõ để tìm)",
           dropdown=True, search=True)
    slicer(page, p + "_sl_instr", ECN, "EC Instruction", x_fg + fg_w + 12, SY, 260, SH, 230,
           "EC Instruction", dropdown=True)
    clear_button(page, p + "_clear", clear_x, SY, 202, SH, 240)


def filters(page, p, week_default):
    slicer(page, p + "_sl_week", ECN, "Tuần so với nay", 40, SY, 800, SH, 200,
           "Tuần chạy (từ Date, dời lịch theo Note)", horizontal=True, default=week_default)
    clear_and_form(page, p)


# ---------------------------- TAB 1 : LOG - tuần này (trái) | tuần sau (phải) ----
band(P1, "a", "ECN HÔM NAY CHẠY GÌ", "Tuần này (bên trái)  ·  Tuần sau (bên phải)")
clear_and_form(P1, "a", x_fg=40, fg_w=700, clear_x=1024)
LW = 905
for side, (x, meas_name, fillc, week, fname) in enumerate([
        (40, "Tiêu đề tuần này", BRAND, "Tuần này", "fWeekNow"),
        (975, "Tiêu đề tuần sau", BLUE, "Tuần sau", "fWeekNext")]):
    visual(P1, "a_h%d" % side, "card", x, 198, LW, 66, 300 + side,
        query={"Values": {"projections": [meas(meas_name, ECN)]}},
        objects={"labels": [{"properties": {"color": solid("#FFFFFF"), "fontSize": lit("30L"),
                                            "bold": lit("true")}}],
                 "categoryLabels": [{"properties": {"show": lit("false")}}]},
        vco=container(fill=fillc))
    loglist(P1, "a_log%d" % side, x, 274, LW, 782, 400 + side,
            [(ECN, "ECN hiển thị", 240, "ECN number"), (ECN, "FG hiển thị", 150, "FG"),
             (ECN, "Thay đổi gọn", 335, "Description of Change"),
             (ECN, "EC Instruction", 150, "EC Instruction")],
            sort=[desc_(ECN, "Ngày ban hành"), desc_(ECN, "ECN Number")],
            fsize="15D", big="17D", head="13D", pad="6D",
            filter_config=only_filter(ECN, "Tuần so với nay", week, fname))

# ---------------------------- TAB 2 : LỊCH SỬ - dời lịch / sự kiện khác ---------
band(P3, "c", "LỊCH SỬ ECN", "Mỗi dòng Note là một sự kiện  ·  Dời lịch = có Wxx/week xx + change")
slicer(P3, "c_sl_kind", NOT, "Loại sự kiện", 40, SY, 420, SH, 200, "Loại sự kiện", horizontal=True)
slicer(P3, "c_sl_cat", NOT, "Phân loại", 472, SY, 330, SH, 205, "Phân loại ghi chú", dropdown=True)
clear_and_form(P3, "c", x_fg=814)
kpi(P3, "c_k0", "Dòng ghi chú", KX[0], KY, 445, KH, 300, TXT, NOT, label="Sự kiện (dòng Note)")
kpi(P3, "c_k1", "Dòng dời lịch", KX[1], KY, 445, KH, 301, AMBER, NOT, label="Dòng dời lịch")
kpi(P3, "c_k2", "ECN có dời lịch", KX[2], KY, 445, KH, 302, BLUE, NOT, label="ECN có dời lịch")
kpi(P3, "c_k3", "Dòng chưa nhận dạng", KX[3], KY, 445, KH, 303, RED, NOT, label="Dòng chưa nhận dạng")
loglist(P3, "c_log", 40, BY, 1840, BH, 400,
        [(ECN, "ECN Number", 190, "ECN number"), (ECN, "Model / FG", 230, "Model / FG"),
         (NOT, "Cập nhật", 130, "Tuần/ngày ghi"), (NOT, "Loại sự kiện", 130, "Loại"),
         (NOT, "Dời sang", 120, "Dời sang"), (NOT, "Người phụ trách", 130, "Người"),
         (NOT, "Phân loại", 210, "Phân loại"), (NOT, "Nội dung", 640, "Nội dung")],
        sort=[desc_(ECN, "ECN Number"), asc(NOT, "Seq")], fsize="13D", big="14D", head="12D", pad="6D",
        fill_col=(NOT, "Loại sự kiện", NOT, "Event Fill"))

# ---------------------------- TAB 3 : PERFORMANCE -----------------------------
band(P2, "b", "PERFORMANCE ECN", "Hiệu suất xử lý ECN  ·  theo model, tuần, giấy ECN")
filters(P2, "b", None)
kpi(P2, "b_k0", "Số ECN", KX[0], KY, 445, KH, 300, TXT, label="Tổng ECN")
kpi(P2, "b_k1", "FG đang mở", KX[1], KY, 445, KH, 301, BLUE, label="FG đang mở")
kpi(P2, "b_k2", "Giấy ECN đã gửi", KX[2], KY, 445, KH, 302, AMBER, label="Giấy ECN đã gửi")
kpi(P2, "b_k3", "Ngày xử lý (trung vị)", KX[3], KY, 445, KH, 303, GREEN,
    label="Ngày từ ban hành đến hoàn tất")
CX, CW, RY, RH = [40, 667, 1294], 586, [334, 700], 356
chart(P2, "b_model", "barChart", CX[0], RY[0], CW, RH, 400,
      column_("Model", MOD, active=True), [meas("Top 10 Model")], [((ECN, "Top 10 Model"), BLUE)],
      "ECN theo Model (top 10)", sort=[mdesc(ECN, "Top 10 Model")])
chart(P2, "b_lead", "columnChart", CX[1], RY[0], CW, RH, 401,
      column_("Nhóm xử lý", ECN, active=True), [meas("Số ECN")], [((ECN, "Số ECN"), GREEN)],
      "Từ ngày ban hành đến ngày hoàn tất", sort=[asc(ECN, "Nhóm Order")])
chart(P2, "b_form", "columnChart", CX[2], RY[0], CW, RH, 402,
      column_("Giấy ECN", ECN, active=True), [meas("ECN đang mở")], [((ECN, "ECN đang mở"), AMBER)],
      "ECN đang mở: giấy đã gửi / chưa gửi")
chart(P2, "b_line", "lineChart", CX[0], RY[1], CW, RH, 403,
      column_("Tuần", CAL, active=True), [meas("ECN ban hành")],
      [((ECN, "ECN ban hành"), BLUE)], "ECN mở mới mỗi tuần (xu hướng 26 tuần)",
      sort=[asc(CAL, "Week Start")], label_size="12D")
chart(P2, "b_cmp", "clusteredColumnChart", CX[1], RY[1], CW, RH, 404,
      column_("Tuần", CAL, active=True), [meas("ECN ban hành"), meas("ECN hoàn tất")],
      [((ECN, "ECN ban hành"), BLUE), ((ECN, "ECN hoàn tất"), GREEN)],
      "ECN mỗi tuần: mở mới vs hoàn tất (13 tuần)", sort=[asc(CAL, "Week Start")], legend=True,
      filter_config=only_filter(CAL, "Gần đây", "Yes", "fRecent"))
visual(P2, "b_fg", "treemap", CX[2], RY[1], CW, RH, 405,
       query={"Group": {"projections": [column_("FG Code", FGT, active=True)]},
              "Values": {"projections": [meas("Top 5 FG")]}},
       objects={"labels": [{"properties": {"show": lit("true"), "fontSize": lit("18D")}}],
                "legend": [{"properties": {"show": lit("false")}}]},
       vco=container("Top 5 FG đổi nhiều nhất"))

# ---------------------------- write report files ---------------------------
DEF = os.path.join(RP, "definition")
if os.path.isdir(os.path.join(DEF, "pages")):
    shutil.rmtree(os.path.join(DEF, "pages"))
w(os.path.join(RP, "definition.pbir"), {
    "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definitionProperties/1.0.0/schema.json",
    "version": "4.0", "datasetReference": {"byPath": {"path": "../" + NAME + ".SemanticModel"}}})
w(os.path.join(RP, ".platform"), {
    "$schema": "https://developer.microsoft.com/json-schemas/fabric/gitIntegration/platformProperties/2.0.0/schema.json",
    "metadata": {"type": "Report", "displayName": NAME},
    "config": {"version": "2.0", "logicalId": tag()}})
w(os.path.join(DEF, "version.json"), {
    "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/versionMetadata/1.0.0/schema.json",
    "version": "2.0.0"})
VER = {"visual": "2.9.0", "report": "3.3.0", "page": "2.3.1"}
w(os.path.join(DEF, "report.json"), {
    "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/report/3.3.0/schema.json",
    "themeCollection": {"baseTheme": {"name": "CY26SU05", "type": "SharedResources",
                                      "reportVersionAtImport": VER}},
    "objects": {"outspacePane": [{"properties": {"expanded": lit("false")}}]},
    "resourcePackages": [
        {"name": "SharedResources", "type": "SharedResources",
         "items": [{"name": "CY26SU05", "path": "BaseThemes/CY26SU05.json", "type": "BaseTheme"}]},
        {"name": "RegisteredResources", "type": "RegisteredResources",
         "items": [{"name": LOGO, "path": LOGO, "type": "Image"}]}],
    "settings": {"useStylableVisualContainerHeader": True, "defaultDrillFilterOtherVisuals": True}})
w(os.path.join(DEF, "pages", "pages.json"), {
    "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/pagesMetadata/1.1.0/schema.json",
    "pageOrder": [P1] if VARIANT == "assembly" else [P1, P3, P2], "activePageName": P1})
for pname, disp in [(P1, "Log - hôm nay chạy gì")] + ([] if VARIANT == "assembly" else
                                                    [(P3, "Lịch sử"), (P2, "Performance")]):
    w(os.path.join(DEF, "pages", pname, "page.json"), {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/page/2.1.0/schema.json",
        "name": pname, "displayName": disp, "displayOption": "FitToPage",
        "height": 1080, "width": 1920,
        "objects": {"background": [{"properties": {"color": solid(BG), "transparency": lit("0L")}}]}})
for pname, vid, v in VISUALS:
    if VARIANT == "assembly" and pname != P1:
        continue
    w(os.path.join(DEF, "pages", pname, "visuals", vid, "visual.json"), v)

theme = zipfile.ZipFile(
    r"C:\Users\nguyenhoang88502\Desktop\NPI_CAPEX_Report\OPEX\NPI_OPEX_Dashboard.pbix"
).read("Report/StaticResources/SharedResources/BaseThemes/CY26SU05.json")
dst = os.path.join(RP, "StaticResources", "SharedResources", "BaseThemes", "CY26SU05.json")
os.makedirs(os.path.dirname(dst), exist_ok=True)
open(dst, "wb").write(theme)
reg = os.path.join(RP, "StaticResources", "RegisteredResources", LOGO)
os.makedirs(os.path.dirname(reg), exist_ok=True)
shutil.copyfile(LOGO_SRC, reg)
w(os.path.join(ROOT, NAME + ".pbip"), {
    "$schema": "https://developer.microsoft.com/json-schemas/fabric/pbip/pbipProperties/1.0.0/schema.json",
    "version": "1.0", "artifacts": [{"report": {"path": NAME + ".Report"}}],
    "settings": {"enableAutoRecovery": True}})


# ---------------- readable copies -----------------------------------------
parts = ["/* READABLE COPY of the M inside %s.SemanticModel\\model.bim.\n"
         "   Power BI reads model.bim, not this file.\n"
         "   Edit keywords in the Functions group (fnTags, fnComps, fnNoteLine, fnParseDesc)\n"
         "   and names in parameters pPeople / pVendors. */\n\n" % NAME]
parts.append("/* PARAMETERS: pSiteUrl, pFolderPath, pFileName (SharePoint), pSheetName, pTodayOverride (yyyy-MM-dd to test), pPeople, pVendors */\n\n")
for title, body in [("TodayDate", M_TODAY)] + FUNCS + [
        ("ECN Raw  (not loaded)", M_RAW), ("Week Map  (not loaded)", M_WEEKMAP),
        ("ECN Base  (not loaded)", M_BASE), ("ECN Notes Base  (not loaded)", M_NOTES_BASE),
        ("ECN Weeks Base  (not loaded)", M_WEEKS_BASE), ("Line Stock  (not loaded)", M_LINESTOCK),
        ("ECN Weeks", M_WEEKS), ("ECN Calendar", M_CAL),
        ("ECN Lines", M_LINES), ("ECN Notes", M_NOTES), ("ECN Models", M_MODELS),
        ("ECN FG", M_FG), ("ECN", M_ECN)]:
    parts.append("/*" + "-" * 70 + "\n  " + title + "\n" + "-" * 70 + "*/\n")
    parts.append(body.strip("\n") + "\n\n")
w(os.path.join(ROOT, "PowerQuery_M_Code.m"), "".join(parts), raw=True)

dax = ["/* DAX measures. READABLE COPY of model.bim. */\n\n"]
for tbl, lst in [(ECN, MEASURES + NEW_MEASURES), (NOT, NOTE_MEASURES), (WKS, WEEK_MEASURES)]:
    for mname, fmt, expr, desc in lst:
        dax.append("-- [%s] %s" % (tbl, desc))
        dax.append("[" + mname + "] =" + expr.rstrip("\n") + "\n\n")
w(os.path.join(ROOT, "DAX_Measures.dax"), "\n".join(dax), raw=True)

# ==========================================================================
# VALIDATION
# ==========================================================================
tbl = {t["name"]: t for t in model_bim["model"]["tables"]}
cols = {t: {c["name"] for c in tbl[t]["columns"]} for t in tbl}
meas_ = {t: {m["name"] for m in tbl[t].get("measures", [])} for t in tbl}
exprs = {e["name"] for e in expressions}
errors = []

for pname, vid, v in VISUALS:
    blob = json.dumps(v, ensure_ascii=False)
    for m in re.finditer(r'"(Column|Measure)": \{"Expression": \{"SourceRef": \{"Entity": "([^"]+)"\}\}, "Property": "([^"]+)"', blob):
        kind, ent, prop = m.groups()
        pool = cols if kind == "Column" else meas_
        if ent not in pool or prop not in pool[ent]:
            errors.append("%s/%s: %s %s[%s] not found" % (pname, vid, kind, ent, prop))
    for m in re.finditer(r'"metadata": "([^"]+)"', blob):
        ent, prop = m.group(1).split(".", 1)
        if ent not in cols or (prop not in cols[ent] and prop not in meas_[ent]):
            errors.append("%s/%s: selector %s not found" % (pname, vid, m.group(1)))
for t in tbl.values():
    for mm in t.get("measures", []):
        e = "\n".join(mm["expression"])
        for ref in re.findall(r"'(\w[\w ]*)'\[([^\]]+)\]", e):
            if ref[1] not in cols.get(ref[0], set()):
                errors.append("DAX %s -> %s[%s] missing" % (mm["name"], ref[0], ref[1]))
# every function / expression used in M must exist
allm = "\n".join(["\n".join(e["expression"]) for e in expressions] +
                 ["\n".join(t["partitions"][0]["source"]["expression"]) for t in tbl.values()])
for used in set(re.findall(r"\b(fn[A-Z]\w+)\(", allm)):
    if used not in exprs:
        errors.append("M uses undefined function " + used)
for used in set(re.findall(r'#"(ECN Raw|Week Map|Line Stock|ECN (?:Notes |Weeks )?Base)"', allm)):
    if used not in exprs:
        errors.append('M references undefined query "%s"' % used)
# every model column must be produced by its M
for t, m in [(ECN, M_ECN), (LIN, M_LINES), (NOT, M_NOTES), (MOD, M_MODELS), (FGT, M_FG), (WKS, M_WEEKS), (CAL, M_CAL)]:
    for c in cols[t]:
        if ('"%s"' % c) not in m and ('#"%s"' % c) not in m and (c + " =") not in m \
                and (c + ",") not in m and (c + " ") not in m:
            errors.append("%s column %s not found in its M" % (t, c))
# balanced brackets in every M text
for name, m in [("ECN", M_ECN), ("NOTES_BASE", M_NOTES_BASE), ("NOTELINE", M_FN_NOTELINE),
                ("PARSEDESC", M_FN_PARSEDESC), ("BASE", M_BASE), ("COMPS", M_FN_COMPS),
                ("WEEKMAP", M_WEEKMAP), ("WEEKS_BASE", M_WEEKS_BASE), ("LINESTOCK", M_LINESTOCK),
                ("WEEKS", M_WEEKS), ("CAL", M_CAL), ("WEEKTOKENS", M_FN_WEEKTOKENS),
                ("WEEKSTART", M_FN_WEEKSTART)]:
    s = re.sub(r'"[^"]*"', '""', m)
    for a, b in ["()", "[]", "{}"]:
        if s.count(a) != s.count(b):
            errors.append("%s: unbalanced %s%s %d/%d" % (name, a, b, s.count(a), s.count(b)))

print("PBIP written. visuals:", len(VISUALS), " tables:", len(tbl),
      " expressions:", len(expressions))
print("VALIDATION:", "OK" if not errors else errors)
