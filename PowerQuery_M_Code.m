/* READABLE COPY of the M inside ECN_Assembly.SemanticModel\model.bim.
   Power BI reads model.bim, not this file.
   Edit keywords in the Functions group (fnTags, fnComps, fnNoteLine, fnParseDesc)
   and names in parameters pPeople / pVendors. */

/* PARAMETERS: pSiteUrl, pFolderPath, pFileName (SharePoint), pSheetName, pTodayOverride (yyyy-MM-dd to test), pPeople, pVendors */

/*----------------------------------------------------------------------
  TodayDate
----------------------------------------------------------------------*/
let
    // "Today" is evaluated at refresh time in Vietnam time (UTC+7).
    // Type a date as yyyy-MM-dd into pTodayOverride to test another day.
    Result =
        if pTodayOverride <> null and Text.Trim(pTodayOverride) <> ""
        then Date.FromText(Text.Trim(pTodayOverride), [Format = "yyyy-MM-dd", Culture = "en-US"])
        else Date.From(DateTimeZone.SwitchZone(DateTimeZone.UtcNow(), 7))
in
    Result

/*----------------------------------------------------------------------
  fnIsoWeek
----------------------------------------------------------------------*/
(d as date) as number =>
let
    dow = Date.DayOfWeek(d, Day.Monday),
    thu = Date.AddDays(d, 3 - dow)
in
    Number.IntegerDivide(Date.DayOfYear(thu) - 1, 7) + 1

/*----------------------------------------------------------------------
  fnWeekNo
----------------------------------------------------------------------*/
(t as nullable text) as nullable number =>
let
    s = if t = null then "" else Text.Upper(Text.Trim(t)),
    n = if Text.StartsWith(s, "W") then try Number.FromText(Text.Middle(s, 1)) otherwise null else null
in
    n

/*----------------------------------------------------------------------
  fnPlanWk
----------------------------------------------------------------------*/
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

/*----------------------------------------------------------------------
  fnEffWk
----------------------------------------------------------------------*/
(wk as nullable number, todayWk as number) as nullable number =>
    if wk = null then null
    else if wk <= 12 and todayWk >= 40 then wk + 53
    else wk

/*----------------------------------------------------------------------
  fnIsDigits
----------------------------------------------------------------------*/
(x as text) as logical =>
    x <> "" and Text.Select(x, {"0".."9"}) = x

/*----------------------------------------------------------------------
  fnTokens
----------------------------------------------------------------------*/
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

/*----------------------------------------------------------------------
  fnSplitList
----------------------------------------------------------------------*/
(s as nullable text) as list =>
    if s = null then {}
    else
        let
            x = Text.Replace(Text.Replace(Text.Replace(s, ";", ","), "#(lf)", ","), "#(cr)", ","),
            p = List.Transform(Text.Split(x, ","), Text.Trim)
        in
            List.Distinct(List.Select(p, each _ <> ""))

/*----------------------------------------------------------------------
  fnCodeKind
----------------------------------------------------------------------*/
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

/*----------------------------------------------------------------------
  fnTags
----------------------------------------------------------------------*/
(desc as nullable text) as text =>
let
    d   = if desc = null then "" else Text.Lower(desc),
    Pkg = {"hộp", "thùng", "carton", "túi", "blister", "mút", "packaging", "color box", "graphic card", "card", "quai", "bvld", "mailer", "vỉ", "bao bì", "pe bag", "foam", "khay", "lót", "chèn"},
    Mec = {"thân", "trimmer", "lưỡi", "blade", "vít", "screw", "nắp", "housing", "cơ khí", "lò xo", "spring", "gá", "khuôn", "linh kiện nhựa", "motor", "mô tơ", "đầu cắt", "comb", "lược", "dao", "ốc", "phôi", "jack", "đầu cắm", "cụm"},
    Pcb = {"mạch", "pcb", "bo mạch", "board", "chip", "dây điện", "wire", "pin", "battery", "adapter", "sạc", "charger", "led", "switch", "công tắc", "mach", "usb", "type c", "bóng đèn", "4v", "5v"},
    Lbl = {"nhãn", "label", "tem", "laser", "sticker", "sách", "manual", "hướng dẫn", "in ấn", "bản vẽ", "barcode", "mã vạch", "decal", "logo", "print", "nội dung in", "phiên bản", "version", "warning", "warranty", "instruction", "ib #", "fill"},
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

/*----------------------------------------------------------------------
  fnComps
----------------------------------------------------------------------*/
(desc as nullable text) as text =>
let
    // Which physical part changes. Keywords of 3 characters or less must equal a whole word.
    d     = if desc = null then "" else Text.Lower(desc),
    ltoks = List.Transform(fnTokens(desc), Text.Lower),
    Defs  = {{"Hộp màu", {"hộp màu", "color box", "clor box", "hộp", "box"}}, {"Thùng carton", {"thùng", "carton"}}, {"Túi / pouch", {"túi", "pouch", "zip"}}, {"Nhãn", {"nhãn", "label", "tem"}}, {"Sách HD", {"sách", "book", "manual", "instruction"}}, {"Graphic card", {"graphic card", "card"}}, {"Blister / khay", {"blister", "khay"}}, {"Mút / chèn", {"mút", "chèn", "foam"}}, {"Bản mạch", {"mạch", "pcb", "mach"}}, {"Thân / vỏ", {"thân", "vỏ", "housing"}}, {"Dao / lưỡi", {"dao", "blade", "lưỡi"}}, {"Trimmer", {"trimmer", "đầu tỉa"}}, {"Pin / sạc", {"pin", "sạc", "battery", "4v", "5v", "usb"}}, {"Laser / in", {"laser", "in xoay", "logo", "print"}}, {"Chống ẩm", {"desiccant", "chống ẩm"}}, {"Mô tơ", {"mô tơ", "motor"}}},
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

/*----------------------------------------------------------------------
  fnParseDesc
----------------------------------------------------------------------*/
(desc as nullable text) as record =>
let
    // Reads "Replace 1006982 by 1011819", "Remove 1003942 ... Add item 1011829",
    // "Thay hộp 3029144 bằng hộp 3031758, Bỏ nhãn 3031690" word by word.
    Rem = {"bỏ", "remove", "romove", "hủy", "xóa", "scrap"},
    Add = {"thêm", "add", "bổ"},
    Rep = {"thay", "replace", "rplace", "đổi", "change"},
    Upd = {"update", "updete", "cập"},
    Con = {"bằng", "by", "thành", "to", "sang", "->", "→", "➔"},
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

/*----------------------------------------------------------------------
  fnEstWk
----------------------------------------------------------------------*/
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

/*----------------------------------------------------------------------
  fnWeekTokens
----------------------------------------------------------------------*/
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
            isWk  = List.Contains({"week", "wk", "w", "tuần", "tuan"}, t) and fnIsDigits(nxt)
                    and Text.Length(nxt) <= 2
        in
            if isW and not List.IsEmpty(parts) and fnIsDigits(List.Last(parts)) then Number.FromText(List.Last(parts))
            else if isWk then Number.FromText(nxt)
            else null
in
    List.RemoveNulls(List.Transform(List.Positions(ltoks), one))

/*----------------------------------------------------------------------
  fnWeekStart
----------------------------------------------------------------------*/
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

/*----------------------------------------------------------------------
  fnWeekStartY
----------------------------------------------------------------------*/
(wk as number, yr as number) as date =>
    // Monday of ISO week wk of year yr (week 1 = the week that holds 4 January)
    Date.AddDays(Date.StartOfWeek(#date(yr, 1, 4), Day.Monday), (wk - 1) * 7)

/*----------------------------------------------------------------------
  fnNoteLine
----------------------------------------------------------------------*/
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
    CatDefs = {{"Chưa có nhu cầu / đơn", {"no more demand", "không thấy có demand", "chưa có kh", "chưa tạo mo", "no order", "không cần clear"}}, {"Chờ Buyer / giấy tờ", {"not yet email", "buyer not action", "wrote mail", "đã làm giấy", "buyer informed", "buyer noted", "buyer updated", "buyer est"}}, {"Chờ khách / thiết kế", {"customer", "cus.", "design", "khách"}}, {"Chờ NCC / hàng về", {"vendor", "delivery", "arrives", "last order", "quick connect", "viet may", "deliver"}}, {"Xử lý tồn kho", {"stock", "tồn", "scrap", "clear", "wh ", "inventory", "keep old", "available"}}, {"Setup / BOM", {"setup", "bom", "mo ", "level"}}, {"Hạn hoàn tất", {"finished by"}}, {"Theo dõi", {"follow", "recheck", "check", "inform", "update", "est", "wait", "will change"}}},
    vl      = List.Transform(vendors, Text.Lower),
    Defs    = List.Transform(CatDefs, each if _{0} = "Chờ NCC / hàng về" then {_{0}, _{1} & vl} else _),
    found   = List.First(
                  List.Select(Defs, each List.AnyTrue(List.Transform(_{1}, (k) => Text.Contains(lt, k)))),
                  null),
    isDone  = List.Count(ltoks) > 0 and List.Count(ltoks) <= 3 and List.Contains({"done", "ok", "updated"}, ltoks{0}),
    cat     = if isDone then "Đã xong"
              else if lt = "cancel" then "Hủy"
              else if found <> null then found{0}
              else "Chưa nhận dạng",

    key     = if dateFrom <> null then Number.From(dateFrom)
              else if wkFrom <> null then Number.From(Date.AddDays(#date(Date.Year(T), 1, 1), (wkFrom - 1) * 7))
              else -1,

    // ---- schedule move: the line has a week (Wxx / week xx) AND the word "change" ----
    chg     = List.AnyTrue(List.Transform({"change", "dời", "đổi"}, each Text.Contains(lt, _))),
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

/*----------------------------------------------------------------------
  ECN Raw  (not loaded)
----------------------------------------------------------------------*/
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

/*----------------------------------------------------------------------
  Week Map  (not loaded)
----------------------------------------------------------------------*/
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

/*----------------------------------------------------------------------
  ECN Base  (not loaded)
----------------------------------------------------------------------*/
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
                "New item (est MP week)", "New item (est MP year)", "Note"
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
                {"New item (est MP year)", each ToText(_), type nullable text},
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
                {"New item (est MP year)", "Est Year Text"},
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

    // "W44" or a typed "W44 --> W43" (last week number wins); #N/A became null above.
    AddEstWk =
        Table.AddColumn(
            AddIssue, "Est Wk",
            each if [Est MP Text] = null then null
                 else List.Last(fnWeekTokens(List.Transform(fnTokens([Est MP Text]), Text.Lower)), null),
            type nullable number),
    AddEst =
        Table.AddColumn(
            AddEstWk, "Est Year",
            each let d = if [Est Year Text] = null then "" else Text.Select([Est Year Text], {"0".."9"})
                 in if Text.Length(d) >= 4 then Number.FromText(Text.Start(d, 4)) else null,
            type nullable number),

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

/*----------------------------------------------------------------------
  ECN Notes Base  (not loaded)
----------------------------------------------------------------------*/
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

/*----------------------------------------------------------------------
  ECN Weeks Base  (not loaded)
----------------------------------------------------------------------*/
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

/*----------------------------------------------------------------------
  Line Stock  (not loaded)
----------------------------------------------------------------------*/
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

/*----------------------------------------------------------------------
  ECN Weeks
----------------------------------------------------------------------*/
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

/*----------------------------------------------------------------------
  ECN Calendar
----------------------------------------------------------------------*/
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

/*----------------------------------------------------------------------
  ECN Lines
----------------------------------------------------------------------*/
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

/*----------------------------------------------------------------------
  ECN Notes
----------------------------------------------------------------------*/
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

/*----------------------------------------------------------------------
  ECN Models
----------------------------------------------------------------------*/
let
    Rows  = Table.SelectColumns(#"ECN Base", {"ECN Number", "Model / Customer No."}),
    Split = Table.AddColumn(Rows, "Model", each fnSplitList([#"Model / Customer No."])),
    Flat  = Table.ExpandListColumn(Split, "Model"),
    Pick  = Table.SelectRows(Table.SelectColumns(Flat, {"ECN Number", "Model"}), each [Model] <> null),
    Out   = Table.TransformColumnTypes(Table.Distinct(Pick), {{"ECN Number", type text}, {"Model", type text}})
in
    Out

/*----------------------------------------------------------------------
  ECN FG
----------------------------------------------------------------------*/
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

/*----------------------------------------------------------------------
  ECN
----------------------------------------------------------------------*/
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

            // ---- week the ECN runs: est MP week + year (cols R, S) > Note move (Wxx + change) > Date --
            planStart = if issue = null then null else Date.StartOfWeek(issue, Day.Monday),
            mv        = Table.Sort(
                            Table.SelectRows(ns, each [IsMove] = true),
                            {{"Cell", Order.Ascending}, {"Pos", Order.Ascending}}),
            mvWk      = List.RemoveNulls(mv[MoveWk]),
            moves     = Table.RowCount(mv),
            noteStart = if not List.IsEmpty(mvWk) then fnWeekStart(List.First(mvWk), T) else null,
            estStarts = List.Transform(
                            Table.ToRecords(Table.SelectRows(rs, each [Est Wk] <> null)),
                            each if [Est Year] <> null then fnWeekStartY([Est Wk], [Est Year])
                                 else fnWeekStart([Est Wk], T)),
            estStart  = if List.IsEmpty(estStarts) then null else List.Min(estStarts),
            yearOnly  = List.RemoveNulls(Table.SelectRows(rs, each [Est Wk] = null)[Est Year]),
            yearStart = if List.IsEmpty(yearOnly) then null else #date(List.Min(yearOnly), 1, 1),
            isYearOnly = estStart = null and noteStart = null and yearStart <> null
                         and (planStart = null or Date.Year(yearStart) > Date.Year(planStart)),
            curStart  = if estStart <> null then estStart
                        else if noteStart <> null then noteStart
                        else if isYearOnly then yearStart
                        else planStart,
            pad2      = (n as number) as text => "W" & Text.PadStart(Text.From(n), 2, "0"),
            runNo     = if curStart = null or isYearOnly then null else fnIsoWeek(curStart),
            runTxt    = if isYearOnly then Text.From(Date.Year(curStart))
                        else if runNo = null then "—" else pad2(runNo),
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
            chainNums = List.Combine({
                            if planStart = null then {} else {fnIsoWeek(planStart)},
                            List.Reverse(mvWk),
                            if runNo = null then {} else {runNo}}),
            chainU    = List.Accumulate(
                            chainNums, {},
                            (st, x) => if not List.IsEmpty(st) and List.Last(st) = x then st else st & {x}),
            chainTxt  = List.Transform(chainU, pad2) & (if isYearOnly then {runTxt} else {}),
            histTxt   = if List.IsEmpty(chainTxt) then "—"
                        else if List.Count(chainTxt) = 1 then "Đúng kế hoạch " & chainTxt{0}
                        else Text.Combine(chainTxt, " → "),
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

