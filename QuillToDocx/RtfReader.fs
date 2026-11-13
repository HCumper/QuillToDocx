namespace QuillToDocx

open System
open System.Collections.Generic
open System.Text

/// Reads the RTF Word hands an export converter, keeping what a Quill document can hold:
/// paragraphs with their indents, alignment and tab stops, bold/underline/sub/superscript,
/// tabs, soft hyphens and page breaks, the default header and footer (with PAGE fields), the
/// page size and margins, and document variables. Everything else is skipped.
module RtfReader =

    type RtfParagraph =
        { LeftIndent: int
          RightIndent: int
          FirstIndent: int
          Alignment: QuillLayout.Alignment
          Tabs: (int * QuillLayout.TabKind) list
          PageBreakBefore: bool
          Content: Inline list }

    type RtfDocument =
        { Body: RtfParagraph list
          Header: RtfParagraph list
          Footer: RtfParagraph list
          PaperWidth: int option
          PaperHeight: int option
          MarginLeft: int option
          MarginRight: int option
          MarginTop: int option
          MarginBottom: int option
          HeaderY: int option
          FooterY: int option
          PageStart: int option
          DocVars: Map<string, string> }

    type private Destination =
        | Body
        | Header
        | Footer
        | FieldInstruction
        | DocVariable
        | FontTable
        | Skip

    type private ParaProps =
        { Left: int
          Right: int
          First: int
          Align: QuillLayout.Alignment
          /// Newest first.
          Tabs: (int * QuillLayout.TabKind) list
          PendingTab: QuillLayout.TabKind
          PageBreakBefore: bool }

    /// A field being read. A PAGE field's result is replaced by Quill's placeholder, written once
    /// with the formatting of the result's first text (Word puts the bold inside the result).
    type private Field =
        { Instruction: StringBuilder
          mutable Placeholder: string option
          mutable Emitted: bool }

    type private GroupState =
        { Dest: Destination
          Style: TextStyle
          Hidden: bool
          Para: ParaProps
          Uc: int
          InListText: bool
          PageField: Field option
          Font: int }

    let private plain =
        { Bold = false
          Italic = false
          Underline = false
          Sub = false
          Super = false }

    let private defaultPara =
        { Left = 0
          Right = 0
          First = 0
          Align = QuillLayout.AlignLeft
          Tabs = []
          PendingTab = QuillLayout.TabLeft
          PageBreakBefore = false }

    /// Destinations whose content is never text of the document.
    let private skipped =
        HashSet<string>(
            [ "colortbl"; "stylesheet"; "info"; "listtable"; "listoverridetable"; "rsidtbl"
              "generator"; "pict"; "object"; "shp"; "shpinst"; "shppict"; "nonshppict"; "xmlnstbl"
              "mmathPr"; "themedata"; "colorschememapping"; "latentstyles"; "datastore"; "pgdsctbl"
              "revtbl"; "footnote"; "annotation"; "atnid"; "atnauthor"; "atrfstart"; "atrfend"
              "bkmkstart"; "bkmkend"; "headerl"; "headerf"; "footerl"; "footerf"; "ftnsep"; "ftnsepc"
              "ftncn"; "aftnsep"; "aftnsepc"; "aftncn"; "pnseclvl"; "userprops"; "xe"; "tc"; "txe"
              "template"; "filetbl"; "falt"; "panose"; "fname"; "objdata"; "blipuid"; "wgrffmtfilter"
              "listpicture"; "protusertbl"; "fchars"; "lchars"; "pn"; "pntxta"; "pntxtb"; "sp"; "sn"
              "sv"; "formfield"; "ffname"; "datafield"; "do"; "picprop"; "background"; "passwordhash"
              "defchp"; "defpap"; "ud"; "upr"; "category"; "comment"; "doccomm"; "keywords"; "operator"
              "subject"; "title"; "author"; "company"; "manager"; "hlinkbase"; "nesttableprops"
              "nonesttables"; "dptxbxtext"; "factoidname"; "oldcprops"; "oldpprops"; "oldsprops"
              "oldtprops"; "mhtmltag"; "htmltag"; "fontemb"; "fontfile"; "leveltext"; "levelnumbers" ])

    let private underlineWords =
        HashSet<string>(
            [ "ul"; "uld"; "uldash"; "uldashd"; "uldashdd"; "uldb"; "ulhwave"; "ulldash"; "ulth"; "ulthd"
              "ulthdash"; "ulthdashd"; "ulthdashdd"; "ulthldash"; "ululdbwave"; "ulw"; "ulwave" ])

    /// List numbers and bullets: keep digits, letters and simple punctuation; bullet glyphs from
    /// symbol fonts become '*'.
    let private listText (text: string) =
        text
        |> String.map (fun c -> if int c > 0x7E then '*' else c)

    /// Windows code page for an RTF font charset. Word writes Greek, Cyrillic and Central European
    /// text as \'hh bytes in the font's own code page, not the document's.
    let private charsetCodePage charset =
        match charset with
        | 128 -> Some 932
        | 129 -> Some 949
        | 130 -> Some 1361
        | 134 -> Some 936
        | 136 -> Some 950
        | 161 -> Some 1253
        | 162 -> Some 1254
        | 163 -> Some 1258
        | 177 -> Some 1255
        | 178 -> Some 1256
        | 186 -> Some 1257
        | 204 -> Some 1251
        | 222 -> Some 874
        | 238 -> Some 1250
        | 255 -> Some 437
        | _ -> None

    /// What Quill prints for a PAGE field: nnn, or rrr / aaa for roman numerals and letters.
    let private pagePlaceholder (instruction: string) =
        let words = instruction.Split([| ' '; '\t' |], StringSplitOptions.RemoveEmptyEntries)

        if words.Length = 0 || not (String.Equals(words[0], "PAGE", StringComparison.OrdinalIgnoreCase)) then
            None
        else
            let lower = instruction.ToLowerInvariant()
            if lower.Contains "roman" then Some "rrr"
            elif lower.Contains "alphabetic" then Some "aaa"
            else Some "nnn"

    let parse (rtf: string) =
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance)
        let mutable codePage = Encoding.GetEncoding 1252

        let mutable state =
            { Dest = Body
              Style = plain
              Hidden = false
              Para = defaultPara
              Uc = 1
              InListText = false
              PageField = None
              Font = 0 }

        let stack = Stack<GroupState * (unit -> unit) list>()
        let mutable onClose: (unit -> unit) list = []
        let fields = Stack<Field>()
        let docVars = Dictionary<string, string>()
        let mutable docVarText = StringBuilder()

        let content = Dictionary<Destination, ResizeArray<Inline>>()
        let paragraphs = Dictionary<Destination, ResizeArray<RtfParagraph>>()

        for d in [ Body; Header; Footer ] do
            content[d] <- ResizeArray()
            paragraphs[d] <- ResizeArray()

        let mutable headerTaken = false
        let mutable footerTaken = false
        let mutable sectionBreakPending = false
        let mutable starPending = false
        let mutable fallbackToSkip = 0
        let pending = StringBuilder()

        let fontCharsets = Dictionary<int, int>()
        let encodings = Dictionary<int, Encoding>()
        let mutable fontBeingDefined = 0
        let mutable defaultFont = 0

        let decodeByte (b: byte) =
            let encoding =
                match fontCharsets.TryGetValue state.Font with
                | true, charset ->
                    match charsetCodePage charset with
                    | Some cp ->
                        match encodings.TryGetValue cp with
                        | true, e -> e
                        | _ ->
                            let e = try Encoding.GetEncoding cp with _ -> codePage
                            encodings[cp] <- e
                            e
                    | None -> codePage
                | _ -> codePage

            encoding.GetString([| b |])

        let docProps = Dictionary<string, int>()
        let firstSection = Dictionary<string, int>()

        let isTextDest d =
            match d with
            | Body
            | Header
            | Footer -> true
            | _ -> false

        let addItem item =
            if isTextDest state.Dest && not state.Hidden then
                content[state.Dest].Add item

        let appendText (dest: Destination) (text: string) (style: TextStyle) =
            let items = content[dest]

            match (if items.Count > 0 then Some items[items.Count - 1] else None) with
            | Some (RunText (previous, s)) when s = style -> items[items.Count - 1] <- RunText(previous + text, style)
            | _ -> items.Add(RunText(text, style))

        let flush () =
            if pending.Length > 0 then
                let text = pending.ToString()
                pending.Clear() |> ignore

                match state.Dest with
                | d when isTextDest d && not state.Hidden ->
                    match state.PageField with
                    | Some field ->
                        if not field.Emitted then
                            field.Emitted <- true
                            appendText d (defaultArg field.Placeholder "nnn") state.Style
                    | None -> appendText d (if state.InListText then listText text else text) state.Style
                | FieldInstruction when fields.Count > 0 -> fields.Peek().Instruction.Append(text) |> ignore
                | DocVariable -> docVarText.Append(text) |> ignore
                | _ -> ()

        let text (s: string) =
            if fallbackToSkip > 0 then
                fallbackToSkip <- fallbackToSkip - 1
            else
                pending.Append(s) |> ignore

        let finishParagraph (dest: Destination) =
            flush ()

            if isTextDest dest then
                let items = content[dest]
                let p = state.Para

                paragraphs[dest].Add
                    { LeftIndent = p.Left
                      RightIndent = p.Right
                      FirstIndent = p.First
                      Alignment = p.Align
                      Tabs = List.rev p.Tabs
                      PageBreakBefore = p.PageBreakBefore || (dest = Body && sectionBreakPending)
                      Content = List.ofSeq items }

                if dest = Body then sectionBreakPending <- false
                items.Clear()

        let setStyle f =
            flush ()
            state <- { state with Style = f state.Style }

        let setPara f =
            state <- { state with Para = f state.Para }

        let on (param: int option) = param <> Some 0

        let controlWord (word: string) (param: int option) =
            let value = defaultArg param 0

            if starPending then
                starPending <- false

                match word with
                | "fldinst"
                | "docvar" -> ()
                | _ ->
                    flush ()
                    state <- { state with Dest = Skip }

            if state.Dest = Skip then
                ()
            elif state.Dest = FontTable then
                match word with
                | "f" -> fontBeingDefined <- value
                | "fcharset" -> fontCharsets[fontBeingDefined] <- value
                | _ -> ()
            else
                match word with
                | "fonttbl" ->
                    flush ()
                    state <- { state with Dest = FontTable }
                | "deff" -> defaultFont <- value
                | "f" ->
                    flush ()
                    state <- { state with Font = value }
                | "ansicpg" ->
                    try
                        codePage <- Encoding.GetEncoding value
                    with _ ->
                        ()
                | "uc" -> state <- { state with Uc = max 0 value }
                | "u" ->
                    fallbackToSkip <- 0
                    text (string (char (if value < 0 then value + 65536 else value)))
                    fallbackToSkip <- state.Uc
                // Word writes typographic characters as control words.
                | "lquote" -> text "‘"
                | "rquote" -> text "’"
                | "ldblquote" -> text "“"
                | "rdblquote" -> text "”"
                | "emdash" -> text "—"
                | "endash" -> text "–"
                | "bullet" -> text "•"
                | "emspace"
                | "enspace"
                | "qmspace" -> text " "
                | "par"
                | "line" -> finishParagraph state.Dest
                | "tab"
                | "cell"
                | "nestcell" ->
                    flush ()
                    addItem (RunTab state.Style)
                | "row"
                | "nestrow" ->
                    flush ()

                    if isTextDest state.Dest then
                        let items = content[state.Dest]

                        match (if items.Count > 0 then Some items[items.Count - 1] else None) with
                        | Some (RunTab _) -> items.RemoveAt(items.Count - 1)
                        | _ -> ()

                    finishParagraph state.Dest
                | "page" ->
                    flush ()
                    addItem RunPageBreak
                | "sect" ->
                    flush ()
                    if content[Body].Count > 0 then finishParagraph Body
                    sectionBreakPending <- true
                | "pard" -> setPara (fun _ -> defaultPara)
                | "plain" ->
                    flush ()
                    state <- { state with Style = plain; Hidden = false; Font = defaultFont }
                | "b" -> setStyle (fun s -> { s with Bold = on param })
                | "i" -> setStyle (fun s -> { s with Italic = on param })
                | w when underlineWords.Contains w -> setStyle (fun s -> { s with Underline = on param })
                | "ulnone" -> setStyle (fun s -> { s with Underline = false })
                | "sub" -> setStyle (fun s -> { s with Sub = on param; Super = false })
                | "super" -> setStyle (fun s -> { s with Super = on param; Sub = false })
                | "nosupersub" -> setStyle (fun s -> { s with Sub = false; Super = false })
                | "v"
                | "deleted" ->
                    flush ()
                    state <- { state with Hidden = on param }
                | "li" -> setPara (fun p -> { p with Left = value })
                | "ri" -> setPara (fun p -> { p with Right = value })
                | "fi" -> setPara (fun p -> { p with First = value })
                | "ql" -> setPara (fun p -> { p with Align = QuillLayout.AlignLeft })
                | "qc" -> setPara (fun p -> { p with Align = QuillLayout.AlignCenter })
                | "qr" -> setPara (fun p -> { p with Align = QuillLayout.AlignRight })
                | "qj"
                | "qd" -> setPara (fun p -> { p with Align = QuillLayout.AlignJustify })
                | "tqc" -> setPara (fun p -> { p with PendingTab = QuillLayout.TabCenter })
                | "tqr" -> setPara (fun p -> { p with PendingTab = QuillLayout.TabRight })
                | "tqdec" -> setPara (fun p -> { p with PendingTab = QuillLayout.TabDecimal })
                | "tx" ->
                    setPara (fun p ->
                        { p with
                            Tabs = (value, p.PendingTab) :: p.Tabs
                            PendingTab = QuillLayout.TabLeft })
                | "pagebb" -> setPara (fun p -> { p with PageBreakBefore = on param })
                | "paperw"
                | "paperh"
                | "margl"
                | "margr"
                | "margt"
                | "margb" -> docProps[word] <- value
                | "pgwsxn"
                | "pghsxn"
                | "marglsxn"
                | "margrsxn"
                | "margtsxn"
                | "margbsxn"
                | "headery"
                | "footery"
                | "pgnstarts" -> if not (firstSection.ContainsKey word) then firstSection[word] <- value
                | "header"
                | "headerr"
                | "footer"
                | "footerr" ->
                    flush ()
                    let isHeader = word.StartsWith "header"
                    let dest = if isHeader then Header else Footer
                    let taken = if isHeader then headerTaken else footerTaken
                    if isHeader then headerTaken <- true else footerTaken <- true
                    state <- { state with Dest = (if taken then Skip else dest) }

                    // Text left without a closing \par still belongs to this header or footer.
                    if not taken then
                        onClose <- (fun () -> if content[dest].Count > 0 then finishParagraph dest) :: onClose
                | "field" ->
                    flush ()
                    fields.Push { Instruction = StringBuilder(); Placeholder = None; Emitted = false }

                    onClose <-
                        (fun () ->
                            if fields.Count > 0 then
                                let field = fields.Pop()

                                // A PAGE field with an empty result still prints the page number.
                                match field.Placeholder with
                                | Some placeholder when not field.Emitted && isTextDest state.Dest && not state.Hidden ->
                                    appendText state.Dest placeholder state.Style
                                | _ -> ())
                        :: onClose
                | "fldinst" ->
                    flush ()
                    state <- { state with Dest = FieldInstruction }
                | "fldrslt" ->
                    flush ()

                    if fields.Count > 0 && isTextDest state.Dest then
                        let field = fields.Peek()
                        field.Placeholder <- pagePlaceholder (field.Instruction.ToString().Trim())

                        if field.Placeholder.IsSome then
                            state <- { state with PageField = Some field }
                | "docvar" ->
                    flush ()
                    docVarText <- StringBuilder()
                    state <- { state with Dest = DocVariable }
                    let captured = docVarText

                    onClose <-
                        (fun () ->
                            match captured.ToString().Split('\u0001') with
                            | [| _; name; value |] -> docVars[name] <- value
                            | _ -> ())
                        :: onClose
                | "listtext"
                | "pntext" ->
                    flush ()
                    state <- { state with InListText = true }
                | w when skipped.Contains w ->
                    flush ()
                    state <- { state with Dest = Skip }
                | _ -> ()

        let n = rtf.Length
        let mutable i = 0

        while i < n do
            let c = rtf[i]

            match c with
            | '{' ->
                flush ()
                if state.Dest = DocVariable then docVarText.Append('\u0001') |> ignore
                stack.Push((state, onClose))
                onClose <- []
                i <- i + 1
            | '}' ->
                flush ()
                for action in onClose do action ()

                // The document's closing brace: text after the last \par is a final paragraph (as
                // Word reads it), finished here while its properties still apply.
                if stack.Count = 1 && content[Body].Count > 0 then
                    state <- { state with Dest = Body }
                    finishParagraph Body

                if stack.Count > 0 then
                    let saved, savedOnClose = stack.Pop()
                    state <- saved
                    onClose <- savedOnClose

                starPending <- false
                i <- i + 1
            | '\\' when i + 1 < n ->
                let next = rtf[i + 1]

                if Char.IsAsciiLetter next then
                    let start = i + 1
                    let mutable j = start
                    while j < n && Char.IsAsciiLetter rtf[j] do j <- j + 1
                    let word = rtf.Substring(start, j - start)
                    let numberStart = j
                    if j < n && rtf[j] = '-' then j <- j + 1
                    while j < n && Char.IsAsciiDigit rtf[j] do j <- j + 1

                    let param =
                        if j > numberStart && rtf[j - 1] <> '-' then
                            match Int32.TryParse(rtf.AsSpan(numberStart, j - numberStart)) with
                            | true, v -> Some v
                            | _ -> None
                        else
                            None

                    if j < n && rtf[j] = ' ' then j <- j + 1

                    if word = "bin" then
                        // Raw binary data follows; it is never text.
                        j <- j + defaultArg param 0
                    else
                        // Fallback text for \uN is plain characters; a control word ends it.
                        if word <> "u" then fallbackToSkip <- 0
                        flush ()
                        controlWord word param

                    i <- j
                else
                    i <- i + 2

                    match next with
                    | '\\'
                    | '{'
                    | '}' -> text (string next)
                    | '\'' when i + 1 < n ->
                        match Byte.TryParse(rtf.AsSpan(i, 2), Globalization.NumberStyles.HexNumber, null) with
                        | true, b -> text (decodeByte b)
                        | _ -> ()

                        i <- i + 2
                    | '~' -> text " "
                    | '_' -> text "-"
                    | '-' ->
                        flush ()
                        addItem (RunSoftHyphen state.Style)
                    | '*' -> starPending <- true
                    | '\r'
                    | '\n' -> controlWord "par" None
                    | _ -> ()
            | '\r'
            | '\n' -> i <- i + 1
            | _ ->
                text (if int c < 0x80 then string c else decodeByte (byte c))
                i <- i + 1

        flush ()

        for d in [ Body; Header; Footer ] do
            if content[d].Count > 0 then
                state <- { state with Dest = d }
                finishParagraph d

        let prop (sectionWord: string) (docWord: string) =
            match firstSection.TryGetValue sectionWord with
            | true, v -> Some v
            | _ ->
                match docProps.TryGetValue docWord with
                | true, v -> Some v
                | _ -> None

        { Body = List.ofSeq paragraphs[Body]
          Header = List.ofSeq paragraphs[Header]
          Footer = List.ofSeq paragraphs[Footer]
          PaperWidth = prop "pgwsxn" "paperw"
          PaperHeight = prop "pghsxn" "paperh"
          MarginLeft = prop "marglsxn" "margl"
          MarginRight = prop "margrsxn" "margr"
          MarginTop = prop "margtsxn" "margt"
          MarginBottom = prop "margbsxn" "margb"
          HeaderY = prop "headery" "headery"
          FooterY = prop "footery" "footery"
          PageStart = prop "pgnstarts" "pgnstarts"
          DocVars = docVars |> Seq.map (fun kv -> kv.Key, kv.Value) |> Map.ofSeq }
