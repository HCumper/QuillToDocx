namespace QuillToDocx

open System
open RtfReader

/// Turns the RTF Word saves into a Quill document. Indents and tab stops go back to character
/// columns (120 twips each, as the import wrote them). What Word can't hold, such as the display
/// mode, the tab tables' numbering, and a header that is switched off, comes from the original's
/// layout, which the import left in a document variable.
module QuillFromRtf =

    [<Literal>]
    let private TwipsPerColumn = 120

    [<Literal>]
    let private TwipsPerLine = 240

    /// Quill's own defaults, for documents that did not start as Quill files.
    let private defaultLayout: LayoutTable =
        { BottomMarg = 6uy
          DispMode = 0uy
          LineGap = 0uy
          PageLen = 66uy
          StartPage = 1uy
          Color = 0uy
          TopMargin = 6uy
          Dummy1 = 0uy
          WordCount = 0us
          MaxTabSize = 64us
          TabSize = 0us
          HeaderF = 0uy
          FooterF = 0uy
          HeaderMarg = 2uy
          FooterMarg = 2uy
          HeaderBold = 0uy
          FooterBold = 0uy }

    let private toByte (value: int) = byte (max 0 (min 255 value))

    let private round (twips: int) (unit: int) =
        int (Math.Round(float twips / float unit, MidpointRounding.AwayFromZero))

    let private columns twips = round twips TwipsPerColumn

    let private isVisible item =
        match item with
        | RunText (text, _) -> not (String.IsNullOrWhiteSpace text)
        | _ -> false

    /// Quill's one-line header or footer from Word's: the text, the layout flag for its alignment
    /// (1 left, 2 centre, 3 right) and whether it is all bold, which Quill keeps as a flag.
    let private headerLine (paragraphs: RtfParagraph list) =
        match paragraphs |> List.filter (fun p -> List.exists isVisible p.Content) with
        | [] -> None
        | visible ->
            let lineOf (p: RtfParagraph) =
                p.Content
                |> List.choose (function
                    | RunText (text, style) -> Some(RunText(text, style))
                    | RunTab style -> Some(RunText(" ", style))
                    | RunSoftHyphen _
                    | RunPageBreak -> None)

            let items =
                visible
                |> List.map lineOf
                |> List.reduce (fun a b -> a @ [ RunText(" ", { Bold = false; Italic = false; Underline = false; Sub = false; Super = false }) ] @ b)

            let allBold =
                items
                |> List.forall (function
                    | RunText (text, style) -> style.Bold || String.IsNullOrWhiteSpace text
                    | _ -> true)

            let items =
                if allBold then
                    items
                    |> List.map (function
                        | RunText (text, style) -> RunText(text, { style with Bold = false })
                        | other -> other)
                else
                    items

            let flag =
                match visible.Head.Alignment with
                | QuillLayout.AlignCenter -> 2uy
                | QuillLayout.AlignRight -> 3uy
                | _ -> 1uy

            Some(items, flag, allBold)

    let fromRtf (rtf: RtfDocument) =
        let original =
            rtf.DocVars.TryFind QuillOriginal.VariableName |> Option.bind QuillOriginal.tryDecode

        let baseLayout =
            original |> Option.map (fun o -> o.Layout) |> Option.defaultValue defaultLayout

        // Keep a stored value for as long as Word still shows what the import made of it, so an
        // unchanged setting survives exactly (including values the import had to clamp).
        // Only the heights are compared, so the body width doesn't matter here.
        let imported = QuillLayout.pageGeometry baseLayout 80

        let keep (fromWord: int option) importedTwips (stored: byte) =
            match fromWord with
            | Some twips when twips <> importedTwips -> toByte (round twips TwipsPerLine)
            | _ -> stored

        let startPage =
            match rtf.PageStart with
            | Some 1 when baseLayout.StartPage <= 1uy -> baseLayout.StartPage
            | Some page -> toByte page
            | None -> if baseLayout.StartPage > 1uy then 1uy else baseLayout.StartPage

        // RTF's own defaults when Word leaves the page size out.
        let textWidth =
            defaultArg rtf.PaperWidth 12240 - defaultArg rtf.MarginLeft 1800 - defaultArg rtf.MarginRight 1800

        let tables =
            ResizeArray<TabTableDef>(original |> Option.map (fun o -> o.TabTables) |> Option.defaultValue [])

        let tableFor (stops: (int * QuillLayout.TabKind) list) =
            let entries =
                stops
                |> List.map (fun (position, kind) ->
                    { Pos = toByte (columns position)
                      Type =
                        match kind with
                        | QuillLayout.TabCenter -> 1uy
                        | QuillLayout.TabRight -> 2uy
                        | QuillLayout.TabDecimal -> 3uy
                        | QuillLayout.TabLeft -> 0uy })
                |> List.distinctBy (fun e -> e.Pos)
                |> List.sortBy (fun e -> e.Pos)

            match tables |> Seq.tryFind (fun t -> List.sortBy (fun e -> e.Pos) t.Entries = entries) with
            | Some table -> table.EntryNumber
            | None ->
                let number =
                    if tables.Count = 0 then 1uy
                    else (tables |> Seq.map (fun t -> t.EntryNumber) |> Seq.max) + 1uy

                tables.Add { EntryNumber = number; Entries = entries }
                number

        let toQuill (p: RtfParagraph) =
            let left = max 0 (columns p.LeftIndent)
            let first = max 0 (columns (p.LeftIndent + p.FirstIndent))

            // Quill centres between the margins and has no right alignment.
            let justif, indent =
                match p.Alignment with
                | QuillLayout.AlignCenter -> 1uy, left
                | QuillLayout.AlignJustify -> 2uy, first
                | _ -> 0uy, first

            let right = max (columns (textWidth - p.RightIndent)) (max left indent + 1)

            { LeftMarg = toByte left
              IndentMarg = toByte indent
              RightMarg = toByte right
              Justif = justif
              TabTable = tableFor p.Tabs
              Content = p.Content }

        let body =
            [ for p in rtf.Body do
                  let paragraph = toQuill p
                  // Quill keeps a page break as a paragraph holding just the form feed.
                  if p.PageBreakBefore then yield { paragraph with Content = [ RunPageBreak ] }
                  yield paragraph ]

        let body =
            if body.IsEmpty then
                [ { LeftMarg = 9uy; IndentMarg = 9uy; RightMarg = 72uy; Justif = 0uy; TabTable = tableFor []; Content = [] } ]
            else
                body

        let header = headerLine rtf.Header
        let footer = headerLine rtf.Footer

        let text line (stored: Inline list) =
            match line with
            | Some (items, _, _) -> items
            | None -> stored

        // With nothing to show, the header is off, unless the original's was on with no text in it:
        // that prints nothing too, so it stays as it was.
        let flag line (stored: byte) (storedText: Inline list option) =
            match line, storedText with
            | Some (_, f, _), _ -> f
            | None, Some text when not (List.exists isVisible text) -> stored
            | None, _ -> 0uy

        let bold line = line |> Option.map (fun (_, _, b) -> if b then 1uy else 0uy) |> Option.defaultValue 0uy

        let layout =
            { baseLayout with
                PageLen = keep rtf.PaperHeight imported.Height baseLayout.PageLen
                TopMargin = keep rtf.MarginTop imported.TopMargin baseLayout.TopMargin
                BottomMarg = keep rtf.MarginBottom imported.BottomMargin baseLayout.BottomMarg
                HeaderMarg = keep rtf.HeaderY imported.HeaderMargin baseLayout.HeaderMarg
                FooterMarg = keep rtf.FooterY imported.FooterMargin baseLayout.FooterMarg
                StartPage = startPage
                HeaderF = flag header baseLayout.HeaderF (original |> Option.map (fun o -> o.Header))
                FooterF = flag footer baseLayout.FooterF (original |> Option.map (fun o -> o.Footer))
                HeaderBold = (if header.IsSome then bold header else baseLayout.HeaderBold)
                FooterBold = (if footer.IsSome then bold footer else baseLayout.FooterBold) }

        { Layout = layout
          TabTables = List.ofSeq tables
          Header = text header (original |> Option.map (fun o -> o.Header) |> Option.defaultValue [])
          Footer = text footer (original |> Option.map (fun o -> o.Footer) |> Option.defaultValue [])
          Paragraphs = body }

    let convert (rtfText: string) = fromRtf (RtfReader.parse rtfText)
