namespace QuillToDocx

open System
open System.IO
open System.Text

module PdfWriter =

    [<Literal>]
    let private FontName = "Courier"

    [<Literal>]
    let private FontSizePt = 10.0

    [<Literal>]
    let private DefaultPageMarginTwips = 1440

    [<Literal>]
    let private DefaultPageWidthTwips = 12480

    [<Literal>]
    let private DefaultPageHeightTwips = 15840

    [<Literal>]
    let private DefaultBodyWidthCols = 80

    let private twipsFromCols cols =
        cols * 120

    let private clampTwips defaultValue minValue maxValue rawValue =
        if rawValue >= minValue && rawValue <= maxValue then rawValue else defaultValue

    let private validatedPageMargin rawValue =
        clampTwips DefaultPageMarginTwips 720 2880 rawValue

    let private lineSpacingTwips (layout: LayoutTable) =
        clampTwips 240 200 480 (240 + (int layout.LineGap * 10))

    let private paragraphAfterTwips (layout: LayoutTable) =
        clampTwips 80 0 240 (lineSpacingTwips layout / 3)

    let private pageHeightTwips (layout: LayoutTable) topMargin bottomMargin =
        let candidate = topMargin + bottomMargin + (int layout.PageLen * lineSpacingTwips layout)
        clampTwips DefaultPageHeightTwips 10080 20160 candidate

    let private pageWidthTwips leftMargin rightMargin =
        clampTwips DefaultPageWidthTwips 10080 15840 (leftMargin + rightMargin + twipsFromCols DefaultBodyWidthCols)

    let private pointsFromTwips twips =
        float twips / 20.0

    let private charWidthPt =
        FontSizePt * 0.6

    let private mapJustification j =
        match int j with
        | 1
        | 5 -> "center"
        | 2
        | 6 -> "right"
        | _ -> "left"

    let private tabStopsForParagraph (qf: QuillFile) (p: ParaTable) =
        qf.TabTables
        |> List.tryFind (fun t -> t.EntryNumber = p.TabTable)
        |> Option.map (fun t -> t.Entries |> List.map (fun e -> int e.Pos))
        |> Option.defaultValue []

    let private nextTabColumn current tabStops =
        tabStops
        |> List.tryFind (fun stop -> stop > current)
        |> Option.defaultValue (((current / 8) + 1) * 8)

    let private flattenParagraphText (qf: QuillFile) (p: DecodedParagraph) =
        let sb = StringBuilder()
        let mutable currentCol = max (int p.Descriptor.LeftMarg) (int p.Descriptor.IndentMarg)
        let tabStops = tabStopsForParagraph qf p.Descriptor

        for item in p.Content do
            match item with
            | RunText (text, _) ->
                sb.Append(text) |> ignore
                currentCol <- currentCol + text.Length
            | RunTab _ ->
                let nextCol = nextTabColumn currentCol tabStops
                let spaces = max 1 (nextCol - currentCol)
                sb.Append(' ', spaces) |> ignore
                currentCol <- currentCol + spaces
            | RunSoftHyphen _ ->
                sb.Append('-') |> ignore
                currentCol <- currentCol + 1

        sb.ToString()

    let rec private wrapText width (text: string) =
        if String.IsNullOrEmpty(text) || text.Length <= width then
            [ text ]
        else
            let breakAt =
                let candidate = text.LastIndexOf(' ', width - 1, width)
                if candidate > 0 then candidate else width

            let head = text.Substring(0, breakAt).TrimEnd()
            let tail = text.Substring(breakAt).TrimStart()
            head :: wrapText width tail

    let private paragraphLines (qf: QuillFile) (p: DecodedParagraph) =
        let leftCols = int p.Descriptor.LeftMarg
        let firstCols = max leftCols (int p.Descriptor.IndentMarg)
        let rightCols =
            if int p.Descriptor.RightMarg > leftCols then
                int p.Descriptor.RightMarg
            else
                DefaultBodyWidthCols

        let firstWidth = max 10 (rightCols - firstCols)
        let otherWidth = max 10 (rightCols - leftCols)
        let rawText = flattenParagraphText qf p
        let initialLines = wrapText firstWidth rawText

        initialLines
        |> List.mapi (fun index line ->
            if index = 0 then
                firstCols, firstWidth, line
            else
                leftCols, otherWidth, line)

    let private pdfEscape (text: string) =
        let sb = StringBuilder()

        for ch in text do
            match ch with
            | '\\' -> sb.Append(@"\\") |> ignore
            | '(' -> sb.Append(@"\(") |> ignore
            | ')' -> sb.Append(@"\)") |> ignore
            | '\r'
            | '\n' -> sb.Append(' ') |> ignore
            | _ when int ch < 32 -> sb.Append('?') |> ignore
            | _ when int ch > 126 -> sb.Append('?') |> ignore
            | _ -> sb.Append(ch) |> ignore

        sb.ToString()

    let private lineCommand (_pageWidth: float) (leftMargin: float) (_rightMargin: float) (justification: string) (indentCols: int) (widthCols: int) (text: string) (y: float) =
        let indentX = leftMargin + (float indentCols * charWidthPt)
        let availableWidth = max 0.0 (float widthCols * charWidthPt)
        let textWidth = float text.Length * charWidthPt
        let x =
            match justification with
            | "center" -> indentX + max 0.0 ((availableWidth - textWidth) / 2.0)
            | "right" -> indentX + max 0.0 (availableWidth - textWidth)
            | _ -> indentX

        sprintf "BT /F1 %.1f Tf 1 0 0 1 %.2f %.2f Tm (%s) Tj ET" FontSizePt x y (pdfEscape text)

    let private renderPages (qf: QuillFile) (doc: DecodedDocument) =
        let topMarginTwips = validatedPageMargin (int qf.Layout.TopMargin * 240)
        let bottomMarginTwips = validatedPageMargin (int qf.Layout.BottomMarg * 240)
        let leftMarginTwips = DefaultPageMarginTwips
        let rightMarginTwips = DefaultPageMarginTwips

        let pageWidthPt = pointsFromTwips (pageWidthTwips leftMarginTwips rightMarginTwips)
        let pageHeightPt = pointsFromTwips (pageHeightTwips qf.Layout topMarginTwips bottomMarginTwips)
        let leftMarginPt = pointsFromTwips leftMarginTwips
        let rightMarginPt = pointsFromTwips rightMarginTwips
        let topMarginPt = pointsFromTwips topMarginTwips
        let bottomMarginPt = pointsFromTwips bottomMarginTwips
        let lineAdvancePt = pointsFromTwips (lineSpacingTwips qf.Layout)
        let paragraphAfterPt = pointsFromTwips (paragraphAfterTwips qf.Layout)

        let pages = ResizeArray<string>()
        let currentCommands = ResizeArray<string>()
        let mutable currentY = pageHeightPt - topMarginPt - FontSizePt

        let flushPage () =
            if currentCommands.Count > 0 then
                pages.Add(String.Join("\n", currentCommands))
                currentCommands.Clear()
                currentY <- pageHeightPt - topMarginPt - FontSizePt

        let ensureSpace requiredHeight =
            if currentY - requiredHeight < bottomMarginPt then
                flushPage ()

        for paragraph in doc.BodyParagraphs do
            let lines = paragraphLines qf paragraph
            let justification = mapJustification paragraph.Descriptor.Justif
            let paragraphHeight =
                if List.isEmpty lines then
                    lineAdvancePt + paragraphAfterPt
                else
                    (float lines.Length * lineAdvancePt) + paragraphAfterPt

            ensureSpace paragraphHeight

            match lines with
            | [] ->
                currentY <- currentY - lineAdvancePt - paragraphAfterPt
            | _ ->
                for indentCols, widthCols, text in lines do
                    currentCommands.Add(lineCommand pageWidthPt leftMarginPt rightMarginPt justification indentCols widthCols text currentY)
                    currentY <- currentY - lineAdvancePt

                currentY <- currentY - paragraphAfterPt

        flushPage ()

        if pages.Count = 0 then
            pages.Add("")

        pageWidthPt, pageHeightPt, pages |> Seq.toList

    let private appendObject (buffer: StringBuilder) (offsets: ResizeArray<int>) (id: int) (body: string) =
        offsets.Add(buffer.Length) |> ignore
        buffer.Append(sprintf "%d 0 obj\n" id) |> ignore
        buffer.Append(body) |> ignore
        if not (body.EndsWith("\n")) then
            buffer.Append('\n') |> ignore
        buffer.Append("endobj\n") |> ignore

    let writePdf (qf: QuillFile) (doc: DecodedDocument) (outPath: string) =
        if File.Exists(outPath) then
            File.Delete(outPath)

        let pageWidthPt, pageHeightPt, pageStreams = renderPages qf doc
        let pageCount = pageStreams.Length
        let pageObjectId index = 4 + (index * 2)
        let contentObjectId index = 5 + (index * 2)
        let maxObjectId = 3 + (pageCount * 2)
        let buffer = StringBuilder()
        let offsets = ResizeArray<int>()

        buffer.Append("%PDF-1.4\n") |> ignore

        appendObject buffer offsets 1 "<< /Type /Catalog /Pages 2 0 R >>"

        let kids =
            [ for index in 0 .. pageCount - 1 -> sprintf "%d 0 R" (pageObjectId index) ]
            |> String.concat " "

        appendObject buffer offsets 2 (sprintf "<< /Type /Pages /Count %d /Kids [ %s ] >>" pageCount kids)
        appendObject buffer offsets 3 (sprintf "<< /Type /Font /Subtype /Type1 /BaseFont /%s >>" FontName)

        for index, stream in pageStreams |> List.indexed do
            let streamBytes = Encoding.ASCII.GetByteCount(stream)
            let pageBody =
                sprintf
                    "<< /Type /Page /Parent 2 0 R /MediaBox [ 0 0 %.2f %.2f ] /Resources << /Font << /F1 3 0 R >> >> /Contents %d 0 R >>"
                    pageWidthPt
                    pageHeightPt
                    (contentObjectId index)

            appendObject buffer offsets (pageObjectId index) pageBody
            appendObject buffer offsets (contentObjectId index) (sprintf "<< /Length %d >>\nstream\n%s\nendstream" streamBytes stream)

        let xrefStart = buffer.Length
        buffer.AppendFormat("xref\n0 {0}\n", maxObjectId + 1) |> ignore
        buffer.Append("0000000000 65535 f \n") |> ignore

        for offset in offsets do
            buffer.AppendFormat("{0:0000000000} 00000 n \n", offset) |> ignore

        buffer.AppendFormat("trailer\n<< /Size {0} /Root 1 0 R >>\nstartxref\n{1}\n%%EOF\n", maxObjectId + 1, xrefStart)
        |> ignore

        File.WriteAllText(outPath, buffer.ToString(), Encoding.ASCII)
