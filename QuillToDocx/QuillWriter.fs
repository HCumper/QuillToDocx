namespace QuillToDocx

open System
open System.IO

/// A paragraph as Quill stores it: margins and tab stops in character columns, text as runs.
type QuillParagraph =
    { LeftMarg: byte
      IndentMarg: byte
      RightMarg: byte
      Justif: byte
      TabTable: byte
      Content: Inline list }

/// Everything a Quill document file holds. The writer recomputes the layout's word count and tab
/// area size; the other layout fields are written as given.
type QuillDocument =
    { Layout: LayoutTable
      TabTables: TabTableDef list
      Header: Inline list
      Footer: Inline list
      Paragraphs: QuillParagraph list }

/// Writes QL Quill document files (big-endian). The structure follows what Quill itself leaves on
/// disk across 1,497 real documents: contiguous text from offset 20, tables that carry one dummy
/// entry more than they allocate, and a 38-byte free space table.
module QuillWriter =

    [<Literal>]
    let private HeaderLength = 20

    [<Literal>]
    let private ParaEntrySize = 14

    [<Literal>]
    let private ParaGranularity = 8

    [<Literal>]
    let private PageBreakFlag = 0x20uy

    let private plain =
        { Bold = false
          Italic = false
          Underline = false
          Sub = false
          Super = false }

    let private putU16 (out: ResizeArray<byte>) (value: int) =
        out.Add(byte ((value >>> 8) &&& 0xFF))
        out.Add(byte (value &&& 0xFF))

    let private putU32 (out: ResizeArray<byte>) (value: int) =
        putU16 out ((value >>> 16) &&& 0xFFFF)
        putU16 out (value &&& 0xFFFF)

    /// Paragraph text with Quill's toggle codes, ending in the 00 that also clears every toggle;
    /// like Quill, toggles still on at the end are left to it. Quill has no italic, so italic text
    /// is written plain.
    let encodeContent (content: Inline list) =
        let out = ResizeArray<byte>()
        let mutable current = plain

        let switchTo (style: TextStyle) =
            if style.Bold <> current.Bold then out.Add QuillDecode.Bold
            if style.Underline <> current.Underline then out.Add QuillDecode.Underline
            // Subscript and superscript exclude each other: switch one off before the other goes on.
            if current.Sub && not style.Sub then out.Add QuillDecode.SubScript
            if current.Super && not style.Super then out.Add QuillDecode.SuperScript
            if style.Sub && not current.Sub then out.Add QuillDecode.SubScript
            if style.Super && not current.Super then out.Add QuillDecode.SuperScript
            current <- { style with Italic = false }

        for item in content do
            match item with
            | RunText (text, style) ->
                switchTo style
                QuillCharset.appendEncoded out text
            | RunTab style ->
                switchTo style
                out.Add QuillDecode.Tab
            | RunSoftHyphen style ->
                switchTo style
                out.Add QuillDecode.SoftHyphen
            | RunPageBreak -> out.Add QuillDecode.FormFeed

        out.Add QuillDecode.EndPara
        out.ToArray()

    let private tabStopColumns (tables: TabTableDef list) (number: byte) =
        tables
        |> List.tryFind (fun t -> t.EntryNumber = number)
        |> Option.map (fun t -> t.Entries |> List.map (fun e -> int e.Pos) |> List.sort)
        |> Option.defaultValue []

    /// Lines the paragraph fills: the first from the indent margin, the rest from the left margin,
    /// wrapped at spaces. Quill keeps this count in each paragraph entry.
    let private lineCount (tables: TabTableDef list) (p: QuillParagraph) (text: byte[]) =
        let stops = tabStopColumns tables p.TabTable
        let right = max (int p.RightMarg) (max (int p.LeftMarg) (int p.IndentMarg) + 1)
        let mutable lines = 1
        let mutable col = int p.IndentMarg
        let mutable lastSpaceCol = -1

        for b in text do
            if b = QuillDecode.Tab then
                col <-
                    stops
                    |> List.tryFind (fun s -> s > col)
                    |> Option.defaultValue (((col / 8) + 1) * 8)
            elif b = 0x20uy then
                col <- col + 1
                lastSpaceCol <- col
            elif b > 0x20uy && b < 0xC0uy then
                col <- col + 1

                if col > right then
                    lines <- lines + 1
                    let carried = if lastSpaceCol >= 0 then col - lastSpaceCol else 1
                    col <- int p.LeftMarg + carried
                    lastSpaceCol <- -1

        min lines 255

    let private wordCount (texts: byte[] seq) =
        let mutable words = 0

        for text in texts do
            let mutable inWord = false

            for b in text do
                let letter = b > 0x20uy && b < 0xC0uy
                if letter && not inWord then words <- words + 1
                inWord <- letter

        words

    let tabArea (tables: TabTableDef list) =
        let out = ResizeArray<byte>()

        for table in tables |> List.sortBy (fun t -> t.EntryNumber) do
            out.Add table.EntryNumber
            out.Add(byte (2 + 2 * table.Entries.Length))

            for e in table.Entries do
                out.Add e.Pos
                out.Add e.Type

        // An entry number of zero ends the list; the terminator counts towards the size used.
        out.Add 0uy
        out.Add 0uy
        out.ToArray()

    let toBytes (doc: QuillDocument) =
        let header = encodeContent doc.Header
        let footer = encodeContent doc.Footer
        let body = doc.Paragraphs |> List.map (fun p -> p, encodeContent p.Content)

        let out = ResizeArray<byte>()

        // File header; the four lengths are patched once the areas are written.
        putU16 out HeaderLength
        for c in "vrm1qdf0" do out.Add(byte c)
        putU32 out 0
        putU16 out 0
        putU16 out 0
        putU16 out 0

        // Text area.
        let offsets = ResizeArray<int>()

        for text in Seq.append [ header; footer ] (body |> Seq.map snd) do
            offsets.Add out.Count
            out.AddRange text

        let textLen = out.Count

        // Paragraph table: entry 0 is a dummy, then header, footer and body; allocated in eights.
        let used = 2 + body.Length
        let alloc = ((used + 1 + ParaGranularity - 1) / ParaGranularity) * ParaGranularity
        putU16 out ParaEntrySize
        putU16 out ParaGranularity
        putU16 out used
        putU16 out alloc
        out.AddRange(Array.zeroCreate<byte> ParaEntrySize)

        let putEntry offset (length: int) flags left indent right justif tabTable lineByte =
            putU32 out offset
            putU16 out length
            out.Add flags
            out.Add left
            out.Add indent
            out.Add right
            out.Add justif
            out.Add tabTable
            out.Add 0uy
            out.Add lineByte

        // The header and footer entries' margins are never used; Quill's defaults go there.
        putEntry offsets[0] header.Length 0uy 9uy 9uy 72uy 0uy 1uy 1uy
        putEntry offsets[1] footer.Length 0uy 9uy 9uy 72uy 0uy 1uy 1uy

        body
        |> List.iteri (fun i (p, text) ->
            let flags = if Array.contains QuillDecode.FormFeed text then PageBreakFlag else 0uy
            let lines = if text.Length <= 1 then 1 else lineCount doc.TabTables p text
            putEntry offsets[i + 2] text.Length flags p.LeftMarg p.IndentMarg p.RightMarg p.Justif p.TabTable (byte lines))

        out.AddRange(Array.zeroCreate<byte> ((alloc - used) * ParaEntrySize))
        let paraLen = out.Count - textLen

        // Free space table: always the same 38 bytes in Quill's files, with nothing in use.
        putU16 out 6
        putU16 out 4
        putU16 out 0
        putU16 out 4
        out.AddRange(Array.zeroCreate<byte> 30)
        let freeLen = out.Count - textLen - paraLen

        // Layout table, then the tab area padded to its allocated size.
        let tabs = tabArea doc.TabTables
        let maxTabSize = max (max (int doc.Layout.MaxTabSize) 64) (((tabs.Length + 15) / 16) * 16)
        let layoutStart = out.Count
        let l = doc.Layout
        out.Add l.BottomMarg
        out.Add l.DispMode
        out.Add l.LineGap
        out.Add l.PageLen
        out.Add l.StartPage
        out.Add l.Color
        out.Add l.TopMargin
        out.Add 0uy
        putU16 out (min 0xFFFF (wordCount (body |> Seq.map snd)))
        putU16 out maxTabSize
        putU16 out tabs.Length
        out.Add l.HeaderF
        out.Add l.FooterF
        out.Add l.HeaderMarg
        out.Add l.FooterMarg
        out.Add l.HeaderBold
        out.Add l.FooterBold
        out.AddRange tabs
        // Quill's layout length is four bytes more than the fixed part and the tab area.
        out.AddRange(Array.zeroCreate<byte> (maxTabSize - tabs.Length + 4))
        let layoutLen = out.Count - layoutStart

        // Quill rounds files up to 2048 bytes, and beyond that to a multiple of 512.
        let size = max 2048 (((out.Count + 511) / 512) * 512)
        out.AddRange(Array.zeroCreate<byte> (size - out.Count))

        let bytes = out.ToArray()

        let patchU16 offset (value: int) =
            bytes[offset] <- byte ((value >>> 8) &&& 0xFF)
            bytes[offset + 1] <- byte (value &&& 0xFF)

        patchU16 10 (textLen >>> 16)
        patchU16 12 (textLen &&& 0xFFFF)
        patchU16 14 paraLen
        patchU16 16 freeLen
        patchU16 18 layoutLen
        bytes

    /// A Quill document read back for rewriting: every paragraph as stored, page breaks in place.
    let fromQuillFile (qf: QuillFile) =
        let content index =
            if index < qf.Paras.Length then
                try QuillDecode.entryContent qf index with _ -> []
            else
                []

        { Layout = qf.Layout
          TabTables = qf.TabTables
          Header = content 1
          Footer = content 2
          Paragraphs =
            [ for i in 3 .. qf.Paras.Length - 1 do
                  let p = qf.Paras[i]
                  // PC Quill counts justification from 4 and keeps line spacing in the top four bits.
                  let justif =
                      if qf.IsPcFile && p.Justif &&& 0x0Fuy >= 4uy then (p.Justif &&& 0x0Fuy) - 4uy
                      else p.Justif

                  { LeftMarg = p.LeftMarg
                    IndentMarg = p.IndentMarg
                    RightMarg = p.RightMarg
                    Justif = justif
                    TabTable = p.TabTable
                    Content = content i } ] }

    /// Tab tables from a tab area as stored after the layout table.
    let parseTabArea (area: byte[]) =
        let tables = ResizeArray<TabTableDef>()
        let mutable pos = 0

        while pos + 1 < area.Length && area[pos] <> 0uy && area[pos + 1] >= 2uy do
            let length = int area[pos + 1]

            tables.Add
                { EntryNumber = area[pos]
                  Entries =
                    [ for off in pos + 2 .. 2 .. min (pos + length) area.Length - 2 ->
                          { Pos = area[off]; Type = area[off + 1] } ] }

            pos <- pos + length

        List.ofSeq tables

    /// Writes through a temporary file, so a failure never leaves half a document behind.
    let write (path: string) (doc: QuillDocument) =
        let bytes = toBytes doc
        let temp = path + ".quilltmp"
        File.WriteAllBytes(temp, bytes)
        File.Move(temp, path, true)

/// What Word can't hold of a Quill document, carried through Word in a document variable so that
/// saving back to Quill restores it: the layout and tab tables exactly as stored, and the header
/// and footer text, kept even when the layout switches them off.
module QuillOriginal =

    [<Literal>]
    let VariableName = "QuillLayout"

    [<Literal>]
    let private FormatVersion = 1uy

    type Original =
        { Layout: LayoutTable
          TabTables: TabTableDef list
          Header: Inline list
          Footer: Inline list }

    let private layoutBytes (l: LayoutTable) =
        [| l.BottomMarg; l.DispMode; l.LineGap; l.PageLen; l.StartPage; l.Color; l.TopMargin; l.Dummy1
           byte (l.WordCount >>> 8); byte l.WordCount
           byte (l.MaxTabSize >>> 8); byte l.MaxTabSize
           byte (l.TabSize >>> 8); byte l.TabSize
           l.HeaderF; l.FooterF; l.HeaderMarg; l.FooterMarg; l.HeaderBold; l.FooterBold |]

    let private layoutFromBytes (b: byte[]) =
        let u16 i = (uint16 b[i] <<< 8) ||| uint16 b[i + 1]

        { BottomMarg = b[0]; DispMode = b[1]; LineGap = b[2]; PageLen = b[3]; StartPage = b[4]; Color = b[5]
          TopMargin = b[6]; Dummy1 = b[7]; WordCount = u16 8; MaxTabSize = u16 10; TabSize = u16 12
          HeaderF = b[14]; FooterF = b[15]; HeaderMarg = b[16]; FooterMarg = b[17]; HeaderBold = b[18]; FooterBold = b[19] }

    let private withLength (bytes: byte[]) =
        Array.append [| byte (bytes.Length >>> 8); byte bytes.Length |] bytes

    /// The document variable's value for a Quill file: hex, since Word keeps it as plain text.
    let encode (qf: QuillFile) =
        let doc = QuillWriter.fromQuillFile qf

        Array.concat
            [ [| FormatVersion |]
              layoutBytes qf.Layout
              withLength (QuillWriter.tabArea qf.TabTables)
              withLength (QuillWriter.encodeContent doc.Header)
              withLength (QuillWriter.encodeContent doc.Footer) ]
        |> Convert.ToHexString

    let tryDecode (hex: string) =
        try
            let b = Convert.FromHexString(hex.Trim())

            if b.Length < 21 || b[0] <> FormatVersion then
                None
            else
                let mutable pos = 21

                let next () =
                    let length = (int b[pos] <<< 8) ||| int b[pos + 1]
                    let part = b[pos + 2 .. pos + 1 + length]
                    pos <- pos + 2 + length
                    part

                let tabs = next ()
                let header = next ()
                let footer = next ()

                Some
                    { Layout = layoutFromBytes b[1..20]
                      TabTables = QuillWriter.parseTabArea tabs
                      Header = QuillDecode.decodeText header
                      Footer = QuillDecode.decodeText footer }
        with _ ->
            None
