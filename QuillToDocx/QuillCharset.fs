namespace QuillToDocx

open System
open System.Collections.Generic
open System.Globalization
open System.Text

/// The Sinclair QL character set: ASCII except 0x60 (pound) and 0x7F (copyright), then accented
/// letters, Greek letters and arrows from 0x80 to 0xBF.
module QuillCharset =

    /// QL characters 0x80-0xBF. 0xB5 was an odd script mark in the original ROM; SMSQ/E made it
    /// the euro sign, which is how documents use it.
    let private highCharacters =
        "äãåéöõøüçñæœáàâë" // 80
        + "èêïíìîóòôúùûß¢¥`" // 90
        + "ÄÃÅÉÖÕØÜÇÑÆŒαδθλ" // A0
        + "µπφ¡¿€§¤«»°÷←→↑↓" // B0

    /// Unicode for a printable QL byte (0x20 and up); control codes are the caller's business.
    let toUnicode (b: byte) =
        match int b with
        | 0x60 -> '£'
        | 0x7F -> '©'
        | x when x < 0x7F -> char x
        | x when x <= 0xBF -> highCharacters[x - 0x80]
        // Not QL characters; the few files that use them hold Latin-1 text (0xE9 for e-acute).
        | x -> char x

    let private byCharacter =
        let map = Dictionary<char, byte>()
        for b in 0x20 .. 0xBF do
            map[toUnicode (byte b)] <- byte b
        map

    /// Characters the QL lacks, written the way a QL user would have typed them.
    let private substitutes =
        dict
            [ '‘', "'"; '’', "'"; '‚', "'"; '′', "'"
              '“', "\""; '”', "\""; '„', "\""; '″', "\""
              '‐', "-"; '‑', "-"; '‒', "-"; '–', "-"; '—', "-"; '―', "-"; '−', "-"
              '…', "..."; '•', "*"; '▪', "*"; '●', "*"; '·', "*"
              ' ', " "; ' ', " "; ' ', " "; ' ', " "; ' ', " "
              '™', "TM"; '®', "(R)"; '½', "1/2"; '¼', "1/4"; '¾', "3/4"
              '×', "x"; '±', "+/-"; 'μ', "µ"; 'Π', "π"; 'Φ', "φ"; 'ϕ', "φ" ]

    /// Appends the QL bytes for a Unicode string. Characters the QL can't show become the nearest
    /// ASCII (smart quotes, dashes, unaccented letters) or '?'. Control characters are dropped.
    let appendEncoded (out: ResizeArray<byte>) (text: string) =
        let rec put (ch: char) =
            match byCharacter.TryGetValue ch with
            | true, b -> out.Add b
            | _ ->
                match substitutes.TryGetValue ch with
                | true, s -> for c in s do put c
                | _ ->
                    // Strip accents the QL doesn't have: a-macron -> a.
                    let baseLetter = (string ch).Normalize(NormalizationForm.FormD)[0]

                    if baseLetter <> ch && byCharacter.ContainsKey baseLetter then
                        out.Add byCharacter[baseLetter]
                    else
                        out.Add(byte '?')

        let mutable i = 0

        while i < text.Length do
            let ch = text[i]

            if Char.IsHighSurrogate ch && i + 1 < text.Length then
                out.Add(byte '?')
                i <- i + 1
            elif ch >= ' ' then
                put ch

            i <- i + 1
