namespace QuillToDocx

open System
open System.IO

module Program =

    // QL names such as "report_doc" have no extension, so only add ".doc" to a basename
    // when the name as given doesn't exist.
    let private resolveInputPath (inputPath: string) =
        let ext = Path.GetExtension(inputPath)
        if String.IsNullOrWhiteSpace(ext) && not (File.Exists inputPath) then inputPath + ".doc" else inputPath

    let private usage () =
        eprintfn "Usage: QuillToDocx <input.doc> [output.docx|output.rtf|output.pdf|output.html|output.md]"
        eprintfn "       QuillToDocx <basename>"
        eprintfn "       QuillToDocx <input.rtf|input.doc> <output.doc> [--quill]"
        eprintfn "       Writes a Quill document; --quill forces that for names without .doc (report_doc)."

    [<Literal>]
    let private QuillFlag = "--quill"

    let private isRtfFile (path: string) =
        use stream = File.OpenRead path
        let head = Array.zeroCreate<byte> 5
        stream.Read(head, 0, 5) = 5 && Text.Encoding.ASCII.GetString(head) = @"{\rtf"

    let private isQuillFile (path: string) =
        use stream = File.OpenRead path
        let head = Array.zeroCreate<byte> 10
        stream.Read(head, 0, 10) = 10 && Text.Encoding.ASCII.GetString(head, 2, 8) = "vrm1qdf0"

    /// Writes a Quill document from RTF (what Word saves through the converter) or from another
    /// Quill document. The first time a Quill document is replaced it is kept as <name>.bak.
    let private writeQuill (inputPath: string) (outputPath: string) =
        let doc =
            if isRtfFile inputPath then
                QuillFromRtf.convert (File.ReadAllText(inputPath, Text.Encoding.Latin1))
            else
                QuillWriter.fromQuillFile (QuillParser.parseFile inputPath)

        let backup = outputPath + ".bak"

        if File.Exists outputPath && not (File.Exists backup) && isQuillFile outputPath then
            File.Copy(outputPath, backup)

        QuillWriter.write outputPath doc

    let private resolvePaths (argv: string array) =
        match argv with
        | [| singleArg |] ->
            let ext = Path.GetExtension(singleArg)

            if
                String.Equals(ext, ".docx", StringComparison.OrdinalIgnoreCase)
                || String.Equals(ext, ".rtf", StringComparison.OrdinalIgnoreCase)
                || String.Equals(ext, ".pdf", StringComparison.OrdinalIgnoreCase)
                || String.Equals(ext, ".html", StringComparison.OrdinalIgnoreCase)
                || String.Equals(ext, ".md", StringComparison.OrdinalIgnoreCase)
            then
                Error "Single-argument mode expects a basename or a .doc input file."
            else
                let inputPath = resolveInputPath singleArg

                let outputPath =
                    match Path.ChangeExtension(inputPath, ".docx") with
                    | null -> failwithf "Unable to resolve output path for '%s'" inputPath
                    | path -> path

                Ok(inputPath, outputPath)

        | [| inputPath; outputPath |] -> Ok(resolveInputPath inputPath, outputPath)
        | _ -> Error "Invalid arguments."

    [<EntryPoint>]
    let main argv =
        try
            let forceQuill = argv |> Array.exists (fun a -> a = QuillFlag)

            match resolvePaths (argv |> Array.filter (fun a -> a <> QuillFlag)) with
            | Error message ->
                eprintfn "%s" message
                usage ()
                1
            | Ok(inputPath, _) when not (File.Exists inputPath) ->
                eprintfn "Input file not found: %s" inputPath
                1
            | Ok(inputPath, outputPath) when
                forceQuill || String.Equals(Path.GetExtension outputPath, ".doc", StringComparison.OrdinalIgnoreCase)
                ->
                writeQuill inputPath outputPath
                printfn "Converted '%s' to Quill document '%s'" inputPath outputPath
                0
            | Ok(inputPath, outputPath) ->
                let parsed = QuillParser.parseFile inputPath
                let decoded = QuillDecode.decodeDocument parsed
                let outputExtension =
                    match Path.GetExtension(outputPath) with
                    | null -> ""
                    | extension -> extension

                match outputExtension.ToLowerInvariant() with
                | ".docx" -> DocxWriter.writeDocx parsed decoded outputPath
                | ".rtf" -> RtfWriter.writeRtf parsed decoded outputPath
                | ".pdf" -> PdfWriter.writePdf parsed decoded outputPath
                | ".html" -> HtmlWriter.writeHtml parsed decoded outputPath
                | ".md" -> MarkdownWriter.writeMarkdown parsed decoded outputPath
                | _ -> failwithf "Unsupported output format '%s'. Use .docx, .rtf, .pdf, .html, .md or .doc." outputExtension

                printfn "Converted '%s' to '%s'" inputPath outputPath
                0
        with ex ->
            eprintfn "ERROR: %s" ex.Message
            2
