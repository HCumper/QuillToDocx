namespace QuillToDocx

open System
open System.IO

module Program =

    let private resolveInputPath (inputPath: string) =
        let ext = Path.GetExtension(inputPath)
        if String.IsNullOrWhiteSpace(ext) then inputPath + ".doc" else inputPath

    let private usage () =
        eprintfn "Usage: QuillToDocx <input.doc> [output.docx|output.pdf|output.html]"
        eprintfn "       QuillToDocx <basename>"

    let private resolvePaths (argv: string array) =
        match argv with
        | [| singleArg |] ->
            let ext = Path.GetExtension(singleArg)

            if String.Equals(ext, ".docx", StringComparison.OrdinalIgnoreCase) then
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
            match resolvePaths argv with
            | Error message ->
                eprintfn "%s" message
                usage ()
                1
            | Ok(inputPath, outputPath) ->
                if not (File.Exists inputPath) then
                    eprintfn "Input file not found: %s" inputPath
                    1
                else
                    let parsed = QuillParser.parseFile inputPath
                    let decoded = QuillDecode.decodeDocument parsed
                    let outputExtension =
                        match Path.GetExtension(outputPath) with
                        | null -> ""
                        | extension -> extension

                    match outputExtension.ToLowerInvariant() with
                    | ".docx" -> DocxWriter.writeDocx parsed decoded outputPath
                    | ".pdf" -> PdfWriter.writePdf parsed decoded outputPath
                    | ".html" -> HtmlWriter.writeHtml parsed decoded outputPath
                    | _ -> failwithf "Unsupported output format '%s'. Use .docx, .pdf, or .html." outputExtension

                    printfn "Converted '%s' to '%s'" inputPath outputPath
                    0
        with ex ->
            eprintfn "ERROR: %s" ex.Message
            2
