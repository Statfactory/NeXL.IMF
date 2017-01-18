namespace NeXL.IMF
open NeXL.ManagedXll
open NeXL.XlInterop
open System
open System.IO
open System.Runtime.InteropServices
open System.Data
open FSharp.Data
open Newtonsoft.Json
open Newtonsoft.Json.Linq
open Deedle

[<XlQualifiedName(true)>]
module IMF =

    let private frameToDataTable (rowIndexName : string) (frame : Frame<'T, string>) : DataTable =
        let dbTable = new DataTable()
        let dateCol = new DataColumn(rowIndexName, typeof<'T>)
        let colNames = frame.Columns.Keys |> Seq.toArray
        let cols = colNames |> Array.map (fun colName -> new DataColumn(colName, typeof<decimal>)) 
        dbTable.Columns.Add(dateCol)
        dbTable.Columns.AddRange(cols)
        frame.RowKeys |> Seq.iter (fun date ->
                                       let frameRow = frame.Rows.[date]
                                       let dbRow = dbTable.NewRow()
                                       dbRow.[rowIndexName] <- date
                                       colNames |> Array.iter (fun colName ->
                                                                   let v = frameRow.TryGet(colName)
                                                                   if v.HasValue then
                                                                       dbRow.[colName] <- v.Value
                                                                   else
                                                                       dbRow.[colName] <- DBNull.Value
                                                              )
                                       dbTable.Rows.Add(dbRow)
                                  )
        dbTable

    let private toDatasetInfo (v : DataflowItemJson) : DatasetInfo =
        {
            DatasetId = v.KeyFamilyRef.KeyFamilyID
            Description = v.Name.``#text``
        }

    let private toDimInfo (v : DimensionItemJson) : DimensionInfo =
        {
            Name = v.``@conceptRef``
            CodeList = v.``@codelist``
        }

    let private toCode (v : CodeItemJson) : Code =
        {
            Value = v.``@value``
            Description = v.Description.``#text``
        }

    let private nmult n =
        match n with  
            | 1 -> 10m
            | 2 -> 100m
            | 3 -> 1000m
            | 4 -> 10000m
            | 5 -> 100000m
            | 6 -> 1000000m
            | 7 -> 10000000m
            | 8 -> 100000000m
            | 9 -> 1000000000m
            | 10 -> 10000000000m
            | 11 -> 100000000000m
            | 12 -> 1000000000000m
            | 13 -> 10000000000000m
            | 14 -> 100000000000000m
            | 15 -> 1000000000000000m
            | 16 -> 10000000000000000m
            | _ -> raise (new NotImplementedException("Multiplier > 10 ^ 16"))

    let private dataflowUrl = "http://dataservices.imf.org/REST/SDMX_JSON.svc/Dataflow/"

    let private getDataStructureUrl datasetId = sprintf "http://dataservices.imf.org/REST/SDMX_JSON.svc/DataStructure/%s" datasetId

    let private getCodeListUrl codeList = sprintf "http://dataservices.imf.org/REST/SDMX_JSON.svc/CodeList/%s" codeList

    let private getCompactDataUrl datasetId freq dims = sprintf "http://dataservices.imf.org/REST/SDMX_JSON.svc/CompactData/%s/%s.%s" datasetId freq dims


    [<XlFunctionHelp("This function will asynchronously return a list of datasets")>]
    let getDatasetList(
                        [<XlArgHelp("True if headers should be returned (optional, default is TRUE)")>] headers : bool option,
                        [<XlArgHelp("True if table should be returned as transposed (optional, default is FALSE)")>] transposed : bool option,
                        [<XlArgHelp("Timestamp to force refresh on recalc. You can use Excel Today() but not Now() (optional)")>] timestamp : DateTime option
                      ) =
        async  
            {
            let transposed = defaultArg transposed false

            let headers = defaultArg headers true

            let! response = Http.AsyncRequest(dataflowUrl, [], silentHttpErrors = true)

            match response.Body with  
                | Text(json) -> 
                    if response.StatusCode >= 400 then
                        let doc = HtmlDocument.Parse(json)
                        let err = doc.Descendants ["string"] 
                                    |> Seq.map (fun v -> v.InnerText())
                                    |>  String.concat "."
                        raise (new ArgumentException(err))
                        return XlTable.Empty
                    else
                        let dataflow = JsonConvert.DeserializeObject<DataflowStructureResponse>(json)
                        let datasetList = dataflow.Structure.Dataflows.Dataflow |> Array.map toDatasetInfo
                        return XlTable.Create(datasetList, String.Empty, String.Empty, false, transposed, headers)
                | Binary(_) -> 
                    raise (new ArgumentException("Binary response received, json expected"))
                    return XlTable.Empty
             }   

    [<XlFunctionHelp("This function will asynchronously return a list of dataset dimensions")>]
    let getDimensions(
                        [<XlArgHelp("Dataset Id")>] datasetId : string,
                        [<XlArgHelp("True if headers should be returned (optional, default is TRUE)")>] headers : bool option,
                        [<XlArgHelp("True if table should be returned as transposed (optional, default is FALSE)")>] transposed : bool option,
                        [<XlArgHelp("Timestamp to force refresh on recalc. You can use Excel Today() but not Now() (optional)")>] timestamp : DateTime option
                      ) =
        async  
            {
            let transposed = defaultArg transposed false

            let headers = defaultArg headers true

            let! response = Http.AsyncRequest(getDataStructureUrl datasetId, [], silentHttpErrors = true)

            match response.Body with  
                | Text(json) -> 
                    if response.StatusCode >= 400 then
                        let doc = HtmlDocument.Parse(json)
                        let err = doc.Descendants ["string"] 
                                    |> Seq.map (fun v -> v.InnerText())
                                    |>  String.concat "."
                        raise (new ArgumentException(err))
                        return XlTable.Empty
                    else
                        let datasetStructure = JsonConvert.DeserializeObject<DatasetStructureResponse>(json)
                        let dimInfo = datasetStructure.Structure.KeyFamilies.KeyFamily.Components.Dimension |> Array.map toDimInfo
                        return XlTable.Create(dimInfo, String.Empty, String.Empty, false, transposed, headers)
                | Binary(_) -> 
                    raise (new ArgumentException("Binary response received, json expected"))
                    return XlTable.Empty
             }   

    [<XlFunctionHelp("This function will asynchronously return a codelist")>]
    let getCodeList(
                        [<XlArgHelp("CodeList")>] codeList : string,
                        [<XlArgHelp("True if headers should be returned (optional, default is TRUE)")>] headers : bool option,
                        [<XlArgHelp("True if table should be returned as transposed (optional, default is FALSE)")>] transposed : bool option,
                        [<XlArgHelp("Timestamp to force refresh on recalc. You can use Excel Today() but not Now() (optional)")>] timestamp : DateTime option
                      ) =
        async  
            {
            let transposed = defaultArg transposed false

            let headers = defaultArg headers true

            let! response = Http.AsyncRequest(getCodeListUrl codeList, [], silentHttpErrors = true)

            match response.Body with  
                | Text(json) -> 
                    if response.StatusCode >= 400 then
                        let doc = HtmlDocument.Parse(json)
                        let err = doc.Descendants ["string"] 
                                    |> Seq.map (fun v -> v.InnerText())
                                    |>  String.concat "."
                        raise (new ArgumentException(err))
                        return XlTable.Empty
                    else
                        let codelist = JsonConvert.DeserializeObject<CodeListResponse>(json)
                        let codes = codelist.Structure.CodeLists.CodeList.Code |> Array.map toCode
                        return XlTable.Create(codes, String.Empty, String.Empty, false, transposed, headers)
                | Binary(_) -> 
                    raise (new ArgumentException("Binary response received, json expected"))
                    return XlTable.Empty
             }   

    [<XlFunctionHelp("This function will asynchronously return data series for given frequency and dimensions")>]
    let getSeriesData(
                        [<XlArgHelp("Dataset Id")>] datasetId : string,
                        [<XlArgHelp("Frequency")>] frequency : string,
                        [<XlArgHelp("Dimensions")>] dimensions : string option[],
                        [<XlArgHelp("Start Period")>] startPeriod : string,
                        [<XlArgHelp("End Period")>] endPeriod : string,
                        [<XlArgHelp("Apply unit multiplier (optional, default is FALSE)")>] applyUnitMult : bool option,
                        [<XlArgHelp("True if headers should be returned (optional, default is TRUE)")>] headers : bool option,
                        [<XlArgHelp("True if table should be returned as transposed (optional, default is FALSE)")>] transposed : bool option,
                        [<XlArgHelp("Timestamp to force refresh on recalc. You can use Excel Today() but not Now() (optional)")>] timestamp : DateTime option
                      ) =
        async  
            {
            let transposed = defaultArg transposed false

            let headers = defaultArg headers true

            let applyUnitMult = defaultArg applyUnitMult false

            let dims = dimensions |> Array.map (fun d -> match d with | Some(v) -> v | None -> String.Empty) |> String.concat "."

            let startPrm = ("startPeriod", startPeriod)

            let endPrm = ("endPeriod", endPeriod)

            let! response = Http.AsyncRequest(getCompactDataUrl datasetId frequency dims, [startPrm; endPrm], silentHttpErrors = true)

            match response.Body with  
                | Text(json) -> 
                    if response.StatusCode >= 400 then
                        let doc = HtmlDocument.Parse(json)
                        let err = doc.Descendants ["string"] 
                                    |> Seq.map (fun v -> v.InnerText())
                                    |>  String.concat "."
                        raise (new ArgumentException(err))
                        return XlTable.Empty
                    else
                        let jval = JsonValue.Parse json
                        match jval.GetProperty("CompactData").GetProperty("DataSet").TryGetProperty("Series") with
                            | Some(series) ->
                                let series = 
                                    match series with
                                        | JsonValue.Array(v) -> v
                                        | _ -> [|series|]

                                if series.Length > 0 then
                                    let dims = 
                                        match series.[0] with 
                                            | JsonValue.Record(r) ->
                                                 r |> Array.map fst |> Array.filter (fun x -> x <> "@FREQ" && x <> "@UNIT_MULT" && x <> "@TIME_FORMAT" && x <> "Obs")
                                            | _ -> raise (new InvalidOperationException())
                                    if startPeriod = endPeriod && dims.Length = 2 then
                                        let data = series |> Seq.map (fun v -> 
                                                                            let dim0 = v.GetProperty(dims.[0]).AsString()
                                                                            let dim1 = v.GetProperty(dims.[1]).AsString()
                                                                            let n = v.GetProperty("@UNIT_MULT").AsInteger()
                                                                            match v.TryGetProperty("Obs") with
                                                                                | Some(jv) ->
                                                                                    let obs = jv.GetProperty("@OBS_VALUE").AsDecimal()
                                                                                    Some(dim0, dim1, if applyUnitMult then obs * nmult n else obs)
                                                                                | None -> None
                                                                    )
                                                        |> Seq.choose id
                                                        |> Frame.ofValues 
                                                        |> frameToDataTable (dims.[0].Substring(1))
                                        return XlTable(data, String.Empty, String.Empty, false, transposed, headers)
                                    else
                                        let data = series |> Seq.collect (fun v -> 
                                                                            let dim = dims |> Array.map (fun x -> v.GetProperty(x).AsString())
                                                                                           |> String.concat "|"
                                                                            let n = v.GetProperty("@UNIT_MULT").AsInteger()
                                                                            match v.TryGetProperty("Obs") with
                                                                                | Some(jv) ->
                                                                                    let obsArr = jv.AsArray()
                                                                                    obsArr |> Seq.filter (fun obs -> obs.TryGetProperty("@OBS_VALUE").IsSome && obs.TryGetProperty("@TIME_PERIOD").IsSome)
                                                                                           |> Seq.map (fun obs -> 
                                                                                                        let obsVal = obs.GetProperty("@OBS_VALUE").AsDecimal()
                                                                                                        let period = obs.GetProperty("@TIME_PERIOD").AsString()
                                                                                                        period, dim, if applyUnitMult then obsVal * nmult n else obsVal
                                                                                                    ) 
                                                                                | None -> Seq.empty
                                                                        )
                                                        |> Frame.ofValues 
                                                        |> frameToDataTable "Period"
                                        return XlTable(data, String.Empty, String.Empty, false, transposed, headers)

                                else
                                    raise (new ArgumentException("No series data returned."))
                                    return XlTable.Empty

                            | None ->
                                raise (new ArgumentException("No series data returned."))
                                return XlTable.Empty
                | Binary(_) -> 
                    raise (new ArgumentException("Binary response received, json expected"))
                    return XlTable.Empty
             }   

    let getErrors newOnTop =
        UdfErrorHandler.OnError |> Event.scan (fun s e -> e :: s) []
                                |> Event.map (fun errs ->
                                                  let errs = if newOnTop then errs |> List.toArray else errs |> List.rev |> List.toArray
                                                  XlTable.Create(errs, "", "", false, false, true)
                                             )
