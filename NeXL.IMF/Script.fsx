#r "../packages/DocumentFormat.OpenXml/lib/DocumentFormat.OpenXml.dll"
#r "../packages/NeXL/lib/net45/NeXL.XlInterop.dll"
#r "../packages/NeXL/lib/net45/NeXL.ManagedXll.dll"
#r "../packages/FSharp.Data/lib/net40/FSharp.Data.dll"
#r "../packages/Newtonsoft.Json/lib/net45/Newtonsoft.Json.dll"
#r "../packages/Deedle/lib/net40/Deedle.dll"

#load "Types.fs"

open NeXL.ManagedXll
open NeXL.XlInterop
open System
open System.IO
open System.Runtime.InteropServices
open System.Data
open FSharp.Data
open FSharp.Data.JsonExtensions
open FSharp.Data.HtmlExtensions
open Newtonsoft.Json
open Newtonsoft.Json.Linq
open Deedle
open NeXL.IMF

let response = Http.Request("http://dataservices.imf.org/REST/SDMX_JSON.svc/CompactData/BOP/A.DE+FR.BXSOFI_BP6_EUR?startPeriod=2000&endPeriod=2001", [], silentHttpErrors = true)

let res =
    match response.Body with  
        | Text(json) -> 
            if response.StatusCode >= 400 then
                let doc = HtmlDocument.Parse(json)
                let body = doc.Body()
                let err = body.Descendants ["p"] 
                            |> Seq.map (fun v -> v.InnerText())
                            |>  String.concat "."
                printfn "%A" err
                None
            else
                //let dataflow = JsonConvert.DeserializeObject<CompactDataResponse>(json) 
                let data = JsonValue.Parse json
                Some data
                //return XlTable.Create(countries, String.Empty, String.Empty, false, transposed, headers)
        | Binary(_) -> None

let v = res.Value
let vv = v.GetProperty("CompactData").GetProperty("DataSet").GetProperty("Series")
let props = vv.AsArray().[0].Properties
