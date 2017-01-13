namespace NeXL.IMF
open NeXL.ManagedXll
open NeXL.XlInterop
open System
open Newtonsoft.Json
open Newtonsoft.Json.Linq


[<XlInvisible>]
type NameJson =
    {
     ``@xml:lang`` : string
     ``#text`` : string
    }

[<XlInvisible>]
type KeyFamilyRefJson =
    {
     KeyFamilyID : string
     KeyFamilyAgencyID : string
    }

[<XlInvisible>]
type DataflowItemJson =
    {
     ``@id`` : string
     ``@version`` : string
     ``@agencyId`` : string
     Name : NameJson
     KeyFamilyRef : KeyFamilyRefJson
    }

[<XlInvisible>]
type DataflowJson =
    {
     Dataflow : DataflowItemJson[]
    }

[<XlInvisible>]
type DataflowsJson =
    {
     Dataflows : DataflowJson
    }

[<XlInvisible>]
type StructureJson =
    {
     Structure : DataflowsJson
    }