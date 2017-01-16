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
type DatasetInfo =
    {
     DatasetId : string
     Description : string
    }

[<XlInvisible>]
type DataflowJson =
    {
     Dataflow : DataflowItemJson[]
    }

[<XlInvisible>]
type DataflowStructureJson =
    {
     Dataflows : DataflowJson
    }

[<XlInvisible>]
type DataflowStructureResponse =
    {
     Structure : DataflowStructureJson
    }

[<XlInvisible>]
type DimensionItemJson =
    {
     ``@conceptRef`` : string
     ``@codelist`` : string
    }

[<XlInvisible>]
type DimensionInfo =
    {
     Name : string
     CodeList : string
    }

[<XlInvisible>]
type ComponentsJson =
    {
        Dimension : DimensionItemJson[]
    }

[<XlInvisible>]
type KeyFamilyJson =
    {
     ``@id`` : string
     ``@version`` : string
     ``@agencyId`` : string
     Name : NameJson
     Components : ComponentsJson
    }

[<XlInvisible>]
type KeyFamiliesJson =
    {
     KeyFamily : KeyFamilyJson
    }

[<XlInvisible>]
type DatasetStructureJson =
    {
     KeyFamilies : KeyFamiliesJson
    }

[<XlInvisible>]
type DatasetStructureResponse =
    {
     Structure : DatasetStructureJson
    }

[<XlInvisible>]
type CodeItemJson =
    {
     ``@value`` : string
     Description : NameJson
    }

[<XlInvisible>]
type Code =
    {
     Value : string
     Description : string
    }

[<XlInvisible>]
type CodeListJson =
    {
     ``@id`` : string
     ``@version`` : string
     ``@agencyId`` : string
     Name : NameJson
     Code : CodeItemJson[]
    }

[<XlInvisible>]
type CodeListsJson =
    {
     CodeList : CodeListJson
    }

[<XlInvisible>]
type CodeListStructureJson =
    {
     CodeLists : CodeListsJson
    }

[<XlInvisible>]
type CodeListResponse =
    {
     Structure : CodeListStructureJson
    }












