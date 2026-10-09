// Query: ERP_Sales
// Loads the ERP sales extract, types columns, cleans values and enriches
// with product/region master data. Set the SourceFolder parameter first.
let
    SourceFolder = Excel.CurrentWorkbook(){[Name="prmSourceFolder"]}[Content]{0}[Column1],
    Source       = Csv.Document(File.Contents(SourceFolder & "erp_sales.csv"),
                     [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
    Promoted     = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    Typed        = Table.TransformColumnTypes(Promoted, {
                     {"Invoice", type text}, {"Date", type date}, {"RegionID", type text},
                     {"ProductID", type text}, {"Qty", Int64.Type}, {"Discount", type number}}),
    Trimmed      = Table.TransformColumns(Typed, {{"RegionID", Text.Trim}, {"ProductID", Text.Upper}}),
    NoBlanks     = Table.SelectRows(Trimmed, each [Invoice] <> null and [Qty] > 0),
    Deduped      = Table.Distinct(NoBlanks, {"Invoice"}),
    AddMonth     = Table.AddColumn(Deduped, "Month", each Date.ToText([Date], "MMM"), type text),
    MergeProduct = Table.NestedJoin(AddMonth, {"ProductID"}, Products, {"ProductID"}, "P", JoinKind.LeftOuter),
    ExpandProd   = Table.ExpandTableColumn(MergeProduct, "P",
                     {"ProductName", "Category", "UnitPrice", "GrossMargin"}),
    MergeRegion  = Table.NestedJoin(ExpandProd, {"RegionID"}, Regions, {"RegionID"}, "R", JoinKind.LeftOuter),
    ExpandRegion = Table.ExpandTableColumn(MergeRegion, "R", {"RegionName"}),
    NetRevenue   = Table.AddColumn(ExpandRegion, "NetRevenue",
                     each [Qty] * [UnitPrice] * (1 - [Discount]), Currency.Type),
    GrossProfit  = Table.AddColumn(NetRevenue, "GrossProfit",
                     each [NetRevenue] * [GrossMargin], Currency.Type),
    Final        = Table.RemoveColumns(GrossProfit, {"GrossMargin"})
in
    Final
