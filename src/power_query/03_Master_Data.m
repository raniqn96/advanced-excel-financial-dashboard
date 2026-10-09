// Queries: Products and Regions (connection-only reference tables)
// Create two blank queries and paste each section separately.

// ---- Products ----
let
    Source = Excel.CurrentWorkbook(){[Name="tblProducts"]}[Content],
    Renamed = Table.RenameColumns(Source, {
                {"Product ID", "ProductID"}, {"Product Name", "ProductName"},
                {"Unit Price ($)", "UnitPrice"}, {"Gross Margin %", "GrossMargin"}}),
    Typed = Table.TransformColumnTypes(Renamed, {
                {"ProductID", type text}, {"ProductName", type text}, {"Category", type text},
                {"UnitPrice", Currency.Type}, {"GrossMargin", type number}})
in
    Typed

// ---- Regions ----
// Region master is stored horizontally (for HLOOKUP) - transpose it here.
let
    Source     = Excel.CurrentWorkbook(){[Name="rngRegions"]}[Content],
    Transposed = Table.Transpose(Source),
    Promoted   = Table.PromoteHeaders(Transposed, [PromoteAllScalars=true]),
    Renamed    = Table.RenameColumns(Promoted, {
                   {"Region ID", "RegionID"}, {"Region Name", "RegionName"}, {"Regional Manager", "Manager"}})
in
    Renamed
