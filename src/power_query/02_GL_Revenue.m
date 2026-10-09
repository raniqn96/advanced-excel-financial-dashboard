// Query: GL_Revenue
// Loads General Ledger revenue postings and filters to revenue accounts.
let
    SourceFolder = Excel.CurrentWorkbook(){[Name="prmSourceFolder"]}[Content]{0}[Column1],
    Source       = Csv.Document(File.Contents(SourceFolder & "gl_revenue.csv"),
                     [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
    Promoted     = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    Typed        = Table.TransformColumnTypes(Promoted, {
                     {"JournalID", type text}, {"PostingDate", type date},
                     {"RegionID", type text}, {"Account", type text}, {"Amount", Currency.Type}}),
    RevenueOnly  = Table.SelectRows(Typed, each Text.StartsWith([Account], "4")),
    AddMonth     = Table.AddColumn(RevenueOnly, "Month", each Date.ToText([PostingDate], "MMM"), type text),
    AddMonthNo   = Table.AddColumn(AddMonth, "MonthNo", each Date.Month([PostingDate]), Int64.Type)
in
    AddMonthNo
