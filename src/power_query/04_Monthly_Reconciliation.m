// Query: Monthly_Reconciliation
// Aggregates ERP and GL by region & month, joins them and flags breaks.
let
    Tolerance = 0.01,
    ErpAgg = Table.Group(ERP_Sales, {"RegionID", "Month"},
               {{"ERP_Revenue", each List.Sum([NetRevenue]), Currency.Type}}),
    GlAgg  = Table.Group(GL_Revenue, {"RegionID", "Month"},
               {{"GL_Revenue", each List.Sum([Amount]), Currency.Type}}),
    Joined = Table.NestedJoin(ErpAgg, {"RegionID", "Month"}, GlAgg, {"RegionID", "Month"}, "GL", JoinKind.FullOuter),
    Expand = Table.ExpandTableColumn(Joined, "GL", {"GL_Revenue"}),
    Nulls  = Table.ReplaceValue(Expand, null, 0, Replacer.ReplaceValue, {"ERP_Revenue", "GL_Revenue"}),
    Diff   = Table.AddColumn(Nulls, "Difference", each [GL_Revenue] - [ERP_Revenue], Currency.Type),
    Pct    = Table.AddColumn(Diff, "DiffPct",
               each if [ERP_Revenue] = 0 then null else [Difference] / [ERP_Revenue], Percentage.Type),
    Status = Table.AddColumn(Pct, "Status",
               each if [DiffPct] <> null and Number.Abs([DiffPct]) <= Tolerance then "Reconciled" else "Investigate", type text),
    Sorted = Table.Sort(Status, {{"RegionID", Order.Ascending}, {"Month", Order.Ascending}})
in
    Sorted
