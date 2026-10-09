Attribute VB_Name = "modDashboardAutomation"
'==============================================================================
' Module : modDashboardAutomation
' Purpose: One-click monthly refresh, native PivotTable build, reconciliation
'          checks and PDF export for the Financial Dashboard model.
' Usage  : Import via VBA Editor (Alt+F11) > File > Import File...
'          Save workbook as .xlsm. Run "RefreshAll_Model" from Macros (Alt+F8).
'==============================================================================
Option Explicit

Private Const SHEET_DASH As String = "Dashboard"
Private Const SHEET_SALES As String = "Data_Sales"
Private Const SHEET_RECON As String = "Reconciliation"
Private Const SHEET_PIVOT As String = "Pivot_Native"

'--- Master routine: refresh Power Query, pivots, checks, and timestamp -------
Public Sub RefreshAll_Model()
    Dim t As Double: t = Timer
    On Error GoTo ErrHandler
    SpeedUp True

    ThisWorkbook.RefreshAll                      ' Power Query connections
    Application.CalculateUntilAsyncQueriesDone
    BuildNativePivot
    RefreshAllPivots
    Application.CalculateFull

    With ThisWorkbook.Worksheets(SHEET_DASH)
        .Range("J3").Value = "Last refreshed: " & Format(Now, "dd-mmm-yyyy hh:nn")
    End With

    SpeedUp False
    Dim exceptions As Long
    exceptions = ReconciliationExceptions()
    MsgBox "Model refreshed in " & Format(Timer - t, "0.0") & "s." & vbCrLf & _
           "Reconciliation exceptions: " & exceptions, _
           IIf(exceptions = 0, vbInformation, vbExclamation), "Refresh complete"
    Exit Sub
ErrHandler:
    SpeedUp False
    MsgBox "Refresh failed: " & Err.Description, vbCritical
End Sub

'--- Build (or rebuild) a native PivotTable from tblSales ---------------------
Public Sub BuildNativePivot()
    Dim ws As Worksheet, pc As PivotCache, pt As PivotTable

    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets(SHEET_PIVOT).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = SHEET_PIVOT

    Set pc = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:="tblSales")
    Set pt = pc.CreatePivotTable(TableDestination:=ws.Range("A3"), TableName:="ptRevenue")

    With pt
        .PivotFields("Region").Orientation = xlRowField
        .PivotFields("Category").Orientation = xlRowField
        .PivotFields("Month").Orientation = xlColumnField
        With .AddDataField(.PivotFields("Net Revenue"), "Sum of Net Revenue", xlSum)
            .NumberFormat = "$#,##0"
        End With
        .PivotFields("Product").Orientation = xlPageField
        .RowAxisLayout xlTabularRow
        .TableStyle2 = "PivotStyleMedium2"
    End With
    ws.Range("A1").Value = "Native PivotTable - Revenue by Region / Category / Month"
    ws.Range("A1").Font.Bold = True
End Sub

Public Sub RefreshAllPivots()
    Dim ws As Worksheet, pt As PivotTable
    For Each ws In ThisWorkbook.Worksheets
        For Each pt In ws.PivotTables
            pt.PivotCache.Refresh
        Next pt
    Next ws
End Sub

'--- Count reconciliation rows flagged "Investigate" -------------------------
Public Function ReconciliationExceptions() As Long
    ReconciliationExceptions = Application.WorksheetFunction.CountIf( _
        ThisWorkbook.Worksheets(SHEET_RECON).Range("F5:F16"), "Investigate")
End Function

'--- Highlight exception rows and list them in the Immediate window ----------
Public Sub ReviewReconciliation()
    Dim r As Long, ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_RECON)
    For r = 5 To 16
        If ws.Cells(r, 6).Value = "Investigate" Then
            Debug.Print ws.Cells(r, 1).Value & ": diff " & Format(ws.Cells(r, 4).Value, "$#,##0")
        End If
    Next r
    ws.Activate
End Sub

'--- Set dashboard region filter programmatically ----------------------------
Public Sub SetRegion(ByVal regionName As String)
    ThisWorkbook.Worksheets(SHEET_DASH).Range("C4").Value = regionName
End Sub

'--- Export the dashboard + variance report to PDF ---------------------------
Public Sub ExportDashboardPDF()
    Dim path As String
    path = ThisWorkbook.Path & Application.PathSeparator & _
           "Executive_Dashboard_" & Format(Date, "yyyymmdd") & ".pdf"
    ThisWorkbook.Worksheets(Array(SHEET_DASH, "Budget_vs_Actual", SHEET_RECON)).Select
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=path, _
        Quality:=xlQualityStandard, IgnorePrintAreas:=False, OpenAfterPublish:=False
    ThisWorkbook.Worksheets(SHEET_DASH).Select
    MsgBox "Exported: " & path, vbInformation
End Sub

Private Sub SpeedUp(ByVal turnOn As Boolean)
    With Application
        .ScreenUpdating = Not turnOn
        .EnableEvents = Not turnOn
        .Calculation = IIf(turnOn, xlCalculationManual, xlCalculationAutomatic)
    End With
End Sub
