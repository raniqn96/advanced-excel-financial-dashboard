Attribute VB_Name = "modDataImport"
'==============================================================================
' Module : modDataImport
' Purpose: Import monthly CSV extracts into the source tables, validate them,
'          and log the import. Complements the Power Query pipeline for users
'          who prefer a macro-driven load.
'==============================================================================
Option Explicit

Public Sub ImportSalesCSV()
    Dim f As Variant
    f = Application.GetOpenFilename("CSV Files (*.csv),*.csv", , "Select ERP sales extract")
    If VarType(f) = vbBoolean Then Exit Sub
    LoadCsvIntoTable CStr(f), "Data_Sales", "tblSales", 6
End Sub

Public Sub ImportGLCSV()
    Dim f As Variant
    f = Application.GetOpenFilename("CSV Files (*.csv),*.csv", , "Select GL revenue extract")
    If VarType(f) = vbBoolean Then Exit Sub
    LoadCsvIntoTable CStr(f), "Data_GL", "tblGL", 5
End Sub

' Appends raw columns (1..rawCols) from a CSV into a ListObject; calculated
' columns inside the table auto-fill because they are structured formulas.
Private Sub LoadCsvIntoTable(ByVal path As String, ByVal sheetName As String, _
                             ByVal tableName As String, ByVal rawCols As Long)
    Dim lo As ListObject, fnum As Integer, lineTxt As String, parts() As String
    Dim newRow As ListRow, i As Long, added As Long, rejected As Long

    Set lo = ThisWorkbook.Worksheets(sheetName).ListObjects(tableName)
    Application.ScreenUpdating = False

    fnum = FreeFile
    Open path For Input As #fnum
    If Not EOF(fnum) Then Line Input #fnum, lineTxt   ' skip header
    Do While Not EOF(fnum)
        Line Input #fnum, lineTxt
        parts = Split(lineTxt, ",")
        If UBound(parts) + 1 >= rawCols And Len(Trim$(parts(0))) > 0 Then
            Set newRow = lo.ListRows.Add
            For i = 0 To rawCols - 1
                If IsNumeric(parts(i)) Then
                    newRow.Range(1, i + 1).Value = CDbl(parts(i))
                ElseIf IsDate(parts(i)) Then
                    newRow.Range(1, i + 1).Value = CDate(parts(i))
                Else
                    newRow.Range(1, i + 1).Value = Trim$(parts(i))
                End If
            Next i
            added = added + 1
        Else
            rejected = rejected + 1
        End If
    Loop
    Close #fnum

    Application.ScreenUpdating = True
    LogImport tableName, path, added, rejected
    MsgBox added & " rows imported, " & rejected & " rejected.", vbInformation, tableName
End Sub

Private Sub LogImport(ByVal tbl As String, ByVal src As String, ByVal ok As Long, ByVal bad As Long)
    Dim ws As Worksheet, r As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Import_Log")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = "Import_Log"
        ws.Range("A1:E1").Value = Array("Timestamp", "Table", "Source File", "Rows Added", "Rows Rejected")
        ws.Range("A1:E1").Font.Bold = True
    End If
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(r, 1).Value = Now
    ws.Cells(r, 2).Value = tbl
    ws.Cells(r, 3).Value = src
    ws.Cells(r, 4).Value = ok
    ws.Cells(r, 5).Value = bad
End Sub

' Removes duplicate invoices from tblSales based on the Invoice column.
Public Sub RemoveDuplicateInvoices()
    With ThisWorkbook.Worksheets("Data_Sales").ListObjects("tblSales")
        .Range.RemoveDuplicates Columns:=1, Header:=xlYes
    End With
End Sub
