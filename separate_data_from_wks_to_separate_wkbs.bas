Option Explicit

' ----------------------------------------------------------------------
Sub main_process()

Call copy_main_table
Call delete_inappropriate_data
Call save_data_to_separate_worksheets

End Sub

' ----------------------------------------------------------------------
Sub copy_main_table()

Dim lngRow, lngRowMax As Long
Dim lngCol, lngColMax As Long

Dim rngData As Range

Tabelle1.UsedRange.Copy Destination:=Tabelle2.Range("A1")
Tabelle1.UsedRange.Copy Destination:=Tabelle3.Range("A1")
Tabelle1.UsedRange.Copy Destination:=Tabelle4.Range("A1")

Tabelle2.UsedRange.Columns.AutoFit
Tabelle3.UsedRange.Columns.AutoFit
Tabelle4.UsedRange.Columns.AutoFit

End Sub

' ----------------------------------------------------------------------
Sub delete_inappropriate_data()

Dim lngRow, lngRowMax As Long
Dim lngCol, lngColMax As Long
Dim strName, strSourceData, strWksName As String
Dim wksSheet As Worksheet

strSourceData = "Data"

For Each wksSheet In ThisWorkbook.Sheets
  If wksSheet.Name <> strSourceData Then
    strName = wksSheet.Name
    strName = Replace(strName, "_", "/")
    With wksSheet
      lngRowMax = .Cells(.Rows.Count, 1).End(xlUp).Row
      For lngRow = lngRowMax To 2 Step -1
        If strName <> .Cells(lngRow, 10).Value Then
          .Rows(lngRow).Delete
        End If
      Next lngRow
    End With
  End If
Next wksSheet

End Sub

' ----------------------------------------------------------------------
Sub save_data_to_separate_worksheets()

Dim lngRow, lngRowMax As Long
Dim lngCol, lngColMax As Long
Dim strName, strSourceData, strWksName As String
Dim wksSheet As Worksheet

strSourceData = "Data"
strWksName = "Rev_"

For Each wksSheet In ThisWorkbook.Sheets
  If wksSheet.Name <> strSourceData Then
    strName = wksSheet.Name
    With wksSheet
      .Copy
      ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count).Name = strWksName & strName
      ActiveWorkbook.SaveAs ThisWorkbook.Path & "\" & strWksName & strName & ".xlsx"
      ActiveWorkbook.Close
    End With
  End If
Next wksSheet

End Sub
