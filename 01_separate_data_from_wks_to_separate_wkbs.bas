Option Explicit

' ______________________________________________________________________
Dim wksMainSheet As Worksheet
Dim wksFirst As Worksheet
Dim wksSecond As Worksheet
Dim wksThird As Worksheet
Dim wksFourth As Worksheet

' ______________________________________________________________________
Sub main_process()

Call set_worksheets
Call copy_main_table
Call delete_inappropriate_data
Call save_data_to_separate_worksheets

End Sub

' ______________________________________________________________________
Sub set_worksheets()

Dim intYear
Dim wksAddSheet As Worksheet

intYear = 2022

With ThisWorkbook
  Set wksMainSheet = .Sheets(1)

  Do While .Sheets.Count < 5
    Set wksAddSheet = .Sheets.Add(After:=.Worksheets(.Worksheets.Count))
  Loop

  Set wksFirst = .Sheets(2)
  Set wksSecond = .Sheets(3)
  Set wksThird = .Sheets(4)
  Set wksFourth = .Sheets(5)
  
  wksMainSheet.Name = "Data"
  wksFirst.Name = "Q1_" & intYear
  wksSecond.Name = "Q2_" & intYear
  wksThird.Name = "Q3_" & intYear
  wksFourth.Name = "Q4_" & intYear
  
End With

End Sub

' ______________________________________________________________________
Sub copy_main_table()

Dim lngRow, lngRowMax As Long
Dim lngCol, lngColMax As Long

Dim rngData As Range

wksMainSheet.UsedRange.Copy Destination:=wksFirst.Range("A1")
wksMainSheet.UsedRange.Copy Destination:=wksSecond.Range("A1")
wksMainSheet.UsedRange.Copy Destination:=wksThird.Range("A1")
wksMainSheet.UsedRange.Copy Destination:=wksFourth.Range("A1")

wksFirst.UsedRange.Columns.AutoFit
wksSecond.UsedRange.Columns.AutoFit
wksThird.UsedRange.Columns.AutoFit
wksFourth.UsedRange.Columns.AutoFit


End Sub

' ______________________________________________________________________
Sub delete_inappropriate_data()

Dim lngRow, lngRowMax As Long
Dim lngCol, lngColMax As Long
Dim strName, strSourceData, strWksName As String
Dim wksSheet As Worksheet

strSourceData = "Data"

For Each wksSheet In ThisWorkbook.Sheets
  If wksSheet.Name <> strSourceData Then
    strName = wksSheet.Name ' e.g. strName = 'Q1/2022'
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

' ______________________________________________________________________
Sub save_data_to_separate_worksheets()

Dim lngRow, lngRowMax As Long
Dim lngCol, lngColMax As Long
Dim strName, strWksName As String
Dim wksSheet As Worksheet

strWksName = "Rev_"

Application.ScreenUpdating = False
For Each wksSheet In ThisWorkbook.Sheets
  If wksSheet.Name <> wksMainSheet.Name Then
    strName = wksSheet.Name
    With wksSheet
      .Copy
      ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count).Name = strWksName & strName
      ActiveWorkbook.SaveAs ThisWorkbook.Path & "\" & strWksName & strName & ".xlsx"
      ActiveWorkbook.Close
    End With
  End If
Next wksSheet
Application.ScreenUpdating = True

End Sub
