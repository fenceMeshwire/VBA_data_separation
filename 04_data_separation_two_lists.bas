Option Explicit

' ________________________________________________________________________________________________
Sub distinguish_differences()

Dim lngRow As Long, lngRowMax As Long
Dim lngRowComp As Long, lngRowCompMax As Long
Dim strProperty As String, strPropertyComp As String
Dim wkSheetMain As Worksheet, wkSheetComp As Worksheet

Set wkSheetMain = Sheet1
Set wkSheetComp = Sheet2

Application.ScreenUpdating = False

lngRowMax = wkSheetMain.Cells(wkSheetMain.Rows.Count, 1).End(xlUp).Row
lngRowCompMax = wkSheetComp.Cells(wkSheetComp.Rows.Count, 1).End(xlUp).Row

For lngRow = 2 To lngRowMax
  strProperty = wkSheetMain.Cells(lngRow, 1).Value
  For lngRowComp = 2 To lngRowCompMax
    strPropertyComp = wkSheetComp.Cells(lngRowComp, 1).Value
    Call color_cells(wkSheetMain, lngRow, 4)
    If strProperty = strPropertyComp Then
      Call color_cells(wkSheetMain, lngRow, 3)
      GoTo nextline
    End If
  Next lngRowComp
nextline:
Next lngRow

Application.DisplayAlerts = True

End Sub

' ________________________________________________________________________________________________
Function color_cells(ByRef wkSheetMain As Worksheet, ByVal lngRow As Long, ByVal intColor As Integer)

Dim intCol As Integer, intColMax As Integer

With wkSheetMain
  intColMax = .UsedRange.Columns.Count
  For intCol = 1 To intColMax
    .Cells(lngRow, intCol).Interior.ColorIndex = intColor
  Next intCol
End With

End Function
