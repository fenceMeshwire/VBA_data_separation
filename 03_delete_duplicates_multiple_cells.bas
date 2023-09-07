Option Explicit

' _______________________________________________________________________________________
Sub delete_duplicates_multiple_cells()

Dim rngSelect as Range
Dim wksSheet As Worksheet

Set wksSheet as Sheet1

With wksSheet
  ' Select the range as follows:
  Set rngSelect = Range("A1", Range("D1").End(xlDown))
  ' Select the affected columns with an array:
  rngSelect.RemoveDuplicates Columns:=Array(1, 2, 3, 4), Header:=xlYes
End With

End Sub
