Option Explicit

' Create two tables previously: sample, lookup

' Columns of the sample table have the following structure: 
' A1234/B2345/E1711,A1234/D3456,B2345/C8976,E1711/A1234,E1711/C8976,C8976

' The lookup table holds the decipherment for the codes: column1 = code, column2 = decipherment
' {'A1234': 'alpha', 'B2345':'beta', 'C8976':'gamma', 'D3456':'delta', 'E1711':'epsilon'}

' The program creates the decipherment for the column in the sample table:
' alpha/beta/epsilon,alpha/delta,beta/gamma,epsilon/alpha,epsilon/gamma,gamma

'_____________________________________________________________________________________________________________
Sub decipher_legend()

Dim lngCol&, lngColMax&
Dim strCombination$, strChecked$
Dim varDat

lngColMax& = sample.UsedRange.Columns.Count

For lngCol& = 1 To lngColMax&

  strCombination$ = sample.Cells(1, lngCol&).Value
  varDat = Split(strCombination, "/")
  strChecked$ = check_array$(varDat)
  sample.Cells(2, lngCol&).Value = strChecked$
  
Next lngCol&

End Sub

'_____________________________________________________________________________________________________________
Function check_array$(ByRef varDat)

Dim intCounter%
Dim lngRow&, lngRowMax&
Dim strCompare$, strLookup$, strResult$, strStorage$

lngRowMax& = lookup.UsedRange.Rows.Count

For intCounter% = LBound(varDat) To UBound(varDat)
  strLookup$ = varDat(intCounter)
  
  For lngRow& = 2 To lngRowMax&
    strCompare$ = lookup.Cells(lngRow, 1).Value
    
    If strCompare$ = strLookup$ And UBound(varDat) < 1 Then
      check_array$ = lookup.Cells(lngRow, 2).Value
      Exit Function
    ElseIf strCompare$ = strLookup And UBound(varDat) > 0 Then
      strResult$ = lookup.Cells(lngRow, 2).Value
      If strStorage$ = "" Then
        strStorage = strResult
      Else
        strStorage$ = strStorage$ & "/" & strResult$
      End If
    End If
    
  Next lngRow&
  
Next intCounter%

check_array$ = strStorage$

End Function
