Attribute VB_Name = "mDeleteBlanks"
Option Explicit

Public Sub delete_blanks()
On Error GoTo ErrorHandler
    Dim lastrow As Long
    Dim rngFirstCell As Range
    Dim i As Integer
    Dim arrData As Variant
    Dim numerator As Long
    Dim numRow As Integer, numColumn As Integer
    
    Set rngFirstCell = Application.InputBox("Please select first cell of column " & _
      "with empty range", Type:=8)
    lastrow = rngFirstCell.CurrentRegion.Rows.Count
    numRow = rngFirstCell.Row
    numColumn = rngFirstCell.Column
    arrData = Range(Cells(numRow, numColumn), _
              Cells(lastrow + numRow - 1, numColumn)).Value
    numerator = 0
    i = lastrow
    
    Do Until i = 1
        If Len(arrData(i, 1)) = 0 Then
            rngFirstCell.Offset(i - 1, 0).EntireRow.Delete
            numerator = numerator + 1
        End If
        i = i - 1
    Loop

MsgBox prompt:="well! delete" & " " & numerator & " " & "empty range"

Exit Sub
ErrorHandler:

MsgBox prompt:=Err.Description & " " & Err.Number

End Sub


