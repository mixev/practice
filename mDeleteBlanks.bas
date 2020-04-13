Attribute VB_Name = "mDeleteBlanks"
Option Explicit

Public Sub new_delete_blanks()
On Error GoTo ErrorHandler
    Dim lastrow As Long
    Dim rngFirstCell As Range
    Dim i As Integer, s As Long
    Dim cellValue As Variant
    Dim numerator As Long
    
    Set rngFirstCell = Application.InputBox("Please select first cell of column " & _
      "with empty range", Type:=8)
    lastrow = rngFirstCell.CurrentRegion.Rows.Count
    cellValue = rngFirstCell.Value
    numerator = 0
    i = 0
    
    Do Until i > lastrow
      cellValue = rngFirstCell.Offset(i, 0).Value
        If Len(cellValue) = 0 Then
          s = i
          Do Until Len(cellValue) > 0 Or s > lastrow
            rngFirstCell.Offset(i, 0).EntireRow.Delete
            cellValue = rngFirstCell.Offset(i, 0).Value
            s = s + 1
            numerator = numerator + 1
          Loop
          s = 0
        End If
        lastrow = rngFirstCell.CurrentRegion.Rows.Count
        i = i + 1
    Loop

MsgBox prompt:="well! delete" & " " & numerator & " " & "empty range"

Exit Sub
ErrorHandler:

MsgBox prompt:=Err.Description & " " & Err.Number

End Sub


