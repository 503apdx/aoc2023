Attribute VB_Name = "Module1"
Sub AOCday1baybay()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet

    Dim cell As Range
    Dim i As Integer
    Dim totalSum As Long
    For Each cell In ws.UsedRange
        Dim cellValue As String
        cellValue = cell.Value

       'replace characters with blanks'
        For i = 65 To 90
            cellValue = Replace(cellValue, Chr(i), "")
        
        Next i
        
        For i = 97 To 122
            cellValue = Replace(cellValue, Chr(i), "")
            
        Next i
            
        'set values'
        cell.Value = Left(cellValue, 1) & Right(cellValue, 1)
            
        Next cell
    'sum values'
    totalSum = Application.WorksheetFunction.Sum(ws.UsedRange)
    MsgBox "Total Sum of Used Range: " & totalSum

End Sub
