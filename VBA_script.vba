Option Explicit
Sub cleandata()
Dim Nr As Integer, i As Integer
Dim cell As Range
Dim CommaPosition As Integer
Dim ws As Worksheet
Set ws = ThisWorkbook.Sheets("SalesData")

'Activeate the worksheet containing the data
ws.Activate

Application.ScreenUpdating = False

If Range("A1").Value = "Customer FirstName" And Range("B1").Value = "Customer LastName" Then
    MsgBox "Data is already Cleaned"
    Exit Sub
End If


'Insert  two new columns and name them
ws.Range("B1:C1").EntireColumn.Insert
ws.Range("B1").Value = "Customer FirstName"
ws.Range("C1").Value = "Customer LastName"

'Extract the First and last Name from Column A in to B and C respectively
Nr = WorksheetFunction.CountA(Columns("A:A"))
For i = 2 To Nr
 CommaPosition = InStr(ws.Cells(i, 1).Value, ",")
 ws.Cells(i, 2).Value = Left(ws.Cells(i, 1).Value, CommaPosition - 1)
 ws.Cells(i, 3).Value = WorksheetFunction.Trim(Right(ws.Cells(i, 1).Value, Len(ws.Cells(i, 1).Value) - CommaPosition))
Next i

'Delete Column A
Range("A1").EntireColumn.Delete
Range("A1").Select

'Extract Email from Referrer Email column
For Each cell In Range("D2:D" & Nr).Cells
    cell.Value = EmailExtract(cell.Value)
Next cell

'Remove Blank values on Id column
Range("A1").AutoFilter Field:=5, Criteria1:=""
Range("A1").CurrentRegion.Offset(1, 0).EntireRow.Delete
Range("A1").AutoFilter

'Extract the Id numbers
Nr = WorksheetFunction.CountA(Columns("A:A"))
For Each cell In Range("E2:E" & Nr).Cells
    cell.Value = Right(cell.Value, Len(cell.Value) - InStr(cell.Value, "D"))
Next cell

'Change all dateFormat to dd/mm/yyyy
For Each cell In Range("G2:G" & Nr).Cells
         Dim DateValue As Date
         If IsDate(cell.Value) Then
            ' Convert the value to a date
            DateValue = CDate(cell.Value)
            
            ' Apply the standard date format "dd/mm/yyyy"
            cell.Value = Format(DateValue, "dd/mm/yyyy")
            cell.NumberFormat = "dd/mm/yyyy"
        End If
Next cell

'Remove Outliers in the Sales Amount column
' Calculate Q1 and Q3
    Dim Q1 As Double, Q3 As Double, IQR As Double, OutlierCount As Integer
    Dim LowerBound As Double, UpperBound As Double
    
    Q1 = WorksheetFunction.Percentile(ws.Range("F2:F" & Nr), 0.25)
    Q3 = WorksheetFunction.Percentile(ws.Range("F2:F" & Nr), 0.75)
    
    ' Calculate IQR
    IQR = Q3 - Q1
    
    ' Calculate the lower and upper bounds
    LowerBound = Q1 - 1.5 * IQR
    UpperBound = Q3 + 1.5 * IQR
    
    ' Loop through each cell in the range E2:ELastRow
    For Each cell In ws.Range("F2:F" & Nr).Cells
        If cell.Value < LowerBound Or cell.Value > UpperBound Then
            'Count the Outlier
            OutlierCount = OutlierCount + 1
            ' Delete the outlier and entire row
            cell.EntireRow.Delete
        End If
    Next cell
    
MsgBox OutlierCount & " Outliers Deleted"

End Sub

Function EmailExtract(Email As String) As String
Dim L As Integer, AtLoc As Integer, StartLoc As Integer, EndLoc As Integer, i As Integer

Application.ScreenUpdating = False
L = Len(Email)
AtLoc = InStr(Email, "@")

For i = AtLoc - 1 To 1 Step -1
    If Mid(Email, i, 1) = "[" Or Mid(Email, i, 1) = "<" Or Mid(Email, i, 1) = ":" Or Mid(Email, i, 1) = " " Then
        StartLoc = i + 1
        Exit For
    ElseIf i = 1 Then
        StartLoc = 1
    End If
Next i

For i = AtLoc + 1 To L
    If Mid(Email, i, 1) = "]" Or Mid(Email, i, 1) = ">" Then
        EndLoc = i - 1
        Exit For
    ElseIf i = L Then
        EndLoc = L
    End If
Next i

EmailExtract = Mid(Email, StartLoc, EndLoc - StartLoc + 1)
End Function

