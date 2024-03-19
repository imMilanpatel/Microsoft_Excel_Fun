Private Sub Worksheet_Change(ByVal Target As Range)
    Dim ws As Worksheet
    Dim lastRow As Long
    
    ' Define the worksheet where the change occurred
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    ' Check if the changed cell is within the specified range
    If Not Intersect(Target, ws.Range("C2:J11")) Is Nothing Then ' Assuming your data range is within C3:J11, adjust if necessary
        ' MsgBox "Change detected within the specified range." ' Debug message
        
        ' Disable events to prevent infinite loop
        Application.EnableEvents = False
        
        ' Determine the last row of data
        lastRow = ws.Cells(ws.Rows.Count, "J").End(xlUp).Row
        
        ' Sort the data based on column J (Point Percentage) in descending order
        With ws.Sort
            .SortFields.Clear
            .SortFields.Add2 Key:=ws.Range("J3:J" & lastRow), _
                SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
            .SetRange ws.Range("C2:J" & lastRow) ' Adjust the range as needed
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
        
        ' MsgBox "Sorting completed." ' Debug message
        
        ' Enable events back
        Application.EnableEvents = True
    End If
End Sub


