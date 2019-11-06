Attribute VB_Name = "TrackerFunctions"
'Returns all assest rows in the Tracker
Function GetProjectRng() As Range
    
    Dim RowFirst As Integer
    Dim RowLast As Integer
    Dim ProjectColumn As Long
    
    With Tracker_WS
    
        ProjectColumn = Range("project_list").Column
        RowFirst = Range("project_list").Row + 1
        RowLast = Cells(Range("labels").Row - 1, ProjectColumn).End(xlUp).Row
    
        Set GetProjectRng = Range(Cells(RowFirst, ProjectColumn), Cells(RowLast, ProjectColumn))
    
    End With
    
End Function

'Return all month columns in the Tracker
Function GetDateRange() As Range
   
   Dim ColumnFirst, ColumnLast, MonthRow As Integer
   
   With Tracker_WS
   
        MonthRow = Range("project_list").Row
        ColumnFirst = Range("project_list").Column + 2
        ColumnLast = Cells(Range("project_list").Row, Tracker_WS.Columns.Count).End(xlToLeft).Column
        Set GetDateRange = Range(Cells(MonthRow, ColumnFirst), Cells(MonthRow, ColumnLast))
   
   End With
    
End Function


Function GetTrackerRng() As Range

    Call Initalise
    
    Dim FirstCol, LastCol, FirstRow, LastRow As Integer
    FirstRow = Tracker_WS.Range("project_list").Row + 1
    LastRow = Tracker_WS.Cells(Range("labels").Row - 1, Range("project_list").Column).End(xlUp).Row
    FirstCol = Tracker_WS.Range("project_list").Column + 2
    LastCol = Tracker_WS.Cells(FirstRow - 1, Tracker_WS.Columns.Count).End(xlToLeft).Column
    
    Set GetTrackerRng = Range(Cells(FirstRow, FirstCol), Cells(LastRow, LastCol))

End Function
