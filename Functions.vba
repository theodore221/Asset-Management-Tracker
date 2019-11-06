Attribute VB_Name = "Functions"
Public Const TRACK As String = "Tracker"
''Public Const LIST_WS As String = "List"


Public Function GetNextID()
    
    ''Dim Context As Worksheet
   '' Set Context = Application.Workbooks(ThisWorkbook.Name).Sheets(DATA)
    GetNextID = WorksheetFunction.Max(LIST_WS.[Table2[Outage ID]]) + 1
End Function



Public Function GetAssetArray()

    Dim Context As Worksheet
    Set Context = Application.Workbooks(ThisWorkbook.Name).Sheets(TRACK)
    
    Dim Temp(50) As Variant
    Dim Count As Integer
    Dim NextRow As Range
    
    Count = 0
    Set NextRow = Range("project_list")
    
    While (NextRow.Offset(1, 0).Value2 <> "")
        NextRow = NextRow.Offset(1, 0)
        
        If NextRow.Value2 <> NextRow.Offset(-1, 0).Value2 Then
            Temp(Count) = NextRow.Value2
            Count = Count + 1
        End If
    Wend
    
    ReDim AssetArr(Count) As String
    
    For i = 0 To Count - 1
        AssetArr(i) = Temp(i)
    Next i
    
    GetAssetArray = AssetArr

End Function
