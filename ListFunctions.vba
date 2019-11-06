Attribute VB_Name = "ListFunctions"
Function AddNewRow()

    Dim tbl As ListObject
    Set tbl = LIST_WS.ListObjects("Table2")
    tbl.ListRows.Add
    AddNewRow = tbl.ListRows.Count

End Function
    
    
'Finds Outage ID given a ProjectName from LIST

Function GetOutageID(TargetProject As String) As Integer

    Dim FirstRow, LastRow As Integer
    Dim SearchRange As Range
    
    
    With LIST_WS
        FirstRow = .Range("list_outagename_hdr").Row + 1
        LastRow = .Cells(.Rows.Count, "C").End(xlUp).Row
        Set SearchRange = .Range(.Cells(FirstRow, 3), .Cells(LastRow, 3))
        
        For Each Cell In SearchRange
            If Cell.value = TargetProject Then
                GetOutageID = .Cells(Cell.Row, 2).value
                Exit For
            End If
            GetOutageID = 0
        Next Cell
        
    End With

End Function



'Go through the Outage List and finds the position of required outage ID
'@param id: Outage ID number
'@return returns the table row number of id, returns 0 if not found
Public Function FindOutage(ByVal id As Integer)
    
    Dim vArr As Variant
    Dim i As Integer
    Dim FOUND As Boolean
  
    vArr = LIST_WS.Range("Table2[Outage ID]").Value2
    FindOutage = 0
    
    'Table is empty
    If LIST_WS.Range("B5").Value2 = "" Then
        FOUND = False
    End If
    
    For i = 1 To UBound(vArr)
        If vArr(i, 1) = id Then
            FindOutage = i
            FOUND = True
            Exit For
        End If
    Next

End Function


Sub UpdateList()

    Dim FirstRow, LastRow As Integer
    Dim UpdateRange As Range
    
    Call Initalise

    
    With LIST_WS
        FirstRow = .Range("list_outagename_hdr").Row + 1
        LastRow = .Cells(.Rows.Count, "C").End(xlUp).Row
        Set UpdateRange = .Range(.Cells(FirstRow, 4), .Cells(LastRow, 4))
        
        For Each Cell In UpdateRange
            Cell.Offset(0, 2).value = FindCountry(Cell.value, Cell.Offset(0, 1).value)
            Cell.Offset(0, 3).value = FindType(Cell.value, Cell.Offset(0, 1).value)
        Next Cell
        
    End With


End Sub

