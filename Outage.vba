
Private Sub delete_button_Click()

    Application.DisplayAlerts = False
    Dim OldOutageCell As Range
    
    If CurrentRow <> 0 Then
        Set OldOutageCell = Range(Cells(CurrentRow, CurrentSCol), Cells(CurrentRow, CurrentECol))
        
        If OldOutageCell.Columns.Count > 1 Then
            OldOutageCell.UnMerge
        End If
        
        OldOutageCell.Cells(1, 1) = ""
        
        If Not OldOutageCell.Cells(1, 1).Comment Is Nothing Then
            OldOutageCell.ClearComments
        End If
        
        OldOutageCell.Interior.Color = xlNone
        With OldOutageCell.Borders
            .LineStyle = xlContinous
            .Weight = xlThin
            .ColorIndex = 1
        End With
        
        
        Dim DeleteRow As Integer
        DeleteRow = FindOutage(outageid_tb.value)
        
        If DeleteRow = 0 Then
        
        Else
            LIST_WS.Range("Table2").Rows(DeleteRow).Delete
        End If
        
        MsgBox ("Outage Deleted")
        Unload Me
    
    End If
    Application.DisplayAlerts = True

End Sub

Private Sub enddate_button_Click()

    Dim TempDate, DateToAdd As Date
    Dim SelectDate As Date
    
    If enddate_tb <> "" Then
        TempDate = DateValue(enddate_tb.value)
        SelectDate = TempDate
    Else
        SelectDate = Date
    End If
    
    DateToAdd = CalendarForm.GetDate(SelectedDate:=SelectDate)
    Me.enddate_tb = Format(DateToAdd, "dd/mm/yyyy")
    
    If enddate_tb = "30/12/1899" Then
        enddate_tb = Format(TempDate, "dd/mm/yyyy")
    End If

End Sub

Private Sub startdate_button_Click()

    Dim TempDate, DateToAdd As Date
    Dim SelectDate As Date
    
    If startdate_tb <> "" Then
        TempDate = DateValue(startdate_tb.value)
        SelectDate = TempDate
    Else
        SelectDate = Date
    End If
    
    DateToAdd = CalendarForm.GetDate(SelectedDate:=SelectDate)
    startdate_tb = Format(DateToAdd, "dd/mm/yyyy")
    
    If startdate_tb = "30/12/1899" Then
        startdate_tb = Format(TempDate, "dd/mm/yyyy")
    End If

End Sub

Function CreateAssets()

Call Initalise


Dim Unique As Collection
Dim Assets() As String

With Tracker_WS

    Set Unique = New Collection
    For Each Cell In GetProjectRng()
        On Error Resume Next
        Unique.Add Cell.value, CStr(Cell.value)
        On Error GoTo 0
    Next Cell
    
End With




ReDim Assets(1 To Unique.Count)

For i = LBound(Assets) To UBound(Assets)
    Assets(i) = Unique(i)
Next i

CreateUniqueAssets = Assets


End Function


Private Sub UserForm_Initialize()

    'If Outage ID is not passed in then it generates the next available ID number
    outageid_tb.value = GetNextID()
    
    'Change to add Assets more dynamically
    'site_cb.List = Array("Bayswater", "Liddell", "Callide C", "Vales Point", "Mt. Piper", "Tallawarra", "Yallourn", "Eraring", "Tarong Nth", "Muja", "Calaca")
    
    'Dim Assets() As String
    'Assets = CreateAssets()
    
    'For i = 1 To Assets.Length - 1
    '    site_cb.AddItem (Assets(i))
    'Next
    
    
    Dim Unique As Collection
    Dim Assets() As String

    With Tracker_WS

        Set Unique = New Collection
        For Each Cell In GetProjectRng()
            On Error Resume Next
            Unique.Add Cell.value, CStr(Cell.value)
            On Error GoTo 0
        Next Cell
    
    End With

    For i = 1 To Unique.Count
        site_cb.AddItem (Unique(i))
    Next
    


'ReDim Assets(1 To Unique.Count)

'For i = LBound(Assets) To UBound(Assets)
    'Assets(i) = Unique(i)
'Next i
    
    
    
    
    
    
    'site_cb.LIST = GetAssetArray()
    outagecat_cb.List = Array("Major", "Minor", "Retrofit", "Retrofit + AVR")
    ticinvolve_cb.List = Array("Heavy Involvement", "Minor Involvement", "No Involvement")

End Sub



'Frees up userform data and exits current userform
Private Sub cancel_button_Click()

    Unload Me

End Sub

Private Sub outagecat_cb_Change()

End Sub

'Saves data entered into userform into outage list and tracker
Private Sub save_button_Click()

    If DateValue(Me.startdate_tb) > DateValue(Me.enddate_tb) Then
        MsgBox ("selected end date is earlier then start date, please try again")
        Exit Sub
    End If

    'Checks to see that all information is filled in correctly
    If site_cb <> "" And unit_cb <> "" And startdate_tb <> "" And enddate_tb <> "" And outagecat_cb <> "" And ticinvolve_cb <> "" Then
    
        'Check to ensure information has been saved to both list and tracker
        Dim Success As Integer
        Success = SaveOutageToTracker()
        
        If Success Then
            Success = SaveOutageToList(outageid_tb.value)
        End If
        
        
        'If Save was successful then displays a message and exits userform
        If Success Then
            MsgBox ("Outage Details Saved Successfully")
            Unload Me
        End If
        
    Else
        MsgBox ("Please fill in all required fields before saving")
    End If

End Sub

Private Sub scope_lbl_Click()

End Sub

Private Sub site_cb_Change()

    unit_cb.Clear
    
    Dim AssetString As String
    AssetString = site_cb.value
    
    For Each Cell In GetProjectRng()
        If Cell.value = AssetString Then
            unit_cb.AddItem Cells(Cell.Row, Cell.Column + 1).value
        End If
    Next Cell

End Sub



Private Sub outageid_tb_Change()

    If IsNumeric(outageid_tb.value) And outageid_tb.value > 0 Then
        LoadData (outageid_tb.value)
    End If

End Sub


Function SaveOutageToList(OutageID As Integer) As Integer
    
    Dim SaveRow As Integer
    SaveRow = FindOutage(OutageID)
    
    If SaveRow = 0 Then
        SaveRow = AddNewRow
    End If
    
    With LIST_WS.Range("Table2")
        .Cells(SaveRow, 1) = OutageID
        .Cells(SaveRow, 2) = site_cb.value & " Unit " & unit_cb.value & ", " & outagecat_cb & " (" & UCase(MonthName(DatePart("m", startdate_tb), True)) & DatePart("yyyy", startdate_tb) & ")"
        .Cells(SaveRow, 3) = site_cb.value
        .Cells(SaveRow, 4) = unit_cb.value
        
        .Cells(SaveRow, 5) = FindCountry(site_cb.value, unit_cb.value)
        .Cells(SaveRow, 6) = FindType(site_cb.value, unit_cb.value)
        
        .Cells(SaveRow, 7) = DateValue(startdate_tb)
        .Cells(SaveRow, 8) = DateValue(enddate_tb)
        .Cells(SaveRow, 9) = DateDiff("d", startdate_tb.value, enddate_tb.value)
        .Cells(SaveRow, 10) = outagecat_cb.value
        .Cells(SaveRow, 11) = scope_tb.value
        .Cells(SaveRow, 12) = ticinvolve_cb.value
    End With
    
    'returns 1 if saved sucessfully
    SaveOutageToList = 1

End Function

Function SaveOutageToTracker() As Integer
    
    'Disables Alerts when merging
    Application.DisplayAlerts = False

    'Tracker Variables
    Dim AssetRange, MonthRange As Range
    Dim NewRow, NewSCol, NewECol As Integer
    Dim CompareVal, MonthStr, YearStr As String
    Dim DateFound As Integer
    
    
    With Tracker_WS
    
        Set AssetRange = GetProjectRng()
        Set MonthRange = GetDateRange()
        
        'Find the Asset Row for the Outage
        For Each asset In AssetRange
            If StrComp(asset.value, site_cb) = 0 Then
                If StrComp(asset.Offset(0, 1), unit_cb) = 0 Then
                    NewRow = asset.Row
                    Exit For
                End If
            End If
        Next asset
        
        'Finds the Month where the Outage Starts
        MonthStr = MonthName(DatePart("m", startdate_tb.value), True)
        YearStr = DatePart("yyyy", startdate_tb.value)
        DateFound = 0
        
        'Finds the start month column
        For Each Mth In MonthRange
            If StrComp(Mth.value, MonthStr) = 0 Then
                If StrComp(Mth.Offset(-1, 0).MergeArea.Cells(1, 1), YearStr) = 0 Then
                    NewSCol = Mth.Column
                    DateFound = 1
                    Exit For
                End If
            End If
        Next Mth
        
        If DateFound = 0 Then
            MsgBox ("Selected dates are not within current Tracker range")
            SaveOutageToTracker = 0
            Exit Function
        End If
        
        'Find the Month where the Outage Ends
        MonthStr = MonthName(DatePart("m", enddate_tb.value), True)
        YearStr = DatePart("yyyy", enddate_tb.value)
        DateFound = 0
        
        'Find the end month column
        For Each EMth In MonthRange
            If StrComp(EMth, MonthStr) = 0 Then
                If StrComp(EMth.Offset(-1, 0).MergeArea.Cells(1, 1), YearStr) = 0 Then
                    NewECol = EMth.Column
                    DateFound = 1
                    Exit For
                End If
            End If
        Next EMth
        
        If DateFound = 0 Then
            MsgBox ("Selected dates are not within current Tracker range")
            SaveOutageToTracker = 0
            Exit Function
        End If
        
        'Variable to check if the position of the outage cell should change
        Dim Same As Boolean
        Same = True
        
        'Variables to store the value of old outage incase things fail
        Dim OldOutageCell As Range
        Dim TempValue As String
        Dim TempColor As Long
        Dim TempComment As String
        
        If CurrentRow <> NewRow Or CurrentSCol <> NewSCol Or CurrentECol <> NewECol Then
            Same = False
        End If
        
        Dim NewOutageCell As Range
        Set NewOutageCell = Range(Cells(NewRow, NewSCol), Cells(NewRow, NewECol))
        
        If Not Same Then
            
            'If theres a old outage cell then removes the cell and clears formatting
            If CurrentRow <> 0 Then
                Set OldOutageCell = Range(Cells(CurrentRow, CurrentSCol), Cells(CurrentRow, CurrentECol))
                
                If OldOutageCell.Columns.Count > 1 Then
                    OldOutageCell.UnMerge
                End If
        
                TempValue = OldOutageCell.Cells(1, 1).Value2
                OldOutageCell.Cells(1, 1) = ""
                TempColor = OldOutageCell.Cells(1, 1).Interior.Color
                
                
                If OldOutageCell.Cells(1, 1).Comment Is Nothing Then
                    TempComment = ""
                Else
                    TempComment = OldOutageCell.Cells(1, 1).Comment.Text
                    OldOutageCell.ClearComments
                End If
                
            
                OldOutageCell.Interior.Color = xlNone
                With OldOutageCell.Borders
                    .LineStyle = xlContinous
                    .Weight = xlThin
                    .ColorIndex = 1
                End With
            End If
            
            
            'Checks to see if there is a outage already in that cell location
            'If yes then errors out
            For Each Cell In NewOutageCell
                If Cell.Value2 <> "" Then
                    
                    If CurrentRow <> 0 Then
                        OldOutageCell.Value2 = TempValue
                        OldOutageCell.Merge
                        OldOutageCell.Interior.Color = TempColor
                        If TempComment <> "" Then
                            OldOutageCell.AddComment TempComment
                        End If
                    End If
                    
                    MsgBox ("There's already an outage at that time, please fix conflict try again")
                    
                    SaveOutageToTracker = 0
                    Exit Function
                End If
            Next Cell
            
        End If
            
            
        'Add New outage to required position
        NewOutageCell.Value2 = outagecat_cb
        NewOutageCell.Merge
        
        Select Case ticinvolve_cb
        
            Case "Heavy Involvement"
                If outagecat_cb = "Major" Then
                    NewOutageCell.Interior.Color = RGB(190, 235, 250)
                End If
                
                If outagecat_cb = "Minor" Then
                    NewOutageCell.Interior.Color = RGB(199, 204, 228)
                End If
                
                If outagecat_cb = "Retrofit" Then
                    NewOutageCell.Interior.Color = RGB(241, 65, 36)
                End If
                
                If outagecat_cb = "Retrofit + AVR" Then
                    NewOutageCell.Interior.Color = RGB(241, 65, 36)
                End If
                
            Case "Minor Involvement"
                NewOutageCell.Interior.Color = RGB(201, 242, 151)
            
            Case "No Involvement"
                NewOutageCell.Interior.Color = RGB(217, 217, 217)
            
        End Select
    
        
        If scope_tb <> "" Then
            Dim CommentText As String
            CommentText = scope_tb.value
            If NewOutageCell.Cells(1, 1).Comment Is Nothing Then
                NewOutageCell.Cells(1, 1).AddComment CommentText
            Else
                NewOutageCell.Cells(1, 1).Comment.Text (CommentText)
            End If
        End If
    
    End With
    
    SaveOutageToTracker = 1
    
    Application.DisplayAlerts = True

End Function

Private Sub LoadData(id As Integer)

Dim DataRow As Integer
Dim TargetSheet As Worksheet

Set TargetSheet = Worksheets("List")

DataRow = FindOutage(id)

If DataRow = 0 Then
    GoTo cleanup
End If

With TargetSheet.Range("Table2")
    projectname_lbl.Caption = .Cells(DataRow, 2)
    site_cb = .Cells(DataRow, 3)
    unit_cb = .Cells(DataRow, 4)
    startdate_tb = Format(.Cells(DataRow, 7).value, "dd/mm/yyyy")
    enddate_tb = Format(.Cells(DataRow, 8).value, "dd/mm/yyyy")
    outagecat_cb = .Cells(DataRow, 10)
    ticinvolve_cb = .Cells(DataRow, 12)
    scope_tb = .Cells(DataRow, 11)
End With

cleanup:

End Sub

