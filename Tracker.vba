Attribute VB_Name = "Tracker"
Global OutArry() As Variant
Global ArrTracker As Integer
Global permission As Boolean
Public Const DATA As String = "Sheet4"
Public Const Tracker As String = "Sheet1"

Sub LoadOutageArray()

    Call LoadOutageArrayFromSheet
    Call LoadOutageArrayFromTracker

End Sub

Sub LoadOutageArrayFromSheet()

    Worksheets(DATA).Visible = True
    Worksheets(DATA).Select

    ReDim OutArry(30, 8)
    ArrTracker = 0
    
    Dim ActiveCell As Range
    
    Set ActiveCell = Cells(Range("data_projectname_hdr").Row + 1, Range("data_projectname_hdr").Column)
    
    Do While ActiveCell.value <> ""
        
        If ArrTracker > UBound(OutArry, 1) Then
            OutArry = ReDimPreserve(OutArry, 1.5 * UBound(OutArry, 1), 8)
        End If
        
        OutArry(ArrTracker, 0) = ActiveCell.value                     'ProjectName
        OutArry(ArrTracker, 1) = ActiveCell.Offset(0, 1).value        'Project Start Date
        OutArry(ArrTracker, 2) = ActiveCell.Offset(0, 2).value        'Project End Date
        OutArry(ArrTracker, 3) = ActiveCell.Offset(0, 3).value        'Project Description
        OutArry(ArrTracker, 4) = ActiveCell.Offset(0, 4).value        'Outage Category
        OutArry(ArrTracker, 5) = ActiveCell.Offset(0, 5).value        'Project Asset
        OutArry(ArrTracker, 6) = ActiveCell.Offset(0, 6).value        'Asset Unit
        OutArry(ArrTracker, 7) = ActiveCell.Offset(0, 7).value        'Cell String
        
        ArrTracker = ArrTracker + 1
        Set ActiveCell = ActiveCell.Offset(1, 0)
        
    Loop

End Sub

Sub LoadOutageArrayFromTracker()

    Worksheets(Tracker).Select

    'Tracker Variables
    Dim TrackerRange, Cell As Range
    Dim AssetCol, UnitCol As Integer
    
    'Project Variables
    Dim ProjectStr As String
    
    'Array Variables
    Dim InArray As Boolean
    Dim ArrPlaceVal As Integer
    
    AssetCol = Range("project_list").Column
    UnitCol = Range("project_list").Column + 1

    Set TrackerRange = GetTrackerRng()

    For Each Cell In TrackerRange                                               'Loops Thru each Cell in the Tracker
    
        'Extends the Size of the Array If Required
        If ArrTracker > UBound(OutArry, 1) Then
            OutArry = ReDimPreserve(OutArry, 1.5 * UBound(OutArry, 1), 8)
        End If
        
        'If Cell is not empty check value inside
        If Cell <> "" Then
        
            'Returns Project Name
            ProjectStr = GetProjectStr(Cell)
            
            If Not IsInArray(ProjectStr) Then
            
                ArrPlaceVal = ArrTracker
                ArrTracker = ArrTracker + 1
                
                                                                              
                OutArry(ArrPlaceVal, 0) = ProjectStr                             'Project Name
                OutArry(ArrPlaceVal, 1) = GetStartDate(Cell)                     'Project Start Date
                OutArry(ArrPlaceVal, 2) = GetEndDate(Cell)                       'Project End Date

                
                'If there is no comments then comment is empty
                If Not Cell.Comment Is Nothing Then
                    OutArry(ArrPlaceVal, 3) = Cell.Comment.Text                  'Project Comments
                Else
                    OutArry(ArrPlaceVal, 3) = ""
                End If
                
                
                Select Case (Cell.Interior.Color)                                     'Outage Category
                
                Case RGB(190, 235, 250)
                    OutArry(ArrPlaceVal, 4) = "Heavy Involvement"
                Case RGB(199, 204, 228)
                    OutArry(ArrPlaceVal, 4) = "Heavy Involvement"
                Case RGB(241, 65, 36)
                    OutArry(ArrPlaceVal, 4) = "Heavy Involvement"
                Case RGB(201, 242, 151)
                    OutArry(ArrPlaceVal, 4) = "Minor Involvement"
                Case RGB(217, 217, 217)
                    OutArry(ArrPlaceVal, 4) = "No Involvement"
                
                Case Else
                    OutArry(ArrPlaceVal, 4) = Cell.value
                    
                End Select
                
                OutArry(ArrPlaceVal, 5) = Cells(Cell.Row, AssetCol).value        'Outage Asset
                OutArry(ArrPlaceVal, 6) = Cells(Cell.Row, UnitCol).value         'Outage Asset Unit
                OutArry(ArrPlaceVal, 7) = Cell.value                              'Cell String
            
            End If
        End If
    Next Cell
    
End Sub

Sub SaveArraytoSheet()

    Worksheets(DATA).Select
    
    Dim SaveRange As Range
    Dim FirstRow, LastRow, FirstCol, LastCol As Integer
    
    FirstRow = LBound(OutArry, 1) + 2
    FirstCol = LBound(OutArry, 1) + 1
    LastRow = UBound(OutArry, 1) + 1
    LastCol = UBound(OutArry, 2) + 1

    Set SaveRange = Range(Cells(FirstRow, FirstCol), Cells(LastRow, LastCol))
    
    SaveRange.ClearContents
    SaveRange.value = OutArry
    
    Worksheets(Tracker).Select
    
    Application.ScreenUpdating = True


End Sub

Sub UpdateTracker()

    Dim WS_Tracker As Worksheet
    Set WS_Tracker = Worksheets("Sheet1")

    Application.ScreenUpdating = False
    
    Worksheets(Tracker).Select
    
    Call ClearTracker
    
    Dim TargetCellRow, TargetCellSCol, TargetCellECol As Integer
    
    Dim TargetAsset As String
    Dim TargetUnit As String
    Dim CellColor As Long
    Dim CommentString As String
        
        For i = LBound(OutArry, 1) To UBound(OutArry, 1)
        
            If OutArry(i, 0) <> "" Then
            
                TargetAsset = OutArry(i, 5)
                TargetUnit = OutArry(i, 6)
                TargetCellRow = FindProjectAsset(TargetAsset, TargetUnit)
                TargetCellSCol = FindMonth(DateValue(OutArry(i, 1)))
                TargetCellECol = FindMonth(DateValue(OutArry(i, 2)))
            
                Application.DisplayAlerts = False
                
                Select Case (OutArry(i, 4))
                
                Case "Heavy Involvement"
                
                    If InStr(OutArry(i, 7), "Major") Then
                        CellColor = RGB(190, 235, 250)
                    ElseIf InStr(OutArry(i, 7), "Minor") Then
                        CellColor = RGB(199, 204, 228)
                    ElseIf InStr(OutArry(i, 7), "Retro") Then
                        CellColor = RGB(241, 65, 36)
                    End If
                
                Case "No Involvement"
                
                    CellColor = RGB(217, 217, 217)
                    
                Case "Minor Involvement"
                
                    CellColor = RGB(201, 242, 151)
                    
                End Select
                
                Cells(TargetCellRow, TargetCellSCol).value = OutArry(i, 7)
            
                With Range(Cells(TargetCellRow, TargetCellSCol), Cells(TargetCellRow, TargetCellECol))
                    .Merge
                    .Interior.Color = CellColor
                    .HorizontalAlignment = (xlCenterAcrossSelection)
                End With
                
                
                If OutArry(i, 3) <> "" Then
                    CommentString = OutArry(i, 3)
                    WS_Tracker.Cells(TargetCellRow, TargetCellECol).AddComment CommentString
                End If
                
            End If
            
        Next i
        
    'Application.ScreenUpdating = True
    
End Sub

Sub ClearTracker()

    'Application.ScreenUpdating = False
    
    Dim Tracker As Range
    Set Tracker = GetTrackerRng()
    
    With Tracker
        .ClearContents
        .ClearComments
        .UnMerge
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinous
        .Borders.Weight = xlThin
        .Borders(xlEdgeLeft).Weight = xlThick

        
    End With
    
    'Application.ScreenUpdating = True

End Sub

Sub SaveTrackerData()

Call LoadOutageArray

Dim Dbug As Integer

'DBug = ""

Worksheets("Sheet4").Select

'This needs to be dyanmic range creatged thru the array size
Range("A1:I80").value = OutArry

End Sub
'SUB GETS RID OF PAST MONTHS FROM THE TRACKER EARILER THEN 6 MONTHS FROM CURRENT MONTH
Sub ClearOldMonths()

    Dim ActiveCell As Range
    Dim ActiveDate As Date
    Dim YearRow, MonthRow As Integer
    
    Dim TargetYearString As String
    Dim StartRow, EndRow As Integer
    
    Dim StartMonthCol, EndMonthCol As Integer
    
    YearRow = Range("project_list").Row - 1
    MonthRow = Range("project_list").Row

    Set ActiveCell = Cells(MonthRow, Range("unit_list").Column + 1)
    StartMonthCol = ActiveCell.Column
    ActiveDate = DateValue(ActiveCell.value & " 01, " & Cells(ActiveCell.Row - 1, ActiveCell.Column).value)
    Dim Test As Integer
    Test = DateDiff("m", ActiveDate, Date)
        
    'Sub Only Runs if there are dates greater then 6 months
    If Test > 6 Then
    
        'Loops and find all the months to delete
        Do While Test > 6
    
            Set ActiveCell = ActiveCell.Offset(0, 1)
            ActiveDate = DateValue(ActiveCell.value & " 01, " & Cells(ActiveCell.Row - 1, ActiveCell.Column).MergeArea.Cells(1, 1).value)
            Test = DateDiff("m", ActiveDate, Date)
    
        Loop
    
        EndMonthCol = ActiveCell.Column - 1
    
        'Sub that goes thru the final column of month to delete and moves outages that may get deleted in that column
        ClearMonthColumn (EndMonthCol)
    
        TargetYearString = Cells(YearRow, StartMonthCol).value
        StartRow = Range("project_list").Row + 1
        EndRow = Range("project_list").End(xlDown).Row
    
        Range(Cells(YearRow, StartMonthCol), Cells(EndRow, EndMonthCol)).Delete
    
        Cells(YearRow, StartMonthCol).value = TargetYearString
    
   End If

End Sub


Sub ClearMonthColumn(TargetCol As Integer)

    Dim StartRow, EndRow As Integer
    Dim TargetColumn As Range
    
    Dim OutageSCol, OutageECol As Integer
    
    StartRow = Range("project_list").Row + 1
    EndRow = Range("project_list").End(xlDown).Row
    
    Set TargetColumn = Range(Cells(StartRow, TargetCol), Cells(EndRow, TargetCol))
    
    For Each Cell In TargetColumn
    
        If Cell.value <> "" Then
        
            OutageSCol = Cell.MergeArea.Cells(1, 1).Column
            OutageECol = Cell.MergeArea.Cells(Cell.MergeArea.Rows.Count, Cell.MergeArea.Columns.Count).Column
            
            If Not Cell.MergeArea.Cells(Cell.MergeArea.Rows.Count, Cell.MergeArea.Columns.Count).Column > TargetCol Then
            
            Else
            
                If Not OutageSCol = OutageECol Then
            
                    Cell.UnMerge
                    OutageSCol = OutageSCol + 1
                    Cells(Cell.Row, OutageSCol).value = Cells(Cell.Row, OutageSCol - 1).value
                    Range(Cells(Cell.Row, OutageSCol), Cells(Cell.Row, OutageECol)).Merge
            
                End If
            End If
        End If
    
    Next Cell

End Sub

Function FindMonth(TDate As Date) As Integer

    Dim Year, Month As String
    Dim YearRow, MonthRow As Integer
    Dim ActiveCell, DateRange As Range
    
    Year = DatePart("yyyy", TDate)
    Month = MonthName(DatePart("m", TDate), True)
    
    YearRow = Range("project_list").Row - 1
    MonthRow = Range("project_list").Row
    
    Dim TestString
 
    Set DateRange = GetDateRange()
    
    For Each Cell In DateRange
        If Cell.value = Month Then
            If Cells(YearRow, Cell.Column).MergeArea.Cells(1, 1).value = Year Then
                FindMonth = Cell.Column
                Exit For
            End If
        End If
    Next Cell
    
'    Set ActiveCell = DateRange.Find(What:=Month, SearchOrder:=xlByColumns)
'    Dim TestStr As String
'
'    TestStr = Cells(YearRow, ActiveCell.Column).MergeArea.Cells(1, 1).value
'    Do While TestStr <> Year
'
'        Set ActiveCell = DateRange.FindNext(ActiveCell)
'        TestStr = Cells(YearRow, ActiveCell.Column).MergeArea.Cells(1, 1).value
'    Loop
'
'    FindMonth = ActiveCell.Column

End Function

Sub TestSub()

Dim Month As Integer
Dim GDate As Date

GDate = "21 / 2 / 2019"

Month = FindMonth(GDate)


End Sub


Function FindProjectAsset(TargetAsset As String, TargetUnit As String) As Integer

    Dim AssetRange As Range
    
    Set AssetRange = Range(Cells(Range("project_list").Row, Range("project_list").Column), Cells(Cells(Range("labels").Row - 1, Range("labels").Column).End(xlUp).Row, Range("project_list").Column))
    
    For Each Cell In AssetRange
        If Cell.value = TargetAsset Then
            If Cells(Cell.Row, Cell.Column + 1).value = TargetUnit Then
                FindProjectAsset = Cell.Row
                Exit For
            End If
        End If
    Next Cell

End Function

Function FindNextPlace() As Integer

    Dim ArrayPlaces As Integer
    ArrayPlaces = 0
    
    
    For i = LBound(OutArry, 1) To UBound(OutArry, 1)
        If (OutArry(i, 0) = "") Then
            Exit For
        End If
        
        ArrayPlaces = ArrayPlaces + 1
        
    Next i
    
    FindNextPlace = ArrayPlaces

End Function

Function IsInArray(ProjectStr As String) As Boolean
    
    For i = LBound(OutArry, 1) To UBound(OutArry, 1)
        If InStr(OutArry(i, 0), ProjectStr) Then
            IsInArray = True
            Exit For
        End If
    IsInArray = False
    Next i

End Function


Function FindInArray(ProjectStr As String) As Integer

Dim InArray As Boolean
Dim ArrPlaceVal As Integer

For i = LBound(OutArry, 1) To UBound(OutArry, 1)
    If OutArry(i, 1) = ProjectStr Then
        ArrPlaceVal = i
        InArray = True
        Exit For
    End If
    
Next i
                
If InArray = True Then
                
Else
    ArrPlaceVal = ArrTracker
    ArrTracker = ArrTracker + 1
End If

FindInArray = ArrPlaceVal

End Function
'RETURN A STRING OF THE PROJECT NAME GIVEN A OUTAGE CELL
Function GetProjectStr(Cell As Range) As String

    Dim AssetStr, UnitStr, OutageStr, MonthStr, YearStr As String
    
    Dim CellCol As Integer
    OutageStr = Cell.value
    AssetStr = Cells(Cell.Row, Range("project_list").Column).value
    UnitStr = CStr(Cells(Cell.Row, Range("project_list").Column + 1).value)
    CellCol = Cell.MergeArea.Cells(1, 1).Column
    MonthStr = Cells(Range("project_list").Row, CellCol).value
    YearStr = Cells(Range("project_list").Row - 1, CellCol).MergeArea.Cells(1, 1).value
    
    GetProjectStr = AssetStr + " Unit " + UnitStr + ", " + OutageStr + " (" & UCase(MonthStr) & YearStr & ")"

End Function



Function CalculateDate(TargetRng As Range) As Date

    SelectSheet (Tracker)
        
    Dim MonthRow, YearRow As Integer
    
    MonthRow = Range("project_list").Row
    YearRow = Range("project_list").Row - 1
    
    CalculateDate = DateValue(Cells(MonthRow, TargetRng.Column).value & " 01, " & Cells(YearRow, TargetRng.Column).MergeArea.Cells(1, 1).value)
    
End Function










Sub Test1Sub()

Call LoadOutageArray
Dim Dbug As Integer

Call UpdateTracker
Call SaveArraytoSheet


End Sub
