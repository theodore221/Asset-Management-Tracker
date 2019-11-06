Attribute VB_Name = "OutageFunctions"
'Array / List Management
Public OutArr() As Variant
Public ArrTracker As Integer

'Form Managment
Public OutageID As Integer

'For Tracker Mangement
Public CurrentRow, CurrentSCol, CurrentECol As Integer

'For Spreadsheet Management
Public Tracker_WS, LIST_WS, ASSET_WS As Worksheet

' Initalises Worksheet Names for future Functions & Subs
Sub Initalise()

Set Tracker_WS = Worksheets("Tracker")
Set LIST_WS = Worksheets("List")
Set ASSET_WS = Worksheets("Asset Reference")

End Sub

'Updates all line formatting for the tracker
Sub UpdateFormat()

'Finds new bounds for first row, last row, and columns and make new grids lines

    Call Initalise
    
    Tracker_WS.Select
    

    ColumnFirst = Range("project_list").Column
    ColumnLast = Cells(Range("project_list").Row, Tracker_WS.Columns.Count).End(xlToLeft).Column
    RowFirst = Range("tracker_start").Row
    RowLast = Range("project_list").End(xlDown).Row
    
    TypeCol = Range("tracker_type_hdr").Column
    MonthRow = Range("project_list").Row
    
    
    With Range(Cells(RowFirst, ColumnFirst), Cells(RowLast, ColumnLast))
    
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        
        
    End With
    
    'Thick Borders across entire tracker
    Range(Cells(RowFirst, ColumnFirst), Cells(RowLast, ColumnLast)).BorderAround LineStyle:=xlContinuous, Weight:=xlThick
    'Lines around the column headings
    Range(Cells(RowFirst, ColumnFirst), Cells(MonthRow, ColumnLast)).BorderAround LineStyle:=xlContinuous, Weight:=xlThick
    'Lines acrous the asset section
    Range(Cells(RowFirst, ColumnFirst), Cells(RowLast, TypeCol)).BorderAround LineStyle:=xlContinuous, Weight:=xlThick
    
    Dim Done As Boolean
    Dim ActiveCell As Range
    
    
    'SubRoutine, Groups Assets together and fills borders
    
    Done = False
    Set ActiveCell = Range("project_list").Offset(1, 0)
        
    While Not Done
        StartRow = ActiveCell.Row
        StartWord = ActiveCell.value
    
        While Not InStr(ActiveCell.value, StartWord) = 0
           Set ActiveCell = ActiveCell.Offset(1, 0)
        Wend
        
        EndRow = ActiveCell.Offset(-1, 0).Row
        Range(Cells(StartRow, ColumnFirst), Cells(EndRow, ColumnLast)).BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
        
        If ActiveCell.Offset(-1, 0).Row = RowLast Then
            Done = True
        End If
    Wend
    
    
    'SubRoutine, Groups and Fills Borders into Years
    
    Done = False
    Set ActiveCell = Range("month_start")
        
    While Not Done
        StartCol = ActiveCell.Column
        StartYear = ActiveCell.Offset(-1, 0).value
    
        While ActiveCell.Offset(-1, 0).MergeArea.Cells(1, 1).value = StartYear
           Set ActiveCell = ActiveCell.Offset(0, 1)
        Wend
        
        EndCol = ActiveCell.Offset(0, -1).Column
        Range(Cells(RowFirst, StartCol), Cells(RowLast, EndCol)).BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
        
        If ActiveCell.Offset(0, -1).Column = ColumnLast Then
            Done = True
        End If
    Wend
    
    
    With Range(Cells(RowFirst, ColumnFirst), Cells(RowLast, ColumnLast))
    
        .HorizontalAlignment = xlHAlignCenter
    
    End With

End Sub

'Refreshs the Tracker, Updates the Formating
'TODO: ADD IN ABILITY TO GENERATE NEW YEARS AND HIDE YEARS GONE PAST
Sub Refresh()

Application.ScreenUpdating = False


Call UpdateFormat


Application.ScreenUpdating = True

End Sub

Sub EditAssetList()

Call Initalise

Set ASSET_WS = Worksheets("Asset Reference")
ASSET_WS.Visible = xlSheetVisible
ASSET_WS.Select

End Sub


Sub AddAssetToTracker()

    Application.ScreenUpdating = False
    
    Call Initalise
    'Adding new assets to Tracker from Asset Reference List
    
    ASSET_WS.Select
    
    Dim AssetList As Range
    Dim FirstRow_AL, LastRow_AL As Integer
    Dim AssetCol_AL, UnitCol_AL, CountryCol_AL, TypeCol_AL As Integer
    Dim Asset_AL, Unit_AL, Country_AL, Type_AL As String
    Dim ActiveCell As Range
    Dim FOUND As Boolean
    Dim ColumnLast, ColumnFirst, RowFirst, RowLast As Integer
    
    
    FOUND = False
    AssetCol_AL = Range("al_assetname_hdr").Column
    FirstRow_AL = Range("al_assetname_hdr").Row + 1
    LastRow_AL = Range("al_assetname_hdr").End(xlDown).Row
    
    
    'Adds all Assets from this list into the tracker if not already added.
    
    Set AssetList = Range(Cells(FirstRow_AL, AssetCol_AL), Cells(LastRow_AL, AssetCol_AL))
    
    For Each asset In AssetList
    
        FOUND = False
        INSERTED = False
        Asset_AL = asset.value
        Unit_AL = asset.Offset(0, 1).value
        Country_AL = asset.Offset(0, 2).value
        Type_AL = asset.Offset(0, 3).value
        
        
        Tracker_WS.Select
        Set ActiveCell = Range("project_list").Offset(1, 0)
        
        ColumnFirst = Range("project_list").Column
        ColumnLast = Cells(Range("project_list").Row, Tracker_WS.Columns.Count).End(xlToLeft).Column
        RowLast = ActiveCell.End(xlDown).Row
        
        While Not FOUND
                    
            If Not InStr(ActiveCell.value, Asset_AL) = 0 Then
                If ActiveCell.Offset(0, 1).value = Unit_AL Then
                    FOUND = True
                    
                Else
                    If InStr(ActiveCell.Offset(1, 0).value, Asset_AL) = 0 Then
                    
                        With Range(Cells(ActiveCell.Offset(1, 0).Row, ColumnFirst), Cells(ActiveCell.Offset(1, 0).Row, ColumnLast))
                            .Insert (xlShiftDown)
                        End With
                        
                        Range(Cells(ActiveCell.Offset(1, 0).Row, ColumnFirst), Cells(ActiveCell.Offset(1, 0).Row, ColumnLast)).ClearFormats
                        Cells(ActiveCell.Offset(1, 0).Row, Range("project_list").Column).value = Asset_AL
                        Cells(ActiveCell.Offset(1, 0).Row, Range("project_list").Column + 1).value = Unit_AL
                        Cells(ActiveCell.Offset(1, 0).Row, Range("project_list").Column + 2).value = Country_AL
                        Cells(ActiveCell.Offset(1, 0).Row, Range("project_list").Column + 3).value = Type_AL
                        
                        FOUND = True
                        INSERTED = True
    
                    End If
                End If
            End If
            
            'Set to Last Row If ActiveCell.Row =
            If ActiveCell.Row = RowLast And INSERTED = False Then
                With Range(Cells(ActiveCell.Offset(1, 0).Row, ColumnFirst), Cells(ActiveCell.Offset(1, 0).Row, ColumnLast))
                    .Insert (xlShiftDown)
                    .ClearFormats
                End With
                
                Cells(ActiveCell.Offset(1, 0).Row, Range("project_list").Column).value = Asset_AL
                Cells(ActiveCell.Offset(1, 0).Row, Range("project_list").Column + 1).value = Unit_AL
                Cells(ActiveCell.Offset(1, 0).Row, Range("project_list").Column + 2).value = Country_AL
                Cells(ActiveCell.Offset(1, 0).Row, Range("project_list").Column + 3).value = Type_AL
                
                FOUND = True
                INSERTED = True
            End If
            
        Set ActiveCell = ActiveCell.Offset(1, 0)
            
        Wend
        
        ASSET_WS.Select
    
    Next asset
    
    Tracker_WS.Select
    ASSET_WS.Visible = xlSheetHidden
    
    Call Refresh
    
    Application.ScreenUpdating = True


End Sub


'Initialization to load data into list from Tracker
'Sub only needs to be run once, unless tracker is manually edited

Sub LoadList()

    Call Initalise
    
    Tracker_WS.Select
    
    Dim TrackerRange, Cell As Range
    Dim AssetCol, UnitCol As Integer
    Dim ProjectStr As String
    
    OutageID = 0
    
    ReDim OutArr(30, 10)
    ArrTracker = 0
    
    AssetCol = Range("project_list").Column
    UnitCol = Range("project_list").Column + 1
    
    Set TrackerRange = GetTrackerRng()
    
    For Each Cell In TrackerRange
        
        If ArrTracker > UBound(OutArr, 1) Then
            OutArr = ReDimPreserve(OutArr, 1.5 * UBound(OutArr, 1), 10)
        End If
        
        If Cell <> "" Then
            ProjectStr = GetProjectStr(Cell)
            OutageID = OutageID + 1
            OutArr(ArrTracker, 0) = OutageID
            
            OutArr(ArrTracker, 1) = GetProjectStr(Cell)
            OutArr(ArrTracker, 2) = Cells(Cell.Row, AssetCol).value
            OutArr(ArrTracker, 3) = Cells(Cell.Row, UnitCol).value
            OutArr(ArrTracker, 4) = GetStartDate(Cell)
            OutArr(ArrTracker, 5) = GetEndDate(Cell)
            OutArr(ArrTracker, 6) = DateDiff("d", GetStartDate(Cell), GetEndDate(Cell))
            OutArr(ArrTracker, 7) = Cell.value
            OutArr(ArrTracker, 8) = OutageID
            
            If Not Cell.Comment Is Nothing Then
                OutArr(ArrTracker, 8) = Cell.Comment.Text
            Else
                OutArr(ArrTracker, 8) = ""
            End If
            
            
            Select Case (Cell.Interior.Color)
            Case RGB(190, 235, 250)
                OutArr(ArrTracker, 9) = "Heavy Involvement"
            Case RGB(199, 204, 228)
                OutArr(ArrTracker, 9) = "Heavy Involvement"
            Case RGB(241, 65, 36)
                OutArr(ArrTracker, 9) = "Heavy Involvement"
            Case RGB(201, 242, 151)
                OutArr(ArrTracker, 9) = "Minor Involvement"
            Case RGB(217, 217, 217)
                OutArr(ArrTracker, 9) = "No Involvement"
            Case Else
                OutArr(ArrTracker, 9) = Cell.value
            End Select
            
            
        ArrTracker = ArrTracker + 1
            
        End If
        
    Next Cell
    
    'Load Array to a Sheet
    
    LIST_WS.Select
    
    
    Dim SaveRange As Range
    
    
    FirstRow = LBound(OutArr, 1) + 5
    FirstCol = LBound(OutArr, 1) + 2
    LastRow = UBound(OutArr, 1) + 1
    LastCol = UBound(OutArr, 2) + 1
    
    Set SaveRange = Range(Cells(FirstRow, FirstCol), Cells(LastRow, LastCol))
        
    SaveRange.ClearContents
    SaveRange.value = OutArr


End Sub

Function GetStartDate(Cell As Range) As Date

 Dim MonthRow, YearRow As Integer
 
 Dim Month, Year As String
 Dim StartCellCol As Integer
 
 MonthRow = Range("project_list").Row
 YearRow = Range("project_list").Row - 1
 
 StartCellCol = Cell.MergeArea.Cells(1, 1).Column
 
 Month = Cells(MonthRow, StartCellCol).value
 Year = Cells(YearRow, StartCellCol).MergeArea.Cells(1, 1).value
 
 GetStartDate = DateValue(Month & " 01, " & Year)

End Function

Function GetEndDate(Cell As Range) As Date

 Dim MonthRow, YearRow As Integer
 
 Dim Month, Year As String
 Dim CellCol As Integer
 
 MonthRow = Range("project_list").Row
 YearRow = Range("project_list").Row - 1
 
 CellCol = Cell.MergeArea.Cells(1, Cell.MergeArea.Count).Column
 
 Month = Cells(MonthRow, CellCol).value
 Year = Cells(YearRow, CellCol).MergeArea.Cells(1, 1).value
 
 GetEndDate = DateValue(Month & " 28, " & Year)

End Function

'Calls a new userform to create new Outage
Sub AddNewOutage()

    Call Initalise

    CurrentRow = 0
    CurrentSCol = 0
    CurrentECol = 0

    Dim frm As New Outage
    frm.Show

End Sub

Sub QuickEdit()

    If ActiveCell.value <> "" Then

        Call Initalise
    
        Dim SiteCol, UnitCol, MonthRow, YearRow As Integer
        Dim SearchString As String
        Dim TargetSite, TargetUnit, TargetMonth, TargetYear As String
        Dim OutageID As Integer
        
        Dim frm As New Outage
        
        With Tracker_WS
        
            SiteCol = .Range("project_list").Column
            UnitCol = .Range("tracker_unit_hdr").Column
            MonthRow = .Range("project_list").Row
            YearRow = MonthRow - 1
            
            
         
            CurrentRow = ActiveCell.Row
            CurrentSCol = ActiveCell.MergeArea.Cells(1, 1).Column
            CurrentECol = ActiveCell.MergeArea.Cells(ActiveCell.MergeArea.Rows.Count, ActiveCell.MergeArea.Columns.Count).Column
            
            TargetSite = .Cells(CurrentRow, SiteCol)
            TargetUnit = .Cells(CurrentRow, UnitCol)
            TargetMonth = .Cells(MonthRow, CurrentSCol)
            TargetYear = .Cells(YearRow, CurrentSCol).MergeArea.Cells(1, 1)
        
        End With
        
        SearchString = TargetSite & " Unit " & TargetUnit & ", " & ActiveCell.value & " (" & UCase(TargetMonth) & TargetYear & ")"
        OutageID = GetOutageID(SearchString)
    
        Tracker_WS.Select
    
        frm.outageid_tb.value = OutageID
        frm.Show
    
    Else
        MsgBox ("Please Select a Outage from the Tracker before editing")
        
        
    End If


End Sub

Sub QuickEditList()

    If ActiveCell <> "" And ActiveCell.Row <> 4 Then
        
        Call Initalise
        
        Dim frm As New Outage
        Dim IdColumn As Integer
        
        With LIST_WS
            frm.outageid_tb.value = Cells(ActiveCell.Row, Range("outageid_hdr").Column)
        End With
        
        frm.Show
    
    Else
    
        MsgBox ("Please select a outage from the Table before editing")
    
    End If
    

End Sub





