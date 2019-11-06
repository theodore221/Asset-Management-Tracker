Attribute VB_Name = "Navigation"
Sub TrackerButton()

    Application.ScreenUpdating = False
    Call Initalise
    Tracker_WS.Select
    Application.ScreenUpdating = True

End Sub

Sub ListButton()

    Application.ScreenUpdating = False
    Call Initalise
    LIST_WS.Select
    Application.ScreenUpdating = True
    
End Sub
