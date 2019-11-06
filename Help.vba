Attribute VB_Name = "Help"
Public Function ReDimPreserve(aArrayToPreserve, nNewFirstUBound, nNewLastUBound)
    ReDimPreserve = False
    'check if its in array first
    If IsArray(aArrayToPreserve) Then
        'create new array
        ReDim aPreservedArray(nNewFirstUBound, nNewLastUBound)
        'get old lBound/uBound
        nOldFirstUBound = UBound(aArrayToPreserve, 1)
        nOldLastUBound = UBound(aArrayToPreserve, 2)
        'loop through first
        For nFirst = LBound(aArrayToPreserve, 1) To nNewFirstUBound
            For nLast = LBound(aArrayToPreserve, 2) To nNewLastUBound
                'if its in range, then append to new array the same way
                If nOldFirstUBound >= nFirst And nOldLastUBound >= nLast Then
                    aPreservedArray(nFirst, nLast) = aArrayToPreserve(nFirst, nLast)
                End If
            Next
        Next
        'return the array redimmed
        If IsArray(aPreservedArray) Then ReDimPreserve = aPreservedArray
    End If
End Function

