Attribute VB_Name = "AssetListFunctions"

'Returns country based on provided Asset and Unit

Function FindCountry(asset As String, unit As String) As String

    With ASSET_WS

        Dim AssetList As Range
        Dim FirstRow_AL, LastRow_AL As Integer
        Dim AssetCol_AL As Integer
        
        AssetCol_AL = .Range("al_assetname_hdr").Column
        FirstRow_AL = .Range("al_assetname_hdr").Row + 1
        LastRow_AL = .Range("al_assetname_hdr").End(xlDown).Row
        
        Set AssetList = Range(.Cells(FirstRow_AL, AssetCol_AL), .Cells(LastRow_AL, AssetCol_AL))
        
        For Each entry In AssetList
            If InStr(entry.value, asset) = 1 And entry.Offset(0, 1).value = unit Then
                FindCountry = entry.Offset(0, 2).value
                Exit For
            End If
        Next entry
        
    End With

End Function

'returns Asset type based on asset name and unit

Function FindType(asset As String, unit As String) As String

    With ASSET_WS

        Dim AssetList As Range
        Dim FirstRow_AL, LastRow_AL As Integer
        Dim AssetCol_AL As Integer
    
        AssetCol_AL = .Range("al_assetname_hdr").Column
        FirstRow_AL = .Range("al_assetname_hdr").Row + 1
        LastRow_AL = .Range("al_assetname_hdr").End(xlDown).Row
    
        Set AssetList = Range(.Cells(FirstRow_AL, AssetCol_AL), .Cells(LastRow_AL, AssetCol_AL))
    
        For Each entry In AssetList
            If InStr(entry.value, asset) = 1 And entry.Offset(0, 1).value = unit Then
                FindType = entry.Offset(0, 3).value
                Exit For
            End If
        Next entry
        
    End With

End Function
