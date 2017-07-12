Option Explicit
Function urlBuilder(fileNum As Integer) As String()

'
' urlBuilder Macro
' Replace values in the url to navigate
' Returns an array of strings with each file to explore
' Selection.End(xlDown).Select
    Dim i As Integer, arreglo() As String, Path As String
    
    ReDim arreglo(fileNum - 1)
    Path = "http://sacnte335/Reports/Pages/Folder.aspx?ItemPath=%2f"
    
    For i = 1 To fileNum
        arreglo(i - 1) = Cells(i + 1, 5).Value
        arreglo(i - 1) = Replace(arreglo(i - 1), "/", "%2f")
        arreglo(i - 1) = Path & Replace(arreglo(i - 1), " ", "+")
        
    Next i
    
    urlBuilder = arreglo
    
End Function
