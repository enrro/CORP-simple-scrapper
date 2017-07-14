Option Explicit
Function urlBuilder(fileNum As Integer) As String()

'
' urlBuilder Macro
' Replace values in the url to navigate
' Returns an array of strings with each file to explore
' Selection.End(xlDown).Select
    Dim i As Integer, arreglo() As String, Path As String
    
    ReDim arreglo(fileNum - 1)
    Path = "http://sacnte335/Reports/Pages/Report.aspx?ItemPath=%2f"
    
    For i = 0 To fileNum - 1
        arreglo(i) = Cells(i + 2, 5).Value
        arreglo(i) = Replace(arreglo(i), "/", "%2f")
        arreglo(i) = Path & Replace(arreglo(i), " ", "+")
        
    Next i
    
    urlBuilder = arreglo
    
End Function
