Option Explicit

Sub Macro2()
'
'
    
    Dim appIE As Object
    Dim Document As htmlDocument
    Dim Elements As IHTMLElementCollection
    Dim Element As IHTMLElement
    Dim fileNum As Integer
    Dim urlArray() As String
    Dim i As Integer
    Dim errorChain As String
    
    fileNum = WorksheetFunction.CountA(Worksheets(1).Columns(5)) - 1
    urlArray = urlBuilder(fileNum)
    Set appIE = New InternetExplorerMedium
    For i = 0 To fileNum - 1
    
        
        With appIE
            .Navigate urlArray(i)
            .Visible = False
        End With
        
        Do While (appIE.Busy Or appIE.readyState <> 4)
            DoEvents
        Loop
        
        Set Document = appIE.Document
        
        errorChain = ""
        Set Elements = Document.getElementsByTagName("ul")
        For Each Element In Elements
            
            errorChain = errorChain & Element.innerText
            
        Next Element
        Debug.Print i
        Sheet1.Cells(i + 2, 6).Value = errorChain
        
    Next i
    appIE.Quit
    Set appIE = Nothing
    Debug.Print "fin"
    
    
End Sub