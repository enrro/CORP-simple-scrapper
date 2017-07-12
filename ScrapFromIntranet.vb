Option Explicit

Sub Macro2()
'
'
    
    Dim appIE As Object
    Dim Document As htmlDocument
    Dim allRowOfData As Object
    Dim Elements As IHTMLElementCollection
    Dim Element As IHTMLElement
    Dim fileNum As Integer
    Dim urlArray() As String
    Dim i As Integer
    Dim errorChain As String
    
    fileNum = WorksheetFunction.CountA(Worksheets(1).Columns(5)) - 1
    urlArray = urlBuilder(fileNum)
    
    For i = 0 To fileNum - 1
    
        Set appIE = New InternetExplorerMedium
        With appIE
            .Navigate urlArray(i)
            .Visible = True
        End With
        
        Do While appIE.Busy
            DoEvents
        Loop
        
        Set Document = appIE.Document
        
        'Dim myValue As String: myValue = allRowOfData.Cells(7).innerHTML
        Set Elements = Document.getElementsByTagName("ul")
        For Each Element In Elements
            
            errorChain = errorChain & Element.innerText
            
        Next Element
        Debug.Print errorChain
        Sheet1.Cells(i + 2, 6).Value = errorChain
        
        
        appIE.Quit
        Set appIE = Nothing
    Next i
    Debug.Print "fin"
    
    
End Sub