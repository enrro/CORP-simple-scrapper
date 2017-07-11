Sub Macro2()
'
'
    
    Dim appIE As Object
    Dim Document As htmlDocument
    Dim allRowOfData As Object
    Dim Elements As IHTMLElementCollection
    Dim Element As IHTMLElement
    
    Set appIE = New InternetExplorerMedium
    With appIE
        .Navigate "http://sacnte335/Reports/Pages/Report.aspx?ItemPath=%2fAP_Reporting%2fAP_Weekly_Tracker"
        .Visible = True
    End With
    
    Do While appIE.Busy
        DoEvents
    Loop
    
    Set Document = appIE.Document
    
    'Dim myValue As String: myValue = allRowOfData.Cells(7).innerHTML
    Set Elements = Document.getElementsByTagName("ul")
        For Each Element In Elements
        Debug.Print Element.innerText
    Next Element
   
    
    
    appIE.Quit
    Set appIE = Nothing
    'Debug.Print "my value is"
    
End Sub