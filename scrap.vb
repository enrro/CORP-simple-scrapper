Sub Macro2()
'
' Macro2 Macro
'

'
    Dim appIE As InternetExplorer
    Dim Document As htmlDocument
    Dim Elements As IHTMLElementCollection
    Dim Element As IHTMLElement
    
    Set appIE = New InternetExplorer
    
    With appIE
        .Navigate "http://sacnte335/Reports/Pages/Report.aspx?ItemPath=%2fAP_Reporting%2fAP_Weekly_Tracker_v1&ViewMode=Detail"
        .Visible = True
    End With
    
    Do While appIE.Busy And Not appIE.readyState = READYSTATE_COMPLETE
        DoEvents
    Loop
    
    Set Document = appIE.Document
    
    Set Elements = Document.getElementsByTagName("ul")
        For Each Element In Elements
        Debug.Print Element.innerText
    Next Element
    
    Set Document = Nothing
    Set appIE = Nothing
End Sub