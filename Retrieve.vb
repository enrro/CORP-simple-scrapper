Public Function appendURL(folder As String)
    Dim url As String, delimiter As String
    delimiter = "%2f"

    url = "http://sacnt685/Reports/Pages/Folder.aspx?ItemPath=%2feBPM" & folder & "&ViewMode=Detail"
    
    appendURL = url
End Function
Public Function spacesToPluses(x As String) As String

        

End Function

Public Sub calculateQuerie(url As String)


    With ActiveSheet.QueryTables.Add(Connection:= _
        url _
        , Destination:=Range("$A$1"))
        .Name = "Folder.aspx?ItemPath=%2feBPM&ViewMode=Detail_1"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .WebSelectionType = xlSpecifiedTables
        .WebFormatting = xlWebFormattingNone
        .WebTables = """ui_tblHeader"""
        .WebPreFormattedTextToColumns = True
        .WebConsecutiveDelimitersAsOne = True
        .WebSingleBlockTextImport = False
        .WebDisableDateRecognition = False
        .WebDisableRedirections = False
        .Refresh BackgroundQuery:=False
    End With
End Sub
Public Sub removeExtraData()
    
End Sub
Sub Macro1()
'
' Macro1 Macro
' scraping
'

'
    
    Dim folderNum As Integer
    Call calculateQuerie(appendURL(""))
    folderNum = ActiveSheet.Range("d2", ActiveSheet.Range("d2").End(xlDown)).Rows.Count
    
    Cells(1, 1) = folderNum
    
    

End Sub




