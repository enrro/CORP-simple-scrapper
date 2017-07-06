Public Function spacesToPluses(x As String) As String

        Area = x * y

End Function

Public Function calculateQuerie(url As String) As String
    calculateQuerie = ""
End Function

Public Sub removeExtraData()
    
End Sub
Sub Macro1()
'
' Macro1 Macro
' scraping
'

'

    Dim url As String, delimiter As String, folder As String, folderNum As Integer
    delimiter = "%2f"
    folder = ""
    url = "http://sacnt685/Reports/Pages/Folder.aspx?ItemPath=%2feBPM" & folder & "&ViewMode=Detail"


    With ActiveSheet.QueryTables.Add(Connection:= _
        "URL;http://sacnt685/Reports/Pages/Folder.aspx?ItemPath=%2feBPM&ViewMode=Detail" _
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


    folderNum = ActiveSheet.Range("d2", ActiveSheet.Range("d2").End(xlDown)).Rows.Count
    
    Cells(1, 1) = folderNum
    
    

End Sub




