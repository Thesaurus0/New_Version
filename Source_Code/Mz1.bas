Attribute VB_Name = "Mz1"
Sub 宏1()
Attribute 宏1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 宏1 宏
'

'
    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;F:\Github_Local_Repository\Pharmacy_Excel_Tool_Macro\Input_Files\10月流向\国盈10月.csv" _
        , Destination:=Range("$A$1"))
        .CommandType = 0
        .Name = "国盈10月"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 936
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = True
        .TextFileTabDelimiter = False
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileOtherDelimiter = ","
        .TextFileColumnDataTypes = Array(5, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
End Sub
