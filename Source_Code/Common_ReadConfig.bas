Attribute VB_Name = "Common_ReadConfig"
Option Explicit
Option Base 1

Enum InputFile
    [_first] = 1
    ReportID = 1
    FileTag = 2
    FilePath = 3
    Source = 4
    ReLoadOrNot = 5
    FileSpecTag = 6
    Env = 7
    DefaultSheet = 8
    PivotTableTag = 9
    RowNo = 10
    [_last] = 10
End Enum

Function fFindAllColumnsIndexByColNames(rngToFindIn As Range, arrColsName, ByRef arrColsIndex() _
                                , Optional ByRef alHeaderAtRow As Long, Optional bReturnLetter As Boolean = False)
    If fArrayIsEmptyOrNoData(arrColsName) Then fErr "arrColsName is empty."
    If fArrayHasBlankValue(arrColsName) Then fErr "arrColsName has blank element." & vbCr & Join(arrColsName, vbCr)
    If fArrayHasDuplicateElement(arrColsName) Then fErr "arrColsName has duplicate element."
    
    ReDim arrColsIndex(LBound(arrColsName) To UBound(arrColsName))
    
    Dim lColAtRow As Long
    Dim lEachCol As Long
    Dim sEachColName As String
    Dim rngFound As Range
    
    lColAtRow = 0
    For lEachCol = LBound(arrColsName) To UBound(arrColsName)
        sEachColName = Trim(arrColsName(lEachCol))
        sEachColName = Replace(sEachColName, "*", "~*")
        
        Set rngFound = fFindInWorksheet(rngToFindIn, sEachColName)
        
        If lColAtRow <> 0 Then
            If lColAtRow <> rngFound.Row Then
                fErr "Columns are not at the same row."
            End If
        Else
            lColAtRow = rngFound.Row
        End If
        
        If bReturnLetter Then
            arrColsIndex(lEachCol) = fNum2Letter(rngFound.Column)
        Else
            arrColsIndex(lEachCol) = rngFound.Column
        End If
    Next
    
    alHeaderAtRow = lColAtRow
    Set rngFound = Nothing
End Function
'
'Function fValidateDuplicateKeys(arrConfigData(), arrColsIndex(), arrKeyCols, lHeaderAtRow As Long, lStartCol As Long)
'    If fArrayIsEmptyOrNoData(arrKeyCols) Then Exit Function
'
'    Dim lEachRow As Long
'    Dim lEachCol As Long
'    Dim i As Long
'    Dim sKeyStr As String
'    Dim dict As New Dictionary
'
'    For lEachRow = LBound(arrConfigData, 1) To UBound(arrConfigData, 1)
'        sKeyStr = ""
'
'        For i = LBound(arrKeyCols) To UBound(arrKeyCols)
'            'lEachCol = arrColsIndex(arrKeyCols(i) - 1)
'            lEachCol = arrColsIndex(arrKeyCols(i))
'            sKeyStr = sKeyStr & Trim(CStr(arrConfigData(lEachRow, lEachCol)))
'        Next
'
'        If dict.Exists(sKeyStr) Then
'            fErr "Duplicate key " & sKeyStr & " was found " & vbCr & "at row: " & (lHeaderAtRow + lEachRow) _
'                     & ", column: " & fNum2Letter((lStartCol + lEachCol))
'        Else
'            dict.Add sKeyStr, 0
'        End If
'    Next
'
'    Set dict = Nothing
'End Function

Function fValidateDuplicateKeysForConfigBlock(arrConfigData(), arrColsIndex(), arrKeyCols, lHeaderAtRow As Long, lStartCol As Long)
    If fArrayIsEmptyOrNoData(arrKeyCols) Then Exit Function
    
    Dim lEachRow As Long
    Dim lEachCol As Long
    Dim i As Long
    Dim sKeyStr As String
    Dim dict As New Dictionary
    
    For lEachRow = LBound(arrConfigData, 1) To UBound(arrConfigData, 1)
        sKeyStr = ""
        
        For i = LBound(arrKeyCols) To UBound(arrKeyCols)
            'lEachCol = arrColsIndex(arrKeyCols(i) - 1)
            lEachCol = arrColsIndex(arrKeyCols(i))
            sKeyStr = sKeyStr & Trim(CStr(arrConfigData(lEachRow, lEachCol)))
        Next
        
        If dict.Exists(sKeyStr) Then
            fErr "Duplicate key " & sKeyStr & " was found " & vbCr & "at row: " & (lHeaderAtRow + lEachRow) _
                     & ", column: " & fNum2Letter((lStartCol + lEachCol))
        Else
            dict.Add sKeyStr, 0
        End If
    Next
    
    Set dict = Nothing
End Function

'Function fReadConfigBlockToArrayValidated(asTag As String, rngToFindIn As Range, arrColsName
Function fReadConfigBlockToArrayValidated(asTag As String, shtParam As Worksheet, arrColsName _
                                , Optional arrKeyCols _
                                , Optional ByRef lConfigStartRow As Long _
                                , Optional ByRef lConfigStartCol As Long _
                                , Optional ByRef lConfigEndRow As Long _
                                , Optional ByRef lOutConfigHeaderAtRow As Long _
                                , Optional abNoDataConfigThenError As Boolean = False _
                                , Optional bNetValues As Boolean = True) As Variant
    'arrKeyCols:  array(1, 2, 3, 5), or unnecessary: array()
    Dim arrConfigData()
    Dim arrColsIndex()
    Dim arrOut()

    Call fReadConfigBlockToArray(asTag:=asTag, shtParam:=shtParam, arrColsName:=arrColsName _
                                , arrConfigData:=arrConfigData _
                                , arrColsIndex:=arrColsIndex _
                                , lConfigStartRow:=lConfigStartRow _
                                , lConfigStartCol:=lConfigStartCol _
                                , lConfigEndRow:=lConfigEndRow _
                                , lOutConfigHeaderAtRow:=lOutConfigHeaderAtRow _
                                , abNoDataConfigThenError:=abNoDataConfigThenError _
                                )
    
    If fArrayIsEmptyOrNoData(arrConfigData) Then GoTo exit_fun
    
    'Call fValidateDuplicateKeys(arrConfigData, arrColsIndex, arrKeyCols, lOutConfigHeaderAtRow, lConfigStartCol)
    
    If bNetValues Then
        ReDim arrOut(LBound(arrConfigData, 1) To UBound(arrConfigData, 1), 1 To UBound(arrColsIndex) - LBound(arrColsIndex) + 1)
        
        Dim lEachRow As Long
        Dim lEachCol As Long
        Dim i As Long
        
        For lEachRow = LBound(arrConfigData, 1) To UBound(arrConfigData, 1)
            'i = LBound(arrColsIndex) + 1
            i = LBound(arrColsIndex)
            For lEachCol = LBound(arrColsIndex) To UBound(arrColsIndex)
                arrOut(lEachRow, i) = arrConfigData(lEachRow, arrColsIndex(lEachCol))
                i = i + 1
            Next
        Next
    End If
exit_fun:
    Erase arrColsIndex
    
    If bNetValues Then
        fReadConfigBlockToArrayValidated = arrOut
    Else
        fReadConfigBlockToArrayValidated = arrConfigData
    End If
    
    Erase arrConfigData
    Erase arrOut
End Function
'Function fReadConfigBlockToArrayNet(asTag As String, rngToFindIn As Range, arrColsName()
Function fReadConfigBlockToArrayNet(asTag As String, shtParam As Worksheet, arrColsName() _
                                , Optional ByRef lConfigStartRow As Long _
                                , Optional ByRef lConfigStartCol As Long _
                                , Optional ByRef lConfigEndRow As Long _
                                , Optional ByRef lOutConfigHeaderAtRow As Long _
                                , Optional abNoDataConfigThenError As Boolean = False _
                                ) As Variant
    Dim arrOut()
    Dim arrColsIndex()
    Dim arrConfigData()

    Call fReadConfigBlockToArray(asTag:=asTag, shtParam:=shtParam, arrColsName:=arrColsName _
                                , arrConfigData:=arrConfigData _
                                , arrColsIndex:=arrColsIndex _
                                , lConfigStartRow:=lConfigStartRow _
                                , lConfigStartCol:=lConfigStartCol _
                                , lConfigEndRow:=lConfigEndRow _
                                , lOutConfigHeaderAtRow:=lOutConfigHeaderAtRow _
                                , abNoDataConfigThenError:=abNoDataConfigThenError _
                                )
    If fArrayIsEmptyOrNoData(arrConfigData) Then GoTo exit_fun
    
    ReDim arrOut(LBound(arrConfigData, 1) To UBound(arrConfigData, 1), LBound(arrColsIndex) To UBound(arrColsIndex))
    
    Dim lEachRow As Long
    Dim lEachCol As Long
    Dim i As Long
    
    For lEachRow = LBound(arrConfigData, 1) To UBound(arrConfigData, 1)
        i = LBound(arrColsIndex)
        For lEachCol = LBound(arrColsIndex) To UBound(arrColsIndex)
            arrOut(lEachRow, i) = arrConfigData(lEachRow, arrColsIndex(lEachCol))
            i = i + 1
        Next
    Next
exit_fun:
    Erase arrColsIndex
    Erase arrConfigData
    fReadConfigBlockToArrayNet = arrOut
    Erase arrOut
End Function
'Function fReadConfigBlockToArray(asTag As String, rngToFindIn As Range, arrColsName
Function fReadConfigBlockToArray(asTag As String, shtParam As Worksheet, arrColsName _
                                , ByRef arrConfigData() _
                                , ByRef arrColsIndex() _
                                , Optional ByRef lConfigStartRow As Long _
                                , Optional ByRef lConfigStartCol As Long _
                                , Optional ByRef lConfigEndRow As Long _
                                , Optional ByRef lOutConfigHeaderAtRow As Long _
                                , Optional abNoDataConfigThenError As Boolean = False _
                                )
    arrConfigData = Array()
    
    'Dim shtConfig As Worksheet
    'Set shtConfig = rngToFindIn.Parent
    Dim rngToFindIn As Range
    
    Call fReadConfigBlockStartEnd(asTag, shtParam, lConfigStartRow, lConfigStartCol, lConfigEndRow)
    
    If lConfigEndRow < lConfigStartRow + 1 Then
        If abNoDataConfigThenError Then
            fErr "No data is configured under tag " & asTag & " in sheet " & shtParam.Name & vbCr _
                    & "You must leave at least one blank line after the tag."
        End If
    End If
    
    Set rngToFindIn = fGetRangeByStartEndPos(shtParam, lConfigStartRow, lConfigStartCol, lConfigEndRow, Columns.Count)
    Call fFindAllColumnsIndexByColNames(rngToFindIn, arrColsName, arrColsIndex, lOutConfigHeaderAtRow)
    
    Dim lColsMinCol As Long
    Dim lColsMaxCol As Long
    
    lColsMinCol = Application.WorksheetFunction.Min(arrColsIndex)
    lColsMaxCol = Application.WorksheetFunction.Max(arrColsIndex)
    
    lConfigEndRow = fGetValidMaxRowOfRange(fGetRangeByStartEndPos(shtParam, lConfigStartRow, lConfigStartCol, lConfigEndRow, lColsMaxCol))
    
    If lConfigEndRow > lOutConfigHeaderAtRow Then
        arrConfigData = fReadRangeDatatoArrayByStartEndPos(shtParam, lOutConfigHeaderAtRow + 1, lColsMinCol, lConfigEndRow, lColsMaxCol)
    End If
    
    lConfigStartCol = lColsMinCol
    
    Dim lEachCol As Long
    'change 10, 15, 20, to 1, 6, 11
    For lEachCol = UBound(arrColsIndex) To LBound(arrColsIndex) Step -1
        arrColsIndex(lEachCol) = arrColsIndex(lEachCol) - lColsMinCol + 1
    Next
    
    'Set shtConfig = Nothing
End Function

Function fFindRageOfFileSpecConfigBlock(sFileSpecTag As String) As Range
    Dim rngConfigBlock As Range
    Dim lConfigStartRow As Long
    Dim lConfigStartCol As Long
    Dim lConfigEndRow As Long
    
    Call fReadConfigBlockStartEnd(sFileSpecTag, shtFileSpec, lConfigStartRow, lConfigStartCol, lConfigEndRow)
    Set fFindRageOfFileSpecConfigBlock = fGetRangeByStartEndPos(shtFileSpec, lConfigStartRow, lConfigStartCol, lConfigEndRow, Columns.Count)
End Function
'Function fReadConfigBlockStartEnd(sFileSpecTag As String, rngToFindIn As Range
Function fReadConfigBlockStartEnd(sFileSpecTag As String, shtParam As Worksheet _
                                , ByRef lOutBlockStartRow As Long _
                                , ByRef lOutBlockStartCol As Long _
                                , ByRef lOutBlockEndRow As Long)
    
    'Dim shtSource As Worksheet
    Dim lMaxRow As Long
    Dim rngTagFound As Range
    Dim lTagRow As Long
    Dim lTagCol As Long
    
    'Set shtSource = rngToFindIn.Parent
    lMaxRow = fGetValidMaxRow(shtParam)
    
    Set rngTagFound = fFindInWorksheet(shtParam.Cells, sFileSpecTag)
    lTagRow = rngTagFound.Row
    lTagCol = rngTagFound.Column
    
    Set rngTagFound = fFindInWorksheet(fGetRangeByStartEndPos(shtParam, lTagRow + 1, lTagCol, lMaxRow, lTagCol) _
                                    , "[*]", False, True)
    If rngTagFound Is Nothing Then
        lOutBlockEndRow = lMaxRow
    Else
        lOutBlockEndRow = rngTagFound.Row - 1
    End If
    
    lOutBlockStartRow = lTagRow + 1
    lOutBlockStartCol = lTagCol
    
    'Set shtSource = Nothing
    Set rngTagFound = Nothing
End Function

Function fReadConfigInputFiles(Optional asReportID As String = "")
    If asReportID = "" Then asReportID = gsRptID
    
    Dim asTag As String
    Dim arrColsName()
    Dim rngToFindIn As Range
    Dim arrConfigData()
    Dim lConfigStartRow As Long
    Dim lConfigStartCol As Long
    Dim lConfigEndRow As Long
    Dim lConfigHeaderAtRow As Long
                                
    asTag = "[Input Files]"
    ReDim arrColsName(InputFile.ReportID To InputFile.PivotTableTag)
    
    arrColsName(InputFile.ReportID) = "Report ID"
    arrColsName(InputFile.FileTag) = "File Tag"
    arrColsName(InputFile.FilePath) = "File Full Path"
    arrColsName(InputFile.Source) = "Source"
    arrColsName(InputFile.ReLoadOrNot) = "When Data Already Loaded To Sheet"
    arrColsName(InputFile.FileSpecTag) = "File Spec Tag"
    arrColsName(InputFile.Env) = "DEV/UAT/PROD"
    arrColsName(InputFile.DefaultSheet) = "Which Sheet To Import"
    arrColsName(InputFile.PivotTableTag) = "Pivot Table Tag To Be Created From This Data Source"
     
    arrConfigData = fReadConfigBlockToArrayNet(asTag:=asTag, shtParam:=shtSysConf _
                                , arrColsName:=arrColsName _
                                , lConfigStartRow:=lConfigStartRow _
                                , lConfigStartCol:=lConfigStartCol _
                                , lConfigEndRow:=lConfigEndRow _
                                , lOutConfigHeaderAtRow:=lConfigHeaderAtRow _
                                , abNoDataConfigThenError:=True _
                                )
    
    Call fValidateDuplicateInArray(arrConfigData, Array(InputFile.ReportID, InputFile.FileTag, InputFile.Env), False _
        , shtSysConf, lConfigHeaderAtRow, lConfigStartCol, "Report ID + File Tag")
    Call fValidateBlankInArray(arrConfigData, InputFile.ReportID, shtSysConf, lConfigHeaderAtRow, lConfigStartCol, "Report ID")
    Call fValidateBlankInArray(arrConfigData, InputFile.FileTag, shtSysConf, lConfigHeaderAtRow, lConfigStartCol, "File Tag")
    Call fValidateBlankInArray(arrConfigData, InputFile.Source, shtSysConf, lConfigHeaderAtRow, lConfigStartCol, "Source")
    Call fValidateBlankInArray(arrConfigData, InputFile.Env, shtSysConf, lConfigHeaderAtRow, lConfigStartCol, "DEV/UAT/PROD")
    
    Dim lEachRow As Long
    Dim lActualRow As Long
    Dim sRptNameStr As String
    Dim sValueStr As String
    Dim sFileTag As String
    Dim sSource As String
    Dim sEnv As String
    Dim sShtToImport  As String
    Dim sFileName  As String
    
    Dim sPos As String
    sPos = vbCr & vbCr & "Row : $ACTUAL_ROW$" & vbCr & "Column: "
    
    Set gDictInputFiles = New Dictionary
    
    For lEachRow = LBound(arrConfigData, 1) To UBound(arrConfigData, 1)
        If fArrayRowIsBlankHasNoData(arrConfigData, lEachRow) Then GoTo next_row
        
        sRptNameStr = DELIMITER & Trim(arrConfigData(lEachRow, InputFile.ReportID)) & DELIMITER
        If InStr(sRptNameStr, DELIMITER & asReportID & DELIMITER) <= 0 Then GoTo next_row
        
        sEnv = Trim(arrConfigData(lEachRow, InputFile.Env))
        
        If Not (sEnv = gsEnv Or sEnv = "SHARED") Then GoTo next_row
        
        lActualRow = lConfigHeaderAtRow + lEachRow
        
        sSource = Trim(arrConfigData(lEachRow, InputFile.Source))
        'sShtToImport = arrConfigData(lEachRow, InputFile.DefaultSheet)
        sFileName = Trim(arrConfigData(lEachRow, InputFile.FilePath))
        
        If sSource = "READ_PRE_EXISTING_SHEET" Then
            If Not fSheetExists(sFileName) Then
                fErr "READ_PRE_EXISTING_SHEET, but sheet specified does not exists, you may have to run the previous steps first:" _
                & sFileName & Replace(sPos, "$ACTUAL_ROW$", lActualRow) & arrColsName(InputFile.DefaultSheet)
            End If
        ElseIf sSource = "READ_PRE_EXISTING_SHEET" Then
        End If
        
        sFileTag = Trim(arrConfigData(lEachRow, InputFile.FileTag))
        sValueStr = fComposeStrForInputFile(arrConfigData, lEachRow)
        
        gDictInputFiles.Add sFileTag, sValueStr
        Call fUpdateDictionaryItemValueForDelimitedElement(gDictInputFiles, sFileTag, InputFile.RowNo - InputFile.FileTag, lActualRow)
next_row:
    Next
    
    Erase arrConfigData
    Erase arrColsName
End Function

Function fComposeStrForInputFile(arrConfigData, lEachRow As Long) As String
    Dim sOut As String
    Dim i As Integer
    
    For i = InputFile.FilePath To InputFile.PivotTableTag
        sOut = sOut & DELIMITER & Trim(arrConfigData(lEachRow, i))
    Next
    
    fComposeStrForInputFile = Right(sOut, Len(sOut) - 1)
End Function


Function fReadConfigWholeColsToDictionary(shtConfig As Worksheet, asTag As String, asKeyNotNullCol As String, asRtnCol As String) As Dictionary
    If fZero(asTag) Or fZero(asKeyNotNullCol) Or fZero(asRtnCol) Then fErr "Wrong param"

    Dim bRtnColIsKeyCol As Boolean
    bRtnColIsKeyCol = (Trim(asKeyNotNullCol) = Trim(asRtnCol))

    Dim arrColNames()
    ReDim arrColNames(0 To 1)
    arrColNames(0) = asKeyNotNullCol
    arrColNames(1) = asRtnCol

    Dim rngToFindIn As Range
    Dim arrConfigData()
    Dim lConfigStartRow As Long
    Dim lConfigStartCol As Long
    Dim lConfigEndRow As Long
    Dim lConfigHeaderAtRow As Long

    arrKeyColsForValidation = Array(1, 2)

'    arrConfigData = fReadConfigBlockToArray(asTag:=asTag, rngToFindIn:=shtConfig.Cells _
'                                , arrColsName:=arrColsName _
'                                , lConfigStartRow:=lConfigStartRow _
'                                , lConfigStartCol:=lConfigStartCol _
'                                , lConfigEndRow:=lConfigEndRow _
'                                , lOutConfigHeaderAtRow:=lConfigHeaderAtRow _
'                                , abNoDataConfigThenError:=True _
'                                )
End Function

'
Function fReadConfigWholeMultipleColsToArray(shtConfig As Worksheet, asTag As String, arrColsName) As Variant
'arrKeyColsForValidation : Array(1, 2, 5)
    Dim asTag As String
    Dim arrColsName()
    Dim rngToFindIn As Range
    Dim arrConfigData()
    Dim lConfigStartRow As Long
    Dim lConfigStartCol As Long
    Dim lConfigEndRow As Long
    Dim lConfigHeaderAtRow As Long

    arrConfigData = fReadConfigBlockToArrayNet(asTag:=asTag, shtParam:=shtSysConf _
                                , arrColsName:=arrColsName _
                                , lConfigStartRow:=lConfigStartRow _
                                , lConfigStartCol:=lConfigStartCol _
                                , lConfigEndRow:=lConfigEndRow _
                                , lOutConfigHeaderAtRow:=lConfigHeaderAtRow _
                                , abNoDataConfigThenError:=True _
                                )
    Erase arrColsName
    fReadConfigWholeMultipleColsToArray = arrConfigData
    Erase arrConfigData
End Function

'
Function fGetReadConfigWholePairColsValueAsArray(shtConfig As Worksheet, asTag As String, arrFetchCols, Optional arrKeyColsForValidation) As Variant
    Dim dict As Dictionary
    Dim arrOut()
    
    Set dict = fGetReadConfigWholePairColsValueAsDictionary()
    
    Call fCopyDictionaryItems2Array(dict, arrOut)
    
    Set dict = Nothing
    
    fGetReadConfigWholePairColsValueAsArray = arrOut
    Erase arrOut
End Function

Function fGetReadConfigWholePairColsValueAsDictionary(shtConfig As Worksheet, asTag As String _
                    , asKeyNotNullCol As String, asRtnCol As String) As Dictionary
    If fZero(asKeyNotNullCol) Or fZero(asRtnCol) Then fErr "Wrong param"
    
    Dim bRtnColIsKeyCol As Boolean
    bRtnColIsKeyCol = (Trim(asKeyNotNullCol) = Trim(asRtnCol))
    
    Dim asTag As String
    Dim arrColsName()
    Dim rngToFindIn As Range
    Dim arrConfigData()
    Dim lConfigStartRow As Long
    Dim lConfigStartCol As Long
    Dim lConfigEndRow As Long
    Dim lConfigHeaderAtRow As Long

    ReDim arrColsName(1 To 2)
    arrColsName(1) = Trim(asKeyNotNullCol)
    arrColsName(2) = Trim(asRtnCol)
     
    arrConfigData = fReadConfigBlockToArrayNet(asTag:=asTag, shtParam:=shtSysConf _
                                , arrColsName:=arrColsName _
                                , lConfigStartRow:=lConfigStartRow _
                                , lConfigStartCol:=lConfigStartCol _
                                , lConfigEndRow:=lConfigEndRow _
                                , lOutConfigHeaderAtRow:=lConfigHeaderAtRow _
                                , abNoDataConfigThenError:=True _
                                )
    Call fValidateDuplicateInArray(arrConfigData, 1, False, shtSysConf, lConfigHeaderAtRow, lConfigStartCol, asKeyNotNullCol)
    If Not bRtnColIsKeyCol Then
        Call fValidateDuplicateInArray(arrConfigData, 2, False, shtSysConf, lConfigHeaderAtRow, lConfigStartCol, asRtnCol)
    End If
    
'    Call fValidateBlankInArray(arrConfigData, TechTag.Report_ID, shtstaticdata, lConfigHeaderAtRow, lConfigStartCol, "TechTag ID")

    Set fReadConfigTechTagList = fReadArray2DictionaryWithMultipleColsCombined(arrConfigData, TechTag.REPORT_ID _
            , Array(TechTag.ID, TechTag.Name, TechTag.Commission, TechTag.Selected), DELIMITER)
    Erase arrColsName
    Erase arrConfigData
    
End Function

'Function fGetReadConfigWholeSingleColValueAsArray(shtConfig As Worksheet, asTag As String, arrFetchCols, Optional arrKeyColsForValidation) As Variant
'    Dim dict As Dictionary
'    Dim arrOut()
'
'    Set dict = fGetReadConfigWholeSingleColValueAsDictionary()
'
'    Call fCopyDictionaryItems2Array(dict, arrOut)
'
'    Set dict = Nothing
'
'    fGetReadConfigWholePairColsValueAsArray = arrOut
'    Erase arrOut
'End Function
Function fGetReadConfigWholeSingleColValueAsArray(shtConfig As Worksheet, asTag As String _
                    , asColName As String _
                    , Optional IgnoreBlankKeys As Boolean = False _
                    , Optional WhenKeyIsDuplicateError As Boolean = True) As Variant
    If fZero(asColName) Then fErr "Wrong param"
    
    'Dim asTag As String
    Dim arrColsName()
    Dim rngToFindIn As Range
    Dim arrConfigData()
    Dim lConfigStartRow As Long
    Dim lConfigStartCol As Long
    Dim lConfigEndRow As Long
    Dim lConfigHeaderAtRow As Long

    ReDim arrColsName(1 To 1)
    arrColsName(1) = Trim(asColName)

    arrConfigData = fReadConfigBlockToArrayNet(asTag:=asTag, shtParam:=shtConfig _
                                , arrColsName:=arrColsName _
                                , lConfigStartRow:=lConfigStartRow _
                                , lConfigStartCol:=lConfigStartCol _
                                , lConfigEndRow:=lConfigEndRow _
                                , lOutConfigHeaderAtRow:=lConfigHeaderAtRow _
                                , abNoDataConfigThenError:=True _
                                )
    Erase arrColsName
    
    If WhenKeyIsDuplicateError Then
        Call fValidateDuplicateInArray(arrConfigData, 1, True, shtConfig, lConfigHeaderAtRow, lConfigStartCol, asColName)
    End If

    If Not IgnoreBlankKeys Then
        Call fValidateBlankInArray(arrConfigData, 1, shtConfig, lConfigHeaderAtRow, lConfigStartCol, asColName)
    End If

    Dim dict As Dictionary
    Set dict = fRadArray2DictionaryOnlyKeys(arrConfigData, 1, IgnoreBlankKeys, WhenKeyIsDuplicateError)
    
    Dim arrOut()
    Call fCopyDictionaryKeys2Array(dict, arrOut)
    
    Erase arrConfigData
    Set dict = Nothing
    
    fGetReadConfigWholeSingleColValueAsArray = arrOut
End Function

Function fReadSysConfig_InputTxtSheetFile(Optional asReportID As String = "")
    If asReportID = "" Then asReportID = gsRptID
    
    Call fReadConfigInputFiles(asReportID)
    Call fReadTxtFileImportConfig
    
    Dim i As Long
    Dim sFileTag As String
    Dim sSource As String
    Dim sDependantFileTag As String
    Dim sDependantFileOrSheet As String
    
    For i = 0 To gDictInputFiles.Count - 1
        sFileTag = gDictInputFiles.Keys(i)
        sSource = fGetInputFileSourceType(sFileTag)
        
        If sSource = "READ_PREV_STEP_OUTPUT_SHEET" Or sSource = "READ_PREV_STEP_OUTPUT_FILE" Then
            sDependantFileTag = fGetInputFileFileName(sFileTag)
            
            sDependantFileOrSheet = fGetReneratedReport(sDependantFileTag)
            
            If sSource = "READ_PREV_STEP_OUTPUT_SHEET" Then
                If Not fSheetExists(sDependantFileOrSheet) Then
                    fErr "Dependant ourput sheet as below does not exist, you may have to run the previous steps first." _
                        & vbCr & "File Tag: " & sFileTag _
                        & vbCr & "sDependantFileTag Tag: " & sDependantFileTag _
                        & vbCr & "sDependantFileOrSheet: " & sDependantFileOrSheet
                End If
            Else
                If Not fFileExists(sDependantFileOrSheet) Then
                    fErr "Dependant ourput FILE as below does not exist, you may have to run the previous steps first." _
                        & vbCr & "File Tag: " & sFileTag _
                        & vbCr & "sDependantFileTag Tag: " & sDependantFileTag _
                        & vbCr & "sDependantFileOrSheet: " & sDependantFileOrSheet
                End If
            End If
            
            Call fUpdateGDictInputFile_FileName(sFileTag, sDependantFileOrSheet)
        End If
    Next
    
    Call fCrossValidateInputFileTxtSheetFile
    
End Function

Function fReadTxtFileImportConfig(Optional asReportID As String = "")
    If asReportID = "" Then asReportID = gsRptID
    
    Dim asTag As String
    Dim arrColsName()
    Dim rngToFindIn As Range
    Dim arrConfigData()
    Dim lConfigStartRow As Long
    Dim lConfigStartCol As Long
    Dim lConfigEndRow As Long
    Dim lConfigHeaderAtRow As Long

    asTag = "[Input Txt Files Import Config]"
    
    Const REPORT_ID = 1
    Const FILE_TAG = 2
    Const COL_DELIMITER = 3
    Const PLATFORM = 4
    ReDim arrColsName(REPORT_ID To PLATFORM)
    
    arrColsName(REPORT_ID) = "Report ID"
    arrColsName(FILE_TAG) = "File Tag"
    arrColsName(COL_DELIMITER) = "Column Delimiter"
    arrColsName(PLATFORM) = "TextFilePlatForm"
     
    arrConfigData = fReadConfigBlockToArrayNet(asTag:=asTag, shtParam:=shtSysConf _
                                , arrColsName:=arrColsName _
                                , lConfigStartRow:=lConfigStartRow _
                                , lConfigStartCol:=lConfigStartCol _
                                , lConfigEndRow:=lConfigEndRow _
                                , lOutConfigHeaderAtRow:=lConfigHeaderAtRow _
                                , abNoDataConfigThenError:=True _
                                )
    
    Call fValidateDuplicateInArray(arrConfigData, Array(REPORT_ID, FILE_TAG), False _
        , shtSysConf, lConfigHeaderAtRow, lConfigStartCol, "Report ID + File Tag")
    Call fValidateBlankInArray(arrConfigData, REPORT_ID, shtSysConf, lConfigHeaderAtRow, lConfigStartCol, "Report ID")
    Call fValidateBlankInArray(arrConfigData, FILE_TAG, shtSysConf, lConfigHeaderAtRow, lConfigStartCol, "File Tag")
    Call fValidateBlankInArray(arrConfigData, COL_DELIMITER, shtSysConf, lConfigHeaderAtRow, lConfigStartCol, "Column Delimiter")
    Call fValidateBlankInArray(arrConfigData, PLATFORM, shtSysConf, lConfigHeaderAtRow, lConfigStartCol, "TextFilePlatForm")
    
    Dim lEachRow As Long
    Dim lActualRow As Long
    Dim sRptNameStr As String
    Dim sValueStr As String
    Dim sFileTag As String
    
    Dim sPos As String
    sPos = vbCr & vbCr & "Row : $ACTUAL_ROW$" & vbCr & "Column: "
    
    Set gDictTxtFileSpec = New Dictionary
    
    For lEachRow = LBound(arrConfigData, 1) To UBound(arrConfigData, 1)
        If fArrayRowIsBlankHasNoData(arrConfigData, lEachRow) Then GoTo next_row
        
        sRptNameStr = DELIMITER & Trim(arrConfigData(lEachRow, REPORT_ID)) & DELIMITER
        If InStr(sRptNameStr, DELIMITER & asReportID & DELIMITER) <= 0 Then GoTo next_row
        
        lActualRow = lConfigHeaderAtRow + lEachRow
        
        sFileTag = Trim(arrConfigData(lEachRow, FILE_TAG))
        
        sValueStr = Trim(arrConfigData(lEachRow, COL_DELIMITER)) & DELIMITER _
                & Trim(arrConfigData(lEachRow, PLATFORM))
        
        gDictTxtFileSpec.Add sFileTag, sValueStr
next_row:
    Next
    
    Erase arrConfigData
    Erase arrColsName
End Function

Function fReadSysConfig_Output(Optional asReportID As String, Optional asRptType As String) As String
    Dim sRptFileName As String
    Dim asFileSpecTag As String
    
    If asReportID = "" Then asReportID = gsRptID
    sRptFileName = fReadConfigOutputFiles(asReportID, asRptType, asFileSpecTag)
    
    If fNzero(asFileSpecTag) Then
        Call fGetOutputReportColsConfig(asFileSpecTag)
    End If
    
    fReadSysConfig_Output = sRptFileName
End Function

Function fReadConfigOutputFiles(Optional asReportID As String = "" _
            , Optional asRptType As String, Optional asFileSpecTag As String) As String
    If asReportID = "" Then asReportID = gsRptID
    
    Dim asTag As String
    Dim arrColsName()
    Dim rngToFindIn As Range
    Dim arrConfigData()
    Dim lConfigStartRow As Long
    Dim lConfigStartCol As Long
    Dim lConfigEndRow As Long
    Dim lConfigHeaderAtRow As Long
                                
    asTag = "[Output Files]"
    
    ReDim arrColsName(1 To 5)
    arrColsName(1) = "Report ID"
    arrColsName(2) = "Output Type"
    arrColsName(3) = "File Full Path or Sheet Name"
    arrColsName(4) = "File Spec Tag"
    arrColsName(5) = "DEV/UAT/PROD"
     
    arrConfigData = fReadConfigBlockToArrayNet(asTag:=asTag, shtParam:=shtSysConf _
                                , arrColsName:=arrColsName _
                                , lConfigStartRow:=lConfigStartRow _
                                , lConfigStartCol:=lConfigStartCol _
                                , lConfigEndRow:=lConfigEndRow _
                                , lOutConfigHeaderAtRow:=lConfigHeaderAtRow _
                                , abNoDataConfigThenError:=True)
    
    Call fValidateDuplicateInArray(arrConfigData, Array(1, 5), False, shtSysConf, lConfigHeaderAtRow, lConfigStartCol, "Report ID")
    Call fValidateBlankInArray(arrConfigData, 1, shtSysConf, lConfigHeaderAtRow, lConfigStartCol, "Output Type")
    Call fValidateBlankInArray(arrConfigData, 2, shtSysConf, lConfigHeaderAtRow, lConfigStartCol, "Output Type")
    Call fValidateBlankInArray(arrConfigData, 5, shtSysConf, lConfigHeaderAtRow, lConfigStartCol, "Output Type")
    
    Dim lEachRow As Long
    Dim lActualRow As Long
    Dim sRptNameStr As String
    Dim sEnv As String
    Dim sRptFileName As String
    
    Dim sPos As String
    sPos = vbCr & vbCr & "Row : $ACTUAL_ROW$" & vbCr & "Column: "
        
    For lEachRow = LBound(arrConfigData, 1) To UBound(arrConfigData, 1)
        If fArrayRowIsBlankHasNoData(arrConfigData, lEachRow) Then GoTo next_row
        
        sRptNameStr = DELIMITER & Trim(arrConfigData(lEachRow, 1)) & DELIMITER
        If InStr(sRptNameStr, DELIMITER & asReportID & DELIMITER) <= 0 Then GoTo next_row
        
        sEnv = Trim(arrConfigData(lEachRow, 5))
        
        If Not (sEnv = gsEnv Or sEnv = "SHARED") Then GoTo next_row
        
        lActualRow = lConfigHeaderAtRow + lEachRow
        asRptType = Trim(arrConfigData(lEachRow, 2))
        sRptFileName = Trim(arrConfigData(lEachRow, 3))
        asFileSpecTag = Trim(arrConfigData(lEachRow, 4))
        
        If asRptType = "SHEET" Then
            If fZero(asFileSpecTag) Then
                fErr "Sheet name cannot be blnak when SHEET specified  " & Replace(sPos, "$ACTUAL_ROW$", lActualRow) & arrColsName(3)
            End If
        End If
next_row:
    Next
    
    Erase arrConfigData
    Erase arrColsName
    
    fReadConfigOutputFiles = sRptFileName
End Function

Function fGetOutputReportColsConfig(asFileSpecTag As String)
    asFileSpecTag = Trim(asFileSpecTag)
    If Not (Left(asFileSpecTag, 1) = "[" And Right(asFileSpecTag, 1) = "]") Then
        fErr "File Spec Tag is incorrect, which should be like [Output Format - xxx]"
    End If
    
    Call fReadOutputFileSpecConfig(asTag:=asFileSpecTag _
                                    , dictColsIndex:=gDictRptColIndex _
                                    , dictColsName:=gDictRptDisplayName _
                                    , dictRawType:=gDictRptRawType _
                                    , dictCellFormat:=gDictRptCellFormat _
                                    , dictDataFormat:=gDictRptDataFormat _
                                    , dictColWidth:=gDictRptColWidth _
                                    , dictColAttr:=gDictRptColAttr _
                                    )
End Function

Function fReadOutputFileSpecConfig(asTag As String _
                                    , ByRef dictColsIndex As Dictionary _
                                    , ByRef dictColsName As Dictionary _
                                    , ByRef dictRawType As Dictionary _
                                    , ByRef dictCellFormat As Dictionary _
                                    , ByRef dictDataFormat As Dictionary _
                                    , ByRef dictColWidth As Dictionary _
                                    , ByRef dictColAttr As Dictionary _
                                    )
    
    'Dim asTag As String
    Dim arrColsName()
    Dim arrColsIndex()
    Dim rngToFindIn As Range
    Dim arrConfigData()
    Dim lConfigStartRow As Long
    Dim lConfigStartCol As Long
    Dim lConfigEndRow As Long
    Dim lConfigHeaderAtRow As Long
    
    ReDim arrColsName(1 To 9)
    arrColsName(1) = "Column Tech Name"
    arrColsName(2) = "Column Display Name"
    arrColsName(3) = "Column Index"
    arrColsName(4) = "Array Index"
    arrColsName(5) = "Raw Data Type"
    arrColsName(6) = "Cell Format"
    arrColsName(7) = "Data Format"
    arrColsName(8) = "Column Width"
    arrColsName(9) = "Column Attr"
     
    Call fReadConfigBlockToArray(asTag:=asTag, shtParam:=shtFileSpec _
                                , arrConfigData:=arrConfigData _
                                , arrColsName:=arrColsName _
                                , arrColsIndex:=arrColsIndex _
                                , lConfigStartRow:=lConfigStartRow _
                                , lConfigStartCol:=lConfigStartCol _
                                , lConfigEndRow:=lConfigEndRow _
                                , lOutConfigHeaderAtRow:=lConfigHeaderAtRow _
                                , abNoDataConfigThenError:=True)
    Call fValidateDuplicateInArray(arrConfigData, arrColsIndex(1), False, shtSysConf, lConfigHeaderAtRow, lConfigStartCol, arrColsName(1))
    Call fValidateDuplicateInArray(arrConfigData, arrColsIndex(2), False, shtSysConf, lConfigHeaderAtRow, lConfigStartCol, arrColsName(2))
    Call fValidateDuplicateInArray(arrConfigData, arrColsIndex(3), True, shtSysConf, lConfigHeaderAtRow, lConfigStartCol, arrColsName(3))
'    Call fValidateBlankInArray(arrConfigData, 1, shtSysConf, lConfigHeaderAtRow, lConfigStartCol, "Output Type")
'    Call fValidateBlankInArray(arrConfigData, 2, shtSysConf, lConfigHeaderAtRow, lConfigStartCol, "Output Type")
'    Call fValidateBlankInArray(arrConfigData, 5, shtSysConf, lConfigHeaderAtRow, lConfigStartCol, "Output Type")
    
    Dim lEachRow As Long
    Dim lActualRow As Long
    Dim sColTechName As String
'    Dim sEnv As String
    Dim sRptFileName As String
    Dim sColWidth As String
    Dim sColLetterIndex As String
    Dim lColLetter2Num As Long
    Dim sColAttr As String
    Dim lColWidth As Double
    
    Dim sPos As String
    sPos = vbCr & vbCr & "Row : $ACTUAL_ROW$" & vbCr & "Column: "
    
    Set dictColsIndex = New Dictionary
    Set dictColsName = New Dictionary
    Set dictRawType = New Dictionary
    Set dictCellFormat = New Dictionary
    Set dictDataFormat = New Dictionary
    'Set dictNumberFormat = New Dictionary
    Set dictColWidth = New Dictionary
    Set dictColAttr = New Dictionary

    For lEachRow = LBound(arrConfigData, 1) To UBound(arrConfigData, 1)
        If fArrayRowIsBlankHasNoData(arrConfigData, lEachRow) Then GoTo next_row
        
        lActualRow = lConfigHeaderAtRow + lEachRow
        
        sColTechName = Trim(arrConfigData(lEachRow, arrColsIndex(1)))
        sColLetterIndex = Trim(arrConfigData(lEachRow, arrColsIndex(3)))
        sColAttr = Trim(arrConfigData(lEachRow, arrColsIndex(9)))
        
        dictColsName.Add sColTechName, Trim(arrConfigData(lEachRow, arrColsIndex(2)))
        dictRawType.Add sColTechName, UCase(Trim(arrConfigData(lEachRow, arrColsIndex(5))))
        dictCellFormat.Add sColTechName, shtFileSpec.Cells(lActualRow, lConfigStartCol + arrColsIndex(6) - 1).Address
        dictDataFormat.Add sColTechName, Trim(arrConfigData(lEachRow, arrColsIndex(7)))
        
        If sColAttr = "NOT_SHOW_UP" Then
            If Len(sColLetterIndex) > 0 Then
                fErr "Col Letter Index should be blank for NOT_SHOW_UP: " & Replace(sPos, "$ACTUAL_ROW$", lActualRow) & arrColsName(3)
            End If
            
            dictColsIndex.Add sColTechName, -1
        Else
            If Len(sColLetterIndex) <= 0 Then
                fErr "Col Letter Index should NOT be blank: " & Replace(sPos, "$ACTUAL_ROW$", lActualRow) & arrColsName(3)
            End If
            
            lColLetter2Num = fLetter2Num(sColLetterIndex)
            
            If lColLetter2Num <= 0 Or lColLetter2Num > Columns.Count Then
                fErr "Col Letter Index is invalid,should be A - XFD: " & Replace(sPos, "$ACTUAL_ROW$", lActualRow) & arrColsName(3)
            End If
            
            dictColsIndex.Add sColTechName, lColLetter2Num
        End If
        
        sColWidth = Trim(arrConfigData(lEachRow, arrColsIndex(8)))
        If Len(sColWidth) > 0 Then
            If Not IsNumeric(sColWidth) Then
                fErr "Col Width should be numeric: " & Replace(sPos, "$ACTUAL_ROW$", lActualRow) & arrColsName(8)
                
                If CDbl(sColWidth) <= 0 Or CDbl(sColWidth) > 255 Then
                    fErr "Col Width should be > 0 and <= 255: " & Replace(sPos, "$ACTUAL_ROW$", lActualRow) & arrColsName(8)
                End If
            End If
            
            lColWidth = CDbl(sColWidth)
        Else
            lColWidth = 0
        End If
        
        dictColWidth.Add sColTechName, lColWidth
        dictColAttr.Add sColTechName, sColAttr
next_row:
    Next
    
    Dim lNextCol As Long
    lNextCol = WorksheetFunction.Max(dictColsIndex.Items)
    
    For lEachRow = 0 To dictColAttr.Count - 1
        If dictColAttr.Items(lEachRow) = "NOT_SHOW_UP" Then
            lNextCol = lNextCol + 1
            dictColsIndex(dictColAttr.Keys(lEachRow)) = lNextCol
        End If
    Next
    
    Erase arrConfigData
    Erase arrColsName
    Erase arrColsIndex
End Function

Function fCrossValidateInputFileTxtSheetFile()
    Dim i As Long
    Dim sFileTag As String
    Dim sSource As String
    Dim sDependantFileTag As String
    Dim sDependantSheet As String
    Dim lActualRow As Long
    
    For i = 0 To gDictInputFiles.Count - 1
        sFileTag = gDictInputFiles.Keys(i)
        sSource = fGetInputFileSourceType(sFileTag)
        
        If sSource = "FILE_BINDED_IN_MACRO" Or sSource = "READ_FROM_DRIVE" Then
            fIfExcelFileOpenedToCloseIt fGetInputFileFileName(sFileTag)
        ElseIf sSource = "PARSE_AS_TEXT" Then
            If Not gDictTxtFileSpec Is Nothing Then
                If Not gDictTxtFileSpec.Exists(sFileTag) Then
                    lActualRow = fGetInputFileRowNo(sFileTag)
                    fErr "The file Tag below is configured as " & sSource & ", but [TXT File Importing Specification] does not have one, pls check." _
                        & vbCr & "File Tag: " & sFileTag _
                        & vbCr & "Row: " & lActualRow
                End If
            End If
        End If
    Next
    
    If Not gDictTxtFileSpec Is Nothing Then
        For i = 0 To gDictTxtFileSpec.Count - 1
            sFileTag = gDictTxtFileSpec.Keys(i)
            
            If Not gDictInputFiles.Exists(sFileTag) Then
                fErr sFileTag & " is coonfigured under [TXT File Importing Specification], but does not exist in [Input Files], pls check." _
                    & vbCr & "File Tag: " & sFileTag
            End If
        Next
    End If
End Function
Function fGetInputFileFileName(asFileTag As String) As String
    fGetInputFileFileName = Split(gDictInputFiles(asFileTag), DELIMITER)(InputFile.FilePath - InputFile.ReportID - 2)
End Function
Function fGetInputFileSourceType(asFileTag As String) As String
    fGetInputFileSourceType = Split(gDictInputFiles(asFileTag), DELIMITER)(InputFile.Source - InputFile.ReportID - 2)
End Function
Function fGetInputFileFileSpecTag(asFileTag As String) As String
    fGetInputFileFileSpecTag = Split(gDictInputFiles(asFileTag), DELIMITER)(InputFile.FileSpecTag - InputFile.ReportID - 2)
End Function
Function fGetInputFileRowNo(asFileTag As String) As String
    fGetInputFileRowNo = Split(gDictInputFiles(asFileTag), DELIMITER)(InputFile.RowNo - InputFile.ReportID - 2)
End Function
Function fGetInputFileReloadOrNot(asFileTag As String) As String
    fGetInputFileReloadOrNot = Split(gDictInputFiles(asFileTag), DELIMITER)(InputFile.ReLoadOrNot - InputFile.ReportID - 2)
End Function
Function fGetInputFileSheetToImport(asFileTag As String) As String
    fGetInputFileSheetToImport = Split(gDictInputFiles(asFileTag), DELIMITER)(InputFile.DefaultSheet - InputFile.ReportID - 2)
End Function
Function fGetInputFilePivotTableTag(asFileTag As String) As String
    fGetInputFilePivotTableTag = Split(gDictInputFiles(asFileTag), DELIMITER)(InputFile.PivotTableTag - InputFile.ReportID - 2)
End Function

Function fGetReneratedReport(asReportID As String) As String
    fGetReneratedReport = fGetSpecifiedConfigCellValue(shtSysConf, "[Generate Report File / Sheet]" _
                            , "Genearted File Or Sheet", "Report ID=" & asReportID)
End Function


