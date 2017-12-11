Attribute VB_Name = "Common_Local"
Option Explicit
Option Base 1
 

Function fUpdateGDictInputFile_FileName(asFileTag As String, asFileName As String)
    Call fUpdateDictionaryItemValueForDelimitedElement(gDictInputFiles, asFileTag, InputFile.FilePath - InputFile.FileTag, asFileName)
End Function

Function fSetValueBackToSysConf_InputFile_FileName(asFileTag As String, asFileName As String)
    Call fSetSpecifiedConfigCellValue(shtSysConf, "[Input Files]", "File Full Path", "File Tag=" & asFileTag, asFileName)
End Function

Function fGetInputFileSheetAfterLoadingToThisWorkBook(asFileTag As String) As Worksheet
    Set fGetInputFileSheetAfterLoadingToThisWorkBook = ThisWorkbook.Worksheets(fGetInputFileSheetNameAfterLoadingToThisWorkBook(asFileTag))
End Function

Function fGetInputFileSheetNameAfterLoadingToThisWorkBook(asFileTag As String) As String
    Dim sOut As String
    Dim sSource As String
    sSource = fGetInputFileSourceType(asFileTag)
    
    Select Case sSource
        Case "PARSE_AS_TEXT"
            sOut = asFileTag
        Case "FILE_BINDED_IN_MACRO", "READ_FROM_DRIVE", "READ_PREV_STEP_OUTPUT_FILE"
            sOut = asFileTag
        Case "READ_PRE_EXISTING_SHEET", "READ_PREV_STEP_OUTPUT_SHEET"
            sOut = fGetInputFileFileName(asFileTag)
        Case "READ_SHEET_BINDED_IN_MACRO"
            sOut = ""
        Case Else
            fErr "wrong sSource" & sSource
    End Select
    
    fGetInputFileSheetNameAfterLoadingToThisWorkBook = sOut
End Function

Function fConvertFomulaToValueForSheetIfAny(sht As Worksheet)
    Dim rng As Range
    
    On Error Resume Next
    Set rng = sht.Cells.SpecialCells(xlCellTypeFormulas)
    err.Clear
    
    If rng Is Nothing Then Exit Function
    
'    Dim eachRng
'    For Each eachRng In rng.Areas
'        eachRng.Value = eachRng.Value
'    Next

    rng.Parent.UsedRange.Value = rng.Parent.UsedRange.Value
End Function

Function fCloseWorkbookWithoutSave(wb As Workbook)
    wb.Saved = True
    wb.Close savechanges:=False
    Set wb = Nothing
End Function
Function fImportSingleSheetExcelFileToThisWorkbook(sExcelFileFullPath As String, sNewSheet As String _
                        , Optional asShtToImport As String = "", Optional wb As Workbook)
    Call fIfExcelFileOpenedToCloseIt(sExcelFileFullPath)
    
    If wb Is Nothing Then Set wb = ThisWorkbook
    
    Dim wbSource As Workbook
    Set wbSource = Workbooks.Open(sExcelFileFullPath)
    
    asShtToImport = Trim(asShtToImport)
    
    If Len(asShtToImport) <= 0 Then
        wbSource.Worksheets(1).Copy after:=wb.Worksheets(wb.Worksheets.Count)
    Else
        If Not fSheetExists(asShtToImport, , wbSource) Then
            fErr "There is no sheet named """ & asShtToImport & """ in workbook " & sExcelFileFullPath
        End If
        
        wbSource.Worksheets(asShtToImport).Copy after:=wb.Worksheets(wb.Worksheets.Count)
    End If
    
    wb.ActiveSheet.Name = sNewSheet
    ActiveWindow.DisplayGridlines = False
    
    Call fConvertFomulaToValueForSheetIfAny(wb.Worksheets(sNewSheet))
    
    Call fCloseWorkbookWithoutSave(wbSource)
End Function

Function fLoadFileByFileTag(asFileTag As String)
    Dim sFileFullPath As String
    Dim sSource As String
    Dim sReloadOrNot As String
    Dim sShtToImport As String
    Dim sShtToBeAdded As String
    
    sSource = fGetInputFileSourceType(asFileTag)
    If sSource = "READ_SHEET_BINDED_IN_MACRO" Then Exit Function
    
    sFileFullPath = fGetInputFileFileName(asFileTag)
    sReloadOrNot = fGetInputFileReloadOrNot(asFileTag)
    sShtToImport = fGetInputFileSheetToImport(asFileTag)
    
    sShtToBeAdded = fGetInputFileSheetNameAfterLoadingToThisWorkBook(asFileTag)
    
    If fSheetExists(sShtToBeAdded) Then
        If sReloadOrNot = "RELOAD" Or fZero(sReloadOrNot) Then
            Call fDeleteSheet(sShtToBeAdded)
        Else
            Exit Function
        End If
    End If
    
    Select Case sSource
        Case "PARSE_AS_TEXT"
            Call fReadTxtFile2NewSheet(sFileFullPath, sShtToBeAdded, asFileTag)
        Case "FILE_BINDED_IN_MACRO", "READ_FROM_DRIVE", "READ_PREV_STEP_OUTPUT_FILE"
            Call fImportSingleSheetExcelFileToThisWorkbook(sFileFullPath, sShtToBeAdded)
            Call fRemoveFilterForSheet(ThisWorkbook.Worksheets(sShtToBeAdded))
        Case "READ_PRE_EXISTING_SHEET", "READ_PREV_STEP_OUTPUT_SHEET"
        Case "READ_SHEET_BINDED_IN_MACRO"
            Exit Function
        Case Else
            fErr "wrong sSource" & sSource
    End Select
    
End Function

Function fReadTxtColSpec(asFileTag As String) As Variant
    Dim arrOut()
    
    Dim sFileSpecTag As String
    sFileSpecTag = fGetInputFileFileSpecTag(asFileTag)
    
    Dim asTag As String
    Dim arrColsName()
    Dim rngToFindIn As Range
    Dim arrConfigData()
    Dim lConfigStartRow As Long
    Dim lConfigStartCol As Long
    Dim lConfigEndRow As Long
    Dim lConfigHeaderAtRow As Long

    asTag = sFileSpecTag
    ReDim arrColsName(1 To 2)
    arrColsName(1) = "Column Index"
    arrColsName(2) = "TXT Format Only For Text File"
     
    arrConfigData = fReadConfigBlockToArrayNet(asTag:=asTag, shtParam:=shtFileSpec _
                                , arrColsName:=arrColsName _
                                , lConfigStartRow:=lConfigStartRow _
                                , lConfigStartCol:=lConfigStartCol _
                                , lConfigEndRow:=lConfigEndRow _
                                , lOutConfigHeaderAtRow:=lConfigHeaderAtRow _
                                , abNoDataConfigThenError:=True _
                                )
    Call fValidateDuplicateInArray(arrConfigData, 1, False, shtStaticData, lConfigHeaderAtRow, lConfigStartCol, arrColsName(1))
'    Call fValidateBlankInArray(arrConfigData, TechTag.Report_ID, shtstaticdata, lConfigHeaderAtRow, lConfigStartCol, "TechTag ID")

    Dim dict As Dictionary
    Set dict = New Dictionary
    
    Dim lEachRow As Long
    Dim sLetterIndex As String
    Dim sTxtFormat As String
    Dim lTxtFormat As Long
   ' Dim lMaxColNum As Long
'    Dim lColindex As Long
    
'    lMaxColNum = 0
    For lEachRow = LBound(arrConfigData, 1) To UBound(arrConfigData, 1)
        If fArrayRowIsBlankHasNoData(arrConfigData, lEachRow) Then GoTo next_row
        
        sLetterIndex = Trim(arrConfigData(lEachRow, 1))
        sTxtFormat = Trim(arrConfigData(lEachRow, 2))
        
        If Len(sTxtFormat) <= 0 Then
            lTxtFormat = 1
        Else
            lTxtFormat = fGetTxtImportDataFormat(sTxtFormat)
        End If
        
        dict.Add sLetterIndex, lTxtFormat
'        lColindex = fLetter2Num(sLetterIndex)
'        If lColindex > lMaxColNum Then lMaxColNum = lColindex
next_row:
    Next
    
    Erase arrColsName
    Erase arrConfigData
    
    Dim lColindex As Long
    
    ReDim arrOut(1 To dict.Count)
    For lEachRow = 0 To dict.Count - 1
        lColindex = fLetter2Num(dict.Keys(lEachRow))
        arrOut(lColindex) = dict.Items(lEachRow)
    Next
    
    fReadTxtColSpec = arrOut()
    Erase arrOut
    Set dict = Nothing
End Function
Function fImportTxtFile(sFileFullPath, arrColFormat, asDelmiter As String _
                        , alTextFilePlatForm As Long, ByRef shtTo As Worksheet) As Worksheet
    If fArrayIsEmptyOrNoData(arrColFormat) Then arrColFormat = Array(1)
    
    shtTo.Cells.ClearContents
    
    With shtTo.QueryTables.Add(Connection:="TEXT;" & sFileFullPath _
        , Destination:=shtTo.Range("$A$1"))
        '.CommandType = 0
        .Name = shtTo.Name
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
        .TextFilePlatform = alTextFilePlatForm
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = True
        .TextFileTabDelimiter = False
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileOtherDelimiter = ","
        .TextFileColumnDataTypes = arrColFormat
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
    
End Function

Function fReadTxtFile2NewSheet(sFileFullPath As String, sShtToBeAdded As String, asFileTag As String)
    Dim shtToAdd As Worksheet
    Set shtToAdd = fAddNewSheet(sShtToBeAdded)
    
    Dim arrColFormat()
    arrColFormat = fReadTxtColSpec(asFileTag)
    
    Dim sColDelimiter As String
    Dim lPlatForm As Long
    
    sColDelimiter = Trim(Split(gDictTxtFileSpec(asFileTag), DELIMITER)(0))
    lPlatForm = CLng(Split(gDictTxtFileSpec(asFileTag), DELIMITER)(1))
    
    Call fImportTxtFile(sFileFullPath, arrColFormat, sColDelimiter, lPlatForm, shtToAdd)
    
    fDeleteRemoveConnections shtToAdd.Parent
End Function

Function fDeleteRemoveConnections(wb As Workbook)
    Dim i As Long
    
    For i = wb.Connections.Count To 1 Step -1
        wb.Connections(i).Delete
    Next
End Function

Function fCheckIfSheetHasNodata_RaiseErrToStop(arr)
    gbNoData = fArrayIsEmptyOrNoData(arr)
    If gbNoData Then fErr "Input File has no qualified data!"
End Function

Function fFindHeaderAtLineInFileSpec(rngConfigBlock As Range, arrColsName) As Long
    Dim lColAtRow As Long
    Dim lEachCol As Long
    Dim sEachColName As String
    Dim rngFound As Range
    
    lColAtRow = 0
    For lEachCol = LBound(arrColsName) To UBound(arrColsName)
        sEachColName = Trim(arrColsName(lEachCol))
        sEachColName = Replace(sEachColName, "*", "~*")
        
        Set rngFound = fFindInWorksheet(rngConfigBlock, sEachColName)
        
        If lColAtRow <> 0 Then
            If lColAtRow <> rngFound.Row Then
                fErr "Columns are not at the same row."
            End If
        Else
            lColAtRow = rngFound.Row
        End If
    Next
    
    Set rngFound = Nothing
    
    fFindHeaderAtLineInFileSpec = lColAtRow
End Function

Function fFileSpecTemplateHasAdditionalHeader(rngConfigBlock As Range, arrHeadersToFind) As Boolean
    Dim arrColsName(1 To 6)
    arrColsName(1) = "Column Tech Name"
    arrColsName(2) = "Column Display Name"
    arrColsName(3) = "Column Index"
    arrColsName(4) = "Array Index"
    arrColsName(5) = "Raw Data Type"
    arrColsName(6) = "Data Format"
    
    Dim rngHeader As Range
    Dim lHeaderAtLine As Long
    Dim rngFound As Range
    Dim i As Integer
    
    lHeaderAtLine = fFindHeaderAtLineInFileSpec(rngConfigBlock, arrColsName)
    Set rngHeader = fGetRangeByStartEndPos(shtFileSpec, lHeaderAtLine, rngConfigBlock.Column, lHeaderAtLine, Columns.Count)
    
    If IsArray(arrHeadersToFind) Then
        For i = LBound(arrHeadersToFind) To UBound(arrHeadersToFind)
            Set rngFound = fFindInWorksheet(rngHeader, CStr(arrHeadersToFind(i)), False)
            If rngFound Is Nothing Then
                Exit For
            End If
        Next
    Else
        Set rngFound = fFindInWorksheet(rngHeader, CStr(arrHeadersToFind), False)
    End If
    
    fFileSpecTemplateHasAdditionalHeader = (Not rngFound Is Nothing)
    Set rngHeader = Nothing
    Set rngFound = Nothing
End Function
Function fGetTableLevelConfig(rngConfigBlock As Range, asTableLevelConf As String) As String
    Dim shtParent As Worksheet
    Dim rgFound As Range
    Dim rgTarget As Range
    Dim lValueColSameAsDisplayName As Long
    
    Set shtParent = rngConfigBlock.Parent
    
    Set rgFound = fFindInWorksheet(rngConfigBlock, "Column Display Name")
    lValueColSameAsDisplayName = rgFound.Column
    
    Set rgFound = fFindInWorksheet(rngConfigBlock, asTableLevelConf)
    
    Set rgTarget = rgFound.Offset(0, lValueColSameAsDisplayName - rgFound.Column)
    
    If fZero(rgTarget.Value) Then
        fErr "asTableLevelConf cannot be blank in " & shtParent.Name & vbCr & "range: " & rngConfigBlock.Address
    End If
    
    fGetTableLevelConfig = Trim(rgTarget.Value)
    Set rgTarget = Nothing
    Set rgFound = Nothing
    Set shtParent = Nothing
End Function

Function fReadInputFileSpecConfig(sFileSpecTag As String, ByRef dictLetterIndex As Dictionary _
                                , Optional ByRef dictArrayIndex As Dictionary _
                                , Optional ByRef dictDisplayName As Dictionary _
                                , Optional ByRef dictRawType As Dictionary _
                                , Optional ByRef dictDataFormat As Dictionary _
                                , Optional ByRef bReadWholeSheetData As Boolean _
                                , Optional shtData As Worksheet _
                                , Optional alHeaderAtRow As Long = 1)
    'Dim asTag As String
    Dim arrColsName()
    Dim arrColsIndex()
    Dim rngToFindIn As Range
    Dim arrConfigData()
    Dim lConfigStartRow As Long
    Dim lConfigStartCol As Long
    Dim lConfigEndRow As Long
    Dim lConfigHeaderAtRow As Long
    
    Const TECH_NAME = 1
    Const DISPLAY_NAME = 2
    Const LETTER_INDEX = 3
    Const ARRAY_INDEX = 4
    Const RAW_DATA_TYPE = 5
    Const DATA_FORMAT = 6
    
    ReDim arrColsName(TECH_NAME To DATA_FORMAT)
    
    arrColsName(TECH_NAME) = "Column Tech Name"
    arrColsName(DISPLAY_NAME) = "Column Display Name"
    arrColsName(LETTER_INDEX) = "Column Index"
    arrColsName(ARRAY_INDEX) = "Array Index"
    arrColsName(RAW_DATA_TYPE) = "Raw Data Type"
    arrColsName(DATA_FORMAT) = "Data Format"
    
    Dim rngConfigBlock As Range
    Set rngConfigBlock = fFindRageOfFileSpecConfigBlock(sFileSpecTag)
    
    Dim iCol_TxtFormat As Long
    Dim iCol_OutputAsInput As Long
    Dim bTxtTemplate As Boolean
    Dim bOutputAsInput As Boolean
    Dim sGetColIndexBy As String
    Dim sReadSheetDataBy As String
    Dim bDynamic As Boolean
    
    bTxtTemplate = fFileSpecTemplateHasAdditionalHeader(rngConfigBlock, "TXT Format Only For Text File")
    If bTxtTemplate Then iCol_TxtFormat = fEnlargeArayWithValue(arrColsName, "TXT Format Only For Text File")
        
    bOutputAsInput = fFileSpecTemplateHasAdditionalHeader(rngConfigBlock, Array("Column Attr", "Column Width"))
    If bOutputAsInput Then iCol_OutputAsInput = fEnlargeArayWithValue(arrColsName, "Column Attr")
    
    sGetColIndexBy = UCase(fGetTableLevelConfig(rngConfigBlock, "Get Column Index By:"))
    sReadSheetDataBy = UCase(fGetTableLevelConfig(rngConfigBlock, "Read Sheet's Data By:"))
    
    bDynamic = (sGetColIndexBy <> "FIXED_LETTERS")
    bReadWholeSheetData = (sReadSheetDataBy = "READ_WHOLE_SHEET")

    Dim sErrPos As String
    sErrPos = vbCr & vbCr & "Sheet Name: " & shtFileSpec.Name & vbCr & "Range:" & rngConfigBlock.Address
    
    If bDynamic Then
        If iCol_TxtFormat > 0 Then fErr "dynamic (COLUMNS_NAME) cannot be specified for Txt Template" & sErrPos
        If shtData Is Nothing Then fErr "dynamic (COLUMNS_NAME), but shtData is not provided(nothing)." & sErrPos
    End If
    
    Call fReadConfigBlockToArray(asTag:=sFileSpecTag, shtParam:=shtFileSpec _
                                , arrConfigData:=arrConfigData _
                                , arrColsName:=arrColsName _
                                , arrColsIndex:=arrColsIndex _
                                , lConfigStartRow:=lConfigStartRow _
                                , lConfigStartCol:=lConfigStartCol _
                                , lConfigEndRow:=lConfigEndRow _
                                , lOutConfigHeaderAtRow:=lConfigHeaderAtRow _
                                , abNoDataConfigThenError:=True)
    If lConfigHeaderAtRow >= lConfigEndRow Then fErr "No data is configured  " & sErrPos
    Call fValidateDuplicateInArray(arrConfigData, arrColsIndex(TECH_NAME), False, shtSysConf, lConfigHeaderAtRow, lConfigStartCol, arrColsName(1))
    
    If bDynamic Then ' by col dislay name    '"Column Display Name"
        Call fValidateDuplicateInArray(arrConfigData, arrColsIndex(DISPLAY_NAME), False, shtSysConf, lConfigHeaderAtRow, lConfigStartCol, arrColsName(3))
        Call fValidateDuplicateInArray(arrConfigData, arrColsIndex(LETTER_INDEX), True, shtSysConf, lConfigHeaderAtRow, lConfigStartCol, arrColsName(3))
    Else   'by speicified letter  '"Column Index"     Txt Template
        Call fValidateDuplicateInArray(arrConfigData, arrColsIndex(LETTER_INDEX), False, shtSysConf, lConfigHeaderAtRow, lConfigStartCol, arrColsName(3))
    End If
    
    Call fValidateDuplicateInArray(arrConfigData, arrColsIndex(ARRAY_INDEX), True, shtSysConf, lConfigHeaderAtRow, lConfigStartCol, arrColsName(3))
'    Call fValidateBlankInArray(arrConfigData, 1, shtSysConf, lConfigHeaderAtRow, lConfigStartCol, "Output Type")
'    Call fValidateBlankInArray(arrConfigData, 2, shtSysConf, lConfigHeaderAtRow, lConfigStartCol, "Output Type")
    
    Dim lEachRow As Long
    Dim lActualRow As Long
    Dim sColTechName As String
    Dim sDisplayName As String
    Dim sLetterIndex As String
    Dim lColLetter2Num As Long
    Dim sArrayIndex As String
    Dim lColArray2Num As Long
    Dim arrTxtNonImportCol()
    Dim dictActualRow As New Dictionary
    
    sErrPos = sErrPos & vbCr & vbCr & "Row : $ACTUAL_ROW$" & vbCr & "Column: "

    Set dictLetterIndex = New Dictionary
    Set dictArrayIndex = New Dictionary
    Set dictDisplayName = New Dictionary
    Set dictRawType = New Dictionary
    Set dictDataFormat = New Dictionary

    'Dim dictTmpArrayInd As New Dictionary
    For lEachRow = LBound(arrConfigData, 1) To UBound(arrConfigData, 1)
        If fArrayRowIsBlankHasNoData(arrConfigData, lEachRow) Then GoTo next_row
        
        lActualRow = lConfigHeaderAtRow + lEachRow
        
        If bOutputAsInput Then
            If Trim(arrConfigData(lEachRow, arrColsIndex(iCol_OutputAsInput))) = "NOT_SHOW_UP" Then
                If Len(sLetterIndex) > 0 Then
                    fErr "Col Letter Index should be blank for NOT_SHOW_UP: " & Replace(sErrPos, "$ACTUAL_ROW$", lActualRow) & arrColsName(iCol_OutputAsInput)
                End If
                GoTo next_row
            End If
        End If
        
        sColTechName = Trim(arrConfigData(lEachRow, arrColsIndex(TECH_NAME)))
        sDisplayName = Trim(arrConfigData(lEachRow, arrColsIndex(DISPLAY_NAME)))
        sLetterIndex = Trim(arrConfigData(lEachRow, arrColsIndex(LETTER_INDEX)))
        sArrayIndex = Trim(arrConfigData(lEachRow, arrColsIndex(ARRAY_INDEX)))
        
        If bTxtTemplate Then
            If Trim(arrConfigData(lEachRow, arrColsIndex(iCol_TxtFormat))) = "xlSkipColumn" Then
                If Len(sArrayIndex) > 0 Then fErr "ArrayIndex should be blank when xlSkipColumn is specified " & Replace(sErrPos, "$ACTUAL_ROW$", lActualRow) & arrColsName(iCol_TxtFormat)
            
                Call fEnlargeArayWithValue(arrTxtNonImportCol, fLetter2Num(sLetterIndex))
                GoTo next_row
            End If
        End If
        
        If Not bDynamic Then
            'If Len(sLetterIndex) > 0 Then
                lColLetter2Num = fLetter2Num(sLetterIndex)
                
                If lColLetter2Num <= 0 Or lColLetter2Num > Columns.Count Then
                    fErr "Col Letter Index is invalid,should be A - XFD: " & Replace(sErrPos, "$ACTUAL_ROW$", lActualRow) & arrColsName(LETTER_INDEX)
                End If
                dictLetterIndex.Add sColTechName, lColLetter2Num
            'End If
        End If
        If Not bReadWholeSheetData Then
            If Len(sArrayIndex) > 0 Then
                If Not IsNumeric(sArrayIndex) Then
                    fErr "Col Array Index is invalid,should be  1, 2, 3, ...: " & Replace(sErrPos, "$ACTUAL_ROW$", lActualRow) & arrColsName(LETTER_INDEX)
                End If
                lColArray2Num = CLng(sArrayIndex)
                
                If lColArray2Num <= 0 Or lColArray2Num > Columns.Count Then
                    fErr "Col Array Index is invalid,should be A - XFD: " & Replace(sErrPos, "$ACTUAL_ROW$", lActualRow) & arrColsName(ARRAY_INDEX)
                End If
                dictArrayIndex.Add sColTechName, lColArray2Num
            End If
        End If
        
        dictDisplayName.Add sColTechName, sDisplayName
        dictRawType.Add sColTechName, UCase(Trim(arrConfigData(lEachRow, arrColsIndex(RAW_DATA_TYPE))))
        dictDataFormat.Add sColTechName, Trim(arrConfigData(lEachRow, arrColsIndex(DATA_FORMAT)))
        
        dictActualRow.Add sColTechName, lActualRow
next_row:
    Next
    
    If dictActualRow.Count <= 0 Then fErr "Cxxxxxxxxxxx"
    
    If Not bReadWholeSheetData Then
        If dictArrayIndex.Count <= 0 Then
            fErr "READ_SPECIFIED_COLUMNS is specified, but no array index is not specified"
        End If
    End If
    
    If bTxtTemplate Then
        If fRecalculateColumnIndexByRemoveNonImportTxtCol(dictLetterIndex, arrTxtNonImportCol) Then
            For lEachRow = 0 To dictActualRow.Count - 1
                shtFileSpec.Cells(dictActualRow.Items(lEachRow), lConfigStartCol + arrColsIndex(LETTER_INDEX)) = _
                    fNum2Letter(dictLetterIndex.Items(lEachRow)) 'this is for reference
            Next
        End If
    End If
    
    Dim arrDisplayNames()
    Dim arrDynamicColIndex()
    If bDynamic Then
        arrDisplayNames = dictDisplayName.Items
        Call fFindAllColumnsIndexByColNames(shtData.Rows(alHeaderAtRow), arrDisplayNames, arrDynamicColIndex)
        
        If Not Base0(arrDisplayNames) Then fErr "arrDisplayNames is not based from 0"
        For lEachRow = LBound(arrDynamicColIndex) To UBound(arrDynamicColIndex)
            dictLetterIndex.Add dictDisplayName.Keys(lEachRow), arrDynamicColIndex(lEachRow)
        Next
        For lEachRow = 0 To dictActualRow.Count - 1
            shtFileSpec.Cells(dictActualRow.Items(lEachRow), lConfigStartCol + arrColsIndex(LETTER_INDEX) - 1) = _
                    dictLetterIndex.Items(lEachRow)
        Next
    End If
    
    Erase arrDisplayNames
    Erase arrDynamicColIndex
    Erase arrDisplayNames
    Erase arrColsName
    Erase arrColsIndex
    Erase arrConfigData
    Set dictActualRow = Nothing
    Set rngToFindIn = Nothing
End Function

Function fRecalculateColumnIndexByRemoveNonImportTxtCol(ByRef dictLetterIndex As Dictionary, arrTxtNonImportCol()) As Boolean
    Dim bOut As Boolean
    bOut = False
    
    If fArrayIsEmptyOrNoData(arrTxtNonImportCol) Then GoTo exit_fun
    
    Call fSortArayDesc(arrTxtNonImportCol)
    
    Dim iArrayIndex As Long
    Dim iDictIndex As Long
    Dim iColIndex As Long
    
    For iArrayIndex = LBound(arrTxtNonImportCol) To UBound(arrTxtNonImportCol)
        iColIndex = arrTxtNonImportCol(iArrayIndex)
        
        For iDictIndex = 0 To dictLetterIndex.Count - 1
            If dictLetterIndex.Items(iDictIndex) > iColIndex Then
                dictLetterIndex(dictLetterIndex.Keys(iDictIndex)) = dictLetterIndex.Items(iDictIndex) - 1
            ElseIf dictLetterIndex.Items(iDictIndex) = iColIndex Then
                fErr "abnormal in fRecalculateColumnIndexByRemoveNonImportTxtCol"
            Else
                'Debug.Print dictLetterIndex.Items(iDictIndex) & " - " & iColIndex
            End If
        Next iDictIndex
    Next iArrayIndex
    
    bOut = True
exit_fun:
    fRecalculateColumnIndexByRemoveNonImportTxtCol = bOut
End Function

Function fGetTxtImportDataFormat(sDesc As String) As Integer
    Dim iOut As Integer

    Select Case sDesc
        Case "xlGeneralFormat"
            iOut = 1
        Case "xlTextFormat"
            iOut = 2
        Case "xlMDYFormat"
            iOut = 3
        Case "xlDMYFormat"
            iOut = 4
        Case "xlYMDFormat"
            iOut = 5
        Case "xlMYDFormat"
            iOut = 6
        Case "xlDYMFormat"
            iOut = 7
        Case "xlYDMFormat"
            iOut = 8
        Case "xlSkipColumn"
            iOut = 9
        Case "xlEMDFormat"
            iOut = 10
        Case Else
    End Select
    
    fGetTxtImportDataFormat = iOut
End Function

Function fReadSheetDataByConfig(asFileTag As String, ByRef dictColIndex As Dictionary, ByRef arrDataOut() _
                                , Optional ByRef dictColFormat As Dictionary _
                                , Optional ByRef dictRawType As Dictionary _
                                , Optional ByRef dictDisplayName As Dictionary _
                                , Optional alDataFromRow As Long = 2 _
                                , Optional shtData As Worksheet)
    Dim sFileSpecTag As String
    Dim shtToRead As Worksheet
    
    sFileSpecTag = fGetInputFileFileSpecTag(asFileTag)
    
    If shtData Is Nothing Then
        Set shtToRead = fGetInputFileSheetAfterLoadingToThisWorkBook(asFileTag)
    Else
        Set shtToRead = shtToRead
    End If
     
    Dim bReadWholeSheetData As Boolean
    Dim dictLetterIndex As Dictionary
    Dim dictArrayIndex As Dictionary
    Call fReadInputFileSpecConfig(sFileSpecTag:=sFileSpecTag _
                                , dictLetterIndex:=dictLetterIndex _
                                , dictArrayIndex:=dictArrayIndex _
                                , dictDisplayName:=dictDisplayName _
                                , dictRawType:=dictRawType _
                                , dictDataFormat:=dictColFormat _
                                , bReadWholeSheetData:=bReadWholeSheetData _
                                , shtData:=shtToRead _
                                , alHeaderAtRow:=alDataFromRow - 1)
    
    If bReadWholeSheetData Then
        Call fCopyReadWholeSheetData2Array(shtToRead, arrDataOut, dictLetterIndex, alDataFromRow)
        Call fConvertDateCol2RawValue(arrDataOut, dictLetterIndex, dictRawType, dictColFormat)
    Else
        Call fReadSpecifiedColsToArrayByConfig(shtData:=shtToRead, dictLetterIndex:=dictLetterIndex, dictArrayIndex:=dictArrayIndex _
                    , dictRawType:=dictRawType, dictColFormat:=dictColFormat _
                     , arrDataOut:=arrDataOut, alDataFromRow:=alDataFromRow)
    End If
    
    If bReadWholeSheetData Then
        Set dictColIndex = dictLetterIndex
    Else
        Set dictColIndex = dictArrayIndex
    End If
    
    Set shtToRead = Nothing
    Set dictLetterIndex = Nothing
    Set dictArrayIndex = Nothing
End Function

Function fGetReadInputFileSpecConfigItem(sFileSpecTag As String, asItem As String) As Variant
    Dim dictColFormat As Dictionary _
                                , dictRawType As Dictionary _
                                , dictDisplayName As Dictionary
'                                , alDataFromRow As Long _
'                                , shtData As Worksheet
    Dim dictLetterIndex As Dictionary
    Dim dictArrayIndex As Dictionary
    
    Call fReadInputFileSpecConfig(sFileSpecTag:=sFileSpecTag _
                                , dictLetterIndex:=dictLetterIndex _
                                , dictArrayIndex:=dictArrayIndex _
                                , dictDisplayName:=dictDisplayName _
                                , dictRawType:=dictRawType _
                                , dictDataFormat:=dictColFormat _
                                , bReadWholeSheetData:=bReadWholeSheetData _
                                , shtData:=shtToRead _
                                , alHeaderAtRow:=alDataFromRow - 1)
    Select Case asItem
        Case "TXT_COL_FORMAT"
            set fGetReadInputFileSpecConfigItem =
        Case Else
            
    End Select
    
    Set dictColFormat = Nothing
    Set dictRawType = Nothing
    Set dictDisplayName = Nothing
    Set dictLetterIndex = Nothing
    Set dictArrayIndex = Nothing
End Function

Function fReadSpecifiedColsToArrayByConfig(shtData As Worksheet, dictLetterIndex As Dictionary, dictArrayIndex As Dictionary _
                    , dictRawType As Dictionary, dictColFormat As Dictionary _
                     , arrDataOut(), Optional alDataFromRow As Long = 2)

    Dim lColCopyFrom As Long
    Dim lColCopyTo As Long
    Dim lArrayMaxCol As Long
    Dim lShtMaxRow As Long
    
    lArrayMaxCol = WorksheetFunction.Max(dictArrayIndex.Items)
    lShtMaxRow = fGetValidMaxRow(shtData)
    
    If lShtMaxRow < alDataFromRow Then arrDataOut = Array(): Exit Function
    
    ReDim arrDataOut(1 To lShtMaxRow - alDataFromRow + 1, 1 To lArrayMaxCol)
    
    Dim i As Long
    Dim sTechName As String
    Dim sColType As String
    Dim arrEachCol()
    Dim lEachRow As Long
    For i = 0 To dictArrayIndex.Count - 1
        sTechName = dictArrayIndex.Keys(i)
        
        lColCopyFrom = dictLetterIndex(sTechName)
        lColCopyTo = dictArrayIndex(sTechName)
        
        sColType = UCase(dictRawType(sTechName))
        arrEachCol = fReadRangeDatatoArrayByStartEndPos(shtData, alDataFromRow, lColCopyFrom, lShtMaxRow, lColCopyFrom)
        
        If sColType = "DATE" Or sColType = "STRING_PERCENTAGE" Then
            For lEachRow = LBound(arrEachCol, 1) To UBound(arrEachCol, 1)
                arrDataOut(lEachRow, lColCopyTo) = fCType(arrEachCol(lEachRow, 1), sColType, dictColFormat(sTechName))
            Next
        Else
            For lEachRow = LBound(arrEachCol, 1) To UBound(arrEachCol, 1)
                arrDataOut(lEachRow, lColCopyTo) = arrEachCol(lEachRow, 1)
            Next
        End If
    Next
    
End Function

Function fConvertDateCol2RawValue(ByRef arrData(), dictLetterIndex As Dictionary, dictRawType As Dictionary, dictColFormat As Dictionary)
    Dim i As Long
    Dim lEachRow As Long
    Dim lEachCol As Long
    Dim sTechName As String
    Dim sColType As String
    
    For i = 0 To dictLetterIndex.Count - 1
        sTechName = dictLetterIndex.Keys(i)
        sColType = UCase(dictRawType(sTechName))
        
        If sColType = "DATE" Or sColType = "STRING_PERCENTAGE" Then
            lEachCol = dictLetterIndex.Items(i)
            
            For lEachRow = LBound(arrData, 1) To UBound(arrData, 1)
                arrData(lEachRow, lEachCol) = fCType(arrData(lEachRow, lEachCol), sColType, dictColFormat(sTechName))
            Next
        End If
    Next
End Function

Function fCType(aValue, asToType As String, asFormat As String) As Variant
    Dim aOut As Variant
    Dim sDataType As String
    Dim bOrigToAreSame As Boolean
    
    If IsEmpty(aValue) Then fCType = aValue: Exit Function
    
    asToType = UCase(asToType)
    sDataType = UCase(TypeName(aValue))
    
    bOrigToAreSame = False
    Select Case asToType
        Case "STRING", "TEXT"
            If sDataType = "STRING" Then bOrigToAreSame = True
        Case "DATE"
            If aValue = 0 Then fCType = 0: Exit Function
            If sDataType = "DATE" Then bOrigToAreSame = True
        Case "DECIMAL"
            If sDataType = "DECIMAL" Or sDataType = "DOUBLE" Or sDataType = "SINGLE" Or sDataType = "CURRENCY" Then
                bOrigToAreSame = True
            End If
        Case "NUMBER"
            If sDataType = "BYTE" Or sDataType = "INTEGER" Or sDataType = "LONG" Or sDataType = "LONGLONG" Or sDataType = "LONGPRT" Then
                bOrigToAreSame = True
            End If
        Case "STRING_PERCENTAGE"
            
        Case Else
            fErr "wrong param asToType"
    End Select
    
    If bOrigToAreSame Then fCType = aValue: Exit Function
    
    Select Case asToType
        Case "STRING", "TEXT"
            fCType = CStr(aValue)
        Case "DATE"
            Dim dtTmp As Date
            dtTmp = fcdate(CStr(aValue), asFormat)
            
            If dtTmp <= 0 Then
                fErr "Wrong date value: " & aValue & ", please check your data, or contact with IT support."
            End If
            fCType = dtTmp
        Case "DECIMAL"
            fCType = CDbl(aValue)
        Case "NUMBER"
            fCType = CLng(aValue)
        Case "STRING_PERCENTAGE"
            fCType = fCPercentage2Dbl(aValue)
        Case Else
            fErr "wrong param asToType"
    End Select
End Function

Function fCPercentage2Dbl(aValue As String) As Double
    aValue = Trim(aValue)
    aValue = Left(aValue, Len(aValue) - 1)
    fCPercentage2Dbl = Val(aValue) / 100
End Function

Function fCopyReadWholeSheetData2Array(shtToRead As Worksheet, ByRef arrDataOut() _
            , Optional dictLetterIndex As Dictionary, Optional alDataFromRow As Long = 2)
    Dim lMaxRow As Long
    Dim lMaxCol As Long
    
    lMaxRow = fGetValidMaxRow(shtToRead)
    If lMaxRow < alDataFromRow Then arrDataOut = Array(): Exit Function
    
    If dictLetterIndex Is Nothing Then
        lMaxCol = fGetValidMaxCol(shtToRead)
    Else
        lMaxCol = WorksheetFunction.Max(dictLetterIndex.Items)
    End If
    
    arrDataOut = fReadRangeDatatoArrayByStartEndPos(shtToRead, alDataFromRow, 1, lMaxRow, lMaxCol)
End Function

Function fReadMasterSheetData(asFileTag As String, Optional shtData As Worksheet, Optional asDataFromRow As Long = 2 _
        , Optional bNoDataError As Boolean = False)
    Call fReadSheetDataByConfig(asFileTag:=asFileTag, dictColIndex:=gDictMstColIndex, arrDataOut:=arrMaster _
                                , dictColFormat:=gDictMstCellFormat _
                                , dictRawType:=gDictMstRawType _
                                , dictDisplayName:=gDictMstDisplayName _
                                , alDataFromRow:=asDataFromRow _
                                , shtData:=shtData)
    If bNoDataError Then Call fCheckIfSheetHasNodata_RaiseErrToStop(arrMaster)
End Function
