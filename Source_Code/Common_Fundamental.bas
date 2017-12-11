Attribute VB_Name = "Common_Fundamental"
Option Explicit
Option Base 1

Public Function fGetValidMaxRow(shtParam As Worksheet, Optional abCountInMergedCell As Boolean = False) As Long
    Dim lExcelMaxRow As Long
    Dim lUsedMaxRow As Long
    Dim lUsedMaxCol As Long
    
    lExcelMaxRow = shtParam.Rows.Count
    lUsedMaxRow = shtParam.UsedRange.Row + shtParam.UsedRange.Rows.Count - 1
    lUsedMaxCol = shtParam.UsedRange.Column + shtParam.UsedRange.Columns.Count - 1
    
    If lUsedMaxRow = 1 Then
        If shtParam.UsedRange.Address = "$A$1" And Len(shtParam.Range("A1")) <= 0 Then
            fGetValidMaxRow = 0
            Exit Function
        End If
    End If
    
    Dim lEachCol As Long
    Dim lValidMaxRowSaved As Long
    Dim lEachValidMaxRow As Long
    
    lValidMaxRowSaved = 0
    
    For lEachCol = 1 To lUsedMaxCol
        lEachValidMaxRow = shtParam.Cells(lExcelMaxRow, lEachCol).End(xlUp).Row
        
        If lEachValidMaxRow >= lUsedMaxRow Then
            fGetValidMaxRow = lEachValidMaxRow
            Exit Function
        End If
        
        If abCountInMergedCell Then
            If shtParam.Cells(lEachValidMaxRow, lEachCol).MergeCells Then
                lEachValidMaxRow = shtParam.Cells(lEachValidMaxRow, lEachCol).MergeArea.Row _
                                 + shtParam.Cells(lEachValidMaxRow, lEachCol).MergeArea.Rows.Count - 1
            End If
        End If
        
        If lEachValidMaxRow > lValidMaxRowSaved Then lValidMaxRowSaved = lEachValidMaxRow
    Next
    
    fGetValidMaxRow = lValidMaxRowSaved
End Function

Public Function fGetValidMaxCol(shtParam As Worksheet, Optional abCountInMergedCell As Boolean = False) As Long
    Dim lExcelMaxCol As Long
    Dim lUsedMaxRow As Long
    Dim lUsedMaxCol As Long
    
    lExcelMaxCol = shtParam.Columns.Count
    lUsedMaxRow = shtParam.UsedRange.Row + shtParam.UsedRange.Rows.Count - 1
    lUsedMaxCol = shtParam.UsedRange.Column + shtParam.UsedRange.Columns.Count - 1
    
    If lUsedMaxRow = 1 Then
        If shtParam.UsedRange.Address = "$A$1" And Len(shtParam.Range("A1")) <= 0 Then
            fGetValidMaxCol = 0
            Exit Function
        End If
    End If
    
    Dim lEachRow As Long
    Dim lValidMaxColSaved As Long
    Dim lEachValidMaxCol As Long
    
    lValidMaxColSaved = 0
    
    For lEachRow = 1 To lUsedMaxRow
        lEachValidMaxCol = shtParam.Cells(lEachRow, lExcelMaxCol).End(xlToLeft).Column
        
        If lEachValidMaxCol >= lUsedMaxCol Then
            fGetValidMaxCol = lEachValidMaxCol
            Exit Function
        End If
        
        If abCountInMergedCell Then
            If shtParam.Cells(lEachRow, lEachValidMaxCol).MergeCells Then
                lEachValidMaxCol = shtParam.Cells(lEachRow, lEachValidMaxCol).MergeArea.Column _
                                 + shtParam.Cells(lEachRow, lEachValidMaxCol).MergeArea.Columns.Count - 1
            End If
        End If
        
        If lEachValidMaxCol > lValidMaxColSaved Then lValidMaxColSaved = lEachValidMaxCol
    Next
    
    fGetValidMaxCol = lValidMaxColSaved
End Function

Function fGetValidMaxRowOfRange(rngParam As Range, Optional abCountInMergedCell As Boolean = False) As Long
     Dim lOut As Long
     
     'single cell
     If fRangeIsSingleCell(rngParam) Then lOut = rngParam.Row:               GoTo exit_fun
     
     Dim shtParent As Worksheet
     Set shtParent = rngParam.Parent
     
     Dim lExcelMaxRow As Long
     Dim lExcelMaxCol As Long
     Dim lShtValidMaxRow As Long
     Dim lShtValidMaxCol As Long
     Dim lRangeMaxRow As Long
     Dim lRangeMaxCol As Long
     Dim lValidMaxRowSaved As Long
     Dim lEachValidMaxRow As Long
     Dim lEachCol As Long
     
     lExcelMaxRow = shtParent.Rows.Count
     lExcelMaxCol = shtParent.Columns.Count
     lRangeMaxRow = rngParam.Row + rngParam.Rows.Count - 1
     lRangeMaxCol = rngParam.Column + rngParam.Columns.Count - 1
     
     lShtValidMaxRow = fGetValidMaxRow(shtParent, abCountInMergedCell)
     If lShtValidMaxRow < rngParam.Row Then 'blank, out of usedrange
        lOut = rngParam.Row: GoTo exit_fun
     End If
     
     lShtValidMaxCol = fGetValidMaxCol(shtParent, abCountInMergedCell)
     If lShtValidMaxCol < rngParam.Column Then 'blank, out of usedrange
        lOut = rngParam.Row: GoTo exit_fun
     End If
     
     'whole sheet
     If rngParam.Rows.Count = lExcelMaxRow And rngParam.Columns.Count = lExcelMaxCol Then
        lOut = lShtValidMaxRow: GoTo exit_fun
     End If
     
     If lRangeMaxRow > lShtValidMaxRow Then 'shrink row
        lRangeMaxRow = lShtValidMaxRow
     End If
     If lRangeMaxCol > lShtValidMaxCol Then 'shrink col
        lRangeMaxCol = lShtValidMaxCol
     End If
     
'     'several rows
'     If rngParam.Columns.Count = lExcelMaxCol Then
'        lOut = lRangeMaxRow: GoTo exit_fun
'     End If
     
     'several columns
     If rngParam.Rows.Count = lExcelMaxRow Then
        lValidMaxRowSaved = 0
        
        For lEachCol = rngParam.Column To lRangeMaxCol
            lEachValidMaxRow = shtParent.Cells(lExcelMaxRow, lEachCol).End(xlUp).Row
            
            If lEachValidMaxRow >= lShtValidMaxRow Then
                lOut = lShtValidMaxRow
                GoTo exit_fun
            End If
            
            If abCountInMergedCell Then
                If shtParent.Cells(lEachValidMaxRow, lEachCol).MergeCells Then
                    lEachValidMaxRow = shtParent.Cells(lEachValidMaxRow, lEachCol).MergeArea.Row _
                                     + shtParent.Cells(lEachValidMaxRow, lEachCol).MergeArea.Rows.Count - 1
                End If
            End If
            
            If lEachValidMaxRow > lValidMaxRowSaved Then lValidMaxRowSaved = lEachValidMaxRow
        Next
        
        lOut = lValidMaxRowSaved: GoTo exit_fun
    End If
    
    Dim arrShrunk()
    Dim lArrMaxRow As Long
    Dim lArrMaxCol As Long
    arrShrunk = fReadRangeDatatoArrayByStartEndPos(shtParent, rngParam.Row, rngParam.Column, lRangeMaxRow, lRangeMaxCol)
    lArrMaxRow = fGetArrayMaxValidRowCol(arrShrunk, lArrMaxCol)
    Erase arrShrunk
    
    lArrMaxRow = rngParam.Row + lArrMaxRow + IIf(lArrMaxRow > 0, -1, 0)
    lArrMaxCol = rngParam.Column + lArrMaxCol + IIf(lArrMaxCol > 0, -1, 0)
    
    lOut = lArrMaxRow
    lEachValidMaxRow = 0
    If abCountInMergedCell Then
        If shtParent.Cells(lArrMaxRow, lArrMaxCol).MergeCells Then
        '    If shtParent.Cells(lOut, lEachCol).MergeArea.Rows.Count > 1 Then
                lEachValidMaxRow = shtParent.Cells(lArrMaxRow, lArrMaxCol).MergeArea.Row _
                                 + shtParent.Cells(lArrMaxRow, lArrMaxCol).MergeArea.Rows.Count - 1
         '   End If
        End If
        
        If lEachValidMaxRow > lOut Then lOut = lEachValidMaxRow
    End If

exit_fun:
    fGetValidMaxRowOfRange = lOut
End Function

Function fGetArrayMaxValidRowCol(arrParam(), Optional lMaxCol As Long, Optional bReverse As Boolean = True) As Long
    Dim lEachRow As Long
    Dim lEachMaxRow As Long
    Dim lMaxRowSaved As Long
    Dim lEachCol As Long

    lMaxCol = 0
    lMaxRowSaved = 0 'UBound(arrParam, 1) - LBound(arrParam, 1) + 1
    
    If bReverse Then
        For lEachRow = UBound(arrParam, 1) To LBound(arrParam, 1) Step -1
            For lEachCol = LBound(arrParam, 2) To UBound(arrParam, 2)
                If Len(Trim(CStr(arrParam(lEachRow, lEachCol)))) > 0 Then
                    If lEachRow > lMaxRowSaved Then
                        lMaxRowSaved = lEachRow
                        lMaxCol = lEachCol
                        GoTo exit_fun
                    End If
                End If
            Next
        Next
    Else
        For lEachRow = LBound(arrParam, 1) To UBound(arrParam, 1)
            For lEachCol = LBound(arrParam, 2) To UBound(arrParam, 2)
                If Len(Trim(CStr(arrParam(lEachRow, lEachCol)))) > 0 Then
                    If lEachRow > lMaxRowSaved Then
                        lMaxRowSaved = lEachRow
                        lMaxCol = lEachCol
                        GoTo exit_fun
                    End If
                End If
            Next
        Next
    End If
    
exit_fun:
    fGetArrayMaxValidRowCol = lMaxRowSaved
End Function

Function fRangeIsSingleCell(rngParam As Range) As Boolean
    fRangeIsSingleCell = (rngParam.Rows.Count = 1 And rngParam.Columns.Count = 1)
End Function


Function fErr(Optional sMsg As String = "") As VbMsgBoxResult
    If fNzero(sMsg) Then fMsgBox "Error: " & vbCr & vbCr & sMsg, vbCritical
    
    err.Raise vbObjectError + ERROR_NUMBER, "", "Program is to be terminated."
End Function

Function fMsgBox(Optional sMsg As String = "", Optional aVbMsgBoxStyle As VbMsgBoxStyle = vbCritical) As VbMsgBoxResult
    fMsgBox = MsgBox(sMsg, aVbMsgBoxStyle)
End Function

Function fSelectFileDialog(Optional asDefaultFilePath As String = "" _
                         , Optional asFileFilters As String = "" _
                         , Optional asTitle As String = "") As String
    'asFileFilters :   Excel File=*.xls;*.xlsx;*.xlx
    Dim fd As FileDialog
    Dim sFilterDesc As String
    Dim sFilterStr As String
    Dim sDefaultFile As String
    
    If Len(Trim(asFileFilters)) > 0 Then
        sFilterDesc = Trim(Split(asFileFilters, "=")(0))
        sFilterStr = Trim(Split(asFileFilters, "=")(1))
    End If
    
    If Len(Trim(asDefaultFilePath)) > 0 Then
       ' sDefaultFile = fGetFileParentFolder(asDefaultFilePath)
        sDefaultFile = asDefaultFilePath
    Else
        sDefaultFile = ThisWorkbook.Path
    End If
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    fd.InitialFileName = sDefaultFile
    fd.Title = IIf(Len(asTitle) > 0, asTitle, fd.InitialFileName)
    fd.AllowMultiSelect = False
    
    If Len(Trim(sFilterStr)) > 0 Then
        fd.Filters.Clear
        fd.Filters.Add sFilterDesc, sFilterStr, 1
        fd.FilterIndex = 1
        fd.InitialView = msoFileDialogViewDetails
    Else
        If fd.Filters.Count > 0 Then fd.Filters.Delete
    End If

    If fd.Show = -1 Then
        fSelectFileDialog = fd.SelectedItems(1)
    Else
        fSelectFileDialog = ""
    End If
        
    Set fd = Nothing
End Function

Function fGetFileParentFolder(asFileFullPath As String) As String
    fGetFSO
    fGetFileParentFolder = gFSO.GetParentFolderName(asFileFullPath)
End Function

Function fGetFileBaseName(asFileFullPath As String) As String
    fGetFSO
    fGetFileBaseName = gFSO.GetFileName(asFileFullPath)
End Function
Function fGetFileNetName(asFileFullPath As String) As String
    fGetFSO
    fGetFileNetName = gFSO.GetBaseName(asFileFullPath)
End Function
Function fGetFileExtension(asFileFullPath As String, Optional bDot As Boolean = False) As String
    fGetFSO
    fGetFileExtension = IIf(bDot, ".", "") & gFSO.GetExtensionName(asFileFullPath)
End Function

'Function fGetFileNamePart(asFileFullPath As String _
'                        , Optional ByRef sParentFolder As String _
'                        , Optional ByRef sFileBaseName As String _
'                        , Optional ByRef sFileExtension As String _
'                        , Optional ByRef sFileNetName As String) As String
'    If Len(Trim(asFileFullPath)) <= 0 Then Exit Function
'
'    sParentFolder = fso.GetParentFolderName(asFileFullPath)
'    sFileBaseName = fso.GetFileName(asFileFullPath)
'    sFileExtension = fso.GetExtensionName(asFileFullPath)
'    sFileNetName = fso.GetBaseName(asFileFullPath)
'
'    Set fso = Nothing
'End Function

Function fArrayHasBlankValue(ByRef arrParam) As Boolean
    Dim bOut As Boolean
    
    bOut = False
    If fArrayIsEmpty(arrParam) Then GoTo exit_function
    
    Dim iDimension As Integer
    Dim lEachRow As Long
    Dim lEachCol As Long
    
    iDimension = fGetArrayDimension(arrParam)
    If iDimension <= 0 Then GoTo exit_function
    If iDimension >= 2 Then fErr "2 dimensions is not supported."
    
    For lEachRow = LBound(arrParam, 1) To UBound(arrParam, 1)
        If Len(Trim(CStr(arrParam(lEachRow)))) <= 0 Then
            bOut = True
            GoTo exit_function
        End If
    Next
    
exit_function:
    fArrayHasBlankValue = bOut
End Function
Function fArrayHasDuplicateElement(ByRef arrParam) As Boolean
    Dim bOut As Boolean
    
    bOut = False
    If fArrayIsEmpty(arrParam) Then GoTo exit_function
    
    Dim iDimension As Integer
    Dim lEachRow As Long
    
    iDimension = fGetArrayDimension(arrParam)
    If iDimension <= 0 Then GoTo exit_function
    If iDimension >= 2 Then fErr "2 dimensions is not supported."
    
    Dim dict As New Dictionary
    
    For lEachRow = LBound(arrParam, 1) To UBound(arrParam, 1)
        If dict.Exists(arrParam(lEachRow)) Then
            bOut = True
            GoTo exit_function
        Else
            dict.Add arrParam(lEachRow), 0
        End If
    Next
    
exit_function:
    fArrayHasDuplicateElement = bOut
    Set dict = Nothing
End Function

Function fArrayIsEmptyOrNoData(ByRef arrParam) As Boolean
    Dim bOut As Boolean
    
    bOut = True
    If fArrayIsEmpty(arrParam) Then GoTo exit_function
    
    Dim iDimension As Integer
    Dim lEachRow As Long
    Dim lEachCol As Long
    
    iDimension = fGetArrayDimension(arrParam)
    If iDimension <= 0 Then GoTo exit_function
    If iDimension >= 3 Then fErr "3 dimensions is not supported."
    
    If iDimension = 1 Then
        For lEachRow = LBound(arrParam, 1) To UBound(arrParam, 1)
            If Len(Trim(CStr(arrParam(lEachRow)))) > 0 Then
                bOut = False
                GoTo exit_function
            End If
        Next
    ElseIf iDimension = 2 Then
        For lEachRow = LBound(arrParam, 1) To UBound(arrParam, 1)
            For lEachCol = LBound(arrParam, 2) To UBound(arrParam, 2)
                If Len(Trim(CStr(arrParam(lEachRow, lEachCol)))) > 0 Then
                    bOut = False
                    GoTo exit_function
                End If
            Next
        Next
    End If
    
exit_function:
    fArrayIsEmptyOrNoData = bOut
End Function

Function fArrayIsEmpty(ByRef arrParam) As Boolean
    Dim i As Long
    
    fArrayIsEmpty = True
    
    On Error Resume Next
    
    i = UBound(arrParam, 1)
    If err.Number = 0 Then
        If UBound(arrParam) < LBound(arrParam) Then
            Exit Function
        Else
            fArrayIsEmpty = False
        End If
    Else
        err.Clear
    End If
End Function
Function fGetArrayDimension(arrParam) As Integer
    Dim i As Integer
    Dim tmp As Long
    
    On Error GoTo error_exit
    
    i = 0
    Do While True
        i = i + 1
        tmp = UBound(arrParam, i)
        
        If tmp < 0 Then
            fGetArrayDimension = -1
            Exit Function
        End If
    Loop
    
error_exit:
    err.Clear
    fGetArrayDimension = i - 1
End Function

Function fNum2Letter(ByVal alNum As Long) As String
    fNum2Letter = Replace(Split(Columns(alNum).Address, ":")(1), "$", "")
End Function
Function fLetter2Num(ByVal alLetter As String) As Long
    fLetter2Num = Columns(alLetter).Column
End Function

Function fFileExists(sFilePath As String) As Boolean
    fGetFSO
    fFileExists = gFSO.FileExists(sFilePath)
End Function

Function fDeleteFile(sFilePath As String)
    If fFileExists(sFilePath) Then
        SetAttr sFilePath, vbNormal
        Kill sFilePath
    End If
End Function

Function fArrayRowIsBlankHasNoData(arr, alRow As Long) As Boolean
    Dim bOut As Boolean
    Dim lEachCol As Long
    
    bOut = True
    For lEachCol = LBound(arr, 2) To UBound(arr, 2)
        If Len(Trim(CStr(arr(alRow, lEachCol)))) > 0 Then
            bOut = False
            Exit For
        End If
    Next
    
    fArrayRowIsBlankHasNoData = bOut
End Function

Function fGenRandomUniqueString() As String
    fGenRandomUniqueString = Format(Now(), "yyyymmddhhMMSS") & Rnd()
End Function

Function fSplit(asOrig As String, Optional asSeparators As String = "") As Variant
    If Len(asSeparators) <= 0 Then asSeparators = ":;|, " & vbLf
    
    Dim tDelimiter As String
    tDelimiter = Chr(130)   'a non-printable charactor
    
    Dim sTransFormed As String
    Dim sEachDeli As String
    Dim i As Integer
    
    sTransFormed = asOrig
    For i = 1 To Len(asSeparators)
        sEachDeli = Mid(asSeparators, i, 1)
        sTransFormed = Replace(sTransFormed, sEachDeli, tDelimiter)
    Next
    
    While InStr(sTransFormed, tDelimiter & tDelimiter) > 0
        sTransFormed = Replace(sTransFormed, tDelimiter & tDelimiter, tDelimiter)
    Wend
    
    sTransFormed = fTrim(sTransFormed, tDelimiter)
    
    fSplit = Split(sTransFormed, tDelimiter)
End Function

Function fSplitJoin(asOrig As String, Optional asSeparators As String = "", Optional asNewSep As String = DELIMITER) As String
    If Len(asSeparators) <= 0 Then asSeparators = ":;|, " & vbLf
    
    Dim arr
    arr = fSplit(asOrig, asSeparators)
    fSplitJoin = Join(arr, asNewSep)
    
    Erase arr
End Function

Function fJoin(asOrig As String, Optional asNewSep As String = DELIMITER) As String
    fJoin = fSplitJoin(asOrig, , asNewSep)
End Function

Function fTrim(asOrig As String, Optional asWhatToTrim As String = " " & vbTab & vbCr & vbLf) As String
    Dim sOut As String
    
    sOut = Trim(asOrig)
    While InStr(asWhatToTrim, Left(sOut, 1)) > 0
        sOut = Right(sOut, Len(sOut) - 1)
    Wend
    
    While InStr(asWhatToTrim, Right(sOut, 1)) > 0
        sOut = Left(sOut, Len(sOut) - 1)
    Wend
    
    fTrim = sOut
End Function

Function fRTrim(asOrig As String, Optional asWhatToTrim As String = " " & vbTab & vbCr & vbLf) As String
    Dim sOut As String
    
    sOut = Trim(asOrig)
    While InStr(asWhatToTrim, Right(sOut, 1)) > 0
        sOut = Left(sOut, Len(sOut) - 1)
    Wend
    
    fTrim = sOut
End Function

Function fLTrim(asOrig As String, Optional asWhatToTrim As String = " " & vbTab & vbCr & vbLf) As String
    Dim sOut As String
    
    sOut = Trim(asOrig)
    While InStr(asWhatToTrim, Left(sOut, 1)) > 0
        sOut = Right(sOut, Len(sOut) - 1)
    Wend
    
    fTrim = sOut
End Function

Function fLen(sStr) As Long
    fLen = Len(Trim(sStr))
End Function

Function fZero(sStr) As Boolean
    fZero = (Len(Trim(sStr)) <= 0)
End Function

Function fNzero(sStr) As Boolean
    fNzero = (Len(Trim(sStr)) > 0)
End Function

Function fReadArray2DictionaryWithMultipleColsCombined(arrData, lKeyCol As Long, arrItemCols _
                , Optional asDelimiter As String = "" _
                , Optional IgnoreBlankKeys As Boolean = False _
                , Optional WhenKeyDuplicateThenError As Boolean = True) As Dictionary
    Dim dictOut As Dictionary
    
    If lKeyCol < 0 Then fErr "lKeyCol < 0 to fReadArray2DictionaryWithMultipleColsCombined"
    
    Set dictOut = New Dictionary
    
    If fArrayIsEmptyOrNoData(arrData) Then GoTo exit_fun
    If fArrayIsEmptyOrNoData(arrItemCols) Then fErr "arrItemCols is empty."
    If fArrayHasDuplicateElement(arrItemCols) Then fErr "arrItemCols has duplicate element."
    
    Dim i As Long
    Dim sKey
    Dim sValue As String
    For i = LBound(arrData, 1) To UBound(arrData, 1)
        If fArrayRowIsBlankHasNoData(arrData, i) Then GoTo next_row
        
        sKey = arrData(i, lKeyCol)
        
        If fZero(sKey) Then
            If IgnoreBlankKeys Then GoTo next_row
            fErr "Key column is blank, but program decides not allow blank, pls contact with IT support."
        End If
        
        If dictOut.Exists(sKey) Then
            If WhenKeyDuplicateThenError Then fErr "Duplicate Key was found : " & vbCr & "Key: " & sKey
            GoTo next_row
        End If
        
        sValue = ""
        Dim j As Integer
        For j = LBound(arrItemCols) To UBound(arrItemCols)
            sValue = sValue & asDelimiter & Trim(arrData(i, arrItemCols(j)))
        Next
        
        If fNzero(sValue) And fNzero(asDelimiter) Then
            If Left(sValue, 1) = asDelimiter Then sValue = Right(sValue, Len(sValue) - 1)
        End If
        
        dictOut.Add sKey, sValue
next_row:
    Next
    
exit_fun:
    Set fReadArray2DictionaryWithMultipleColsCombined = dictOut
    Set dictOut = Nothing
End Function

Function fRadArray2DictionaryOnlyKeys(arrParam, lKeyCol As Long _
                            , Optional IgnoreBlankKeys As Boolean = False _
                            , Optional WhenKeyIsDuplicateError As Boolean = True) As Dictionary
'==========================================================================
'lItemCol
'         -1: the item is row number
'          0: get key only, not care the item value, 0 as default
'         >0: the item is specified column
'==========================================================================
    Set fRadArray2DictionaryOnlyKeys = fRadArray2Dictionary(arrParam, lKeyCol, 0, IgnoreBlankKeys, WhenKeyIsDuplicateError)
End Function

Private Function fRadArray2Dictionary(arrParam, lKeyCol As Long _
                            , Optional lItemCol As Long = 0 _
                            , Optional IgnoreBlankKeys As Boolean = False _
                            , Optional WhenKeyIsDuplicateError As Boolean = True) As Dictionary
'==========================================================================
'lItemCol
'         -1: the item is row number
'          0: get key only, not care the item value, 0 as default
'         >0: the item is specified column
'==========================================================================
    If lItemCol < -1 Or lKeyCol <= 0 Then fErr "wrong param"
    
    Dim dictOut As Dictionary
    
    Set dictOut = New Dictionary
    If fArrayIsEmptyOrNoData(arrParam) Then GoTo exit_fun
    
    Dim bGetKeyOnly As Boolean
    Dim bGetRowNo As Boolean
    
    bGetKeyOnly = (lItemCol = 0)
    bGetRowNo = (lItemCol = -1)
    
    Dim i As Long
    Dim sKey As String
    Dim sValue As String
    
    For i = LBound(arrParam, 1) To UBound(arrParam, 1)
        If fArrayRowIsBlankHasNoData(arrParam, i) Then GoTo next_row
        
        sKey = Trim(arrParam(i, lKeyCol))
        
        If Len(sKey) <= 0 Then
            If Not IgnoreBlankKeys Then
                fErr "Key col is blank, but you specified IgnoreBlankKeys = false" & vbCr & lKeyCol
            Else
                GoTo next_row
            End If
        End If
        
        If dictOut.Exists(sKey) Then
            If WhenKeyIsDuplicateError Then
                fErr "duplicate key was found:, but you specified IgnoreBlankKeys = false" & vbCr & lKeyCol & vbCr & sKey
            Else
                GoTo next_row
            End If
        End If
        
        If bGetRowNo Then
            dictOut.Add sKey, i
        Else
            If bGetKeyOnly Then
                dictOut.Add sKey, 0
            Else
                dictOut.Add sKey, arrParam(i, lItemCol)
            End If
        End If
next_row:
    Next
    
exit_fun:
    Set fRadArray2Dictionary = dictOut
    Set dictOut = Nothing
End Function

Function fValidateDuplicateInArray(arrParam, arrKeyColsOrSingle _
                        , Optional bAllowBlank As Boolean = False _
                        , Optional shtAt As Worksheet _
                        , Optional lHeaderAtRow As Long = 1, Optional lStartCol As Long _
                        , Optional ByVal sMsgColHeader As String)
'arrKeyColsOrSingle : should be start from 1, since two dimension array is starting from 1
    If fArrayIsEmptyOrNoData(arrParam) Then Exit Function
    
    If IsArray(arrKeyColsOrSingle) Then
        If fArrayIsEmptyOrNoData(arrKeyColsOrSingle) Then fErr "Wrong param: arrKeyColsOrSingle"
    Else
        If arrKeyColsOrSingle <= 0 Then fErr "Wrong param: arrKeyColsOrSingle"
    End If
    
    If IsArray(arrKeyColsOrSingle) Then
        Call fValidateDuplicateInArrayForCombineCols(arrParam:=arrParam, arrKeyCols:=arrKeyColsOrSingle _
                                                    , bAllowBlankIgnore:=bAllowBlank _
                                                    , shtAt:=shtAt _
                                                    , lHeaderAtRow:=lHeaderAtRow, lStartCol:=lStartCol _
                                                    , sMsgColHeader:=sMsgColHeader)
    Else
        Call fValidateDuplicateInArrayForSingleCol(arrParam:=arrParam, lKeyCol:=CLng(arrKeyColsOrSingle) _
                                                    , bAllowBlankIgnore:=bAllowBlank _
                                                    , shtAt:=shtAt _
                                                    , lHeaderAtRow:=lHeaderAtRow, lStartCol:=lStartCol _
                                                    , sMsgColHeader:=sMsgColHeader)
    End If
End Function

Function fValidateDuplicateInArrayForCombineCols(arrParam, arrKeyCols _
                        , Optional bAllowBlankIgnore As Boolean = False _
                        , Optional shtAt As Worksheet _
                        , Optional lHeaderAtRow As Long = 1, Optional lStartCol As Long _
                        , Optional sMsgColHeader As String)
'MultipleCols: means MultipleCols composed as key
'for MultipleCols that is individually, please refer to function fValidateDuplicateInArrayIndividually
    Const DELI = " " & DELIMITER & " "
    
    Dim lEachRow As Long
    Dim lEachCol As Long
    Dim i As Long
    Dim sKeyStr As String
    Dim sColLetterStr As String
    Dim dict As Dictionary
    Dim sPos As String
    Dim lActualRow As Long
    
    If Not fZero(sMsgColHeader) Then
        sColLetterStr = sMsgColHeader
    Else
        For i = LBound(arrKeyCols) To UBound(arrKeyCols)
            lEachCol = arrKeyCols(i)
            sColLetterStr = sColLetterStr & " + " & fNum2Letter(lStartCol + lEachCol - 1)
        Next
        sColLetterStr = Right(sColLetterStr, Len(sColLetterStr) - 3)
    End If
    
    sPos = vbCr & vbCr & "sheet     : " & shtAt.Name _
         & vbCr & vbCr & "Row, Column: " & " ACTUAL_ROW_NO, [" & sColLetterStr & "]"
            
    Set dict = New Dictionary
    For lEachRow = LBound(arrParam, 1) To UBound(arrParam, 1)
        If fArrayRowIsBlankHasNoData(arrParam, lEachRow) Then GoTo next_row
        
        lActualRow = (lHeaderAtRow + lEachRow)
        
        sKeyStr = ""
        For i = LBound(arrKeyCols) To UBound(arrKeyCols)
            lEachCol = arrKeyCols(i)
            sKeyStr = sKeyStr & DELI & Trim(CStr(arrParam(lEachRow, lEachCol)))
        Next
        
        If fZero(Replace(sKeyStr, DELI, "")) Then
            If Not bAllowBlankIgnore Then
                'sPos = sPos & "[" & lActualRow & ", " & sColLetterStr & "]"
                sPos = Replace(sPos, "ACTUAL_ROW_NO", lActualRow)
                fErr "Keys [" & sKeyStr & "] is blank!" & sPos
            End If
            
            GoTo next_row
        End If
        
        sKeyStr = Right(sKeyStr, Len(sKeyStr) - Len(DELI))
        
        If dict.Exists(sKeyStr) Then
            sPos = Replace(sPos, "ACTUAL_ROW_NO", lActualRow)
            fErr "Duplicate key [" & sKeyStr & " was found:" & sPos
        Else
            dict.Add sKeyStr, 0
        End If
next_row:
    Next
    
    Set dict = Nothing
End Function

Function fValidateDuplicateInArrayForSingleCol(arrParam, lKeyCol As Long _
                                            , Optional bAllowBlankIgnore As Boolean = False _
                                            , Optional shtAt As Worksheet _
                                            , Optional lHeaderAtRow As Long = 1, Optional lStartCol As Long _
                                            , Optional sMsgColHeader As String)
    If lKeyCol <= 0 Then fErr "Wrong param: lKeyCol"
    
    Dim lEachRow As Long
    Dim i As Long
    Dim sKeyStr As String
    Dim sColLetter As String
    Dim dict As Dictionary
    Dim sPos As String
    Dim lActualRow As Long
    
    sColLetter = IIf(fZero(sMsgColHeader), fNum2Letter(lStartCol + lKeyCol - 1), sMsgColHeader)
        
    sPos = vbCr & vbCr & "sheet     : " & shtAt.Name _
         & vbCr & vbCr & "Row, Column: " & " ACTUAL_ROW_NO,  [" & sColLetter & "]"
         
    Set dict = New Dictionary
    For lEachRow = LBound(arrParam, 1) To UBound(arrParam, 1)
        If fArrayRowIsBlankHasNoData(arrParam, lEachRow) Then GoTo next_row
        
        lActualRow = (lHeaderAtRow + lEachRow)
        
        sKeyStr = Trim(CStr(arrParam(lEachRow, lKeyCol)))
        
        If fZero(sKeyStr) Then
            If Not bAllowBlankIgnore Then
                'sPos = sPos & lActualRow & " / " & sColLetter
                sPos = Replace(sPos, "ACTUAL_ROW_NO", lActualRow)
                fErr "Keys [" & sColLetter & "] is blank!" & sPos
            End If
            
            GoTo next_row
        End If
        
        If dict.Exists(sKeyStr) Then
            'sPos = sPos & lActualRow & " / " & sColLetter
            sPos = Replace(sPos, "ACTUAL_ROW_NO", lActualRow)
            fErr "Duplicate key [" & sKeyStr & "] was found " & sPos
        Else
            dict.Add sKeyStr, 0
        End If
next_row:
    Next
    
    Set dict = Nothing
End Function

Function fValidateDuplicateInArrayIndividually(arrParam, arrKeyColsOrSingle _
                        , Optional bAllowBlank As Boolean = False _
                        , Optional shtAt As Worksheet _
                        , Optional lHeaderAtRow As Long = 1, Optional lStartCol As Long _
                        , Optional arrColNames)
'arrKeyColsOrSingle : should be start from 1, since two dimension array is starting from 1
    If fArrayIsEmptyOrNoData(arrParam) Then Exit Function
    
    If IsArray(arrKeyColsOrSingle) Then
        If fArrayIsEmptyOrNoData(arrKeyColsOrSingle) Then fErr "Wrong param: arrKeyColsOrSingle"
    Else
        If arrKeyColsOrSingle <= 0 Then fErr "Wrong param: arrKeyColsOrSingle"
    End If
    
    If IsArray(arrKeyColsOrSingle) Then
        For i = LBound(arrKeyColsOrSingle) To UBound(arrKeyColsOrSingle)
            Call fValidateDuplicateInArrayForSingleCol(arrParam:=arrParam, arrKeyColsOrSingle:=arrKeyColsOrSingle _
                                                    , bAllowBlank:=bAllowBlank _
                                                    , shtAt:=shtAt _
                                                    , lHeaderAtRow:=lHeaderAtRow, lStartCol:=lStartCol _
                                                    , arrColNames:=arrColNames)
        Next
    Else
        Call fValidateDuplicateInArrayForSingleCol(arrParam:=arrParam, arrKeyColsOrSingle:=arrKeyColsOrSingle _
                                                    , bAllowBlank:=bAllowBlank _
                                                    , shtAt:=shtAt _
                                                    , lHeaderAtRow:=lHeaderAtRow, lStartCol:=lStartCol _
                                                    , arrColNames:=arrColNames)
    End If
End Function

Function SheetExists(asShtName As String, Optional wb As Workbook) As Boolean
    asShtName = Trim(asShtName)
    If wb Is Nothing Then Set wb = ThisWorkbook
    
    SheetExists = False
    
    Dim sht As Worksheet
    For Each sht In wb.Worksheets
        If sht.Name = asShtName Then
            SheetExists = True
            Exit For
        End If
    Next
End Function

Function fDeleteSheetIfExists(asShtName As String, Optional wb As Workbook)
    asShtName = Trim(asShtName)
    If wb Is Nothing Then Set wb = ThisWorkbook
    
    If SheetExists(asShtName, wb) Then
        Call fDeleteSheet(asShtName, wb)
    End If
End Function

Function fDeleteSheet(asShtName As String, Optional wb As Workbook)
    Dim bEnableEventsOrig As Boolean
    Dim bDisplayAlertsOrig As Boolean
    
    asShtName = Trim(asShtName)
    If wb Is Nothing Then Set wb = ThisWorkbook
    
    bEnableEventsOrig = Application.EnableEvents
    bDisplayAlertsOrig = Application.DisplayAlerts
    
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    
    wb.Worksheets(asShtName).Delete
    
    Application.EnableEvents = bEnableEventsOrig
    Application.DisplayAlerts = bDisplayAlertsOrig
End Function

Function fAddNewSheet(asShtName As String, Optional wb As Workbook) As Worksheet
    Dim shtOut As Worksheet
    
    asShtName = Trim(asShtName)
    If wb Is Nothing Then Set wb = ThisWorkbook

    Set shtOut = wb.Worksheets.Add(after:=wb.Worksheets(wb.Worksheets.Count))
    shtOut.Name = asShtName
    shtOut.Activate
    ActiveWindow.DisplayGridlines = False
    
    Set fAddNewSheet = shtOut
    Set shtOut = Nothing
End Function

Function fAddNewSheetDeleteFirst(asShtName As String, Optional wb As Workbook) As Worksheet
    asShtName = Trim(asShtName)
    If wb Is Nothing Then Set wb = ThisWorkbook
    
    Call fDeleteSheetIfExists(asShtName, wb)
    Set fAddNewSheetDeleteFirst = fAddNewSheet(asShtName, wb)
End Function

Function fGetSheetWhenNotExistsCreate(asShtName As String, Optional wb As Workbook) As Worksheet
    asShtName = Trim(asShtName)
    If wb Is Nothing Then Set wb = ThisWorkbook
    
    If SheetExists(asShtName, wb) Then
        Set fGetSheetWhenNotExistsCreate = wb.Worksheets(asShtName)
    Else
        Set fGetSheetWhenNotExistsCreate = fAddNewSheet(asShtName, wb)
    End If
End Function

Function fGetFSO()
    If gFSO Is Nothing Then Set gFSO = New FileSystemObject
End Function

Function fDeleteAllFilesInFolder(sFolder As String)
    fGetFSO
    
    Dim aFile As File
    
    If gFSO.FolderExists(sFolder) Then
        For Each aFile In gFSO.GetFolder(sFolder).Files
            aFile.Delete True
        Next
    End If
End Function

Function fDeleteOldFilesInFolder(sFolder As String, lDays As Long)
    
End Function


Function fValidateBlankInArrayCombinedCols(arrParam, arrKeyColsOrSingle _
                        , Optional shtAt As Worksheet _
                        , Optional lHeaderAtRow As Long = 1, Optional lStartCol As Long _
                        , Optional sMsgColHeader As String)
'arrKeyColsOrSingle : should be start from 1, since two dimension array is starting from 1
    If fArrayIsEmptyOrNoData(arrParam) Then Exit Function
    
    If IsArray(arrKeyColsOrSingle) Then
        If fArrayIsEmptyOrNoData(arrKeyColsOrSingle) Then fErr "Wrong param: arrKeyColsOrSingle"
    Else
        If arrKeyColsOrSingle <= 0 Then fErr "Wrong param: arrKeyColsOrSingle"
    End If
    
    If IsArray(arrKeyColsOrSingle) Then
        Call fValidateBlankInArrayForCombineCols(arrParam:=arrParam, arrKeyCols:=arrKeyColsOrSingle _
                                                    , shtAt:=shtAt _
                                                    , lHeaderAtRow:=lHeaderAtRow, lStartCol:=lStartCol _
                                                    , sMsgColHeader:=sMsgColHeader)
    Else
        Call fValidateBlankInArrayForSingleCol(arrParam:=arrParam, lKeyCol:=CLng(arrKeyColsOrSingle) _
                                                    , shtAt:=shtAt _
                                                    , lHeaderAtRow:=lHeaderAtRow, lStartCol:=lStartCol _
                                                    , sMsgColHeader:=sMsgColHeader)
    End If
End Function

'call the parent function: fValidateBlankInArrayCombinedCols, not to call this function
Private Function fValidateBlankInArrayForCombineCols(arrParam, arrKeyCols _
                        , Optional shtAt As Worksheet _
                        , Optional lHeaderAtRow As Long = 1, Optional lStartCol As Long _
                        , Optional sMsgColHeader As String)
'MultipleCols: means MultipleCols composed as key
'for MultipleCols that is individually, please refer to function fValidateBlankInArrayIndividually
    Const DELI = " " & DELIMITER & " "
    
    Dim lEachRow As Long
    Dim lEachCol As Long
    Dim i As Long
    Dim sKeyStr As String
    Dim sColLetterStr As String
    Dim sPos As String
    Dim lActualRow As Long
    
    If Not fZero(sMsgColHeader) Then
        sColLetterStr = sMsgColHeader
    Else
        For i = LBound(arrKeyCols) To UBound(arrKeyCols)
            lEachCol = arrKeyCols(i)
            sColLetterStr = sColLetterStr & " + " & fNum2Letter(lStartCol + lEachCol - 1)
        Next
        sColLetterStr = Right(sColLetterStr, Len(sColLetterStr) - 3)
    End If
    
    sPos = vbCr & vbCr & "sheet     : " & shtAt.Name _
         & vbCr & vbCr & "Row, Column: " & " ACTUAL_ROW_NO, [" & sColLetterStr & "]"
    
    For lEachRow = LBound(arrParam, 1) To UBound(arrParam, 1)
        If fArrayRowIsBlankHasNoData(arrParam, lEachRow) Then GoTo next_row
        
        lActualRow = (lHeaderAtRow + lEachRow)
        
        sKeyStr = ""
        For i = LBound(arrKeyCols) To UBound(arrKeyCols)
            lEachCol = arrKeyCols(i)
            sKeyStr = sKeyStr & CStr(arrParam(lEachRow, lEachCol))
        Next
        
        If fZero(sKeyStr) Then
            sPos = Replace(sPos, "ACTUAL_ROW_NO", lActualRow)
            fErr "Keys [" & sKeyStr & "] is blank!" & sPos
        End If
next_row:
    Next
End Function

Function fValidateBlankInArrayForSingleCol(arrParam, lKeyCol As Long _
                                            , Optional shtAt As Worksheet _
                                            , Optional lHeaderAtRow As Long = 1, Optional lStartCol As Long _
                                            , Optional sMsgColHeader As String)
    If lKeyCol <= 0 Then fErr "Wrong param: lKeyCol"
    
    Dim lEachRow As Long
    Dim i As Long
    Dim sKeyStr As String
    Dim sColLetter As String
    Dim sPos As String
    Dim lActualRow As Long
    
    sColLetter = IIf(fZero(sMsgColHeader), fNum2Letter(lStartCol + lKeyCol - 1), sMsgColHeader)
        
    sPos = vbCr & vbCr & "sheet     : " & shtAt.Name _
         & vbCr & vbCr & "Row, Column: " & " ACTUAL_ROW_NO,  [" & sColLetter & "]"

    For lEachRow = LBound(arrParam, 1) To UBound(arrParam, 1)
        If fArrayRowIsBlankHasNoData(arrParam, lEachRow) Then GoTo next_row
        
        lActualRow = (lHeaderAtRow + lEachRow)
        
        sKeyStr = CStr(arrParam(lEachRow, lKeyCol))
    
        If fZero(sKeyStr) Then
            sPos = Replace(sPos, "ACTUAL_ROW_NO", lActualRow)
            fErr "Keys [" & sColLetter & "] is blank!" & sPos
        End If
next_row:
    Next
End Function

Function fValidateBlankInArray(arrParam, arrKeyColsOrSingle _
                        , Optional shtAt As Worksheet _
                        , Optional lHeaderAtRow As Long = 1, Optional lStartCol As Long _
                        , Optional ByVal sMsgColHeader As String)
'arrKeyColsOrSingle : should be start from 1, since two dimension array is starting from 1
    If fArrayIsEmptyOrNoData(arrParam) Then Exit Function
    
    If IsArray(arrKeyColsOrSingle) Then
        If fArrayIsEmptyOrNoData(arrKeyColsOrSingle) Then fErr "Wrong param: arrKeyColsOrSingle"
    Else
        If arrKeyColsOrSingle <= 0 Then fErr "Wrong param: arrKeyColsOrSingle"
    End If
    
    Dim i As Integer
    If IsArray(arrKeyColsOrSingle) Then
        For i = LBound(arrKeyColsOrSingle) To UBound(arrKeyColsOrSingle)
            Call fValidateBlankInArrayForSingleCol(arrParam:=arrParam, lKeyCol:=CLng(arrKeyColsOrSingle) _
                                                    , shtAt:=shtAt _
                                                    , lHeaderAtRow:=lHeaderAtRow, lStartCol:=lStartCol _
                                                    , sMsgColHeader:=sMsgColHeader)
        Next
    Else
        Call fValidateBlankInArrayForSingleCol(arrParam:=arrParam, lKeyCol:=CLng(arrKeyColsOrSingle) _
                                                    , shtAt:=shtAt _
                                                    , lHeaderAtRow:=lHeaderAtRow, lStartCol:=lStartCol _
                                                    , sMsgColHeader:=sMsgColHeader)
    End If
End Function


Function fEnlargeAray(ByRef arr, Optional aPreserve As Boolean = True, Optional lIncrementNum As Integer = 1) As Long
    fRedim arr, arrlen(arr) + 1, aPreserve
End Function


Function fEnlargeArayWithValue(ByRef arr, aValue, Optional aPreserve As Boolean = True, Optional lIncrementNum As Integer = 1) As Long
'    If fArrayIsEmpty(arr) Then
'        Redim arr
'        Exit Function
'    End If
    
    fRedim arr, ArrayLen(arr) + 1, aPreserve
    arr(UBound(arr)) = aValue
    fEnlargeArayWithValue = UBound(arr)
End Function

Function fRedim(ByRef arr, lNewUbound As Long, Optional aPreserve As Boolean = True)
    If fArrayIsEmpty(arr) Then
        If aPreserve Then
            ReDim arr(lNewUbound)
        End If
        Exit Function
    End If

    If Base0(arr) Then
        If aPreserve Then
            ReDim Preserve arr(0 To lNewUbound - 1)
        Else
            ReDim arr(0 To lNewUbound - 1)
        End If
    Else
        If aPreserve Then
            ReDim Preserve arr(1 To lNewUbound)
        Else
            ReDim arr(1 To lNewUbound)
        End If
    End If
End Function

Function ArrayLen(ByRef arr) As Long
    If fArrayIsEmpty(arr) Then
        ArrayLen = 0
        Exit Function
    '    fErr "Empty array is not allowed."
    End If
     ArrayLen = UBound(arr) - LBound(arr) + 1
End Function

Function Base0(ByRef arr) As Boolean
     Base0 = (LBound(arr) = 0)
End Function

Function fSetArrayValue(ByRef arr, aIndex As Long, aValue)
    If Base0(arr) Then
        arr(aIndex - 1) = aValue
    Else
        arr(aIndex) = aValue
    End If
End Function

Function fUpdateDictionaryItemValueForDelimitedElement(ByRef dict As Dictionary, aKey, iElementIndex As Integer, aNewValue, Optional sDelimiter As String = DELIMITER)
    Dim arr
    If iElementIndex <= 0 Then fErr "iElementIndex <= 0 to fUpdateDictionaryItemValueForDelimitedElement"
    
    If Not dict.Exists(aKey) Then fErr "aKey even does not exists in param dict to fUpdateDictionaryItemValueForDelimitedElement"
    
    arr = Split(dict(aKey), sDelimiter)
    
    If ArrayLen(arr) < iElementIndex Then
        fRedim arr, CLng(iElementIndex), True
    End If
    
    If Base0(arr) Then
        arr(iElementIndex - 1) = aNewValue
    Else
        arr(iElementIndex) = aNewValue
    End If
    
    dict(aKey) = Join(arr, sDelimiter)
    Erase arr
End Function

Function fCopyDictionaryKeys2Array(dict As Dictionary, ByRef arrOut())
    If dict.Count <= 0 Then
        arrOut = Array()
    End If
    
    ReDim arrOut(1 To dict.Count)
    
    Dim i As Long
    
    For i = 0 To dict.Count - 1
        arrOut(i + 1) = dict.Keys(i)
    Next
End Function
Function fCopyDictionaryItemsArray(dict As Dictionary, ByRef arrOut())
    If dict.Count <= 0 Then
        arrOut = Array()
    End If
    
    ReDim arrOut(1 To dict.Count)
    
    Dim i As Long
    
    For i = 0 To dict.Count - 1
        arrOut(i + 1) = dict.Items(i)
    Next
End Function

Function fEnableExcelOptionsAll()
    Call fEnableOrDisableExcelOptionsAll(True)
End Function

Function fDisableExcelOptionsAll()
    Call fEnableOrDisableExcelOptionsAll(False)
End Function
Function fEnableOrDisableExcelOptionsAll(bValue As Boolean)
    Application.ScreenUpdating = bValue
    Application.EnableEvents = bValue
    Application.DisplayAlerts = bValue
    Application.AskToUpdateLinks = bValue
    ThisWorkbook.CheckCompatibility = bValue
    
    If bValue Then
        Application.Calculation = xlCalculationAutomatic
    Else
        Application.Calculation = xlCalculationManual
    End If
End Function

Function fGetRangeFromExternalAddress(asExternalAddr As String) As Range
    If fZero(asExternalAddr) Then fErr "wrong param"
    asExternalAddr = Trim(asExternalAddr)
    
    Dim lFileStart As Long
    Dim lFileEnd As Long
    Dim lShtEnd As Long
    Dim sWbName As String
    Dim sShtName As String
    Dim sNetAddr As String
    
    lFileStart = InStr(asExternalAddr, "[")
    lFileEnd = InStr(asExternalAddr, "]")
    lShtEnd = InStr(asExternalAddr, "!")
    
    If lFileStart <= 0 Or lShtEnd <= 0 Then
        fErr "the address passed does not have the excel file name part ot the sheet name part"
    End If
    
    sWbName = Mid(asExternalAddr, lFileStart + 1, lFileEnd - lFileStart - 1)
    sShtName = Mid(asExternalAddr, lFileEnd + 1, lShtEnd - lFileEnd - 1)
    sNetAddr = Right(asExternalAddr, Len(asExternalAddr) - lShtEnd)
    
    sWbName = Replace(sWbName, "'", "")
    sShtName = Replace(sShtName, "'", "")
    sNetAddr = fReplaceConvertR1C1ToA1(sNetAddr)
    
    Dim wbOut As Workbook
    If fExcelFileIsOpen(sWbName, wbOut) Then
        Set fGetRangeFromExternalAddress = wbOut.Worksheets(sShtName).Range(sNetAddr)
    Else
        fErr "Excel file is not open, pls check your program."
    End If
    
    Set wbOut = Nothing
End Function

Function fReplaceConvertR1C1ToA1(sR1C1Address As String) As String
    fGetGRegExp
    
    Dim matchColl As VBScript_RegExp_55.MatchCollection
    Dim match As VBScript_RegExp_55.match
    
    gRegExp.IgnoreCase = True
    gRegExp.Pattern = "R(\d{1,})C(\d{1,})"
    
    Set matchColl = gRegExp.Execute(sR1C1Address)
    
    Dim sAddrNew As String
    Dim lNextStart As Long
    Dim sReplaced As String
    
    sAddrNew = ""
    lNextStart = 1
    
    For Each match In matchColl
        sReplaced = fNum2Letter(CLng(match.SubMatches(1))) & match.SubMatches(0)
        
        sAddrNew = sAddrNew & Mid(sR1C1Address, lNextStart, match.FirstIndex - lNextStart + 1)
        sAddrNew = sAddrNew & sReplaced
        
        lNextStart = match.FirstIndex + match.Length + 1
    Next
    
    If lNextStart <= Len(sR1C1Address) Then
        sAddrNew = sAddrNew & Mid(sR1C1Address, lNextStart, Len(sR1C1Address) - lNextStart + 1)
    End If
    
    Set match = Nothing
    Set matchColl = Nothing
    
    fReplaceConvertR1C1ToA1 = IIf(fZero(sAddrNew), sR1C1Address, sAddrNew)
End Function

Function fGetGRegExp(Optional asPatten As String = "")
    If gRegExp Is Nothing Then
        Set gRegExp = New VBScript_RegExp_55.RegExp
        gRegExp.IgnoreCase = True
        gRegExp.Global = True
    End If
    
    If fNzero(asPatten) Then gRegExp.Pattern = asPatten
End Function

Function fSortArayDesc(ByRef arr(), Optional UseQuickSort As Boolean = True)
    If Not UseQuickSort Then
        Call fSortArrayBubbleSortDesc(arr)
    Else
        Call fSortArrayQuickSortDesc(arr)
    End If
End Function

Function fSortAray(ByRef arr(), Optional UseQuickSort As Boolean = True)
    If Not UseQuickSort Then
        Call fSortArrayBubbleSort(arr)
    Else
        Call fSortArrayQuickSort(arr)
    End If
End Function
Function fSortArrayBubbleSortDesc(ByRef arr())
    Dim i As Long
    Dim j As Long
    Dim Temp
    
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i) < arr(j) Then
                Temp = arr(j)
                arr(j) = arr(i)
                arr(i) = Temp
            End If
        Next j
    Next i
End Function

Function fSortArrayBubbleSort(ByRef arr)
    Dim i As Long
    Dim j As Long
    Dim Temp
    
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i) > arr(j) Then
                Temp = arr(j)
                arr(j) = arr(i)
                arr(i) = Temp
            End If
        Next j
    Next i
End Function
' Omit plngLeft & plngRight; they are used internally during recursion
Function fSortArrayQuickSort(ByRef pvarArray As Variant, Optional ByVal plngLeft As Long, Optional ByVal plngRight As Long)
    Dim lngFirst As Long
    Dim lngLast As Long
    Dim varMid As Variant
    Dim varSwap As Variant

    If plngRight = 0 Then
        plngLeft = LBound(pvarArray)
        plngRight = UBound(pvarArray)
    End If
    lngFirst = plngLeft
    lngLast = plngRight
    varMid = pvarArray((plngLeft + plngRight) \ 2)
    Do
        Do While pvarArray(lngFirst) < varMid And lngFirst < plngRight
            lngFirst = lngFirst + 1
        Loop
        Do While varMid < pvarArray(lngLast) And lngLast > plngLeft
            lngLast = lngLast - 1
        Loop
        If lngFirst <= lngLast Then
            varSwap = pvarArray(lngFirst)
            pvarArray(lngFirst) = pvarArray(lngLast)
            pvarArray(lngLast) = varSwap
            lngFirst = lngFirst + 1
            lngLast = lngLast - 1
        End If
    Loop Until lngFirst > lngLast
    If plngLeft < lngLast Then fSortArrayQuickSort pvarArray, plngLeft, lngLast
    If lngFirst < plngRight Then fSortArrayQuickSort pvarArray, lngFirst, plngRight
End Function

Function fSortArrayQuickSortDesc(ByRef pvarArray As Variant, Optional ByVal plngLeft As Long, Optional ByVal plngRight As Long)
    Dim lngFirst As Long
    Dim lngLast As Long
    Dim varMid As Variant
    Dim varSwap As Variant

    If plngRight = 0 Then
        plngLeft = LBound(pvarArray)
        plngRight = UBound(pvarArray)
    End If
    lngFirst = plngLeft
    lngLast = plngRight
    varMid = pvarArray((plngLeft + plngRight) \ 2)
    Do While lngFirst <= lngLast
        Do While pvarArray(lngFirst) > varMid And lngFirst < plngRight
            lngFirst = lngFirst + 1
        Loop
        Do While varMid > pvarArray(lngLast) And lngLast > plngLeft
            lngLast = lngLast - 1
        Loop
        If lngFirst <= lngLast Then
            varSwap = pvarArray(lngFirst)
            pvarArray(lngFirst) = pvarArray(lngLast)
            pvarArray(lngLast) = varSwap
            lngFirst = lngFirst + 1
            lngLast = lngLast - 1
        End If
    Loop
    If plngLeft < lngLast Then fSortArrayQuickSortDesc pvarArray, plngLeft, lngLast
    If lngFirst < plngRight Then fSortArrayQuickSortDesc pvarArray, lngFirst, plngRight
End Function
