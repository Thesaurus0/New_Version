Attribute VB_Name = "MC9_RefData"
Option Explicit
Option Base 1

Enum TechTag
    REPORT_ID = 1
    ID = 2
    Name = 3
    Commission = 4
    CheckBoxName = 5
    InputFileTextBoxName = 6
    Selected = 7
End Enum


Function fReadConfigTechTagList() As Dictionary
    Dim asTag As String
    Dim arrColsName()
    Dim rngToFindIn As Range
    Dim arrConfigData()
    Dim lConfigStartRow As Long
    Dim lConfigStartCol As Long
    Dim lConfigEndRow As Long
    Dim lConfigHeaderAtRow As Long

    asTag = "[Sales TechTag List]"
    ReDim arrColsName(TechTag.REPORT_ID To TechTag.Selected)
    
    arrColsName(TechTag.REPORT_ID) = "TechTag ID"
    arrColsName(TechTag.ID) = "TechTag ID In DB"
    arrColsName(TechTag.Name) = "TechTag Name"
    arrColsName(TechTag.Commission) = "Default Commission"
    arrColsName(TechTag.CheckBoxName) = "CheckBox Name"
    arrColsName(TechTag.InputFileTextBoxName) = "Input File TextBox Name"
    arrColsName(TechTag.Selected) = "User Ticked"
     
    arrConfigData = fReadConfigBlockToArrayNet(asTag:=asTag, shtParam:=shtStaticData _
                                , arrColsName:=arrColsName _
                                , lConfigStartRow:=lConfigStartRow _
                                , lConfigStartCol:=lConfigStartCol _
                                , lConfigEndRow:=lConfigEndRow _
                                , lOutConfigHeaderAtRow:=lConfigHeaderAtRow _
                                , abNoDataConfigThenError:=True _
                                )
    Call fValidateDuplicateInArray(arrConfigData, TechTag.REPORT_ID, False, shtStaticData, lConfigHeaderAtRow, lConfigStartCol, "TechTag ID")
    Call fValidateDuplicateInArray(arrConfigData, TechTag.ID, False, shtStaticData, lConfigHeaderAtRow, lConfigStartCol, "TechTag ID In DB")
    Call fValidateDuplicateInArray(arrConfigData, TechTag.Name, False, shtStaticData, lConfigHeaderAtRow, lConfigStartCol, "TechTag Name")
    Call fValidateDuplicateInArray(arrConfigData, TechTag.CheckBoxName, False, shtStaticData, lConfigHeaderAtRow, lConfigStartCol, "TechTag Name")
    Call fValidateDuplicateInArray(arrConfigData, TechTag.InputFileTextBoxName, False, shtStaticData, lConfigHeaderAtRow, lConfigStartCol, "TechTag Name")
    
'    Call fValidateBlankInArray(arrConfigData, TechTag.Report_ID, shtstaticdata, lConfigHeaderAtRow, lConfigStartCol, "TechTag ID")
'    Call fValidateBlankInArray(arrConfigData, TechTag.ID, shtstaticdata, lConfigHeaderAtRow, lConfigStartCol, "TechTag ID In DB")
'    Call fValidateBlankInArray(arrConfigData, TechTag.Name, shtstaticdata, lConfigHeaderAtRow, lConfigStartCol, "TechTag Name")
    
'    Set fReadConfigTechTagList = fReadArray2DictionaryWithMultipleColsCombined(arrConfigData, TechTag.Report_ID _
'            , Array(TechTag.ID, TechTag.Name, TechTag.Commission, TechTag.CheckBoxName, TechTag.InputFileTextBoxName, TechTag.Selected) _
'            , DELIMITER)

    Dim dictOut As Dictionary
    Set dictOut = New Dictionary
    
    Dim lEachRow As Long
    Dim sFileTag As String
    Dim sValueStr As String
    
    For lEachRow = LBound(arrConfigData, 1) To UBound(arrConfigData, 1)
        If fArrayRowIsBlankHasNoData(arrConfigData, lEachRow) Then GoTo next_row
        
'        sRptNameStr = DELIMITER & arrConfigData(lEachRow, 1) & DELIMITER
'        If InStr(sRptNameStr, DELIMITER & asReportID & DELIMITER) <= 0 Then GoTo next_row
        
        'lActualRow = lConfigHeaderAtRow + lEachRow
        
        sFileTag = Trim(arrConfigData(lEachRow, TechTag.REPORT_ID))
        sValueStr = fComposeStrForDictTechTagList(arrConfigData, lEachRow)
        
        dictOut.Add sFileTag, sValueStr
next_row:
    Next
    
    Erase arrColsName
    Erase arrConfigData
    Set fReadConfigTechTagList = dictOut
    Set dictOut = Nothing
End Function

Function fComposeStrForDictTechTagList(arrConfigData, lEachRow As Long) As String
    Dim sOut As String
    Dim i As Integer
    
    For i = TechTag.ID To TechTag.Selected
        sOut = sOut & DELIMITER & Trim(arrConfigData(lEachRow, i))
    Next
    
    fComposeStrForDictTechTagList = Right(sOut, Len(sOut) - 1)
End Function

Function fGetTechTag_InputFileTextBoxName(asTechTagID As String) As String
    fGetTechTag_InputFileTextBoxName = Split(dictCompList(asTechTagID), DELIMITER)(TechTag.InputFileTextBoxName - TechTag.REPORT_ID - 1)
End Function
Function fGetTechTag_CheckBoxName(asTechTagID As String) As String
    fGetTechTag_CheckBoxName = Split(dictCompList(asTechTagID), DELIMITER)(TechTag.CheckBoxName - TechTag.REPORT_ID - 1)
End Function
Function fGetTechTag_UserTicked(asTechTagID As String) As String
    fGetTechTag_UserTicked = Split(dictCompList(asTechTagID), DELIMITER)(TechTag.Selected - TechTag.REPORT_ID - 1)
End Function


