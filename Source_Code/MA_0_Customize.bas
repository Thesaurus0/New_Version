Attribute VB_Name = "MA_0_Customize"
Option Explicit
Option Base 1

Function fSetBackToConfigSheetAndUpdategDict_UserTicket()
    
    Dim ckb As Object
    
    Dim eachObj As Object
    
    'for each eachobj in shtmenu.
    Dim i As Long
    Dim sTechTagID As String
    Dim sTickValue As String
    
    For i = 0 To dictCompList.Count - 1
        sTechTagID = dictCompList.Keys(i)
         
        If Not fActiveXControlExistsInSheet(shtMenu, fGetTechTag_CheckBoxName(sTechTagID), ckb) Then GoTo next_TechTag
        
        sTickValue = IIf(ckb.Value, "Y", "N")
        
        Call fSetSpecifiedConfigCellValue(shtStaticData, "[Sales TechTag List]", "User Ticked", "TechTag ID=" & sTechTagID, sTickValue)
        Call fUpdateDictionaryItemValueForDelimitedElement(dictCompList, sTechTagID, TechTag.Selected - TechTag.REPORT_ID, sTickValue)
next_TechTag:
    Next
End Function

Function fSetBackToConfigSheetAndUpdategDict_InputFiles()
    Dim i As Integer
    Dim sEachTechTagID As String
    Dim sFilePathRange As String
    Dim sEachFilePath  As String
    
    For i = 0 To dictCompList.Count - 1
        sEachTechTagID = dictCompList.Keys(i)
        'sFilePathRange = "rngSalesFilePath_" & sEachTechTagID
        
        If fGetTechTag_UserTicked(sEachTechTagID) = "Y" Then
            sFilePathRange = fGetTechTag_InputFileTextBoxName(sEachTechTagID)
            sEachFilePath = Trim(shtMenu.Range(sFilePathRange).Value)
        Else
            sEachFilePath = "User not selected."
        End If
         
        Call fSetValueBackToSysConf_InputFile_FileName(sEachTechTagID, sEachFilePath)
        Call fUpdateGDictInputFile_FileName(sEachTechTagID, sEachFilePath)
        
        'Call fSetUserProvidedFileToMainConfig(sEachTechTagID, sEachFilePath)
    Next
    
    
'    sFile = Trim(shtMenu.Range("rngSalesFilePath_GY").Value)
'
'    Call fSetValueBackToSysConf_InputFile_FileName("GY", sFile)
'    Call fUpdateGDictInputFile_FileName("GY", sFile)
    
    
End Function

Function fSetIntialValueForShtMenuInitialize()
    
End Function

Function fInitialization()
    err.Clear
    gbNoData = False
    gbBusinessError = False
    gbUserCanceled = False
    
    If fZero(gsEnv) Then gsEnv = fGetEnvFromSysConf
    
    Call fDisableExcelOptionsAll
    
    If fIsDev Then Application.ScreenUpdating = True
    
    Call fRevmoeFilterForAllSheets(ThisWorkbook)
End Function

