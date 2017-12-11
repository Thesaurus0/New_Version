Attribute VB_Name = "MB1_ImportSalesFiles"
Option Explicit
Option Base 1
   
'Dim arrSalesTechTags()

Public gsTechTagID As String
Public dictCompList As Dictionary

Sub subMain_ImportUserProvidedFiles()
    'If Not fIsDev Then On Error GoTo error_handling
    
    fInitialization
    
    gsRptID = "MAIN_REPORT"
    
    Call fReadSysConfig_InputTxtSheetFile
    
    Set dictCompList = fReadConfigTechTagList
    Call fValidationAndSetToConfigSheet
    
    Call fSetBackToConfigSheetAndUpdategDict_UserTicket
    Call fSetBackToConfigSheetAndUpdategDict_InputFiles
    
    gsRptFilePath = fReadSysConfig_Output(, gsRptType)
    
    Dim i As Integer
    For i = 0 To dictCompList.Count - 1
        gsTechTagID = dictCompList.Keys(i)
        
        If fGetTechTag_UserTicked(gsTechTagID) = "Y" Then
            Call fLoadFilesAndRead2Variables
        End If
    Next
    
    
error_handling:
End Sub

'Function fImportAllUserProvidedFiles()
'    Dim i As Integer
'
'    For i = LBound(arrSalesTechTags, 1) To UBound(arrSalesTechTags, 1)
'        Call fImportUserProvidedFileForComapnay(CStr(arrSalesTechTags(i, 0)) _
'                                            , CStr(arrSalesTechTags(i, 1)) _
'                                            , CStr(arrSalesTechTags(i, 2)))
'    Next
'End Function

Function fLoadFilesAndRead2Variables()
    'gsTechTagID
    Call fLoadFileByFileTag(gsTechTagID)
    Call fReadMasterSheetData(gsTechTagID)
End Function
 

Function fImportUserProvidedFileForComapnay(asTechTagID As String, asTechTagName As String, sUserProvidedFile As String)
    Dim sTmpSht As String
    sTmpSht = fGenRandomUniqueString
    
    
    
End Function


Function fValidationAndSetToConfigSheet()
    Dim i As Integer
    Dim sEachTechTagID As String
    Dim sFilePathRange As String
    Dim sEachFilePath  As String
    
    For i = 0 To dictCompList.Count - 1
        sEachTechTagID = dictCompList.Keys(i)
        'sFilePathRange = "rngSalesFilePath_" & sEachTechTagID
        
        sFilePathRange = fGetTechTag_InputFileTextBoxName(sEachTechTagID)
        sEachFilePath = Trim(shtMenu.Range(sFilePathRange).Value)
        
        If Not fFileExists(sEachFilePath) Then
            shtMenu.Activate
            shtMenu.Range(sFilePathRange).Select
            fErr Split(dictCompList(sEachTechTagID), DELIMITER)(1) & ": 输入的文件不存在，请检查：" & vbCr & sEachFilePath
        End If
        
        'Call fSetUserProvidedFileToMainConfig(sEachTechTagID, sEachFilePath)
    Next
End Function

'Function fSetUserProvidedFileToMainConfig(sTechTagId As String, sFile As String)
'    Call fSetSpecifiedConfigCellAddress(shtSysConf, "[Input Files]", "File Full Path", "TechTag ID=" & sTechTagId, sFile)
'End Function
