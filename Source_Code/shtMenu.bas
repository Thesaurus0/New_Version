VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shtMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub btnBatchImportSaleInfoFiles_Click()
    'Call fImportAllUserProvidedFiles
    Call subMain_ImportUserProvidedFiles
    
End Sub

Private Sub btnSelect_CZL_Click()
    Call fOpenFileSelectDialogAndSetToSheetRangeForTechTag("Apollo")
End Sub

Private Sub btnSelect_GKYX_Click()
    Call fOpenFileSelectDialogAndSetToSheetRangeForTechTag("Murex")
End Sub
 
Function fOpenFileSelectDialogAndSetToSheetRangeForTechTag(sTechTag As String)
    Dim sHeader As String
    
    sHeader = LeftB(shtMenu.Range("rngHeader_" & sTechTag).Value, LenB(shtMenu.Range("rngHeader_" & sTechTag).Value) - 2)
    Call fOpenFileSelectDialogAndSetToSheetRange("rngSalesFilePath_" & sTechTag, , sHeader)
End Function

