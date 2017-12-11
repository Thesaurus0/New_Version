VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' version : 201712 - 001 lost the validate blank -- 23:45 -- 0£º11
Option Explicit
Option Base 1

Dim arrAllCmdBarList()



Private Sub Workbook_Open()
    gsEnv = fGetEnvFromSysConf
    
    Call fRefreshGetAllCommandbarsList
    Call sub_WorkBookInitialization
    Call fSetIntialValueForShtMenuInitialize
    
    ThisWorkbook.Saved = True
End Sub

Sub sub_WorkBookInitialization()
    Call fReadConfigRibbonCommandBarMenuAndCreateCommandBarButton
    
    If fIsDev() Then
        shtSysConf.Visible = xlSheetVisible
        shtStaticData.Visible = xlSheetVisible
        shtFileSpec.Visible = xlSheetVisible
    Else
        shtSysConf.Visible = xlSheetVeryHidden
        shtStaticData.Visible = xlSheetVeryHidden
        shtFileSpec.Visible = xlSheetVeryHidden
    End If
End Sub

Function fRefreshGetAllCommandbarsList()
    Dim arrCBarsInfo()
    Dim dict As Dictionary
    
    Debug.Print "fRefreshGetAllCommandbarsList" & Now()
    
    arrCBarsInfo = fReadConfigCommandBarsInfo
    
    Erase arrAllCmdBarList
   
    Set dict = fRadArray2DictionaryOnlyKeys(arrCBarsInfo, 1)
    Call fCopyDictionaryKeys2Array(dict, arrAllCmdBarList)
    
    Set dict = Nothing
End Function

Function fGetThisWorkBookVariable(sVariable As String) As Variant
    Select Case sVariable
        Case "CMDBAR"
            If fArrayIsEmptyOrNoData(arrAllCmdBarList) Then
                Call fRefreshGetAllCommandbarsList
            End If
            
            fGetThisWorkBookVariable = arrAllCmdBarList
        Case Else
            fErr "wrong param"
    End Select
End Function
 

Private Sub Workbook_Activate()
    Call fEnableOrDisableAllCommandBarsByConfig(True)
End Sub
Private Sub Workbook_Deactivate()
    Call fEnableOrDisableAllCommandBarsByConfig(False)
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    Application.EnableEvents = False
    
    On Error GoTo exit_function
    
    If Sh Is shtSysConf Then
        Call fShtSysConf_SheetChange_DevProdChange(Target)
        'Call shtSysConf_SheetChange_CommandBarConfig(Target)
    End If
    
exit_function:
    Application.EnableEvents = True
End Sub
