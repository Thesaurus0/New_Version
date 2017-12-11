Attribute VB_Name = "Common_CommandBar"
Option Explicit
Option Base 1

Sub sub_RemoveAllCommandBars()
    Dim tmpBar As CommandBar

    For Each tmpBar In Application.CommandBars
        If Not tmpBar.BuiltIn Then
            tmpBar.Delete
        End If
    Next
End Sub

Sub sub_ListAllCommandBars()
'    Dim shtOutput As Worksheet
'    If Not fGetTmpSheetInWorkbookWhenNotExistsCreateIt(shtOutput) Then Exit Sub
    
    'Dim arrData()
    Dim sStr As String
    Dim iCnt As Long
    
    Dim tmpBar As CommandBar

    For Each tmpBar In Application.CommandBars
        If Not tmpBar.BuiltIn Then
            iCnt = iCnt + 1
            sStr = sStr & vbCr & tmpBar.Name
        End If
    Next
    
    MsgBox iCnt & "CommandBars: " & sStr
End Sub

'
'Sub sub_add_new_bar(as_bar_name As String)
'    Dim lcb_new_commdbar As CommandBar
'
'    Call sub_RemoveToolBar(as_bar_name)
'
'    Set lcb_new_commdbar = Application.CommandBars.Add(as_bar_name, msoBarTop)
'    lcb_new_commdbar.Visible = True
'End Sub
'
'Public Sub sub_RemoveToolBar(as_toolbar As String)
'    On Error Resume Next
'
'    Dim lcb_commdbar As CommandBar
'
'    Set lcb_commdbar = Nothing
'
'    Application.CommandBars(as_toolbar).Delete
'    Application.CommandBars("Custom 1").Delete
'End Sub
 
'Public Sub sub_add_new_button(as_bar_name As String, as_btn_caption As String, _
'                    as_on_action As String, ai_face_id As Integer, _
'                    Optional as_tip_text As String)
'
'    Dim lcb_commdbar As CommandBar
'    Dim lbtn_new_button As CommandBarButton
'
'    Set lcb_commdbar = Application.CommandBars(as_bar_name)
'
'    Set lbtn_new_button = lcb_commdbar.Controls.Add(msoControlButton)
'    With lbtn_new_button
'        .Caption = as_btn_caption
'        .Style = msoButtonIconAndCaptionBelow
'        '.OnAction = "sub_RemoveToolBar"
'        .OnAction = as_on_action
'        .FaceId = ai_face_id
'        .TooltipText = as_tip_text
'        .BeginGroup = True
'    End With
'
'    'Set lcb_commdbar = Nothing
'    'Set lbtn_new_button = Nothing
'
'End Sub


Sub subAddNewButtonToBarWhenBarNotExistsCreateIt(asBarName As String, asBtnCaption As String _
                                                , asOnAction As String, aiFaceId As Long _
                                                , Optional asTipText As String = "")

    If fZero(asBarName) Or fZero(asBtnCaption) Or fZero(asOnAction) Then fErr "Wron param"
    
    Dim cmdBar As CommandBar
    
    If fCommandBarExists(asBarName, cmdBar) Then
    Else
        Set cmdBar = fAddNewCommandBar(asBarName)
    End If
    
    Call fAddNewButtonToBarWhenExistsUpdateIt(cmdBar, asBtnCaption, asOnAction, aiFaceId, asTipText)
End Sub

Function fCommandBarExists(asBarName As String, ByRef cmdBar As CommandBar) As Boolean
    'Dim cmdBar As CommandBar
    
    On Error Resume Next
    Set cmdBar = Application.CommandBars(asBarName)
    fCommandBarExists = (Not cmdBar Is Nothing)
    err.Clear
End Function

Function fAddNewCommandBar(asBarName As String) As CommandBar
    Dim cmdBar As CommandBar
    
    Call sub_RemoveCommandBar(asBarName)
    Set cmdBar = Application.CommandBars.Add(asBarName, msoBarTop)
    cmdBar.Visible = True
    
    Set fAddNewCommandBar = cmdBar
    Set cmdBar = Nothing
End Function

Sub sub_RemoveCommandBar(ByVal asBarName As String)
    On Error Resume Next
    Application.CommandBars(asBarName).Delete
    err.Clear
End Sub

Function fAddNewButtonToBarWhenExistsUpdateIt(cmdBar As CommandBar, asBtnCaption As String _
                                            , asOnAction As String, aiFaceId As Long _
                                            , Optional asTipText As String)
    'Dim cmdBar As CommandBar
    Dim btn As CommandBarButton
 
    'Set cmdBar = Application.CommandBars(asBarName)
    
    If fButtonExistsInCommandBar(cmdBar, asBtnCaption, btn) Then
    Else
        Set btn = cmdBar.Controls.Add(msoControlButton)
    End If
 
    btn.Caption = asBtnCaption
    btn.Style = msoButtonIconAndCaptionBelow
    btn.OnAction = asOnAction
    btn.FaceId = aiFaceId
    btn.TooltipText = IIf(Len(asTipText) <= 0, asBtnCaption, asTipText)
    btn.BeginGroup = True
    
    Set cmdBar = Nothing
    Set btn = Nothing
End Function

Function fButtonExistsInCommandBar(cmdBar As CommandBar, asBtnCaption As String, ByRef btn As CommandBarButton) As Boolean
    On Error Resume Next
    Set btn = cmdBar.Controls(asBtnCaption)
    
    fButtonExistsInCommandBar = (Not btn Is Nothing)
    err.Clear
End Function

Function fReadConfigCommandBarsInfo() As Variant
    Dim asTag As String
    Dim arrColsName()
    Dim rngToFindIn As Range
    Dim arrConfigData()
    Dim lConfigStartRow As Long
    Dim lConfigStartCol As Long
    Dim lConfigEndRow As Long
    Dim lConfigHeaderAtRow As Long

    asTag = "[Ribbon CommandBar and Menu]"
    ReDim arrColsName(1 To 4)
    arrColsName(1) = "Toolbar Tech Name"
    arrColsName(2) = "Button Caption"
    arrColsName(3) = "Sub/Function/OnAction"
    arrColsName(4) = "DEV/UAT/PROD"
   
    arrConfigData = fReadConfigBlockToArrayNet(asTag:=asTag, shtParam:=shtSysConf _
                                , arrColsName:=arrColsName _
                                , lConfigStartRow:=lConfigStartRow _
                                , lConfigStartCol:=lConfigStartCol _
                                , lConfigEndRow:=lConfigEndRow _
                                , lOutConfigHeaderAtRow:=lConfigHeaderAtRow _
                                , abNoDataConfigThenError:=True _
                                )
    
    Call fValidateDuplicateInArray(arrConfigData, 1, True, shtSysConf, lConfigHeaderAtRow, lConfigStartCol, arrColsName(1))
    Call fValidateDuplicateInArray(arrConfigData, 2, True, shtSysConf, lConfigHeaderAtRow, lConfigStartCol, arrColsName(2))
    Call fValidateDuplicateInArray(arrConfigData, 3, True, shtSysConf, lConfigHeaderAtRow, lConfigStartCol, arrColsName(3))
    
    Call fValidateBlankInArray(arrConfigData, 1, shtSysConf, lConfigHeaderAtRow, lConfigStartCol, arrColsName(1))
    Call fValidateBlankInArray(arrConfigData, 2, shtSysConf, lConfigHeaderAtRow, lConfigStartCol, arrColsName(2))
    Call fValidateBlankInArray(arrConfigData, 3, shtSysConf, lConfigHeaderAtRow, lConfigStartCol, arrColsName(3))
    Call fValidateBlankInArray(arrConfigData, 4, shtSysConf, lConfigHeaderAtRow, lConfigStartCol, arrColsName(4))
    
    fReadConfigCommandBarsInfo = arrConfigData
    Erase arrConfigData
    Erase arrColsName
End Function

Function fReadConfigRibbonCommandBarMenuAndCreateCommandBarButton()
    Dim asTag As String
    Dim arrColsName()
    Dim rngToFindIn As Range
    Dim arrConfigData()
    Dim lConfigStartRow As Long
    Dim lConfigStartCol As Long
    Dim lConfigEndRow As Long
    Dim lConfigHeaderAtRow As Long

    asTag = "[Ribbon CommandBar and Menu]"
    ReDim arrColsName(1 To 6)
    arrColsName(1) = "Toolbar Tech Name"
    arrColsName(2) = "Button Caption"
    arrColsName(3) = "Sub/Function/OnAction"
    arrColsName(4) = "FaceID / Icon"
    arrColsName(5) = "DEV/UAT/PROD"
    arrColsName(6) = "Tip Text"
   
    arrConfigData = fReadConfigBlockToArrayNet(asTag:=asTag, shtParam:=shtSysConf _
                                , arrColsName:=arrColsName _
                                , lConfigStartRow:=lConfigStartRow _
                                , lConfigStartCol:=lConfigStartCol _
                                , lConfigEndRow:=lConfigEndRow _
                                , lOutConfigHeaderAtRow:=lConfigHeaderAtRow _
                                , abNoDataConfigThenError:=True _
                                )
    
'    Call fValidateDuplicateInArray(arrConfigData, 1, True, shtSysConf, lConfigHeaderAtRow, lConfigStartCol, arrColsName(1))
'    Call fValidateDuplicateInArray(arrConfigData, 2, True, shtSysConf, lConfigHeaderAtRow, lConfigStartCol, arrColsName(2))
'    Call fValidateDuplicateInArray(arrConfigData, 3, True, shtSysConf, lConfigHeaderAtRow, lConfigStartCol, arrColsName(3))
'
'    Call fValidateBlankInArray(arrConfigData, 1, shtSysConf, lConfigHeaderAtRow, lConfigStartCol, arrColsName(1))
'    Call fValidateBlankInArray(arrConfigData, 2, shtSysConf, lConfigHeaderAtRow, lConfigStartCol, arrColsName(2))
'    Call fValidateBlankInArray(arrConfigData, 3, shtSysConf, lConfigHeaderAtRow, lConfigStartCol, arrColsName(3))
    Dim i As Long
    Dim sEnv As String
    Dim sCmdBarName As String
    Dim sBtnCap As String
    Dim sSub As String
    Dim lFaceId As Long
    Dim sTip As String
    
    For i = LBound(arrConfigData, 1) To UBound(arrConfigData, 1)
        If fArrayRowIsBlankHasNoData(arrConfigData, i) Then GoTo next_row
        
        sEnv = arrConfigData(i, 5)
        
        If sEnv = gsEnv Or sEnv = "SHARED" Then
            sCmdBarName = arrConfigData(i, 1)
            sBtnCap = arrConfigData(i, 2)
            sSub = arrConfigData(i, 3)
            lFaceId = arrConfigData(i, 4)
            sTip = arrConfigData(i, 6)
            
            Call subAddNewButtonToBarWhenBarNotExistsCreateIt(sCmdBarName, sBtnCap, sSub, lFaceId, sTip)
        End If
next_row:
    Next
    
    Erase arrConfigData
    Erase arrColsName
End Function

Function fRemoveAllCommandbarsByConfig()
    On Error Resume Next
    
    Dim arrAllCmdBarList()
    arrAllCmdBarList = ThisWorkbook.fGetThisWorkBookVariable("CMDBAR")
    
    Dim i As Long
    
    For i = LBound(arrAllCmdBarList) To UBound(arrAllCmdBarList)
        Call sub_RemoveCommandBar(arrAllCmdBarList(i))
    Next
    
    err.Clear
End Function

Function fEnableOrDisableAllCommandBarsByConfig(bValue As Boolean)
    On Error Resume Next
    
    Dim arrAllCmdBarList()
    arrAllCmdBarList = ThisWorkbook.fGetThisWorkBookVariable("CMDBAR")
    
    Dim i As Long
    
    For i = LBound(arrAllCmdBarList) To UBound(arrAllCmdBarList)
        Application.CommandBars(arrAllCmdBarList(i)).Visible = bValue
    Next
    
    err.Clear
End Function
