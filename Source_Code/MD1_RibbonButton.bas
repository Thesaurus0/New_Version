Attribute VB_Name = "MD1_RibbonButton"
Option Explicit
Option Base 1
   
Sub subMain_Ribbon_ImportUserProvidedFiles()
    If shtMenu.Visible = xlSheetVisible Then
        If ActiveSheet.Name <> shtMenu.Name Then
            shtMenu.Visible = xlSheetVisible
            shtMenu.Activate
            Range("A63").Select
        Else
            shtMenu.Visible = xlSheetVeryHidden
        End If
    Else
        shtMenu.Visible = xlSheetVisible
        shtMenu.Activate
        Range("A63").Select
    End If
End Sub

