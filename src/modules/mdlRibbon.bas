Attribute VB_Name = "mdlRibbon"
Option Explicit
Public frmOpenRecipeFileBool As Boolean

' Custom Data Tools

Sub RunCode1(control As IRibbonControl) ' Product Manager
    frmIngredientManager.Show vbModeless
End Sub

Sub RunCode2(control As IRibbonControl) ' Recipe Manager
    frmRecipeManager.Show vbModeless
End Sub

Sub RunCode3(control As IRibbonControl) ' Recipe Book
    frmOpenRecipeFile.Show vbModeless
    If frmOpenRecipeFileBool = False Then
        Unload frmOpenRecipeFile
    End If
End Sub

Sub RunCode4(control As IRibbonControl) ' Recipe Folder
    Call OpenRecipeFolder
End Sub

Sub RunCode5(control As IRibbonControl) ' Recovery
    Call UnprotectAllSheets
    Call recoverRecipeFiles
    Call ProtectAllSheets
End Sub

Sub RunCode6(control As IRibbonControl) ' Help & Instructions
    frmHelp.Show vbModeless
End Sub



