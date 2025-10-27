VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_Open()
    If ThisWorkbook.ProtectStructure Then
        ThisWorkbook.Unprotect Password:="123"
    End If
    
    ThisWorkbook.Sheets(1).Visible = xlSheetVeryHidden
    ThisWorkbook.Sheets(4).Visible = xlSheetVeryHidden
    ThisWorkbook.Sheets("ChartData").Visible = xlSheetVeryHidden
    Call ProtectAllSheets
    
    If Not ThisWorkbook.ProtectStructure Then
        ThisWorkbook.Protect Password:="123", Structure:=True, Windows:=False
    End If
End Sub
