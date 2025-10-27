VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmHelp 
   Caption         =   "UserForm1"
   ClientHeight    =   10080
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14040
   OleObjectBlob   =   "frmHelp.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    Me.Caption = "Help & Instructions"
    FrameHelp.Caption = "Recipe Management System"
    Me.txtHelpContent.ScrollBars = fmScrollBarsVertical ' Enable vertical scrollbar
    Me.txtHelpContent.WordWrap = True  ' Enable word wrapping
    Me.txtHelpContent.SelStart = 0 ' Force scrollbar to start at the top
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub


