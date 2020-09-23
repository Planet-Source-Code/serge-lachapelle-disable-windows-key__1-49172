VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Disable Windows Keys"
   ClientHeight    =   984
   ClientLeft      =   1956
   ClientTop       =   1536
   ClientWidth     =   3468
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   984
   ScaleWidth      =   3468
   StartUpPosition =   2  'CenterScreen
   WindowState     =   1  'Minimized
   Begin VB.CheckBox chkDisable 
      Caption         =   "&Disable Windows Keys"
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   2052
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim hhkLowLevelKybd As Long

Private Sub chkDisable_Click()
  If chkDisable = vbChecked Then
    hhkLowLevelKybd = SetWindowsHookEx(WH_KEYBOARD_LL, AddressOf LowLevelKeyboardProc, App.hInstance, 0)
  Else
    UnhookWindowsHookEx hhkLowLevelKybd
    hhkLowLevelKybd = 0
  End If
End Sub

Private Sub Form_Load()
  chkDisable.Value = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If hhkLowLevelKybd <> 0 Then UnhookWindowsHookEx hhkLowLevelKybd
End Sub
