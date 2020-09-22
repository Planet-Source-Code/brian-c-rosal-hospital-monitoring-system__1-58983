VERSION 5.00
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H80000001&
   BorderStyle     =   0  'None
   ClientHeight    =   3975
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   6930
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":0000
   ScaleHeight     =   3975
   ScaleWidth      =   6930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer timTimer 
      Interval        =   2500
      Left            =   7770
      Top             =   480
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
       Case 13
            Unload Me
            frmLogin.Show
            
End Select
End Sub




Private Sub timTimer_Timer()
Unload Me
frmLogin.Show
End Sub
