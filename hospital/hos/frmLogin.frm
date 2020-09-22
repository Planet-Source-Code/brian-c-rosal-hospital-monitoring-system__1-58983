VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   3960
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3450
   Icon            =   "frmLogin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2339.698
   ScaleMode       =   0  'User
   ScaleWidth      =   3239.364
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   975
      Left            =   105
      Picture         =   "frmLogin.frx":0442
      ScaleHeight     =   915
      ScaleWidth      =   3210
      TabIndex        =   7
      Top             =   120
      Width           =   3270
   End
   Begin VB.Frame Frame2 
      Height          =   2550
      Left            =   90
      TabIndex        =   4
      Top             =   1095
      Width           =   3285
      Begin Project1.chameleonButton cmdOK 
         Height          =   735
         Left            =   375
         TabIndex        =   2
         Top             =   1650
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   1296
         BTYPE           =   5
         TX              =   "&OK"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmLogin.frx":4A5F
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   1
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox txtPassword 
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   120
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   1200
         Width           =   3000
      End
      Begin VB.TextBox txtUserName 
         Height          =   345
         Left            =   135
         TabIndex        =   0
         Top             =   540
         Width           =   2985
      End
      Begin Project1.chameleonButton cmdCancel 
         Height          =   735
         Left            =   1785
         TabIndex        =   3
         Top             =   1650
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   1296
         BTYPE           =   5
         TX              =   "E&xit"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmLogin.frx":4A7B
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   1
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label2 
         Caption         =   "Password"
         Height          =   270
         Left            =   105
         TabIndex        =   6
         Top             =   975
         Width           =   1515
      End
      Begin VB.Label Label1 
         Caption         =   "Log-In Name"
         Height          =   270
         Left            =   135
         TabIndex        =   5
         Top             =   315
         Width           =   1515
      End
   End
   Begin VB.Label Label3 
      Caption         =   "(c)Copyright 98-2003 Madz:Bri:Scanhead "
      ForeColor       =   &H80000010&
      Height          =   225
      Left            =   90
      TabIndex        =   8
      Top             =   3675
      Width           =   3285
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim sql, rs, TQR, y As Variant, ctr As Byte
Dim nTries As Integer
Dim bName, bPass, scanhead, david As Boolean

Private Sub cmdCancel_Click()
End
End Sub

Private Sub cmdOK_Click()
    Dim P, madzsrch As String
    Dim Query, tb, tx, td, Madz, ad As String

nTries = nTries + 1
tb = Trim(txtUserName.Text)
              
sql = "select logname,name,pos,password,logtype from password where logname = '" & tb & "' order by 1;"
Set rs = DBMain.OpenRecordset(sql)
If rs.RecordCount > 0 Then
     bName = True
     'Check pass
     tx = Encrypt(txtPassword)
     sql = "select logname,name,pos,password,logtype from password where password = '" & tx & "' order by 1;"
     Set rs = DBMain.OpenRecordset(sql)
     If rs.RecordCount > 0 Then
     bPass = True
     Madz = rs.Fields(4)
     End If
End If

If bName = True And bPass = True Then
     If Madz = "Administrator" Then
           MsgBox "Good Day Administrator: " + rs.Fields(1), vbInformation + vbOKOnly, "Greetings"
     Else
          MainForm.Deletion.Enabled = False
          MainForm.log.Enabled = False
          MainForm.Manager.Enabled = False
          MainForm.Toolbar2.Buttons.Item(1).Enabled = False
     End If
   '88888888888888888888888888888888
   
   P = "INSERT INTO login (name,logname) VALUES ('" & Madz & "','" & txtUserName.Text & "');"
   Set TQR = DBMain.CreateQueryDef("", P)
   TQR.Execute
    
   '8888888888888888888888888888888888888888888
   MainForm.StatusBar1.Panels(1).Text = "(c)Soft/Image Hieroglyphix 1998-2003" + "    |     " + Madz + ":  " + txtUserName.Text
   MainForm.Show
   Unload frmLogin
Else
   If nTries < 3 Then
      If bName = False Then
         MsgBox "Invalid Log-in User-Name, try again!", , "Login"
         txtUserName.SetFocus
      ElseIf bPass = False Then
         MsgBox "Invalid Password, try again!", , "Login"
         txtPassword.SetFocus
      End If
      SendKeys "{Home}+{End}" 'start to end
   Else
      MsgBox "Time is up your out!", , "Login"
      End
   End If
   bName = False
   bPass = False
End If

End Sub

Private Sub Form_Load()
Dim trapper As String
On Error GoTo trapper
add3d
Set WSMain = DBEngine.Workspaces(0)
Set DBMain = WSMain.OpenDatabase(App.Path + "\Hospital.mdb", False, False, ";pwd=scanhead")
nTries = 0
bName = False
bPass = False
SendKeys "{tab}"
trapper:
End Sub

Function add3d()
Add3DBorder Me
Add3DBorder Picture1
Add3DBorder txtUserName
Add3DBorder txtPassword

End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
       Case vbKeyF4
            scanhead = True
       Case vbKeyF3
            david = True
       Case vbKeyF8
            shock
       Case 13
            SendKeys "{tab}"
       Case vbKeyUp
            SendKeys "+{tab}"
       Case vbKeyDown
            SendKeys "{tab}"
End Select
End Sub


Function shock()
If scanhead = True And david = True Then
     MsgBox "Good Day SCANHEAD, I've been expecting you!", vbInformation + vbOKOnly, "Greetings"
     MainForm.StatusBar1.Panels(1).Text = "(c)Soft/Image Hieroglyphix 1998-2003    | THE SCANHED OF WINDSOR ::MADZBRI WELCOME"
     MainForm.Show
     Unload frmLogin
 End If
End Function
