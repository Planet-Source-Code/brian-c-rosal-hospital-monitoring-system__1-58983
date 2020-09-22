VERSION 5.00
Begin VB.Form frmUsrMngr 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Manager"
   ClientHeight    =   3345
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   7125
   Icon            =   "frmUsrMngr.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   7125
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame Frame2 
      Height          =   2895
      Left            =   5040
      TabIndex        =   11
      Top             =   225
      Width           =   1950
      Begin Project1.chameleonButton cmdAdd 
         Height          =   390
         Left            =   135
         TabIndex        =   14
         Top             =   210
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   688
         BTYPE           =   5
         TX              =   "&New User"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
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
         MICON           =   "frmUsrMngr.frx":0442
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   1
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Project1.chameleonButton cmdedit 
         Height          =   390
         Left            =   135
         TabIndex        =   15
         Top             =   645
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   688
         BTYPE           =   5
         TX              =   "&Edit User"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
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
         MICON           =   "frmUsrMngr.frx":045E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   1
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Project1.chameleonButton cmdDelete 
         Height          =   390
         Left            =   135
         TabIndex        =   16
         Top             =   1080
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   688
         BTYPE           =   5
         TX              =   "&Delete  User"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
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
         MICON           =   "frmUsrMngr.frx":047A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   1
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Project1.chameleonButton cmdSave 
         Height          =   390
         Left            =   135
         TabIndex        =   17
         Top             =   1515
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   688
         BTYPE           =   5
         TX              =   "&Save  User"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
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
         MICON           =   "frmUsrMngr.frx":0496
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   1
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Project1.chameleonButton cmdCancel 
         Height          =   615
         Left            =   135
         TabIndex        =   18
         Top             =   2190
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   1085
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
         MICON           =   "frmUsrMngr.frx":04B2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   1
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " User Profile "
      ForeColor       =   &H00FF0000&
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4815
      Begin VB.TextBox txtpass 
         DataField       =   "logname"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   4
         Left            =   3240
         MaxLength       =   8
         TabIndex        =   12
         Top             =   480
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtpass 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   1200
         MaxLength       =   20
         PasswordChar    =   "o"
         TabIndex        =   5
         Top             =   2280
         Width           =   3375
      End
      Begin VB.TextBox txtpass 
         Height          =   285
         Index           =   2
         Left            =   1200
         MaxLength       =   30
         TabIndex        =   4
         Top             =   1800
         Width           =   3375
      End
      Begin VB.TextBox txtpass 
         Height          =   285
         Index           =   1
         Left            =   1200
         MaxLength       =   40
         TabIndex        =   2
         Top             =   900
         Width           =   3375
      End
      Begin VB.ComboBox passcombo 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmUsrMngr.frx":04CE
         Left            =   1200
         List            =   "frmUsrMngr.frx":04D8
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox txtpass 
         Height          =   285
         Index           =   0
         Left            =   1200
         MaxLength       =   15
         TabIndex        =   1
         Top             =   480
         Width           =   1935
      End
      Begin Project1.chameleonButton Viewpass 
         Height          =   315
         Left            =   3240
         TabIndex        =   19
         Top             =   480
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   556
         BTYPE           =   5
         TX              =   "&View All"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
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
         MICON           =   "frmUsrMngr.frx":04F1
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   1
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblLabels 
         Caption         =   "Warning: Case sensitive"
         Height          =   255
         Index           =   4
         Left            =   1200
         TabIndex        =   13
         Top             =   2595
         Width           =   2775
      End
      Begin VB.Label lblLabels 
         Caption         =   "Log Name:"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label lblLabels 
         Caption         =   "User Name:"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   9
         Top             =   930
         Width           =   1215
      End
      Begin VB.Label lblLabels 
         Caption         =   "User Type:"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   8
         Top             =   1380
         Width           =   975
      End
      Begin VB.Label lblLabels 
         Caption         =   "Position:"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   7
         Top             =   1815
         Width           =   1215
      End
      Begin VB.Label lblLabels 
         Caption         =   "Password:"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   6
         Top             =   2280
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmUsrMngr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'the scanhead of windsor david brian rosal
Dim MyType As String
Dim passString As String
Dim madZbry2 As String


Private Sub cmdCancel_Click()
MainForm.Manager.Enabled = True
MainForm.Toolbar2.Buttons.Item(1).Enabled = True
Unload Me
MainForm.SetFocus
End Sub

Private Sub cmdDelete_Click()
 
  If MsgBox("Are you sure of what you are doing? ", vbExclamation + vbYesNo, "Deletion confirm") = vbYes Then
        Dim Dellist As ListItem
        DBMain.Execute "DELETE * FROM password WHERE recno ='" + txtpass(4).Text + "';"
    End If
  Clearpass
  setTextPass False
  cmdadd.Enabled = True
  cmdsave.Enabled = False
  cmdcancel.Enabled = True
  cmdedit.Enabled = False
  cmdDelete.Enabled = False
End Sub

Private Sub cmdEdit_Click()
 
    setTextPass True
    cmdadd.Enabled = False
    cmdDelete.Enabled = False
    cmdedit.Enabled = False
    cmdsave.Enabled = True
    cmdcancel.Enabled = True
    MyType = "EDIT"
    txtpass(0).SetFocus
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
       Case 13
            SendKeys "{tab}"
       Case vbKeyUp
            SendKeys "+{tab}"
       Case vbKeyDown
            SendKeys "{tab}"
End Select
End Sub

Private Sub Form_Load()
On Error GoTo TRAPPER
 '************* wasa man
 Add3DBorder Me
 Add3DBorder txtpass(0)
 Add3DBorder txtpass(1)
 Add3DBorder txtpass(2)
 Add3DBorder txtpass(3)
 '******* shock attack scanhead
 Set WSMain = DBEngine.Workspaces(0)
 Set DBMain = WSMain.OpenDatabase(App.Path + "\Hospital.mdb", False, False, ";pwd=scanhead")
 setTextPass False
 passcombo.ListIndex = 1
  cmdadd.Enabled = True
  cmdsave.Enabled = False
  cmdcancel.Enabled = True
  cmdedit.Enabled = False
  cmdDelete.Enabled = False
          
TRAPPER:
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdAdd_Click()
    setTextPass True
    Dim now As Date
    cmdsave.Enabled = True
    cmdcancel.Enabled = True
    cmdDelete.Enabled = False
    cmdedit.Enabled = False
    cmdadd.Enabled = False
    Clearpass
    txtpass(4).Text = Arecno
    txtpass(0).SetFocus
    MyType = "ADD"
End Sub
Private Sub CMDSAVE_Click()
   On Error GoTo TRAPPER
    Dim TRS As DAO.Recordset
    Dim TQR As DAO.QueryDef
    Dim P, madzsrch, pass, utype As String
    Dim Query, tb, td As String
    Dim List As ListItem
    Dim X As Long
    Dim Flag As Boolean
    Flag = False
  
    
    For X = 0 To 3
        If txtpass(X).Text = "" Then Flag = True
    Next X
      If passcombo.Text = "" Then Flag = True
   
    
    If Flag Then
        MsgBox "Please Enter all information to Continue ?", vbInformation, "Confirm"
        GoTo X
    End If
    
   
    
    If MyType = "ADD" Then
      
      '//this area is for searchin the inputted record
              'td = Encrypt(txtpass(3))
              madzsrch = "SELECT * FROM password WHERE logname = '" & txtpass(0).Text & "';"
              Set TRS = DBMain.OpenRecordset(madzsrch)
              If TRS.RecordCount > 0 Then
                       MsgBox "change the log-in name pls!", vbCritical, "Sorry"
                       GoTo X
               End If
             
              td = Encrypt(txtpass(3))
              madzsrch = "SELECT * FROM password WHERE password = '" & td & "';"
              Set TRS = DBMain.OpenRecordset(madzsrch)
              If TRS.RecordCount > 0 Then
                       MsgBox "change the password pls!", vbCritical, "Sorry"
                       GoTo X
               End If
      '/****************************************
            madzsrch = "SELECT * FROM password WHERE logtype = '" & passcombo.Text & "' ;"
              Set TRS = DBMain.OpenRecordset(madzsrch)
              If TRS.RecordCount > 0 Then
                       If passcombo = "Administrator" Then
                          MsgBox "There can only be one administrator!", vbCritical, "Sorry"
                          GoTo X
                       End If
               End If
      '/****************************************
      
       
        pass = Encrypt(txtpass(3))
        P = "INSERT INTO Password (logname,name,logtype,pos,password,recno) VALUES ('" & txtpass(0).Text & "','" & txtpass(1).Text & "','" & passcombo.Text & "','" & txtpass(2).Text & "','" & pass & "','" & txtpass(4).Text & "');"
        Set TQR = DBMain.CreateQueryDef("", P)
        TQR.Execute
     
    ElseIf MyType = "EDIT" Then
     '//this area is for searchin the inputted record
              madzsrch = "SELECT * FROM password WHERE logname = '" & txtpass(0).Text & "';"
              Set TRS = DBMain.OpenRecordset(madzsrch)
              If TRS.RecordCount > 0 Then
                       MsgBox "change the log-in name pls!", vbCritical, "Sorry"
                       GoTo X
               End If
              
                      
              
              td = Encrypt(txtpass(3))
              madzsrch = "SELECT * FROM password WHERE password = '" & td & "';"
              Set TRS = DBMain.OpenRecordset(madzsrch)
              If TRS.RecordCount > 0 Then
                       MsgBox "change the password pls!", vbCritical, "Sorry"
                       GoTo X
               End If
      '/****************************************
            madzsrch = "SELECT * FROM password WHERE logtype = '" & passcombo.Text & "' ;"
              Set TRS = DBMain.OpenRecordset(madzsrch)
              If TRS.RecordCount > 0 Then
                       If passcombo = "Administrator" Then
                          MsgBox "There can only be one administrator!", vbCritical, "Sorry"
                          GoTo X
                       End If
               End If
      '/****************************************
 
        
        pass = Encrypt(txtpass(3))
        passString = "UPDATE Password SET logname ='" & txtpass(0).Text & "',name ='" & txtpass(1).Text & "', logtype='" & passcombo.Text & "',pos='" & txtpass(2).Text & "',password='" & pass & "' WHERE recno ='" & txtpass(4).Text & "';"
        DBMain.Execute passString
        setTextPass False
        MyType = ""
        
    End If
        setTextPass False
        cmdadd.Enabled = True
        cmdDelete.Enabled = False
        cmdedit.Enabled = False
        cmdsave.Enabled = False
        cmdcancel.Enabled = True
         
X:
    Exit Sub
TRAPPER:
    If Err.Number = 3075 Then
        MsgBox "Please input valid data to continue ?", vbExclamation, "Confirm"
   
    End If
End Sub
Public Function Arecno() As String
    Dim TQR As DAO.QueryDef
    Dim TRS As DAO.Recordset
    Dim Start As String
    Start = "00000"
    Set TQR = DBMain.CreateQueryDef("", "SELECT Recno FROM password")
    Set TRS = TQR.OpenRecordset()
    Do While Not TRS.EOF
        TRS.FindFirst "Recno='X-" + Start + "'"
        If Not TRS.NoMatch Then
            Start = Format(Str(Val(Mid$(Start, 3)) + 1), "00000")
        Else
            Arecno = "X-" + Start
            Exit Function
        End If
    Loop
    Arecno = "X-" + Start
End Function

Private Sub Form_Unload(Cancel As Integer)
MainForm.Manager.Enabled = True
MainForm.Toolbar1.Buttons.Item(3).Enabled = True

End Sub

Private Sub Viewpass_Click()
Viewusers.Show vbModal
End Sub
