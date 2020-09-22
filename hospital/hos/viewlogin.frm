VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Viewlog 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Logged Users Section"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4800
   Icon            =   "viewlogin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   4800
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   165
      TabIndex        =   1
      Top             =   2520
      Width           =   4455
      Begin Project1.chameleonButton cmdclear 
         Height          =   450
         Left            =   120
         TabIndex        =   2
         Top             =   180
         Width           =   2085
         _ExtentX        =   3598
         _ExtentY        =   847
         BTYPE           =   5
         TX              =   "&Clear All"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
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
         MICON           =   "viewlogin.frx":0442
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   1
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Project1.chameleonButton Command1 
         Height          =   450
         Left            =   2265
         TabIndex        =   3
         Top             =   180
         Width           =   2085
         _ExtentX        =   3598
         _ExtentY        =   847
         BTYPE           =   5
         TX              =   "E&xit"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
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
         MICON           =   "viewlogin.frx":045E
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
   Begin VB.Data Mdata 
      Caption         =   "Log-in entries"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   180
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "login"
      Top             =   2595
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "viewlogin.frx":047A
      Height          =   2295
      Left            =   180
      TabIndex        =   0
      Top             =   195
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   4048
      _Version        =   393216
      FixedCols       =   0
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
      BorderStyle     =   0
      Appearance      =   0
      FormatString    =   "----------------------------------------"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Viewlog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdclear_Click()

If MsgBox("Are you sure you want to Remove all Logged Users?", vbExclamation + vbYesNo, "Clear all Logged Users") = vbYes Then
        Set WSMain = DBEngine.Workspaces(0)
        Set DBMain = WSMain.OpenDatabase(App.Path + "\hospital.mdb", False, False, ";pwd=scanhead")
 
        '***** LOGIN ***********'
        DBMain.Execute "DELETE * FROM LOGIN;"
        '**********************************
        MSFlexGrid1.Clear
        MSFlexGrid1.Refresh
     End If
End Sub

Private Sub Command1_Click()
Unload Me

Mdata.Database.Close
MainForm.SetFocus
End Sub


Private Sub Form_Load()
Add3DBorder Me
Add3DBorder MSFlexGrid1
Add3DBorder Command1


Mdata.DatabaseName = App.Path & "\Hospital.mdb"
Mdata.RecordSource = "login"
Mdata.RecordsetType = 0
Mdata.Refresh
End Sub

