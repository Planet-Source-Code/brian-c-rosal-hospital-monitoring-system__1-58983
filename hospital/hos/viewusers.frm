VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Viewusers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Users"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4245
   Icon            =   "viewusers.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   4245
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   2880
      Width           =   4050
      Begin Project1.chameleonButton Command1 
         Height          =   345
         Left            =   2220
         TabIndex        =   3
         Top             =   180
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   609
         BTYPE           =   5
         TX              =   "&Close"
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
         MICON           =   "viewusers.frx":0BC2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   1
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label1 
         Caption         =   "Select user to Edit/Delete"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2055
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2775
      Left            =   75
      TabIndex        =   0
      Top             =   60
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   4895
      SortKey         =   2
      View            =   3
      SortOrder       =   -1  'True
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483624
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "logName"
         Object.Width           =   1834
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "User Name"
         Object.Width           =   5362
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Position"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Password"
         Object.Width           =   2187
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "User Type"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "Viewusers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sql, rs, y As Variant, ctr As Byte

Private Sub Command1_Click()
    Unload Me
    frmUsrMngr.cmdAdd.Enabled = False
    frmUsrMngr.cmdedit.Enabled = True
    frmUsrMngr.cmdDelete.Enabled = True
    
    
End Sub



Public Sub Form_Load()
    Add3DBorder Me
    Add3DBorder ListView1
    sql = "select logname,name,pos,password,logtype from password order by 1;"
    Set rs = DBMain.OpenRecordset(sql)
    Do Until rs.EOF
        Set y = ListView1.ListItems.Add(, , rs.Fields(0))
        y.SubItems(1) = rs.Fields(1)
        y.SubItems(2) = rs.Fields(2)
        y.SubItems(3) = Decrypt(rs.Fields(3))
        y.SubItems(4) = rs.Fields(4)
        rs.MoveNext
    Loop
    'lbltot.Caption = rs.RecordCount
End Sub




Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
 On Error GoTo TRAPPER
    Dim X As Long
    Dim TQR As DAO.QueryDef
    Dim TRS As DAO.Recordset
    'sql kid scanhead rules
    '*********
   
   Set TQR = DBMain.CreateQueryDef("", "SELECT * FROM PASSWORD WHERE LOGNAME ='" & ListView1.SelectedItem.Text & "' ")
   Set TRS = TQR.OpenRecordset()
   
   frmUsrMngr.txtpass(0).Text = TRS.Fields(0)
   frmUsrMngr.txtpass(1).Text = TRS.Fields(1)
   frmUsrMngr.passcombo.Text = TRS.Fields(2)
   frmUsrMngr.txtpass(2).Text = TRS.Fields(3)
   frmUsrMngr.txtpass(3).Text = TRS.Fields(4)
   frmUsrMngr.txtpass(4).Text = TRS.Fields(5)
   
TRAPPER:
End Sub
