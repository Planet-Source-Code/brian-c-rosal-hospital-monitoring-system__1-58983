VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form listcon2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PH-Med Confined Patients List"
   ClientHeight    =   3390
   ClientLeft      =   2205
   ClientTop       =   2880
   ClientWidth     =   7530
   Icon            =   "listcondelete.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   7530
   Begin Project1.chameleonButton command1 
      Height          =   405
      Left            =   5010
      TabIndex        =   2
      Top             =   2835
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   714
      BTYPE           =   5
      TX              =   "E&xit"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
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
      MICON           =   "listcondelete.frx":0BC2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame framecon 
      Caption         =   "Confined Patient(s)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2655
      Left            =   120
      TabIndex        =   1
      Top             =   75
      Width           =   7215
      Begin MSComctlLib.ListView ListView1 
         Height          =   2295
         Left            =   150
         TabIndex        =   3
         Top             =   255
         Width           =   6915
         _ExtentX        =   12197
         _ExtentY        =   4048
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Last Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "First Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "MI"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Age"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Admission  Diagnosed"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Patient No."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "PH Member ID No."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Admission Date/Time"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "(c) SoftImage Hieroglyphix:1998-03:Scanhead:DBrosal"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   315
      TabIndex        =   0
      Top             =   2895
      Width           =   4215
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000009&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000009&
      FillColor       =   &H00808000&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   105
      Shape           =   4  'Rounded Rectangle
      Top             =   2835
      Width           =   4695
   End
End
Attribute VB_Name = "listcon2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sql, rs, y As Variant, ctr As Byte
Dim MyType, passString, madzsrch As String
Dim formod As String


Private Sub chameleonButton1_Click()

End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim tb, testdate, NOID As String
Add3DBorder Me
Add3DBorder ListView1
testdate = ""
NOID = ""
tb = Trim(PHmemDelete.txtmem(1).Text)
If tb = "" Then
  framecon.Caption = "All Patients Confined"
Else
framecon.Caption = "Patient(s) Confined under PHmember : " + tb
End If
sql = "select plastname,pfirstname,pmi,page,padiagnose,patientno,idno,datec,timein from Patient where idno like '*" & tb & "*' and format(dateD,'mm/dd/yyyy') = '" & testdate & "' AND ISNULL(IDNO)=FALSE order by plastname,pfirstname,pmi;"
           Set rs = DBMain.OpenRecordset(sql)
              If rs.RecordCount < 0 Or rs.RecordCount = 0 Then
                       MsgBox "Record not found!", vbCritical, "Sorry"
                       GoTo TRAPPER
              Else
                Do Until rs.EOF
                    Set y = ListView1.ListItems.Add(, , rs.Fields(0))
                        y.SubItems(1) = rs.Fields(1)
                        y.SubItems(2) = rs.Fields(2)
                        y.SubItems(3) = rs.Fields(3)
                        y.SubItems(4) = rs.Fields(4)
                        y.SubItems(5) = rs.Fields(5)
                        y.SubItems(6) = rs.Fields(6)
                        y.SubItems(7) = Format(rs.Fields(7), "mm/dd/yyyy") + " " + Format(rs.Fields(8), "medium time")
                        
                        rs.MoveNext
                 Loop
              End If
           '*************************

TRAPPER:
End Sub

