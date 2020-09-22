VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form PHmemDelete 
   BackColor       =   &H80000007&
   Caption         =   "PHmember Data Deletion Section"
   ClientHeight    =   5550
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9915
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "PHmemDelete.frx":0000
   ScaleHeight     =   5550
   ScaleWidth      =   9915
   WindowState     =   2  'Maximized
   Begin Project1.chameleonButton butx 
      Height          =   285
      Left            =   120
      TabIndex        =   21
      Top             =   165
      Width           =   285
      _extentx        =   503
      _extenty        =   503
      btype           =   5
      tx              =   "X"
      enab            =   -1
      font            =   "PHmemDelete.frx":A2E5
      coltype         =   2
      focusr          =   -1
      bcol            =   12632256
      bcolo           =   12582912
      fcol            =   0
      fcolo           =   16777215
      mcol            =   12632256
      mptr            =   1
      micon           =   "PHmemDelete.frx":A311
      umcol           =   -1
      soft            =   0
      picpos          =   0
      ngrey           =   0
      fx              =   1
      hand            =   0
      check           =   0
      value           =   0
   End
   Begin VB.Frame Frame2 
      Caption         =   "Data Entry Section"
      ForeColor       =   &H00C00000&
      Height          =   4065
      Left            =   4755
      TabIndex        =   6
      Top             =   90
      Width           =   4830
      Begin VB.TextBox txtmem 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   690
         Index           =   5
         Left            =   165
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   2640
         Width           =   4380
      End
      Begin VB.TextBox txtmem 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   4215
         MaxLength       =   1
         TabIndex        =   11
         Top             =   1905
         Width           =   300
      End
      Begin VB.TextBox txtmem 
         BackColor       =   &H80000018&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "A"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   165
         TabIndex        =   10
         Top             =   1905
         Width           =   3810
      End
      Begin VB.TextBox txtmem 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   165
         TabIndex        =   9
         Top             =   1230
         Width           =   3435
      End
      Begin VB.TextBox txtmem 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   165
         MaxLength       =   15
         TabIndex        =   8
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox txtmem 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   300
         Index           =   0
         Left            =   1800
         TabIndex        =   7
         Top             =   240
         Visible         =   0   'False
         Width           =   795
      End
      Begin Project1.chameleonButton cmdsrch 
         Height          =   540
         Left            =   2550
         TabIndex        =   13
         Top             =   585
         Width           =   1050
         _extentx        =   1429
         _extenty        =   582
         btype           =   5
         tx              =   "Searc&h PH Member"
         enab            =   -1
         font            =   "PHmemDelete.frx":A32F
         coltype         =   2
         focusr          =   -1
         bcol            =   12632256
         bcolo           =   12632256
         fcol            =   0
         fcolo           =   255
         mcol            =   4210752
         mptr            =   1
         micon           =   "PHmemDelete.frx":A353
         umcol           =   -1
         soft            =   0
         picpos          =   0
         ngrey           =   0
         fx              =   1
         hand            =   0
         check           =   0
         value           =   0
      End
      Begin Project1.chameleonButton cmddiag 
         Height          =   915
         Left            =   3720
         TabIndex        =   14
         Top             =   585
         Width           =   975
         _extentx        =   1693
         _extenty        =   529
         btype           =   5
         tx              =   "&View Patient(s)"
         enab            =   -1
         font            =   "PHmemDelete.frx":A371
         coltype         =   2
         focusr          =   -1
         bcol            =   12632256
         bcolo           =   12632256
         fcol            =   0
         fcolo           =   255
         mcol            =   4210752
         mptr            =   1
         micon           =   "PHmemDelete.frx":A395
         umcol           =   -1
         soft            =   0
         picpos          =   0
         ngrey           =   0
         fx              =   1
         hand            =   0
         check           =   0
         value           =   0
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H80000003&
         Height          =   525
         Left            =   2040
         Top             =   3420
         Width           =   2550
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Note: Check the data carefully before deleting."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   420
         Index           =   3
         Left            =   2175
         TabIndex        =   22
         Top             =   3465
         Width           =   2250
      End
      Begin VB.Label Label1 
         Caption         =   "Address"
         Height          =   255
         Index           =   4
         Left            =   150
         TabIndex        =   19
         Top             =   2370
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "MI"
         Height          =   255
         Index           =   3
         Left            =   4215
         TabIndex        =   18
         Top             =   1665
         Width           =   405
      End
      Begin VB.Label Label1 
         Caption         =   "First Name"
         Height          =   255
         Index           =   2
         Left            =   135
         TabIndex        =   17
         Top             =   1665
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Last Name"
         Height          =   255
         Index           =   1
         Left            =   135
         TabIndex        =   16
         Top             =   990
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "PH Member ID No."
         Height          =   255
         Index           =   0
         Left            =   135
         TabIndex        =   15
         Top             =   330
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Viewing Section"
      ForeColor       =   &H000000FF&
      Height          =   4080
      Left            =   495
      TabIndex        =   4
      Top             =   75
      Width           =   4080
      Begin MSComctlLib.ListView ListView 
         Height          =   3720
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   3810
         _ExtentX        =   6720
         _ExtentY        =   6562
         View            =   3
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Recno"
            Object.Width           =   4
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "PH Member ID No."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Last Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "First Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "MI"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Address"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame Frame3 
      Height          =   930
      Left            =   480
      TabIndex        =   1
      Top             =   4200
      Width           =   5520
      Begin Project1.chameleonButton cmdDelete 
         Height          =   630
         Left            =   135
         TabIndex        =   2
         Top             =   195
         Width           =   3225
         _extentx        =   5689
         _extenty        =   1111
         btype           =   8
         tx              =   "D&elete Record"
         enab            =   -1
         font            =   "PHmemDelete.frx":A3B3
         coltype         =   2
         focusr          =   -1
         bcol            =   12632256
         bcolo           =   12632256
         fcol            =   0
         fcolo           =   255
         mcol            =   12632256
         mptr            =   1
         micon           =   "PHmemDelete.frx":A3D7
         picn            =   "PHmemDelete.frx":A3F5
         umcol           =   -1
         soft            =   0
         picpos          =   0
         ngrey           =   0
         fx              =   1
         hand            =   0
         check           =   0
         value           =   0
      End
      Begin Project1.chameleonButton cmdreset 
         Height          =   600
         Left            =   4260
         TabIndex        =   3
         Top             =   210
         Width           =   1110
         _extentx        =   1958
         _extenty        =   1058
         btype           =   5
         tx              =   "&Reset"
         enab            =   -1
         font            =   "PHmemDelete.frx":A849
         coltype         =   2
         focusr          =   -1
         bcol            =   12632256
         bcolo           =   12632256
         fcol            =   0
         fcolo           =   255
         mcol            =   12632256
         mptr            =   1
         micon           =   "PHmemDelete.frx":A86D
         umcol           =   -1
         soft            =   0
         picpos          =   0
         ngrey           =   0
         fx              =   1
         hand            =   0
         check           =   0
         value           =   0
      End
   End
   Begin Project1.chameleonButton cmdexit 
      Height          =   795
      Left            =   8175
      TabIndex        =   0
      Top             =   4305
      Width           =   1425
      _extentx        =   2514
      _extenty        =   1402
      btype           =   5
      tx              =   "E&xit"
      enab            =   -1
      font            =   "PHmemDelete.frx":A88B
      coltype         =   2
      focusr          =   -1
      bcol            =   12632256
      bcolo           =   12632256
      fcol            =   0
      fcolo           =   255
      mcol            =   12632256
      mptr            =   1
      micon           =   "PHmemDelete.frx":A8AF
      umcol           =   -1
      soft            =   0
      picpos          =   0
      ngrey           =   0
      fx              =   1
      hand            =   0
      check           =   0
      value           =   0
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5295
      Index           =   0
      Left            =   -15
      TabIndex        =   20
      Top             =   -15
      Width           =   9810
      _ExtentX        =   17304
      _ExtentY        =   9340
      TabWidthStyle   =   2
      MultiRow        =   -1  'True
      ShowTips        =   0   'False
      HotTracking     =   -1  'True
      Placement       =   2
      TabMinWidth     =   0
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
      EndProperty
      Enabled         =   0   'False
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "PHmemDelete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sql, rs, y As Variant, ctr As Byte
Dim MyType, passString As String
Dim formod As String
Dim mFormRegion As Long



Private Sub butx_Click()

Unload Me
MainForm.MemDel.Enabled = True
End Sub


Private Sub cmdDelete_Click()
'the Scanhead rules :: Madz
Dim testdate, ask As String
testdate = ""
ask = "*"
If txtmem(0).Text = "" Then
   MsgBox "No PH Member to Delete!!", vbCritical, "Sorry"
Else
  If MsgBox("Are you sure you want to delete " + txtmem(2).Text + " ," + txtmem(3).Text + " ." + txtmem(4).Text + ", this will also remove the Patient(s) Confined & Discharged under his/her Account.", vbExclamation + vbYesNo, "Confirm Deletion") = vbYes Then
        Dim Dellist As ListItem
        '***** PHMember ***********'
        DBMain.Execute "DELETE * FROM Phmember WHERE recno ='" + txtmem(0).Text + "';"
        '***** Confined Patients***********'
        DBMain.Execute "DELETE * FROM Patient WHERE Idno = '" + txtmem(1).Text + "' and format(dateD,'mm/dd/yyyy') = '" & testdate & "' ;"
        '****** Discharged Patients **********
        DBMain.Execute "DELETE * FROM Pdischarged WHERE Idno = '" + txtmem(1).Text + "';"
        
        Set Dellist = ListView.FindItem(txtmem(0).Text, , , lvwPartial)
        ListView.ListItems.Remove Dellist.Index
        ClearTextBri
    End If
End If
End Sub

Private Sub cmddiag_Click()
Dim tb, testdate, NOID As String
testdate = ""
NOID = ""
tb = Trim(PHmemDelete.txtmem(1).Text)
sql = "select plastname,pfirstname,pmi,page,padiagnose,patientno,idno,datec from Patient where idno like '*" & tb & "*' and format(dateD,'mm/dd/yyyy') = '" & testdate & "' AND TRIM(IDNO) <> TRIM('" & NOID & "')order by plastname,pfirstname,pmi;"
           Set rs = DBMain.OpenRecordset(sql)
              If rs.RecordCount < 0 Or rs.RecordCount = 0 Then
                       MsgBox "No Admitted Patient(s) under the selected PH member!", vbInformation, "Sorry"
                       Exit Sub
              Else
                       
                       listcon2.Show vbModal
              End If
End Sub

Private Sub cmdExit_Click()
Unload Me
MainForm.MemDel.Enabled = True
End Sub



Private Sub cmdreset_Click()
ClearTextBri
DisplayLst
End Sub



Private Sub cmdsrch_Click()
 Load SFMemDel
 ClearTextBri
 cmdreset.Enabled = True
 SFMemDel.Show vbModal
End Sub



Private Sub Form_Load()
add3d
MadzB = False
 Set WSMain = DBEngine.Workspaces(0)
 Set DBMain = WSMain.OpenDatabase(App.Path + "\hospital.mdb", False, False, ";pwd=scanhead")
  sql = "select * from PHmember order by 2;"
    Set rs = DBMain.OpenRecordset(sql)
    Do Until rs.EOF
        Set y = ListView.ListItems.Add(, , rs.Fields(0))
        y.SubItems(1) = rs.Fields(1)
        y.SubItems(2) = rs.Fields(2)
        y.SubItems(3) = rs.Fields(3)
        y.SubItems(4) = rs.Fields(4)
        y.SubItems(5) = rs.Fields(5)
        rs.MoveNext
    Loop
    
           cmdDelete.Enabled = True
           cmdreset.Enabled = True
           
End Sub

Private Sub ListView_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo TRAPPER
    Dim X As Long
    Dim TQR As DAO.QueryDef
    Dim TRS As DAO.Recordset
    'Scanhead rules
    '*********
   
   Set TQR = DBMain.CreateQueryDef("", "SELECT * FROM PHmember WHERE recno ='" & ListView.SelectedItem.Text & "' ORDER BY IDNO")
   Set TRS = TQR.OpenRecordset()
   PHmemDelete.txtmem(0).Text = TRS.Fields(0)
   PHmemDelete.txtmem(1).Text = TRS.Fields(1)
   PHmemDelete.txtmem(2).Text = TRS.Fields(2)
   PHmemDelete.txtmem(3).Text = TRS.Fields(3)
   PHmemDelete.txtmem(4).Text = TRS.Fields(4)
   PHmemDelete.txtmem(5).Text = TRS.Fields(5)
   
TRAPPER:
End Sub

Public Function AutoRecordNumber() As String
    Dim TQR As DAO.QueryDef
    Dim TRS As DAO.Recordset
    Dim Start As String
    Start = "00000000"
    Set TQR = DBMain.CreateQueryDef("", "SELECT Recno FROM Phmember")
    Set TRS = TQR.OpenRecordset()
    Do While Not TRS.EOF
        TRS.FindFirst "Recno='M-" + Start + "'"
        If Not TRS.NoMatch Then
            Start = Format(Str(Val(Mid$(Start, 3)) + 1), "00000000")
        Else
            AutoRecordNumber = "M-" + Start
            Exit Function
        End If
    Loop
    AutoRecordNumber = "M-" + Start
End Function
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


Public Sub DisplayLst() 'Reset control display: gotcha the scanhedbri
On Error GoTo X
    
    Dim TQR As DAO.QueryDef
    Dim rs As DAO.Recordset
    Dim y As ListItem
     Dim sql As String
    Dim X As Long
   
    sql = "select * from PHmember order by 2;"
    Set rs = DBMain.OpenRecordset(sql)
     ListView.ListItems.Clear
    Do Until rs.EOF
        Set y = ListView.ListItems.Add(, , rs.Fields(0))
        y.SubItems(1) = rs.Fields(1)
        y.SubItems(2) = rs.Fields(2)
        y.SubItems(3) = rs.Fields(3)
        y.SubItems(4) = rs.Fields(4)
        y.SubItems(5) = rs.Fields(5)
        rs.MoveNext
    Loop
   
X:
End Sub

Function add3d()
Dim X As Integer
Add3DBorder Me
Add3DBorder ListView

For X = 1 To 5
 Add3DBorder txtmem(X)
Next X
End Function



Private Sub txtmem_KeyPress(Index As Integer, KeyAscii As Integer)
'the madzbry txt validation
Dim madzbry As String
 Select Case Index
        Case 1
         madzbry = "0123456789"
          If KeyAscii > 26 Then
            If InStr(madzbry, Chr(KeyAscii)) = 0 Then
              KeyAscii = 0
            End If
          End If
         Case 2, 3, 4
          madzbry = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ ."
          If KeyAscii > 26 Then
            If InStr(madzbry, Chr(KeyAscii)) = 0 Then
              KeyAscii = 0
            End If
          End If
    
End Select
End Sub


