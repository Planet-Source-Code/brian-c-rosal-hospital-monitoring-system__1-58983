VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form PHmember 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "Phil Health Members: Data  Section "
   ClientHeight    =   6195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9915
   ControlBox      =   0   'False
   Icon            =   "PHmember.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   Picture         =   "PHmember.frx":0BC2
   ScaleHeight     =   6195
   ScaleMode       =   0  'User
   ScaleWidth      =   9915
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin Project1.chameleonButton butx 
      Height          =   285
      Left            =   90
      TabIndex        =   24
      Top             =   195
      Width           =   285
      _extentx        =   503
      _extenty        =   503
      btype           =   5
      tx              =   "X"
      enab            =   -1
      font            =   "PHmember.frx":AEA7
      coltype         =   2
      focusr          =   -1
      bcol            =   12632256
      bcolo           =   12582912
      fcol            =   0
      fcolo           =   16777215
      mcol            =   12632256
      mptr            =   1
      micon           =   "PHmember.frx":AED3
      umcol           =   -1
      soft            =   0
      picpos          =   0
      ngrey           =   0
      fx              =   1
      hand            =   0
      check           =   0
      value           =   0
   End
   Begin Project1.chameleonButton cmdExit 
      Height          =   825
      Left            =   8130
      TabIndex        =   8
      Top             =   4170
      Width           =   1425
      _extentx        =   2514
      _extenty        =   1455
      btype           =   5
      tx              =   "E&xit"
      enab            =   -1
      font            =   "PHmember.frx":AEF1
      coltype         =   2
      focusr          =   -1
      bcol            =   12632256
      bcolo           =   12632256
      fcol            =   0
      fcolo           =   255
      mcol            =   12632256
      mptr            =   1
      micon           =   "PHmember.frx":AF15
      umcol           =   -1
      soft            =   0
      picpos          =   0
      ngrey           =   0
      fx              =   1
      hand            =   0
      check           =   0
      value           =   0
   End
   Begin Project1.chameleonButton cmdcancel 
      Height          =   825
      Left            =   6600
      TabIndex        =   7
      Top             =   4170
      Width           =   1425
      _extentx        =   2514
      _extenty        =   1455
      btype           =   5
      tx              =   "&Cancel"
      enab            =   -1
      font            =   "PHmember.frx":AF33
      coltype         =   2
      focusr          =   -1
      bcol            =   12632256
      bcolo           =   12632256
      fcol            =   0
      fcolo           =   255
      mcol            =   12632256
      mptr            =   1
      micon           =   "PHmember.frx":AF57
      umcol           =   -1
      soft            =   0
      picpos          =   0
      ngrey           =   0
      fx              =   1
      hand            =   0
      check           =   0
      value           =   0
   End
   Begin VB.Frame Frame3 
      Height          =   885
      Left            =   495
      TabIndex        =   22
      Top             =   4095
      Width           =   5520
      Begin Project1.chameleonButton cmdreset 
         Height          =   585
         Left            =   4425
         TabIndex        =   23
         Top             =   195
         Width           =   990
         _extentx        =   1746
         _extenty        =   1032
         btype           =   5
         tx              =   "&Reset"
         enab            =   -1
         font            =   "PHmember.frx":AF75
         coltype         =   2
         focusr          =   -1
         bcol            =   12632256
         bcolo           =   12632256
         fcol            =   0
         fcolo           =   255
         mcol            =   12632256
         mptr            =   1
         micon           =   "PHmember.frx":AFA1
         umcol           =   -1
         soft            =   0
         picpos          =   0
         ngrey           =   0
         fx              =   1
         hand            =   0
         check           =   0
         value           =   0
      End
      Begin Project1.chameleonButton cmdadd 
         Height          =   555
         Left            =   90
         TabIndex        =   9
         Top             =   210
         Width           =   1260
         _extentx        =   2223
         _extenty        =   979
         btype           =   8
         tx              =   "&New Member"
         enab            =   -1
         font            =   "PHmember.frx":AFBF
         coltype         =   2
         focusr          =   -1
         bcol            =   12632256
         bcolo           =   12632256
         fcol            =   0
         fcolo           =   255
         mcol            =   12632256
         mptr            =   1
         micon           =   "PHmember.frx":AFEB
         picn            =   "PHmember.frx":B009
         umcol           =   -1
         soft            =   0
         picpos          =   0
         ngrey           =   0
         fx              =   1
         hand            =   0
         check           =   0
         value           =   0
      End
      Begin Project1.chameleonButton cmdmod 
         Height          =   555
         Left            =   1395
         TabIndex        =   10
         Top             =   195
         Width           =   1260
         _extentx        =   2223
         _extenty        =   979
         btype           =   8
         tx              =   "&Modify"
         enab            =   -1
         font            =   "PHmember.frx":BBDD
         coltype         =   2
         focusr          =   -1
         bcol            =   12632256
         bcolo           =   12632256
         fcol            =   0
         fcolo           =   255
         mcol            =   12632256
         mptr            =   1
         micon           =   "PHmember.frx":BC09
         picn            =   "PHmember.frx":BC27
         umcol           =   -1
         soft            =   0
         picpos          =   0
         ngrey           =   0
         fx              =   1
         hand            =   0
         check           =   0
         value           =   0
      End
      Begin Project1.chameleonButton cmdsave 
         Height          =   555
         Left            =   2700
         TabIndex        =   6
         Top             =   195
         Width           =   1260
         _extentx        =   2223
         _extenty        =   979
         btype           =   8
         tx              =   "&Save"
         enab            =   -1
         font            =   "PHmember.frx":C07B
         coltype         =   2
         focusr          =   -1
         bcol            =   12632256
         bcolo           =   12632256
         fcol            =   0
         fcolo           =   255
         mcol            =   12632256
         mptr            =   1
         micon           =   "PHmember.frx":C0A7
         picn            =   "PHmember.frx":C0C5
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
   Begin VB.Frame Frame1 
      Caption         =   "Viewing Section"
      ForeColor       =   &H000000FF&
      Height          =   3975
      Left            =   510
      TabIndex        =   21
      Top             =   90
      Width           =   4080
      Begin MSComctlLib.ListView ListView 
         Height          =   3600
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   3810
         _ExtentX        =   6720
         _ExtentY        =   6350
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483624
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Recno"
            Object.Width           =   9
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
   Begin VB.Frame Frame2 
      Caption         =   "Data Entry Section"
      ForeColor       =   &H00C00000&
      Height          =   3975
      Left            =   4725
      TabIndex        =   14
      Top             =   90
      Width           =   4830
      Begin Project1.chameleonButton cmdsrch 
         Height          =   540
         Left            =   2535
         TabIndex        =   11
         Top             =   600
         Width           =   1080
         _extentx        =   1905
         _extenty        =   953
         btype           =   5
         tx              =   "Searc&h PH Member"
         enab            =   -1
         font            =   "PHmember.frx":C519
         coltype         =   2
         focusr          =   -1
         bcol            =   12632256
         bcolo           =   12632256
         fcol            =   0
         fcolo           =   255
         mcol            =   12632256
         mptr            =   1
         micon           =   "PHmember.frx":C545
         umcol           =   -1
         soft            =   0
         picpos          =   0
         ngrey           =   0
         fx              =   1
         hand            =   0
         check           =   0
         value           =   0
      End
      Begin VB.TextBox txtmem 
         BackColor       =   &H80000018&
         Height          =   300
         Index           =   0
         Left            =   1575
         TabIndex        =   15
         Top             =   210
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.TextBox txtmem 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   1
         Left            =   165
         MaxLength       =   15
         TabIndex        =   1
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox txtmem 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   2
         Left            =   165
         TabIndex        =   2
         Top             =   1230
         Width           =   3435
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
         Height          =   285
         Index           =   3
         Left            =   165
         TabIndex        =   3
         Top             =   1905
         Width           =   3810
      End
      Begin VB.TextBox txtmem 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   4
         Left            =   4215
         MaxLength       =   1
         TabIndex        =   4
         Top             =   1905
         Width           =   300
      End
      Begin VB.TextBox txtmem 
         BackColor       =   &H80000018&
         Height          =   1125
         Index           =   5
         Left            =   165
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   2640
         Width           =   4380
      End
      Begin Project1.chameleonButton cmddiag 
         Height          =   930
         Left            =   3690
         TabIndex        =   12
         Top             =   600
         Width           =   1035
         _extentx        =   1905
         _extenty        =   529
         btype           =   5
         tx              =   "&View Patient(s)"
         enab            =   -1
         font            =   "PHmember.frx":C563
         coltype         =   2
         focusr          =   -1
         bcol            =   12632256
         bcolo           =   12632256
         fcol            =   0
         fcolo           =   255
         mcol            =   12632256
         mptr            =   1
         micon           =   "PHmember.frx":C587
         umcol           =   -1
         soft            =   0
         picpos          =   0
         ngrey           =   0
         fx              =   1
         hand            =   0
         check           =   0
         value           =   0
      End
      Begin VB.Label Label1 
         Caption         =   "PH Member ID No."
         Height          =   255
         Index           =   0
         Left            =   135
         TabIndex        =   20
         Top             =   330
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Last Name"
         Height          =   255
         Index           =   1
         Left            =   135
         TabIndex        =   19
         Top             =   990
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "First Name"
         Height          =   255
         Index           =   2
         Left            =   135
         TabIndex        =   18
         Top             =   1665
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "MI"
         Height          =   255
         Index           =   3
         Left            =   4215
         TabIndex        =   17
         Top             =   1665
         Width           =   405
      End
      Begin VB.Label Label1 
         Caption         =   "Address"
         Height          =   255
         Index           =   4
         Left            =   150
         TabIndex        =   16
         Top             =   2370
         Width           =   1575
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5175
      Index           =   0
      Left            =   -30
      TabIndex        =   0
      Top             =   -15
      Width           =   9720
      _ExtentX        =   17145
      _ExtentY        =   9128
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
Attribute VB_Name = "PHmember"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sql, rs, y As Variant, ctr As Byte
Dim MyType, passString As String
Dim formod As String
Dim mFormRegion As Long


Private Sub chameleonButton1_Click()

End Sub

Private Sub butx_Click()
If MadzB = True Then
   Set WSMain = DBEngine.Workspaces(0)
   Set DBMain = WSMain.OpenDatabase(App.Path + "\hospital.mdb", False, False, ";pwd=scanhead")
   sql = "select * from PHmember order by 2;"
    Set rs = DBMain.OpenRecordset(sql)
    PHconfined.ListView.ListItems.Clear
    Do Until rs.EOF
        Set y = PHconfined.ListView.ListItems.Add(, , rs.Fields(0))
        y.SubItems(1) = rs.Fields(1)
        y.SubItems(2) = rs.Fields(2)
        y.SubItems(3) = rs.Fields(3)
        y.SubItems(4) = rs.Fields(4)
        y.SubItems(5) = rs.Fields(5)
        rs.MoveNext
    Loop
 
   PHconfined.txtcon(1).Text = txtmem(1).Text
   PHconfined.memtext1.Text = txtmem(2).Text
   PHconfined.memtext2.Text = txtmem(3).Text
   PHconfined.memtext3.Text = txtmem(4).Text
   
 End If

MadzB = False
Unload Me
MainForm.mem.Enabled = True
MainForm.Toolbar1.Buttons.Item(1).Enabled = True
End Sub

Private Sub cmdAdd_Click()
    SetText True
    cmdsave.Enabled = True
    cmdsrch.Enabled = False
    cmdcancel.Enabled = True
    cmdreset.Enabled = False
    ListView.Enabled = False
    cmdmod.Enabled = False
    cmdadd.Enabled = False
    ClearText
    txtmem(0).Text = AutoRecordNumber
    txtmem(1).SetFocus
    MyType = "ADD"
End Sub


Private Sub cmdCancel_Click()
    SetText False
    cmdsave.Enabled = False
    cmdreset.Enabled = False
    cmdcancel.Enabled = False
    ListView.Enabled = True
    cmdmod.Enabled = True
    cmdadd.Enabled = True
    cmdsrch.Enabled = True
     ClearText
    MyType = ""
End Sub

Private Sub cmddiag_Click()
Dim tb, testdate, NOID As String
testdate = ""
NOID = ""
tb = Trim(PHmember.txtmem(1).Text)
sql = "select plastname,pfirstname,pmi,page,padiagnose,patientno,idno,datec from Patient where idno like '*" & tb & "*' and format(dateD,'mm/dd/yyyy') = '" & testdate & "' AND TRIM(IDNO) <> TRIM('" & NOID & "') order by plastname,pfirstname,pmi;"
           Set rs = DBMain.OpenRecordset(sql)
              If rs.RecordCount < 0 Or rs.RecordCount = 0 Then
                       MsgBox "No Admitted Patient(s) under the selected PH member!", vbInformation, "Sorry"
                       Exit Sub
              Else
                       
                       listcon.Show vbModal
              End If
End Sub

Private Sub cmdExit_Click()
If MadzB = True Then
    Set WSMain = DBEngine.Workspaces(0)
    Set DBMain = WSMain.OpenDatabase(App.Path + "\hospital.mdb", False, False, ";pwd=scanhead")
    sql = "select * from PHmember order by 2;"
    Set rs = DBMain.OpenRecordset(sql)
    PHconfined.ListView.ListItems.Clear
    Do Until rs.EOF
        Set y = PHconfined.ListView.ListItems.Add(, , rs.Fields(0))
        y.SubItems(1) = rs.Fields(1)
        y.SubItems(2) = rs.Fields(2)
        y.SubItems(3) = rs.Fields(3)
        y.SubItems(4) = rs.Fields(4)
        y.SubItems(5) = rs.Fields(5)
        rs.MoveNext
    Loop
 
   PHconfined.txtcon(1).Text = txtmem(1).Text
   PHconfined.memtext1.Text = txtmem(2).Text
   PHconfined.memtext2.Text = txtmem(3).Text
   PHconfined.memtext3.Text = txtmem(4).Text
   
 End If

MadzB = False
Unload Me
MainForm.mem.Enabled = True
MainForm.Toolbar1.Buttons.Item(1).Enabled = True
End Sub

Private Sub cmdmod_Click()
    If txtmem(0).Text = "" Then
        MsgBox "There's no Record to Modify ", vbExclamation, "Confirmation"
        Exit Sub
    End If
    SetText True
    ListView.Enabled = False
    cmdadd.Enabled = False
    cmdmod.Enabled = False
    cmdsave.Enabled = True
    cmdreset.Enabled = False
    cmdcancel.Enabled = True
    MyType = "EDIT"
    txtmem(1).SetFocus
End Sub

Private Sub cmdreset_Click()
DisplayLst
End Sub

Private Sub CMDSAVE_Click()
On Error GoTo TRAPPER
    Dim TRS As DAO.Recordset
    Dim TQR As DAO.QueryDef
    Dim P, madzsrch As String
    Dim Query, tb, td, tx As String
    Dim List As ListItem
    Dim X As Long
    Dim Flag As Boolean
    Flag = False
    For X = 0 To 4
        If txtmem(X).Text = "" Then Flag = True
    Next X
    
    
    
    If Flag Then
        MsgBox "Please Enter all information to Continue ?", vbInformation, "Confirmation"
        GoTo TRAPPER
     End If
     
     
    
    If MyType = "ADD" Then
       '//this area is for searchin the inputted record
              tb = Trim(txtmem(1).Text)
              madzsrch = "SELECT * FROM PHmember WHERE trim(idno) like '*" & tb & "*';"
              Set TRS = DBMain.OpenRecordset(madzsrch)
              If TRS.RecordCount > 0 Then
                       MsgBox "Record already exist or Invalid input!", vbCritical, "Sorry"
                       GoTo TRAPPER
               End If
      '/****************************************
       
       P = "INSERT INTO PHmember (recno,idno,lastname,firstname,MI,address) VALUES ('" & txtmem(0).Text & "','" & txtmem(1).Text & "','" & txtmem(2).Text & "','" & txtmem(3).Text & "','" & txtmem(4).Text & "','" & txtmem(5).Text & "');"
        Set TQR = DBMain.CreateQueryDef("", P)
        TQR.Execute
        Set List = ListView.ListItems.Add(, , txtmem(0).Text)
        With List
            .SubItems(1) = txtmem(1).Text
            .SubItems(2) = txtmem(2).Text '
            .SubItems(3) = txtmem(3).Text '
            .SubItems(4) = txtmem(4).Text '
            .SubItems(5) = txtmem(5).Text 'address
        End With
       '************************************************
    
   ElseIf MyType = "EDIT" Then
           '******** Editing Section : the scanhead *************
            passString = "UPDATE PHmember SET idno='" & txtmem(1).Text & "', lastname='" & txtmem(2).Text & "',firstname='" & txtmem(3).Text & "',MI='" & txtmem(4).Text & "',address='" & txtmem(5).Text & "' WHERE recno='" & txtmem(0).Text & "' ;"
            DBMain.Execute passString
            ListView.Enabled = True
            MyType = ""
            Set List = ListView.FindItem(txtmem(0).Text, , , lvwPartial)
            With List
              .SubItems(1) = txtmem(1).Text
              .SubItems(2) = txtmem(2).Text
              .SubItems(3) = txtmem(3).Text
              .SubItems(4) = txtmem(4).Text
              .SubItems(5) = txtmem(4).Text
            End With
      End If
      
        SetText False
        ListView.Enabled = True
        cmdadd.Enabled = True
        cmdmod.Enabled = True
        cmdsave.Enabled = False
        cmdcancel.Enabled = False
        cmdreset.Enabled = False
        cmdsrch.Enabled = True
TRAPPER:
  Exit Sub
    
End Sub

Private Sub cmdsrch_Click()
 Load srchform
 cmdreset.Enabled = True
 srchform.Show vbModal
End Sub



Private Sub Form_Load()
DisableCloseButton Me
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
    
           SetText False
           cmdadd.Enabled = True
           cmdsave.Enabled = False
           cmdmod.Enabled = True
           cmdcancel.Enabled = False
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
   PHmember.txtmem(0).Text = TRS.Fields(0)
   PHmember.txtmem(1).Text = TRS.Fields(1)
   PHmember.txtmem(2).Text = TRS.Fields(2)
   PHmember.txtmem(3).Text = TRS.Fields(3)
   PHmember.txtmem(4).Text = TRS.Fields(4)
   PHmember.txtmem(5).Text = TRS.Fields(5)
   SetText False
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
    ClearText
    SetText False
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

