VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PHconDelete 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   5070
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11475
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "PHconDelete.frx":0000
   ScaleHeight     =   5070
   ScaleWidth      =   11475
   WindowState     =   2  'Maximized
   Begin Project1.chameleonButton cmdexit 
      Height          =   825
      Left            =   8835
      TabIndex        =   25
      Top             =   3000
      Width           =   1785
      _extentx        =   3149
      _extenty        =   1455
      btype           =   5
      tx              =   "E&xit"
      enab            =   -1
      font            =   "PHconDelete.frx":A2E5
      coltype         =   2
      focusr          =   -1
      bcol            =   12632256
      bcolo           =   12632256
      fcol            =   0
      fcolo           =   255
      mcol            =   12632256
      mptr            =   1
      micon           =   "PHconDelete.frx":A309
      umcol           =   -1
      soft            =   0
      picpos          =   0
      ngrey           =   0
      fx              =   1
      hand            =   0
      check           =   0
      value           =   0
   End
   Begin Project1.chameleonButton butx 
      Height          =   285
      Left            =   105
      TabIndex        =   24
      Top             =   90
      Width           =   285
      _extentx        =   503
      _extenty        =   503
      btype           =   5
      tx              =   "X"
      enab            =   -1
      font            =   "PHconDelete.frx":A327
      coltype         =   2
      focusr          =   -1
      bcol            =   12632256
      bcolo           =   12582912
      fcol            =   0
      fcolo           =   16777215
      mcol            =   12632256
      mptr            =   1
      micon           =   "PHconDelete.frx":A353
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
      Caption         =   "Patient's Confinement Data"
      ForeColor       =   &H00C00000&
      Height          =   2880
      Left            =   4395
      TabIndex        =   6
      Top             =   30
      Width           =   6240
      Begin VB.TextBox txtconM 
         BackColor       =   &H80000016&
         Enabled         =   0   'False
         Height          =   300
         Left            =   2055
         MaxLength       =   15
         TabIndex        =   22
         Top             =   525
         Width           =   1740
      End
      Begin VB.TextBox txtcon 
         Enabled         =   0   'False
         Height          =   300
         Index           =   0
         Left            =   135
         TabIndex        =   13
         Top             =   525
         Width           =   1800
      End
      Begin VB.TextBox txtcon 
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   135
         TabIndex        =   11
         Top             =   1095
         Width           =   2655
      End
      Begin VB.TextBox txtcon 
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   2910
         TabIndex        =   10
         Top             =   1095
         Width           =   2700
      End
      Begin VB.TextBox txtcon 
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   5715
         MaxLength       =   1
         TabIndex        =   9
         Top             =   1095
         Width           =   390
      End
      Begin VB.TextBox txtcon 
         Enabled         =   0   'False
         Height          =   300
         Index           =   5
         Left            =   135
         MaxLength       =   3
         TabIndex        =   8
         Top             =   1740
         Width           =   450
      End
      Begin VB.TextBox txtcon 
         Enabled         =   0   'False
         Height          =   300
         Index           =   6
         Left            =   720
         TabIndex        =   7
         Top             =   1740
         Width           =   5385
      End
      Begin MSMask.MaskEdBox txtdate 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "M/d/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   345
         Left            =   645
         TabIndex        =   12
         Top             =   2370
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   609
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin Project1.chameleonButton cmdsrch 
         Height          =   630
         Left            =   3945
         TabIndex        =   21
         Top             =   210
         Width           =   990
         _extentx        =   1746
         _extenty        =   1111
         btype           =   5
         tx              =   "Searc&h Patient"
         enab            =   -1
         font            =   "PHconDelete.frx":A371
         coltype         =   2
         focusr          =   -1
         bcol            =   12632256
         bcolo           =   12632256
         fcol            =   0
         fcolo           =   255
         mcol            =   4210752
         mptr            =   1
         micon           =   "PHconDelete.frx":A395
         umcol           =   -1
         soft            =   0
         picpos          =   0
         ngrey           =   0
         fx              =   1
         hand            =   0
         check           =   0
         value           =   0
      End
      Begin MSMask.MaskEdBox txttime 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "h:nn AM/PM"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   4
         EndProperty
         Height          =   345
         Left            =   1980
         TabIndex        =   26
         Top             =   2370
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   609
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##:## ?M"
         PromptChar      =   "_"
      End
      Begin Project1.chameleonButton cmdcon 
         Height          =   630
         Left            =   5010
         TabIndex        =   27
         Top             =   210
         Width           =   1110
         _extentx        =   1720
         _extenty        =   1111
         btype           =   5
         tx              =   "Con&fined Patient(s)"
         enab            =   -1
         font            =   "PHconDelete.frx":A3B3
         coltype         =   2
         focusr          =   -1
         bcol            =   12632256
         bcolo           =   12632256
         fcol            =   0
         fcolo           =   255
         mcol            =   4210752
         mptr            =   1
         micon           =   "PHconDelete.frx":A3D7
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
         Height          =   555
         Left            =   4080
         Top             =   2160
         Width           =   2025
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
         Left            =   4200
         TabIndex        =   31
         Top             =   2220
         Width           =   1800
      End
      Begin VB.Label Label1 
         Caption         =   "PH Member ID No."
         Height          =   255
         Index           =   0
         Left            =   1980
         TabIndex        =   23
         Top             =   285
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Last Name:"
         Height          =   255
         Index           =   6
         Left            =   150
         TabIndex        =   20
         Top             =   855
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "First Name:"
         Height          =   255
         Index           =   7
         Left            =   2910
         TabIndex        =   19
         Top             =   855
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "MI:"
         Height          =   255
         Index           =   8
         Left            =   5730
         TabIndex        =   18
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Age:"
         Height          =   255
         Index           =   9
         Left            =   150
         TabIndex        =   17
         Top             =   1485
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Admission Diagnosis:"
         Height          =   255
         Index           =   10
         Left            =   735
         TabIndex        =   16
         Top             =   1500
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Admission Date/Time:"
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
         Height          =   360
         Index           =   11
         Left            =   135
         TabIndex        =   15
         Top             =   2100
         Width           =   2760
      End
      Begin VB.Label Label1 
         Caption         =   "Patient No.:"
         Height          =   255
         Index           =   4
         Left            =   150
         TabIndex        =   14
         Top             =   285
         Width           =   1575
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Confined Patient(s) List"
      ForeColor       =   &H000000FF&
      Height          =   2880
      Left            =   495
      TabIndex        =   4
      Top             =   30
      Width           =   3810
      Begin MSComctlLib.ListView ListView1 
         Height          =   2475
         Left            =   150
         TabIndex        =   5
         Top             =   240
         Width           =   3525
         _ExtentX        =   6218
         _ExtentY        =   4366
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
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Patient No."
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "PH Member ID No."
            Object.Width           =   9
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Admission Date/Time"
            Object.Width           =   9
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Last Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "First Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "MI"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Age"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Admission  Diagnosed"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame Frame5 
      Height          =   915
      Left            =   480
      TabIndex        =   1
      Top             =   2940
      Width           =   5715
      Begin Project1.chameleonButton cmdreset 
         Height          =   570
         Left            =   4650
         TabIndex        =   2
         Top             =   240
         Width           =   930
         _extentx        =   1296
         _extenty        =   873
         btype           =   5
         tx              =   "&Reset"
         enab            =   -1
         font            =   "PHconDelete.frx":A3F5
         coltype         =   2
         focusr          =   -1
         bcol            =   12632256
         bcolo           =   12632256
         fcol            =   0
         fcolo           =   255
         mcol            =   12632256
         mptr            =   1
         micon           =   "PHconDelete.frx":A421
         umcol           =   -1
         soft            =   0
         picpos          =   0
         ngrey           =   0
         fx              =   1
         hand            =   0
         check           =   0
         value           =   0
      End
      Begin Project1.chameleonButton cmdDelete 
         Height          =   585
         Left            =   75
         TabIndex        =   3
         Top             =   225
         Width           =   3975
         _extentx        =   2037
         _extenty        =   1032
         btype           =   8
         tx              =   "D&elete Confined Patient"
         enab            =   -1
         font            =   "PHconDelete.frx":A43F
         coltype         =   2
         focusr          =   -1
         bcol            =   12632256
         bcolo           =   12632256
         fcol            =   0
         fcolo           =   255
         mcol            =   12632256
         mptr            =   1
         micon           =   "PHconDelete.frx":A463
         picn            =   "PHconDelete.frx":A481
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
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4035
      Index           =   0
      Left            =   -15
      TabIndex        =   0
      Top             =   -15
      Width           =   10770
      _ExtentX        =   18997
      _ExtentY        =   7117
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
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000003&
      Height          =   555
      Left            =   0
      Top             =   0
      Width           =   2385
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Note: This Form is exclusive for PH-Med Entries only."
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
      Index           =   2
      Left            =   180
      TabIndex        =   30
      Top             =   60
      Width           =   2085
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000003&
      Height          =   555
      Left            =   0
      Top             =   0
      Width           =   2385
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Note: This Form is exclusive for PH-Med Entries only."
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
      Index           =   0
      Left            =   180
      TabIndex        =   29
      Top             =   60
      Width           =   2085
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000003&
      Height          =   555
      Left            =   0
      Top             =   0
      Width           =   2385
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Note: This Form is exclusive for PH-Med Entries only."
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
      Index           =   1
      Left            =   180
      TabIndex        =   28
      Top             =   60
      Width           =   2085
   End
End
Attribute VB_Name = "PHconDelete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub butx_Click()
Unload Me
MainForm.DelCon.Enabled = True
End Sub

Private Sub chameleonButton1_Click()

End Sub

Private Sub cmdcon_Click()
listcon4.Show vbModal

End Sub

Private Sub cmdDelete_Click()
'the Scanhead rules :: Madz
Dim testdate As String
testdate = ""
If txtcon(0).Text = "" Then
   MsgBox "No Confined Patient to Delete!!", vbCritical, "Sorry"
Else

  If MsgBox("Are you sure you want to delete " + txtcon(3).Text + " ." + txtcon(4).Text + " " + txtcon(2).Text + " .", vbExclamation + vbYesNo, "Confirm Deletion") = vbYes Then
        Dim Dellist As ListItem
        '***** Confined Patients***********'
        DBMain.Execute "DELETE * FROM Patient WHERE PatientNo = '" + txtcon(0).Text + "' and format(dateD,'mm/dd/yyyy') = '" & testdate & "';"
        '**********************************
        Set Dellist = ListView1.FindItem(txtcon(0).Text, , , lvwPartial)
        ListView1.ListItems.Remove Dellist.Index
        ClearText2x
   End If
End If
End Sub

Private Sub cmdExit_Click()
Unload Me
MainForm.DelCon.Enabled = True

End Sub

Private Sub cmdreset_Click()
ClearText2x
DisplayLstx
End Sub

Private Sub cmdsrch_Click()
 Load SFconDEL
 cmdreset.Enabled = True
 SFconDEL.Show vbModal
End Sub

Private Sub Form_Load()
 Dim testdate As String
 add3d
 testdate = ""
 Set WSMain = DBEngine.Workspaces(0)
 Set DBMain = WSMain.OpenDatabase(App.Path + "\hospital.mdb", False, False, ";pwd=scanhead")
          '******** NOID ACCEPTED ************
          sql = "select patientno,idno,datec,plastname,pfirstname,pmi,page,padiagnose,timein from Patient where format(dateD,'mm/dd/yyyy') = '" & testdate & "' order by plastname,pfirstname,pmi;"
           Set rs = DBMain.OpenRecordset(sql)
                Do Until rs.EOF
                    Set y = ListView1.ListItems.Add(, , rs.Fields(0))
                         If IsNull(rs.Fields(1)) = True Then
                          y.SubItems(1) = "Non-Med"
                        Else
                          y.SubItems(1) = rs.Fields(1)
                        End If
                        y.SubItems(2) = Format(rs.Fields(2), "mm/dd/yyyy") + " " + Format(rs.Fields(8), "medium time")
                        y.SubItems(3) = rs.Fields(3)
                        y.SubItems(4) = rs.Fields(4)
                        y.SubItems(5) = rs.Fields(5)
                        y.SubItems(6) = rs.Fields(6)
                        y.SubItems(7) = rs.Fields(7)
                        rs.MoveNext
                 Loop
           cmdDelete.Enabled = True
           cmdreset.Enabled = True
           '*******************************
End Sub



Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo TRAPPER
    Dim X As Long
    Dim testdate As String
    Dim TQR As DAO.QueryDef
    Dim TRS As DAO.Recordset
    'Scanhead rules : NOID IS ACCEPTABLE
    '*********
    testdate = ""
   Set TQR = DBMain.CreateQueryDef("", "SELECT Patientno,idno,plastname,pfirstname,PMI,page,padiagnose,datec,timein FROM Patient WHERE patientno ='" & ListView1.SelectedItem.Text & "' and format(dateD,'mm/dd/yyyy') = '" & testdate & "' ORDER BY plastname,pfirstname,pmi")
   Set TRS = TQR.OpenRecordset()
   PHconDelete.txtcon(0).Text = TRS.Fields(0)
   PHconDelete.txtcon(2).Text = TRS.Fields(2)
   PHconDelete.txtcon(3).Text = TRS.Fields(3)
   PHconDelete.txtcon(4).Text = TRS.Fields(4)
   PHconDelete.txtcon(5).Text = TRS.Fields(5)
   PHconDelete.txtcon(6).Text = TRS.Fields(6)
   If IsNull(TRS.Fields(1)) = True Then
      PHconDelete.txtconM.Text = "Non-Med"
   Else
   PHconDelete.txtconM.Text = TRS.Fields(1)
   End If
   PHconDelete.txtdate.Text = Format(TRS.Fields(7), "mm/dd/yyyy")
   PHconDelete.txttime.Text = Format(TRS.Fields(8), "medium time")
TRAPPER:
End Sub


Public Sub DisplayLstx() 'Reset control display: gotcha the scanhedbri
On Error GoTo X
    
    Dim TQR As DAO.QueryDef
    Dim rs As DAO.Recordset
    Dim y As ListItem
     Dim sql As String
    Dim X As Long
    Dim testdate As String
    
    testdate = ""
   
    '/*************** Confined patients  NOID ACCEPTED***************************
  sql = "select patientno,idno,datec,plastname,pfirstname,pmi,page,padiagnose,timein from Patient where format(dateD,'mm/dd/yyyy') = '" & testdate & "' order by plastname,pfirstname,pmi;"
           Set rs = DBMain.OpenRecordset(sql)
           ListView1.ListItems.Clear
                Do Until rs.EOF
                    Set y = ListView1.ListItems.Add(, , rs.Fields(0))
                        If IsNull(rs.Fields(1)) = True Then
                          y.SubItems(1) = "Non-Med"
                        Else
                          y.SubItems(1) = rs.Fields(1)
                        End If
                        y.SubItems(2) = Format(rs.Fields(2), "mm/dd/yyyy") + " " + Format(rs.Fields(8), "medium time")
                        y.SubItems(3) = rs.Fields(3)
                        y.SubItems(4) = rs.Fields(4)
                        y.SubItems(5) = rs.Fields(5)
                        y.SubItems(6) = rs.Fields(6)
                        y.SubItems(7) = rs.Fields(7)
                        rs.MoveNext
                 Loop
 
X:
End Sub

Public Function add3d()
Add3DBorder ListView1
Add3DBorder txtconM
Add3DBorder txtcon(0)
Add3DBorder txtcon(2)
Add3DBorder txtcon(3)
Add3DBorder txtcon(4)
Add3DBorder txtcon(5)
Add3DBorder txtcon(6)
Add3DBorder txtdate
Add3DBorder txttime
End Function
