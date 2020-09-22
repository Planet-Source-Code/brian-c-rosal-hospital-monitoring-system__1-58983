VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PHconNONMed 
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
   Picture         =   "PHconNONMED.frx":0000
   ScaleHeight     =   5070
   ScaleWidth      =   11475
   WindowState     =   2  'Maximized
   Begin Project1.chameleonButton cmdcancel 
      Height          =   825
      Left            =   7830
      TabIndex        =   28
      Top             =   3000
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   1455
      BTYPE           =   5
      TX              =   "&Cancel"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
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
      MICON           =   "PHconNONMED.frx":A2E5
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.chameleonButton cmdexit 
      Height          =   825
      Left            =   9345
      TabIndex        =   12
      Top             =   3000
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   1455
      BTYPE           =   5
      TX              =   "E&xit"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
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
      MICON           =   "PHconNONMED.frx":A301
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.chameleonButton butx 
      Height          =   285
      Left            =   105
      TabIndex        =   15
      Top             =   90
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   503
      BTYPE           =   5
      TX              =   "X"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   12582912
      FCOL            =   0
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "PHconNONMED.frx":A31D
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame3 
      Caption         =   "Patient's Confinement Data"
      ForeColor       =   &H00C00000&
      Height          =   2880
      Left            =   4395
      TabIndex        =   19
      Top             =   30
      Width           =   6375
      Begin VB.TextBox txtcon 
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   135
         TabIndex        =   2
         Top             =   1095
         Width           =   2655
      End
      Begin VB.TextBox txtcon 
         Enabled         =   0   'False
         Height          =   300
         Index           =   0
         Left            =   135
         TabIndex        =   1
         Top             =   525
         Width           =   1800
      End
      Begin VB.TextBox txtcon 
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   2910
         TabIndex        =   3
         Top             =   1095
         Width           =   2805
      End
      Begin VB.TextBox txtcon 
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   5805
         MaxLength       =   1
         TabIndex        =   4
         Top             =   1095
         Width           =   390
      End
      Begin VB.TextBox txtcon 
         Enabled         =   0   'False
         Height          =   270
         Index           =   5
         Left            =   135
         MaxLength       =   3
         TabIndex        =   5
         Top             =   1740
         Width           =   450
      End
      Begin VB.TextBox txtcon 
         Enabled         =   0   'False
         Height          =   300
         Index           =   6
         Left            =   810
         TabIndex        =   6
         Top             =   1740
         Width           =   5370
      End
      Begin Project1.chameleonButton cmdsrch 
         Height          =   585
         Left            =   3915
         TabIndex        =   13
         Top             =   255
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   1032
         BTYPE           =   5
         TX              =   "Searc&h Patient"
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
         MCOL            =   4210752
         MPTR            =   1
         MICON           =   "PHconNONMED.frx":A339
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   1
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
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
         Height          =   375
         Left            =   345
         TabIndex        =   7
         Top             =   2355
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   661
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
         Height          =   375
         Left            =   1830
         TabIndex        =   8
         Top             =   2355
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   661
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
      Begin Project1.chameleonButton cmddiag 
         Height          =   585
         Left            =   5205
         TabIndex        =   29
         Top             =   255
         Width           =   1020
         _extentx        =   1905
         _extenty        =   529
         btype           =   5
         tx              =   "&Confined Patient(s)"
         enab            =   -1
         font            =   "PHconNONMED.frx":A355
         coltype         =   2
         focusr          =   -1
         bcol            =   12632256
         bcolo           =   12632256
         fcol            =   0
         fcolo           =   255
         mcol            =   12632256
         mptr            =   1
         micon           =   "PHconNONMED.frx":A379
         umcol           =   -1
         soft            =   0
         picpos          =   0
         ngrey           =   0
         fx              =   1
         hand            =   0
         check           =   0
         value           =   0
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Note: This Form is exclusived for Non-Med Entries only."
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
         Left            =   3855
         TabIndex        =   30
         Top             =   2250
         Width           =   2160
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H80000003&
         Height          =   585
         Left            =   3330
         Top             =   2160
         Width           =   2880
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
         Height          =   240
         Index           =   11
         Left            =   135
         TabIndex        =   27
         Top             =   2100
         Width           =   2130
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Non-Med"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   0
         Left            =   2400
         TabIndex        =   26
         Top             =   525
         Width           =   1140
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000005&
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   315
         Left            =   2055
         Top             =   510
         Width           =   1755
      End
      Begin VB.Label Label1 
         Caption         =   "Last Name:"
         Height          =   255
         Index           =   6
         Left            =   150
         TabIndex        =   25
         Top             =   855
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "First Name:"
         Height          =   255
         Index           =   7
         Left            =   2910
         TabIndex        =   24
         Top             =   855
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "MI:"
         Height          =   255
         Index           =   8
         Left            =   5790
         TabIndex        =   23
         Top             =   870
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Age:"
         Height          =   255
         Index           =   9
         Left            =   150
         TabIndex        =   22
         Top             =   1485
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Admission Diagnosis:"
         Height          =   255
         Index           =   10
         Left            =   810
         TabIndex        =   21
         Top             =   1500
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Patient No.:"
         Height          =   255
         Index           =   4
         Left            =   150
         TabIndex        =   20
         Top             =   285
         Width           =   1575
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Confined Patient(s) List"
      ForeColor       =   &H000000FF&
      Height          =   2880
      Left            =   495
      TabIndex        =   18
      Top             =   30
      Width           =   3810
      Begin MSComctlLib.ListView ListView1 
         Height          =   2475
         Left            =   150
         TabIndex        =   14
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
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Patient No."
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Admission Date/Time"
            Object.Width           =   9
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
            Text            =   "Age"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Admission  Diagnosis"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame Frame5 
      Height          =   915
      Left            =   480
      TabIndex        =   17
      Top             =   2940
      Width           =   5715
      Begin Project1.chameleonButton cmdreset 
         Height          =   600
         Left            =   4680
         TabIndex        =   10
         Top             =   210
         Width           =   915
         _extentx        =   1296
         _extenty        =   873
         btype           =   5
         tx              =   "&Reset"
         enab            =   -1
         font            =   "PHconNONMED.frx":A397
         coltype         =   2
         focusr          =   -1
         bcol            =   12632256
         bcolo           =   12632256
         fcol            =   0
         fcolo           =   255
         mcol            =   12632256
         mptr            =   1
         micon           =   "PHconNONMED.frx":A3C3
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
         Left            =   2880
         TabIndex        =   9
         Top             =   240
         Width           =   1260
         _extentx        =   2223
         _extenty        =   979
         btype           =   8
         tx              =   "&Save"
         enab            =   -1
         font            =   "PHconNONMED.frx":A3E1
         coltype         =   2
         focusr          =   -1
         bcol            =   12632256
         bcolo           =   12632256
         fcol            =   0
         fcolo           =   255
         mcol            =   12632256
         mptr            =   1
         micon           =   "PHconNONMED.frx":A40D
         picn            =   "PHconNONMED.frx":A42B
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
         Left            =   1485
         TabIndex        =   11
         Top             =   210
         Width           =   1260
         _extentx        =   2223
         _extenty        =   979
         btype           =   8
         tx              =   "&Modify"
         enab            =   -1
         font            =   "PHconNONMED.frx":A87F
         coltype         =   2
         focusr          =   -1
         bcol            =   12632256
         bcolo           =   12632256
         fcol            =   0
         fcolo           =   255
         mcol            =   12632256
         mptr            =   1
         micon           =   "PHconNONMED.frx":A8AB
         picn            =   "PHconNONMED.frx":A8C9
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
         Left            =   120
         TabIndex        =   0
         Top             =   225
         Width           =   1260
         _extentx        =   2223
         _extenty        =   979
         btype           =   8
         tx              =   "&New Confine"
         enab            =   -1
         font            =   "PHconNONMED.frx":AD1D
         coltype         =   2
         focusr          =   -1
         bcol            =   12632256
         bcolo           =   12632256
         fcol            =   0
         fcolo           =   255
         mcol            =   12632256
         mptr            =   1
         micon           =   "PHconNONMED.frx":AD49
         picn            =   "PHconNONMED.frx":AD67
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
      TabIndex        =   16
      Top             =   -15
      Width           =   10920
      _ExtentX        =   19262
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
End
Attribute VB_Name = "PHconNONMed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sql, rs, y As Variant, ctr As Byte
Dim MyType, passString As String
Dim formod As String

Private Sub butx_Click()
Unload Me
MainForm.Con2.Enabled = True
MainForm.Toolbar1.Buttons.Item(3).Enabled = True
End Sub



Private Sub cmdAdd_Click()
SetText_non True
ClearText2Non
txtcon(0).Text = AutoRecordNumber2
txtdate.Text = Format(now, "mm/dd/yyyy")
txttime.Text = Format(now, "medium time")
txtcon(2).SetFocus
ListView1.Enabled = False
cmdcancel.Enabled = True
cmdmod.Enabled = False
cmdsave.Enabled = True
cmdadd.Enabled = False
MyType = "ADD"
End Sub

Private Sub cmdCancel_Click()
 ListView1.Enabled = True
 SetText_non False
 ClearText2Non
 cmdsrch.Enabled = True
 cmdadd.Enabled = True
 cmdsave.Enabled = False
 cmdmod.Enabled = False
 cmdcancel.Enabled = False
 cmdreset.Enabled = True
End Sub

Private Sub cmddiag_Click()
listcon5.Show vbModal
End Sub

Private Sub cmdExit_Click()
Unload Me
MainForm.Con2.Enabled = True
MainForm.Toolbar1.Buttons.Item(3).Enabled = True
End Sub

Private Sub cmdmod_Click()
    If txtcon(0).Text = "" Then
        MsgBox "There's no Record to Modify ", vbExclamation, "Confirmation"
        Exit Sub
    End If
    SetText_non True
    ListView1.Enabled = False
    cmdadd.Enabled = False
    cmdmod.Enabled = False
    cmdsave.Enabled = True
    cmdreset.Enabled = False
    cmdcancel.Enabled = True
   
    txtcon(2).SetFocus
     MyType = "EDIT"
End Sub

Private Sub cmdreset_Click()
ClearText2Non
DisplayLstx
End Sub

Private Sub CMDSAVE_Click()
On Error GoTo TRAPPER
    Dim TRS As DAO.Recordset
    Dim TQR As DAO.QueryDef
    Dim P, madzsrch, testdate, NOID As String
    Dim Query, tb, td As String
    Dim List As ListItem
    Dim X As Long
    Dim Flag As Boolean
    Flag = False
    NOID = "Non-Med"
    For X = 2 To 6
        If txtcon(X).Text = "" Then Flag = True
    Next X
    If txtdate.Text = "__/__/____" Then
       Flag = True
    End If
     
    If Flag Then
        MsgBox "Please Enter all information to Continue ?", vbInformation, "Confirmation"
        GoTo TRAPPER
     End If
   
 If MyType = "ADD" Then
   
 '****************************
   sql = "select patientno,idno,datec,plastname,pfirstname,pmi,page,padiagnose from Patient Where plastname = '" & txtcon(2).Text & "' and pfirstname ='" & txtcon(3).Text & "' and pmi = '" & txtcon(4).Text & "' and ISNULL(dateD) = TRUE order by plastname,pfirstname,pmi;"
           Set rs = DBMain.OpenRecordset(sql)
            If rs.RecordCount <> 0 Or rs.RecordCount > 0 Then
             MsgBox "The Patient exist in the admission list! [ Note: If not as PH-Med as Non-Med ] Check the Confined Patients List for confirmation.", vbCritical, "Warning"
               GoTo TRAPPER
         End If
   
  
 '****************************
  P = "INSERT INTO Patient (patientno,plastname,pfirstname,PMI,Page,PAdiagnose,[datec],TIMEIN,NOID) VALUES ('" & txtcon(0).Text & "','" & txtcon(2).Text & "','" & txtcon(3).Text & "','" & txtcon(4).Text & "','" & txtcon(5).Text & "','" & txtcon(6).Text & "','" & txtdate.Text & "','" & txttime.Text & "','" & NOID & "');"
   Set TQR = DBMain.CreateQueryDef("", P)
   TQR.Execute
   Set List = ListView1.ListItems.Add(, , txtcon(0).Text)
        With List
            .SubItems(1) = Format(txtdate.Text, "mm/dd/yyyy") + " " + Format(txttime.Text, "medium time")
            .SubItems(2) = txtcon(2).Text
            .SubItems(3) = txtcon(3).Text
            .SubItems(4) = txtcon(4).Text
            .SubItems(5) = txtcon(5).Text
            .SubItems(6) = txtcon(6).Text
        End With
   
   ElseIf MyType = "EDIT" Then
   
             passString = "UPDATE Patient SET plastname='" & txtcon(2).Text & "',pfirstname='" & txtcon(3).Text & "',PMI='" & txtcon(4).Text & "',Page='" & txtcon(5).Text & "',PAdiagnose='" & txtcon(6).Text & "',[DateC]='" & txtdate.Text & "',TIMEIN = '" & txttime.Text & "',NOID = '" & NOID & "' WHERE patientno='" & txtcon(0).Text & "' ;"
            DBMain.Execute passString
            ListView1.Enabled = True
            MyType = ""
            Set List = ListView1.FindItem(txtcon(0).Text, , , lvwPartial)
              With List
                  .SubItems(1) = Format(txtdate.Text, "mm/dd/yyyy") + " " + Format(txttime.Text, "medium time")
                  .SubItems(2) = txtcon(2).Text
                  .SubItems(3) = txtcon(3).Text
                  .SubItems(4) = txtcon(4).Text
                  .SubItems(5) = txtcon(5).Text
                  .SubItems(6) = txtcon(6).Text
              End With
    
   
   End If
   
   '************************************************
    cmdsave.Enabled = False
    cmdadd.Enabled = True
    cmdmod.Enabled = True
    ListView1.Enabled = True
    SetText_non False
   
TRAPPER:
   Exit Sub
End Sub

Private Sub cmdsrch_Click()
 Load SFconNON
 cmdreset.Enabled = True
 SFconNON.Show vbModal
End Sub

Private Sub Form_Load()
 Dim testdate, NOID As String
 add3d
 cmdmod.Enabled = True
 cmdsave.Enabled = False
 testdate = ""
 NOID = ""
 
 Set WSMain = DBEngine.Workspaces(0)
 Set DBMain = WSMain.OpenDatabase(App.Path + "\hospital.mdb", False, False, ";pwd=scanhead")
          sql = "select patientno,datec,TIMEIN,plastname,pfirstname,pmi,page,padiagnose from Patient where  ISNULL(dateD) = TRUE AND isnull(IDNO) = true  order by plastname,pfirstname,pmi;"
           Set rs = DBMain.OpenRecordset(sql)
                Do Until rs.EOF
                    Set y = ListView1.ListItems.Add(, , rs.Fields(0))
                        y.SubItems(1) = Format(rs.Fields(1), "mm/dd/yyyy") + " " + Format(rs.Fields(2), "medium time")
                        y.SubItems(2) = rs.Fields(3)
                        y.SubItems(3) = rs.Fields(4)
                        y.SubItems(4) = rs.Fields(5)
                        y.SubItems(5) = rs.Fields(6)
                        y.SubItems(6) = rs.Fields(7)
                        rs.MoveNext
                 Loop
          
           cmdreset.Enabled = False
           '*******************************
End Sub



Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo TRAPPER
    Dim X As Long
    Dim testdate, NOID As String
    Dim TQR As DAO.QueryDef
    Dim TRS As DAO.Recordset
    'Scanhead rules
    '*********
    cmdmod.Enabled = True
    testdate = ""
    NOID = ""
   Set TQR = DBMain.CreateQueryDef("", "SELECT Patientno,idno,plastname,pfirstname,PMI,page,padiagnose,datec,TIMEIN FROM Patient WHERE patientno ='" & ListView1.SelectedItem.Text & "' and ISNULL(dateD) = TRUE AND ISNULL(IDNO) = TRUE ORDER BY plastname,pfirstname,pmi")
   Set TRS = TQR.OpenRecordset()
   txtcon(0).Text = TRS.Fields(0)
   txtcon(2).Text = TRS.Fields(2)
   txtcon(3).Text = TRS.Fields(3)
   txtcon(4).Text = TRS.Fields(4)
   txtcon(5).Text = TRS.Fields(5)
   txtcon(6).Text = TRS.Fields(6)
   txtdate.Text = Format(TRS.Fields(7), "mm/dd/yyyy")
   txttime.Text = Format(TRS.Fields(8), "MEDIUM TIME")
 
TRAPPER:
End Sub


Public Sub DisplayLstx() 'Reset control display: gotcha the scanhedbri
On Error GoTo X
    
    Dim TQR As DAO.QueryDef
    Dim rs As DAO.Recordset
    Dim y As ListItem
     Dim sql As String
    Dim X As Long
    Dim testdate, NOID As String
    
    testdate = ""
    NOID = ""
   
    '/*************** Confined patients ***************************
      sql = "select patientno,datec,TIMEIN,plastname,pfirstname,pmi,page,padiagnose from Patient where  ISNULL(dateD) = TRUE AND isnull(IDNO) = true  order by plastname,pfirstname,pmi;"
                Set rs = DBMain.OpenRecordset(sql)
           ListView1.ListItems.Clear
                Do Until rs.EOF
                    Set y = ListView1.ListItems.Add(, , rs.Fields(0))
                        y.SubItems(1) = Format(rs.Fields(1), "mm/dd/yyyy") + " " + Format(rs.Fields(2), "medium time")
                        y.SubItems(2) = rs.Fields(3)
                        y.SubItems(3) = rs.Fields(4)
                        y.SubItems(4) = rs.Fields(5)
                        y.SubItems(5) = rs.Fields(6)
                        y.SubItems(6) = rs.Fields(7)
          rs.MoveNext
                 Loop
 
X:
End Sub

Public Function add3d()
Add3DBorder ListView1
Add3DBorder txtcon(0)
Add3DBorder txtcon(2)
Add3DBorder txtcon(3)
Add3DBorder txtcon(4)
Add3DBorder txtcon(5)
Add3DBorder txtcon(6)
Add3DBorder txtdate
Add3DBorder txttime
End Function
Public Function AutoRecordNumber2() As String
    Dim TQR As DAO.QueryDef
    Dim TRS As DAO.Recordset
    Dim Start As String
    Start = "00000000"
    Set TQR = DBMain.CreateQueryDef("", "SELECT PatientNo FROM Patient")
    Set TRS = TQR.OpenRecordset()
    Do While Not TRS.EOF
        TRS.FindFirst "PatientNo ='P-" + Start + "'"
        If Not TRS.NoMatch Then
            Start = Format(Str(Val(Mid$(Start, 3)) + 1), "00000000")
        Else
            AutoRecordNumber2 = "P-" + Start
            Exit Function
        End If
    Loop
    AutoRecordNumber2 = "P-" + Start
End Function
Private Sub txtcon_KeyPress(Index As Integer, KeyAscii As Integer)
'the madzbry txt validation
Dim madzbry As String
 Select Case Index
        Case 5
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

Private Sub txtdate_GotFocus()
With txtdate
.SelStart = 0
.SelLength = Len(txtdate.Text)
End With
End Sub

Private Sub txtTIME_GotFocus()
With txttime
.SelStart = 0
.SelLength = Len(txttime.Text)
End With
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
