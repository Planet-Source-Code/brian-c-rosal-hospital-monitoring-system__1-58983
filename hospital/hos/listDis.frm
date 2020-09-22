VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form listDis 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Discharged Patient(s) List [ PHMed/Non-Med]"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11370
   Icon            =   "listDis.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   11370
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   1665
      Left            =   9480
      TabIndex        =   14
      Top             =   2730
      Width           =   1725
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   135
         TabIndex        =   15
         Top             =   960
         Width           =   1485
      End
   End
   Begin VB.Frame Frame1 
      Height          =   930
      Left            =   5340
      TabIndex        =   8
      Top             =   3465
      Width           =   3975
      Begin VB.CommandButton cmdALL 
         Caption         =   "&ALL Discharged"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   2160
         TabIndex        =   13
         Top             =   240
         Width           =   1620
      End
      Begin VB.CheckBox CMED 
         Caption         =   "Check1"
         Height          =   375
         Left            =   105
         TabIndex        =   10
         Top             =   480
         Width           =   255
      End
      Begin VB.CheckBox CNonMed 
         Caption         =   "Check1"
         Height          =   375
         Left            =   105
         TabIndex        =   9
         Top             =   165
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "PH MED"
         Height          =   255
         Index           =   1
         Left            =   465
         TabIndex        =   12
         Top             =   555
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "NON-MED"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   11
         Top             =   255
         Width           =   855
      End
   End
   Begin VB.Frame framecon 
      Caption         =   "Discharged Patient(s) "
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
      Height          =   2685
      Left            =   120
      TabIndex        =   4
      Top             =   15
      Width           =   11085
      Begin MSComctlLib.ListView ListView1 
         Height          =   2295
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   10845
         _ExtentX        =   19129
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
         NumItems        =   11
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
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Age"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Final Diagnosis"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Patient No."
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "PH Member ID No."
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Admission Date/Time"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Date/Time Discharged"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "PH/Amount Paid"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Balance"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.ComboBox combomonth 
      Height          =   315
      ItemData        =   "listDis.frx":0BC2
      Left            =   120
      List            =   "listDis.frx":0BED
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   2910
      Width           =   1815
   End
   Begin VB.ComboBox combodate 
      Height          =   315
      ItemData        =   "listDis.frx":0C59
      Left            =   2040
      List            =   "listDis.frx":0CBD
      TabIndex        =   2
      Text            =   "Combo2"
      Top             =   2910
      Width           =   735
   End
   Begin VB.CommandButton cmdRep 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   5655
      TabIndex        =   1
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CommandButton cmdPerform 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Caption         =   "Pe&rform"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3990
      TabIndex        =   0
      Top             =   2895
      Width           =   1575
   End
   Begin MSMask.MaskEdBox txtyear 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "M/d/yy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Height          =   315
      Left            =   2895
      TabIndex        =   6
      Top             =   2910
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "####"
      PromptChar      =   "_"
   End
   Begin VB.Label Label2 
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
      Height          =   210
      Left            =   390
      TabIndex        =   7
      Top             =   3615
      Width           =   4215
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000009&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000009&
      FillColor       =   &H00808000&
      FillStyle       =   0  'Solid
      Height          =   420
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   3510
      Width           =   5025
   End
End
Attribute VB_Name = "listDis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sql, rs, y As Variant, ctr As Byte
Dim MyType, passString, madzsrch As String
Dim formod, NOID As String


Private Sub cmdall_Click()
combomonth.Text = "None"
combodate.Text = "None"
txtyear.Text = "____"
CNonMed.Value = 1
CMED.Value = 1
cmdPerform.SetFocus
SendKeys "{Enter}"
End Sub

Private Sub cmdPerform_Click()
NOID = "Non-Med"
'***
If combomonth.Text <> "None" And combodate.Text = "None" And txtyear.Text = "____" And CMED.Value = 1 And CNonMed.Value = 0 Then
sql = "select plast,pfirst,pmi,page,Fdiagnose,patientno,idno,pdatec,dateD,TIMEOUT,PHILPAY,DIFF,TIMEIN from PDISCHARGED where FORMAT(dateD,'MMMM') = '" & combomonth.Text & "'  AND TRIM(IDNO) <> '" & NOID & "' order by DATED,TIMEOUT,plast,pfirst,pmi;"
ElseIf combomonth.Text <> "None" And combodate.Text = "None" And txtyear.Text = "____" And CMED.Value = 0 And CNonMed.Value = 1 Then
sql = "select plast,pfirst,pmi,page,Fdiagnose,patientno,idno,pdatec,dateD,TIMEOUT,PHILPAY,DIFF,TIMEIN from PDISCHARGED where FORMAT(dATED,'MMMM') = '" & combomonth.Text & "'  AND TRIM(IDNO) = '" & NOID & "' order by DATED,TIMEOUT ,plast,pfirst,pmi;"
ElseIf combomonth.Text <> "None" And combodate.Text = "None" And txtyear.Text = "____" And CMED.Value = 0 And CNonMed.Value = 0 Then
sql = "select plast,pfirst,pmi,page,Fdiagnose,patientno,idno,pdatec,dateD,TIMEOUT,PHILPAY,DIFF,TIMEIN from PDISCHARGED where FORMAT(dateD,'MMMM') = '" & combomonth.Text & "' order by DATED,TIMEOUT,plast,pfirst,pmi;"
ElseIf combomonth.Text <> "None" And combodate.Text = "None" And txtyear.Text = "____" And CMED.Value = 1 And CNonMed.Value = 1 Then
sql = "select plast,pfirst,pmi,page,Fdiagnose,patientno,idno,pdatec,dateD,TIMEOUT,PHILPAY,DIFF,TIMEIN from PDISCHARGED where FORMAT(dateD,'MMMM') = '" & combomonth.Text & "'  order by DATED,TIMEOUT ,plast,pfirst,pmi;"

'***
ElseIf combomonth.Text = "None" And combodate.Text <> "None" And txtyear.Text = "____" And CMED.Value = 1 And CNonMed.Value = 0 Then
sql = "select plast,pfirst,pmi,page,Fdiagnose,patientno,idno,pdatec,dateD,TIMEOUT,PHILPAY,DIFF,TIMEIN from PDISCHARGED where DAY(dateD) = '" & combodate.Text & "' AND TRIM(IDNO) <> '" & NOID & "' order by DATED ,TIMEOUT ,plast,pfirst,pmi;"
ElseIf combomonth.Text = "None" And combodate.Text <> "None" And txtyear.Text = "____" And CMED.Value = 0 And CNonMed.Value = 1 Then
sql = "select plast,pfirst,pmi,page,Fdiagnose,patientno,idno,pdatec,dateD,TIMEOUT,PHILPAY,DIFF,TIMEIN from PDISCHARGED where DAY(dateD) = '" & combodate.Text & "' AND TRIM(IDNO) = '" & NOID & "' order by DATED ,TIMEOUT ,plast,pfirst,pmi;"
ElseIf combomonth.Text = "None" And combodate.Text <> "None" And txtyear.Text = "____" And CMED.Value = 0 And CNonMed.Value = 0 Then
sql = "select plast,pfirst,pmi,page,Fdiagnose,patientno,idno,pdatec,dateD,TIMEOUT,PHILPAY,DIFF,TIMEIN from PDISCHARGED where DAY(dateD) = '" & combodate.Text & "' order by DATED ,TIMEOUT ,plast,pfirst,pmi;"
ElseIf combomonth.Text = "None" And combodate.Text <> "None" And txtyear.Text = "____" And CMED.Value = 1 And CNonMed.Value = 1 Then
sql = "select plast,pfirst,pmi,page,Fdiagnose,patientno,idno,pdatec,dateD,TIMEOUT,PHILPAY,DIFF,TIMEIN from PDISCHARGED where DAY(dateD) = '" & combodate.Text & "' order by DATED ,TIMEOUT ,plast,pfirst,pmi;"

'***
ElseIf combomonth.Text = "None" And combodate.Text = "None" And txtyear.Text <> "____" And CMED.Value = 1 And CNonMed.Value = 0 Then
sql = "select plast,pfirst,pmi,page,Fdiagnose,patientno,idno,pdatec,dateD,TIMEOUT,PHILPAY,DIFF,TIMEIN from PDISCHARGED where YEAR(dateD) = '" & txtyear.Text & "' AND TRIM(IDNO) <> '" & NOID & "' order by DATED ,TIMEOUT,plast,pfirst,pmi;"
ElseIf combomonth.Text = "None" And combodate.Text = "None" And txtyear.Text <> "____" And CMED.Value = 0 And CNonMed.Value = 1 Then
sql = "select plast,pfirst,pmi,page,Fdiagnose,patientno,idno,pdatec,dateD,TIMEOUT,PHILPAY,DIFF,TIMEIN from PDISCHARGED where YEAR(dateD) = '" & txtyear.Text & "' AND TRIM(IDNO) = '" & NOID & "' order by DATED ,TIMEOUT,plast,pfirst,pmi;"
ElseIf combomonth.Text = "None" And combodate.Text = "None" And txtyear.Text <> "____" And CMED.Value = 0 And CNonMed.Value = 0 Then
sql = "select plast,pfirst,pmi,page,Fdiagnose,patientno,idno,pdatec,dateD,TIMEOUT,PHILPAY,DIFF,TIMEIN from PDISCHARGED where YEAR(dateD) = '" & txtyear.Text & "' order by DATED ,TIMEOUT ,plast,pfirst,pmi;"
ElseIf combomonth.Text = "None" And combodate.Text = "None" And txtyear.Text <> "____" And CMED.Value = 1 And CNonMed.Value = 1 Then
sql = "select plast,pfirst,pmi,page,Fdiagnose,patientno,idno,pdatec,dateD,TIMEOUT,PHILPAY,DIFF,TIMEIN from PDISCHARGED where YEAR(dateD) = '" & txtyear.Text & "' order by DATED ,TIMEOUT,plast,pfirst,pmi;"

'***
ElseIf combomonth.Text <> "None" And combodate.Text <> "None" And txtyear.Text <> "____" And CMED.Value = 1 And CNonMed.Value = 0 Then
sql = "select plast,pfirst,pmi,page,Fdiagnose,patientno,idno,pdatec,dateD,TIMEOUT,PHILPAY,DIFF,TIMEIN from PDISCHARGED where FORMAT(dateD,'MMMM') = '" & combomonth.Text & "' AND DAY(dateD) = '" & combodate.Text & "' AND  YEAR(dateD) = '" & txtyear.Text & "' AND TRIM(IDNO) <> '" & NOID & "' order by DATED ,TIMEOUT,plast,pfirst,pmi;"
ElseIf combomonth.Text <> "None" And combodate.Text <> "None" And txtyear.Text <> "____" And CMED.Value = 0 And CNonMed.Value = 1 Then
sql = "select plast,pfirst,pmi,page,Fdiagnose,patientno,idno,pdatec,dateD,TIMEOUT,PHILPAY,DIFF,TIMEIN from PDISCHARGED where FORMAT(dateD,'MMMM') = '" & combomonth.Text & "' AND DAY(dateD) = '" & combodate.Text & "' AND  YEAR(dateD) = '" & txtyear.Text & "' AND TRIM(IDNO) = '" & NOID & "' order by DATED ,TIMEOUT ,plast,pfirst,pmi;"
ElseIf combomonth.Text <> "None" And combodate.Text <> "None" And txtyear.Text <> "____" And CMED.Value = 0 And CNonMed.Value = 0 Then
sql = "select plast,pfirst,pmi,page,Fdiagnose,patientno,idno,pdatec,dateD,TIMEOUT,PHILPAY,DIFF,TIMEIN from PDISCHARGED where FORMAT(dateD,'MMMM') = '" & combomonth.Text & "' AND DAY(dateD) = '" & combodate.Text & "' AND  YEAR(dateD) = '" & txtyear.Text & "' order by DATED ,TIMEOUT,plast,pfirst,pmi;"
ElseIf combomonth.Text <> "None" And combodate.Text <> "None" And txtyear.Text <> "____" And CMED.Value = 1 And CNonMed.Value = 1 Then
sql = "select plast,pfirst,pmi,page,Fdiagnose,patientno,idno,pdatec,dateD,TIMEOUT,PHILPAY,DIFF,TIMEIN from PDISCHARGED where FORMAT(dateD,'MMMM') = '" & combomonth.Text & "' AND DAY(dateD) = '" & combodate.Text & "' AND  YEAR(dateD) = '" & txtyear.Text & "' order by DATED ,TIMEOUT ,plast,pfirst,pmi;"

'***
ElseIf combomonth.Text <> "None" And combodate.Text <> "None" And txtyear.Text = "____" And CMED.Value = 1 And CNonMed.Value = 0 Then
sql = "select plast,pfirst,pmi,page,Fdiagnose,patientno,idno,pdatec,dateD,TIMEOUT,PHILPAY,DIFF,TIMEIN from PDISCHARGED where FORMAT(dateD,'MMMM') = '" & combomonth.Text & "' AND DAY(dateD) = '" & combodate.Text & "' AND TRIM(IDNO) <> '" & NOID & "' order by DATED ,TIMEOUT ,plast,pfirst,pmi;"
ElseIf combomonth.Text <> "None" And combodate.Text <> "None" And txtyear.Text = "____" And CMED.Value = 0 And CNonMed.Value = 1 Then
sql = "select plast,pfirst,pmi,page,Fdiagnose,patientno,idno,pdatec,dateD,TIMEOUT,PHILPAY,DIFF,TIMEIN from PDISCHARGED where FORMAT(dateD,'MMMM') = '" & combomonth.Text & "' AND DAY(dateD) = '" & combodate.Text & "' AND TRIM(IDNO) = '" & NOID & "' order by DATED ,TIMEOUT ,plast,pfirst,pmi;"
ElseIf combomonth.Text <> "None" And combodate.Text <> "None" And txtyear.Text = "____" And CMED.Value = 0 And CNonMed.Value = 0 Then
sql = "select plast,pfirst,pmi,page,Fdiagnose,patientno,idno,pdatec,dateD,TIMEOUT,PHILPAY,DIFF,TIMEIN from PDISCHARGED where FORMAT(dateD,'MMMM') = '" & combomonth.Text & "' AND DAY(dateD) = '" & combodate.Text & "' order by DATED ,TIMEOUT ,plast,pfirst,pmi;"
ElseIf combomonth.Text <> "None" And combodate.Text <> "None" And txtyear.Text = "____" And CMED.Value = 1 And CNonMed.Value = 1 Then
sql = "select plast,pfirst,pmi,page,Fdiagnose,patientno,idno,pdatec,dateD,TIMEOUT,PHILPAY,DIFF,TIMEIN from PDISCHARGED where FORMAT(dateD,'MMMM') = '" & combomonth.Text & "' AND DAY(dateD) = '" & combodate.Text & "' order by DATED ,TIMEOUT ,plast,pfirst,pmi;"


'***
ElseIf combomonth.Text <> "None" And combodate.Text = "None" And txtyear.Text <> "____" And CMED.Value = 1 And CNonMed.Value = 0 Then
sql = "select plast,pfirst,pmi,page,Fdiagnose,patientno,idno,pdatec,dateD,TIMEOUT,PHILPAY,DIFF,TIMEIN from PDISCHARGED where FORMAT(dateD,'MMMM') = '" & combomonth.Text & "' AND   YEAR(dateD) = '" & txtyear.Text & "' AND TRIM(IDNO) <> '" & NOID & "' order by DATED ,TIMEOUT ,plast,pfirst,pmi;"
ElseIf combomonth.Text <> "None" And combodate.Text = "None" And txtyear.Text <> "____" And CMED.Value = 0 And CNonMed.Value = 1 Then
sql = "select plast,pfirst,pmi,page,Fdiagnose,patientno,idno,pdatec,dateD,TIMEOUT,PHILPAY,DIFF,TIMEIN from PDISCHARGED where FORMAT(dateD,'MMMM') = '" & combomonth.Text & "' AND   YEAR(dateD) = '" & txtyear.Text & "' AND TRIM(IDNO) = '" & NOID & "' order by DATED ,TIMEOUT ,plast,pfirst,pmi;"
ElseIf combomonth.Text <> "None" And combodate.Text = "None" And txtyear.Text <> "____" And CMED.Value = 0 And CNonMed.Value = 0 Then
sql = "select plast,pfirst,pmi,page,Fdiagnose,patientno,idno,pdatec,dateD,TIMEOUT,PHILPAY,DIFF,TIMEIN from PDISCHARGED where FORMAT(dateD,'MMMM') = '" & combomonth.Text & "' AND   YEAR(dateD) = '" & txtyear.Text & "' order by DATED ,TIMEOUT ,plast,pfirst,pmi;"
ElseIf combomonth.Text <> "None" And combodate.Text = "None" And txtyear.Text <> "____" And CMED.Value = 1 And CNonMed.Value = 1 Then
sql = "select plast,pfirst,pmi,page,Fdiagnose,patientno,idno,pdatec,dateD,TIMEOUT,PHILPAY,DIFF,TIMEIN from PDISCHARGED where FORMAT(dateD,'MMMM') = '" & combomonth.Text & "' AND   YEAR(dateD) = '" & txtyear.Text & "' order by DATED ,TIMEOUT,plast,pfirst,pmi;"

'***
ElseIf combomonth.Text = "None" And combodate.Text <> "None" And txtyear.Text <> "____" And CMED.Value = 1 And CNonMed.Value = 0 Then
sql = "select plast,pfirst,pmi,page,Fdiagnose,patientno,idno,pdatec,dateD,TIMEOUT,PHILPAY,DIFF,TIMEIN from PDISCHARGED where DAY(dateD) = '" & combodate.Text & "'  AND   YEAR(dateD) = '" & txtyear.Text & "' AND TRIM(IDNO) <> '" & NOID & "' order by DATED ,TIMEOUT ,plast,pfirst,pmi;"
ElseIf combomonth.Text = "None" And combodate.Text <> "None" And txtyear.Text <> "____" And CMED.Value = 0 And CNonMed.Value = 1 Then
sql = "select plast,pfirst,pmi,page,Fdiagnose,patientno,idno,pdatec,dateD,TIMEOUT,PHILPAY,DIFF,TIMEIN from PDISCHARGED where DAY(dateD) = '" & combodate.Text & "'  AND   YEAR(dateD) = '" & txtyear.Text & "' AND TRIM(IDNO) = '" & NOID & "' order by DATED ,TIMEOUT ,plast,pfirst,pmi;"
ElseIf combomonth.Text = "None" And combodate.Text <> "None" And txtyear.Text <> "____" And CMED.Value = 0 And CNonMed.Value = 0 Then
sql = "select plast,pfirst,pmi,page,Fdiagnose,patientno,idno,pdatec,dateD,TIMEOUT,PHILPAY,DIFF,TIMEIN from PDISCHARGED where DAY(dateD) = '" & combodate.Text & "'  AND   YEAR(dateD) = '" & txtyear.Text & "'  order by DATED ,TIMEOUT ,plast,pfirst,pmi;"
ElseIf combomonth.Text = "None" And combodate.Text <> "None" And txtyear.Text <> "____" And CMED.Value = 1 And CNonMed.Value = 1 Then
sql = "select plast,pfirst,pmi,page,Fdiagnose,patientno,idno,pdatec,dateD,TIMEOUT,PHILPAY,DIFF,TIMEIN from PDISCHARGED where DAY(dateD) = '" & combodate.Text & "'  AND   YEAR(dateD) = '" & txtyear.Text & "' order by DATED ,TIMEOUT ,plast,pfirst,pmi;"

'***
ElseIf combomonth.Text = "None" And combodate.Text = "None" And txtyear.Text = "____" And CMED.Value = 1 And CNonMed.Value = 0 Then
sql = "select plast,pfirst,pmi,page,Fdiagnose,patientno,idno,pdatec,dateD,TIMEOUT,PHILPAY,DIFF,TIMEIN from PDISCHARGED where TRIM(IDNO) <> '" & NOID & "' order by DATED ,TIMEOUT ,plast,pfirst,pmi;"
ElseIf combomonth.Text = "None" And combodate.Text = "None" And txtyear.Text = "____" And CMED.Value = 0 And CNonMed.Value = 1 Then
sql = "select plast,pfirst,pmi,page,Fdiagnose,patientno,idno,pdatec,dateD,TIMEOUT,PHILPAY,DIFF,TIMEIN from PDISCHARGED where TRIM(IDNO) = '" & NOID & "' order by DATED ,TIMEOUT ,plast,pfirst,pmi;"
ElseIf combomonth.Text = "None" And combodate.Text = "None" And txtyear.Text = "____" And CMED.Value = 0 And CNonMed.Value = 0 Then
sql = "select plast,pfirst,pmi,page,Fdiagnose,patientno,idno,pdatec,dateD,TIMEOUT,PHILPAY,DIFF,TIMEIN from PDISCHARGED  order by DATED ,TIMEOUT ,plast,pfirst,pmi;"
ElseIf combomonth.Text = "None" And combodate.Text = "None" And txtyear.Text = "____" And CMED.Value = 1 And CNonMed.Value = 1 Then
sql = "select plast,pfirst,pmi,page,Fdiagnose,patientno,idno,pdatec,dateD,TIMEOUT,PHILPAY,DIFF,TIMEIN from PDISCHARGED order by DATED ,TIMEOUT ,plast,pfirst,pmi;"
End If



Set rs = DBMain.OpenRecordset(sql)
ListView1.ListItems.Clear
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
                        If IsNull(rs.Fields(6)) = True Then
                          y.SubItems(6) = "Non-Med"
                        Else
                          y.SubItems(6) = rs.Fields(6)
                        End If
                        y.SubItems(7) = Format(rs.Fields(7), "mm/dd/yyyy") + " " + Format(rs.Fields(12), "medium time")
                        y.SubItems(8) = Format(rs.Fields(8), "mm/dd/yyyy") + " " + Format(rs.Fields(9), "medium time")
                        y.SubItems(9) = Format(rs.Fields(10), "###,##0.00")
                        y.SubItems(10) = Format(rs.Fields(11), "###,##0.00")
               
                        rs.MoveNext
                 Loop
              End If
           '*************************
TRAPPER:

End Sub

Private Sub cmdRep_Click()
Dim TRAPPER As Variant
Dim NOID As String

On Error GoTo TRAPPER
DBa.Open "DRIVER={Microsoft Access Driver (*.mdb)};dbq=" & App.Path & "\hospital.mdb", , "student"

'***
NOID = "Non-Med"
'***
If combomonth.Text <> "None" And combodate.Text = "None" And txtyear.Text = "____" And CMED.Value = 1 And CNonMed.Value = 0 Then
sql = "select plast,pfirst,pmi,page,Fdiagnose,patientno,idno,pdatec,dateD,TIMEOUT,PHILPAY,DIFF,TIMEIN,DISTOT from PDISCHARGED where FORMAT(dateD,'MMMM') = '" & combomonth.Text & "'  AND TRIM(IDNO) <> '" & NOID & "' order by DATED,TIMEOUT,plast,pfirst,pmi;"
ElseIf combomonth.Text <> "None" And combodate.Text = "None" And txtyear.Text = "____" And CMED.Value = 0 And CNonMed.Value = 1 Then
sql = "select plast,pfirst,pmi,page,Fdiagnose,patientno,idno,pdatec,dateD,TIMEOUT,PHILPAY,DIFF,TIMEIN,DISTOT from PDISCHARGED where FORMAT(dATED,'MMMM') = '" & combomonth.Text & "'  AND TRIM(IDNO) = '" & NOID & "' order by DATED,TIMEOUT ,plast,pfirst,pmi;"
ElseIf combomonth.Text <> "None" And combodate.Text = "None" And txtyear.Text = "____" And CMED.Value = 0 And CNonMed.Value = 0 Then
sql = "select plast,pfirst,pmi,page,Fdiagnose,patientno,idno,pdatec,dateD,TIMEOUT,PHILPAY,DIFF,TIMEIN,DISTOT from PDISCHARGED where FORMAT(dateD,'MMMM') = '" & combomonth.Text & "' order by DATED,TIMEOUT,plast,pfirst,pmi;"
ElseIf combomonth.Text <> "None" And combodate.Text = "None" And txtyear.Text = "____" And CMED.Value = 1 And CNonMed.Value = 1 Then
sql = "select plast,pfirst,pmi,page,Fdiagnose,patientno,idno,pdatec,dateD,TIMEOUT,PHILPAY,DIFF,TIMEIN,DISTOT from PDISCHARGED where FORMAT(dateD,'MMMM') = '" & combomonth.Text & "'  order by DATED,TIMEOUT ,plast,pfirst,pmi;"

'***
ElseIf combomonth.Text = "None" And combodate.Text <> "None" And txtyear.Text = "____" And CMED.Value = 1 And CNonMed.Value = 0 Then
sql = "select plast,pfirst,pmi,page,Fdiagnose,patientno,idno,pdatec,dateD,TIMEOUT,PHILPAY,DIFF,TIMEIN,DISTOT from PDISCHARGED where DAY(dateD) = '" & combodate.Text & "' AND TRIM(IDNO) <> '" & NOID & "' order by DATED ,TIMEOUT ,plast,pfirst,pmi;"
ElseIf combomonth.Text = "None" And combodate.Text <> "None" And txtyear.Text = "____" And CMED.Value = 0 And CNonMed.Value = 1 Then
sql = "select plast,pfirst,pmi,page,Fdiagnose,patientno,idno,pdatec,dateD,TIMEOUT,PHILPAY,DIFF,TIMEIN,DISTOT from PDISCHARGED where DAY(dateD) = '" & combodate.Text & "' AND TRIM(IDNO) = '" & NOID & "' order by DATED ,TIMEOUT ,plast,pfirst,pmi;"
ElseIf combomonth.Text = "None" And combodate.Text <> "None" And txtyear.Text = "____" And CMED.Value = 0 And CNonMed.Value = 0 Then
sql = "select plast,pfirst,pmi,page,Fdiagnose,patientno,idno,pdatec,dateD,TIMEOUT,PHILPAY,DIFF,TIMEIN,DISTOT from PDISCHARGED where DAY(dateD) = '" & combodate.Text & "' order by DATED ,TIMEOUT ,plast,pfirst,pmi;"
ElseIf combomonth.Text = "None" And combodate.Text <> "None" And txtyear.Text = "____" And CMED.Value = 1 And CNonMed.Value = 1 Then
sql = "select plast,pfirst,pmi,page,Fdiagnose,patientno,idno,pdatec,dateD,TIMEOUT,PHILPAY,DIFF,TIMEIN,DISTOT from PDISCHARGED where DAY(dateD) = '" & combodate.Text & "' order by DATED ,TIMEOUT ,plast,pfirst,pmi;"

'***
ElseIf combomonth.Text = "None" And combodate.Text = "None" And txtyear.Text <> "____" And CMED.Value = 1 And CNonMed.Value = 0 Then
sql = "select plast,pfirst,pmi,page,Fdiagnose,patientno,idno,pdatec,dateD,TIMEOUT,PHILPAY,DIFF,TIMEIN,DISTOT from PDISCHARGED where YEAR(dateD) = '" & txtyear.Text & "' AND TRIM(IDNO) <> '" & NOID & "' order by DATED ,TIMEOUT,plast,pfirst,pmi;"
ElseIf combomonth.Text = "None" And combodate.Text = "None" And txtyear.Text <> "____" And CMED.Value = 0 And CNonMed.Value = 1 Then
sql = "select plast,pfirst,pmi,page,Fdiagnose,patientno,idno,pdatec,dateD,TIMEOUT,PHILPAY,DIFF,TIMEIN,DISTOT from PDISCHARGED where YEAR(dateD) = '" & txtyear.Text & "' AND TRIM(IDNO) = '" & NOID & "' order by DATED ,TIMEOUT,plast,pfirst,pmi;"
ElseIf combomonth.Text = "None" And combodate.Text = "None" And txtyear.Text <> "____" And CMED.Value = 0 And CNonMed.Value = 0 Then
sql = "select plast,pfirst,pmi,page,Fdiagnose,patientno,idno,pdatec,dateD,TIMEOUT,PHILPAY,DIFF,TIMEIN,DISTOT from PDISCHARGED where YEAR(dateD) = '" & txtyear.Text & "' order by DATED ,TIMEOUT ,plast,pfirst,pmi;"
ElseIf combomonth.Text = "None" And combodate.Text = "None" And txtyear.Text <> "____" And CMED.Value = 1 And CNonMed.Value = 1 Then
sql = "select plast,pfirst,pmi,page,Fdiagnose,patientno,idno,pdatec,dateD,TIMEOUT,PHILPAY,DIFF,TIMEIN,DISTOT from PDISCHARGED where YEAR(dateD) = '" & txtyear.Text & "' order by DATED ,TIMEOUT,plast,pfirst,pmi;"

'***
ElseIf combomonth.Text <> "None" And combodate.Text <> "None" And txtyear.Text <> "____" And CMED.Value = 1 And CNonMed.Value = 0 Then
sql = "select plast,pfirst,pmi,page,Fdiagnose,patientno,idno,pdatec,dateD,TIMEOUT,PHILPAY,DIFF,TIMEIN,DISTOT from PDISCHARGED where FORMAT(dateD,'MMMM') = '" & combomonth.Text & "' AND DAY(dateD) = '" & combodate.Text & "' AND  YEAR(dateD) = '" & txtyear.Text & "' AND TRIM(IDNO) <> '" & NOID & "' order by DATED ,TIMEOUT,plast,pfirst,pmi;"
ElseIf combomonth.Text <> "None" And combodate.Text <> "None" And txtyear.Text <> "____" And CMED.Value = 0 And CNonMed.Value = 1 Then
sql = "select plast,pfirst,pmi,page,Fdiagnose,patientno,idno,pdatec,dateD,TIMEOUT,PHILPAY,DIFF,TIMEIN,DISTOT from PDISCHARGED where FORMAT(dateD,'MMMM') = '" & combomonth.Text & "' AND DAY(dateD) = '" & combodate.Text & "' AND  YEAR(dateD) = '" & txtyear.Text & "' AND TRIM(IDNO) = '" & NOID & "' order by DATED ,TIMEOUT ,plast,pfirst,pmi;"
ElseIf combomonth.Text <> "None" And combodate.Text <> "None" And txtyear.Text <> "____" And CMED.Value = 0 And CNonMed.Value = 0 Then
sql = "select plast,pfirst,pmi,page,Fdiagnose,patientno,idno,pdatec,dateD,TIMEOUT,PHILPAY,DIFF,TIMEIN,DISTOT from PDISCHARGED where FORMAT(dateD,'MMMM') = '" & combomonth.Text & "' AND DAY(dateD) = '" & combodate.Text & "' AND  YEAR(dateD) = '" & txtyear.Text & "' order by DATED ,TIMEOUT,plast,pfirst,pmi;"
ElseIf combomonth.Text <> "None" And combodate.Text <> "None" And txtyear.Text <> "____" And CMED.Value = 1 And CNonMed.Value = 1 Then
sql = "select plast,pfirst,pmi,page,Fdiagnose,patientno,idno,pdatec,dateD,TIMEOUT,PHILPAY,DIFF,TIMEIN,DISTOT from PDISCHARGED where FORMAT(dateD,'MMMM') = '" & combomonth.Text & "' AND DAY(dateD) = '" & combodate.Text & "' AND  YEAR(dateD) = '" & txtyear.Text & "' order by DATED ,TIMEOUT ,plast,pfirst,pmi;"

'***
ElseIf combomonth.Text <> "None" And combodate.Text <> "None" And txtyear.Text = "____" And CMED.Value = 1 And CNonMed.Value = 0 Then
sql = "select plast,pfirst,pmi,page,Fdiagnose,patientno,idno,pdatec,dateD,TIMEOUT,PHILPAY,DIFF,TIMEIN,DISTOT from PDISCHARGED where FORMAT(dateD,'MMMM') = '" & combomonth.Text & "' AND DAY(dateD) = '" & combodate.Text & "' AND TRIM(IDNO) <> '" & NOID & "' order by DATED ,TIMEOUT ,plast,pfirst,pmi;"
ElseIf combomonth.Text <> "None" And combodate.Text <> "None" And txtyear.Text = "____" And CMED.Value = 0 And CNonMed.Value = 1 Then
sql = "select plast,pfirst,pmi,page,Fdiagnose,patientno,idno,pdatec,dateD,TIMEOUT,PHILPAY,DIFF,TIMEIN,DISTOT from PDISCHARGED where FORMAT(dateD,'MMMM') = '" & combomonth.Text & "' AND DAY(dateD) = '" & combodate.Text & "' AND TRIM(IDNO) = '" & NOID & "' order by DATED ,TIMEOUT ,plast,pfirst,pmi;"
ElseIf combomonth.Text <> "None" And combodate.Text <> "None" And txtyear.Text = "____" And CMED.Value = 0 And CNonMed.Value = 0 Then
sql = "select plast,pfirst,pmi,page,Fdiagnose,patientno,idno,pdatec,dateD,TIMEOUT,PHILPAY,DIFF,TIMEIN,DISTOT from PDISCHARGED where FORMAT(dateD,'MMMM') = '" & combomonth.Text & "' AND DAY(dateD) = '" & combodate.Text & "' order by DATED ,TIMEOUT ,plast,pfirst,pmi;"
ElseIf combomonth.Text <> "None" And combodate.Text <> "None" And txtyear.Text = "____" And CMED.Value = 1 And CNonMed.Value = 1 Then
sql = "select plast,pfirst,pmi,page,Fdiagnose,patientno,idno,pdatec,dateD,TIMEOUT,PHILPAY,DIFF,TIMEIN,DISTOT from PDISCHARGED where FORMAT(dateD,'MMMM') = '" & combomonth.Text & "' AND DAY(dateD) = '" & combodate.Text & "' order by DATED ,TIMEOUT ,plast,pfirst,pmi;"


'***
ElseIf combomonth.Text <> "None" And combodate.Text = "None" And txtyear.Text <> "____" And CMED.Value = 1 And CNonMed.Value = 0 Then
sql = "select plast,pfirst,pmi,page,Fdiagnose,patientno,idno,pdatec,dateD,TIMEOUT,PHILPAY,DIFF,TIMEIN,DISTOT from PDISCHARGED where FORMAT(dateD,'MMMM') = '" & combomonth.Text & "' AND   YEAR(dateD) = '" & txtyear.Text & "' AND TRIM(IDNO) <> '" & NOID & "' order by DATED ,TIMEOUT ,plast,pfirst,pmi;"
ElseIf combomonth.Text <> "None" And combodate.Text = "None" And txtyear.Text <> "____" And CMED.Value = 0 And CNonMed.Value = 1 Then
sql = "select plast,pfirst,pmi,page,Fdiagnose,patientno,idno,pdatec,dateD,TIMEOUT,PHILPAY,DIFF,TIMEIN,DISTOT from PDISCHARGED where FORMAT(dateD,'MMMM') = '" & combomonth.Text & "' AND   YEAR(dateD) = '" & txtyear.Text & "' AND TRIM(IDNO) = '" & NOID & "' order by DATED ,TIMEOUT ,plast,pfirst,pmi;"
ElseIf combomonth.Text <> "None" And combodate.Text = "None" And txtyear.Text <> "____" And CMED.Value = 0 And CNonMed.Value = 0 Then
sql = "select plast,pfirst,pmi,page,Fdiagnose,patientno,idno,pdatec,dateD,TIMEOUT,PHILPAY,DIFF,TIMEIN,DISTOT from PDISCHARGED where FORMAT(dateD,'MMMM') = '" & combomonth.Text & "' AND   YEAR(dateD) = '" & txtyear.Text & "' order by DATED ,TIMEOUT ,plast,pfirst,pmi;"
ElseIf combomonth.Text <> "None" And combodate.Text = "None" And txtyear.Text <> "____" And CMED.Value = 1 And CNonMed.Value = 1 Then
sql = "select plast,pfirst,pmi,page,Fdiagnose,patientno,idno,pdatec,dateD,TIMEOUT,PHILPAY,DIFF,TIMEIN,DISTOT from PDISCHARGED where FORMAT(dateD,'MMMM') = '" & combomonth.Text & "' AND   YEAR(dateD) = '" & txtyear.Text & "' order by DATED ,TIMEOUT,plast,pfirst,pmi;"

'***
ElseIf combomonth.Text = "None" And combodate.Text <> "None" And txtyear.Text <> "____" And CMED.Value = 1 And CNonMed.Value = 0 Then
sql = "select plast,pfirst,pmi,page,Fdiagnose,patientno,idno,pdatec,dateD,TIMEOUT,PHILPAY,DIFF,TIMEIN,DISTOT from PDISCHARGED where DAY(dateD) = '" & combodate.Text & "'  AND   YEAR(dateD) = '" & txtyear.Text & "' AND TRIM(IDNO) <> '" & NOID & "' order by DATED ,TIMEOUT ,plast,pfirst,pmi;"
ElseIf combomonth.Text = "None" And combodate.Text <> "None" And txtyear.Text <> "____" And CMED.Value = 0 And CNonMed.Value = 1 Then
sql = "select plast,pfirst,pmi,page,Fdiagnose,patientno,idno,pdatec,dateD,TIMEOUT,PHILPAY,DIFF,TIMEIN,DISTOT from PDISCHARGED where DAY(dateD) = '" & combodate.Text & "'  AND   YEAR(dateD) = '" & txtyear.Text & "' AND TRIM(IDNO) = '" & NOID & "' order by DATED ,TIMEOUT ,plast,pfirst,pmi;"
ElseIf combomonth.Text = "None" And combodate.Text <> "None" And txtyear.Text <> "____" And CMED.Value = 0 And CNonMed.Value = 0 Then
sql = "select plast,pfirst,pmi,page,Fdiagnose,patientno,idno,pdatec,dateD,TIMEOUT,PHILPAY,DIFF,TIMEIN,DISTOT from PDISCHARGED where DAY(dateD) = '" & combodate.Text & "'  AND   YEAR(dateD) = '" & txtyear.Text & "'  order by DATED ,TIMEOUT ,plast,pfirst,pmi;"
ElseIf combomonth.Text = "None" And combodate.Text <> "None" And txtyear.Text <> "____" And CMED.Value = 1 And CNonMed.Value = 1 Then
sql = "select plast,pfirst,pmi,page,Fdiagnose,patientno,idno,pdatec,dateD,TIMEOUT,PHILPAY,DIFF,TIMEIN,DISTOT from PDISCHARGED where DAY(dateD) = '" & combodate.Text & "'  AND   YEAR(dateD) = '" & txtyear.Text & "' order by DATED ,TIMEOUT ,plast,pfirst,pmi;"

'***
ElseIf combomonth.Text = "None" And combodate.Text = "None" And txtyear.Text = "____" And CMED.Value = 1 And CNonMed.Value = 0 Then
sql = "select plast,pfirst,pmi,page,Fdiagnose,patientno,idno,pdatec,dateD,TIMEOUT,PHILPAY,DIFF,TIMEIN,DISTOT from PDISCHARGED where TRIM(IDNO) <> '" & NOID & "' order by DATED ,TIMEOUT ,plast,pfirst,pmi;"
ElseIf combomonth.Text = "None" And combodate.Text = "None" And txtyear.Text = "____" And CMED.Value = 0 And CNonMed.Value = 1 Then
sql = "select plast,pfirst,pmi,page,Fdiagnose,patientno,idno,pdatec,dateD,TIMEOUT,PHILPAY,DIFF,TIMEIN,DISTOT from PDISCHARGED where TRIM(IDNO) = '" & NOID & "' order by DATED ,TIMEOUT ,plast,pfirst,pmi;"
ElseIf combomonth.Text = "None" And combodate.Text = "None" And txtyear.Text = "____" And CMED.Value = 0 And CNonMed.Value = 0 Then
sql = "select plast,pfirst,pmi,page,Fdiagnose,patientno,idno,pdatec,dateD,TIMEOUT,PHILPAY,DIFF,TIMEIN,DISTOT from PDISCHARGED  order by DATED ,TIMEOUT ,plast,pfirst,pmi;"
ElseIf combomonth.Text = "None" And combodate.Text = "None" And txtyear.Text = "____" And CMED.Value = 1 And CNonMed.Value = 1 Then
sql = "select plast,pfirst,pmi,page,Fdiagnose,patientno,idno,pdatec,dateD,TIMEOUT,PHILPAY,DIFF,TIMEIN from PDISCHARGED order by DATED ,TIMEOUT ,plast,pfirst,pmi;"
End If



Set rsa = DBa.Execute(sql)
Set rs = DBMain.OpenRecordset(sql)
ListView1.ListItems.Clear
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
                        If IsNull(rs.Fields(6)) = True Then
                          y.SubItems(6) = "Non-Med"
                        Else
                          y.SubItems(6) = rs.Fields(6)
                        End If
                        y.SubItems(7) = Format(rs.Fields(7), "mm/dd/yyyy") + " " + Format(rs.Fields(12), "medium time")
                        y.SubItems(8) = Format(rs.Fields(8), "mm/dd/yyyy") + " " + Format(rs.Fields(9), "medium time")
                        y.SubItems(9) = Format(rs.Fields(10), "###,##0.00")
                        y.SubItems(10) = Format(rs.Fields(11), "###,##0.00")
                        rs.MoveNext
                 Loop
                 
                 
                 
                 Set PHdisRep.DataSource = rsa
                 
                 PHdisRep.Show vbModal
                
              End If
           '*************************

TRAPPER:
DBa.Close
Set DBa = Nothing
End Sub





Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim tb As String
Add3DBorder Me
Add3DBorder ListView1
combomonth.Text = MonthName(Month(now))
combodate.Text = Day(now)
txtyear.Text = Year(now)
CNonMed.Value = 1
CMED.Value = 1

framecon.Caption = "All Patients Discharged"
sql = "select plast,pfirst,pmi,page,Fdiagnose,patientno,idno,pdatec,dateD,TIMEOUT,PHILPAY,DIFF,TIMEIN from PDISCHARGED where FORMAT(dateD,'MMMM') = '" & combomonth.Text & "' AND DAY(dateD) = '" & combodate.Text & "' AND  YEAR(dateD) = '" & txtyear.Text & "' order by DATED,TIMEOUT,plast,pfirst,pmi;"
Set WSMain = DBEngine.Workspaces(0)
Set DBMain = WSMain.OpenDatabase(App.Path + "\hospital.mdb", False, False, ";pwd=scanhead")

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
                         If IsNull(rs.Fields(6)) = True Then
                          y.SubItems(6) = "Non-Med"
                        Else
                          y.SubItems(6) = rs.Fields(6)
                        End If
                       y.SubItems(7) = Format(rs.Fields(7), "mm/dd/yyyy") + " " + Format(rs.Fields(12), "medium time")
                       y.SubItems(8) = Format(rs.Fields(8), "mm/dd/yyyy") + " " + Format(rs.Fields(9), "medium time")
                       y.SubItems(9) = Format(rs.Fields(10), "###,##0.00")
                       y.SubItems(10) = Format(rs.Fields(11), "###,##0.00")
                    
            rs.MoveNext
                 Loop
              End If
           '*************************
TRAPPER:
End Sub






Private Sub txtyear_GotFocus()
With txtyear
.SelStart = 0
.SelLength = Len(txtyear.Text)
End With
End Sub
