VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form listcon5 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Non-Med Confined Patient(s) List"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10530
   Icon            =   "listcon5.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   10530
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   1665
      Left            =   8700
      TabIndex        =   14
      Top             =   2775
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
   Begin VB.Frame framecon 
      Caption         =   "Non-Med Confined Patient(s)"
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
      TabIndex        =   10
      Top             =   45
      Width           =   10290
      Begin MSComctlLib.ListView ListView1 
         Height          =   2295
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   10050
         _ExtentX        =   17727
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
         NumItems        =   9
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
            Text            =   "Admission  Diagnosed"
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
            Text            =   "Admission Date"
            Object.Width           =   2295
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Admission Time"
            Object.Width           =   2117
         EndProperty
      End
   End
   Begin VB.ComboBox combomonth 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "listcon5.frx":0BC2
      Left            =   165
      List            =   "listcon5.frx":0BED
      TabIndex        =   9
      Top             =   3000
      Width           =   1815
   End
   Begin VB.ComboBox combodate 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "listcon5.frx":0C59
      Left            =   2085
      List            =   "listcon5.frx":0CBD
      TabIndex        =   8
      Top             =   3000
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
      Left            =   5700
      TabIndex        =   7
      Top             =   3000
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
      Height          =   390
      Left            =   4035
      TabIndex        =   6
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Height          =   915
      Left            =   5325
      TabIndex        =   0
      Top             =   3525
      Width           =   3270
      Begin VB.CheckBox CMED 
         Caption         =   "Check1"
         Height          =   375
         Left            =   105
         TabIndex        =   3
         Top             =   465
         Width           =   255
      End
      Begin VB.CheckBox CNonMed 
         Caption         =   "Check1"
         Height          =   375
         Left            =   105
         TabIndex        =   2
         Top             =   165
         Width           =   255
      End
      Begin VB.CommandButton cmdALL 
         Caption         =   "&ALL Confined"
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
         Left            =   1725
         TabIndex        =   1
         Top             =   225
         Width           =   1425
      End
      Begin VB.Label Label1 
         Caption         =   "PH MED"
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   5
         Top             =   540
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "NON-MED"
         Height          =   255
         Index           =   0
         Left            =   465
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
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
      Left            =   2940
      TabIndex        =   12
      Top             =   3000
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
      Left            =   540
      TabIndex        =   13
      Top             =   3690
      Width           =   4215
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000009&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000009&
      FillColor       =   &H00808000&
      FillStyle       =   0  'Solid
      Height          =   405
      Left            =   165
      Shape           =   4  'Rounded Rectangle
      Top             =   3600
      Width           =   5025
   End
End
Attribute VB_Name = "listcon5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sql, rs, y As Variant, ctr As Byte
Dim MyType, passString, madzsrch As String
Dim formod As String



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
'***
If combomonth.Text <> "None" And combodate.Text = "None" And txtyear.Text = "____" And CMED.Value = 1 And CNonMed.Value = 0 Then
sql = "select plastname,pfirstname,pmi,page,padiagnose,patientno,idno,datec,TIMEIN from Patient where FORMAT(datec,'MMMM') = '" & combomonth.Text & "' AND ISNULL(dateD) = TRUE AND ISNULL(IDNO)=FALSE order by DATEC,TIMEIN ,plastname,pfirstname,pmi;"
ElseIf combomonth.Text <> "None" And combodate.Text = "None" And txtyear.Text = "____" And CMED.Value = 0 And CNonMed.Value = 1 Then
sql = "select plastname,pfirstname,pmi,page,padiagnose,patientno,idno,datec,TIMEIN from Patient where FORMAT(datec,'MMMM') = '" & combomonth.Text & "' AND ISNULL(dateD) = TRUE AND ISNULL(IDNO)=TRUE order by DATEC,TIMEIN ,plastname,pfirstname,pmi;"
ElseIf combomonth.Text <> "None" And combodate.Text = "None" And txtyear.Text = "____" And CMED.Value = 0 And CNonMed.Value = 0 Then
sql = "select plastname,pfirstname,pmi,page,padiagnose,patientno,idno,datec,TIMEIN from Patient where FORMAT(datec,'MMMM') = '" & combomonth.Text & "' AND ISNULL(dateD) = TRUE order by DATEC,TIMEIN ,plastname,pfirstname,pmi;"
ElseIf combomonth.Text <> "None" And combodate.Text = "None" And txtyear.Text = "____" And CMED.Value = 1 And CNonMed.Value = 1 Then
sql = "select plastname,pfirstname,pmi,page,padiagnose,patientno,idno,datec,TIMEIN from Patient where FORMAT(datec,'MMMM') = '" & combomonth.Text & "' AND ISNULL(dateD) = TRUE order by DATEC,TIMEIN ,plastname,pfirstname,pmi;"

'***
ElseIf combomonth.Text = "None" And combodate.Text <> "None" And txtyear.Text = "____" And CMED.Value = 1 And CNonMed.Value = 0 Then
sql = "select plastname,pfirstname,pmi,page,padiagnose,patientno,idno,datec,TIMEIN from Patient where DAY(datec) = '" & combodate.Text & "' AND ISNULL(dateD) = TRUE AND ISNULL(IDNO)=FALSE order by DATEC ,TIMEIN ,plastname,pfirstname,pmi;"
ElseIf combomonth.Text = "None" And combodate.Text <> "None" And txtyear.Text = "____" And CMED.Value = 0 And CNonMed.Value = 1 Then
sql = "select plastname,pfirstname,pmi,page,padiagnose,patientno,idno,datec,TIMEIN from Patient where DAY(datec) = '" & combodate.Text & "' AND ISNULL(dateD) = TRUE AND ISNULL(IDNO)=TRUE order by DATEC ,TIMEIN ,plastname,pfirstname,pmi;"
ElseIf combomonth.Text = "None" And combodate.Text <> "None" And txtyear.Text = "____" And CMED.Value = 0 And CNonMed.Value = 0 Then
sql = "select plastname,pfirstname,pmi,page,padiagnose,patientno,idno,datec,TIMEIN from Patient where DAY(datec) = '" & combodate.Text & "' AND ISNULL(dateD) = TRUE  order by DATEC ,TIMEIN ,plastname,pfirstname,pmi;"
ElseIf combomonth.Text = "None" And combodate.Text <> "None" And txtyear.Text = "____" And CMED.Value = 1 And CNonMed.Value = 1 Then
sql = "select plastname,pfirstname,pmi,page,padiagnose,patientno,idno,datec,TIMEIN from Patient where DAY(datec) = '" & combodate.Text & "' AND ISNULL(dateD) = TRUE  order by DATEC ,TIMEIN ,plastname,pfirstname,pmi;"

'***
ElseIf combomonth.Text = "None" And combodate.Text = "None" And txtyear.Text <> "____" And CMED.Value = 1 And CNonMed.Value = 0 Then
sql = "select plastname,pfirstname,pmi,page,padiagnose,patientno,idno,datec,TIMEIN from Patient where YEAR(datec) = '" & txtyear.Text & "' AND ISNULL(dateD) = TRUE AND ISNULL(IDNO)=FALSE order by DATEC ,TIMEIN ,plastname,pfirstname,pmi;"
ElseIf combomonth.Text = "None" And combodate.Text = "None" And txtyear.Text <> "____" And CMED.Value = 0 And CNonMed.Value = 1 Then
sql = "select plastname,pfirstname,pmi,page,padiagnose,patientno,idno,datec,TIMEIN from Patient where YEAR(datec) = '" & txtyear.Text & "' AND ISNULL(dateD) = TRUE AND ISNULL(IDNO)=TRUE order by DATEC ,TIMEIN ,plastname,pfirstname,pmi;"
ElseIf combomonth.Text = "None" And combodate.Text = "None" And txtyear.Text <> "____" And CMED.Value = 0 And CNonMed.Value = 0 Then
sql = "select plastname,pfirstname,pmi,page,padiagnose,patientno,idno,datec,TIMEIN from Patient where YEAR(datec) = '" & txtyear.Text & "' AND ISNULL(dateD) = TRUE  order by DATEC ,TIMEIN ,plastname,pfirstname,pmi;"
ElseIf combomonth.Text = "None" And combodate.Text = "None" And txtyear.Text <> "____" And CMED.Value = 1 And CNonMed.Value = 1 Then
sql = "select plastname,pfirstname,pmi,page,padiagnose,patientno,idno,datec,TIMEIN from Patient where YEAR(datec) = '" & txtyear.Text & "' AND ISNULL(dateD) = TRUE  order by DATEC ,TIMEIN ,plastname,pfirstname,pmi;"

'***
ElseIf combomonth.Text <> "None" And combodate.Text <> "None" And txtyear.Text <> "____" And CMED.Value = 1 And CNonMed.Value = 0 Then
sql = "select plastname,pfirstname,pmi,page,padiagnose,patientno,idno,datec,TIMEIN from Patient where FORMAT(datec,'MMMM') = '" & combomonth.Text & "' AND DAY(datec) = '" & combodate.Text & "' AND  YEAR(datec) = '" & txtyear.Text & "' AND ISNULL(dateD) = TRUE AND ISNULL(IDNO)=FALSE order by DATEC ,TIMEIN ,plastname,pfirstname,pmi;"
ElseIf combomonth.Text <> "None" And combodate.Text <> "None" And txtyear.Text <> "____" And CMED.Value = 0 And CNonMed.Value = 1 Then
sql = "select plastname,pfirstname,pmi,page,padiagnose,patientno,idno,datec,TIMEIN from Patient where FORMAT(datec,'MMMM') = '" & combomonth.Text & "' AND DAY(datec) = '" & combodate.Text & "' AND  YEAR(datec) = '" & txtyear.Text & "' AND ISNULL(dateD) = TRUE AND ISNULL(IDNO)=TRUE order by DATEC ,TIMEIN ,plastname,pfirstname,pmi;"
ElseIf combomonth.Text <> "None" And combodate.Text <> "None" And txtyear.Text <> "____" And CMED.Value = 0 And CNonMed.Value = 0 Then
sql = "select plastname,pfirstname,pmi,page,padiagnose,patientno,idno,datec,TIMEIN from Patient where FORMAT(datec,'MMMM') = '" & combomonth.Text & "' AND DAY(datec) = '" & combodate.Text & "' AND  YEAR(datec) = '" & txtyear.Text & "' AND ISNULL(dateD) = TRUE  order by DATEC ,TIMEIN ,plastname,pfirstname,pmi;"
ElseIf combomonth.Text <> "None" And combodate.Text <> "None" And txtyear.Text <> "____" And CMED.Value = 1 And CNonMed.Value = 1 Then
sql = "select plastname,pfirstname,pmi,page,padiagnose,patientno,idno,datec,TIMEIN from Patient where FORMAT(datec,'MMMM') = '" & combomonth.Text & "' AND DAY(datec) = '" & combodate.Text & "' AND  YEAR(datec) = '" & txtyear.Text & "' AND ISNULL(dateD) = TRUE  order by DATEC ,TIMEIN ,plastname,pfirstname,pmi;"

'***
ElseIf combomonth.Text <> "None" And combodate.Text <> "None" And txtyear.Text = "____" And CMED.Value = 1 And CNonMed.Value = 0 Then
sql = "select plastname,pfirstname,pmi,page,padiagnose,patientno,idno,datec,TIMEIN from Patient where FORMAT(datec,'MMMM') = '" & combomonth.Text & "' AND DAY(datec) = '" & combodate.Text & "'  AND ISNULL(dateD) = TRUE AND ISNULL(IDNO)=FALSE order by DATEC ,TIMEIN ,plastname,pfirstname,pmi;"
ElseIf combomonth.Text <> "None" And combodate.Text <> "None" And txtyear.Text = "____" And CMED.Value = 0 And CNonMed.Value = 1 Then
sql = "select plastname,pfirstname,pmi,page,padiagnose,patientno,idno,datec,TIMEIN from Patient where FORMAT(datec,'MMMM') = '" & combomonth.Text & "' AND DAY(datec) = '" & combodate.Text & "'  AND ISNULL(dateD) = TRUE AND ISNULL(IDNO)=TRUE order by DATEC ,TIMEIN ,plastname,pfirstname,pmi;"
ElseIf combomonth.Text <> "None" And combodate.Text <> "None" And txtyear.Text = "____" And CMED.Value = 0 And CNonMed.Value = 0 Then
sql = "select plastname,pfirstname,pmi,page,padiagnose,patientno,idno,datec,TIMEIN from Patient where FORMAT(datec,'MMMM') = '" & combomonth.Text & "' AND DAY(datec) = '" & combodate.Text & "'  AND ISNULL(dateD) = TRUE  order by DATEC ,TIMEIN ,plastname,pfirstname,pmi;"
ElseIf combomonth.Text <> "None" And combodate.Text <> "None" And txtyear.Text = "____" And CMED.Value = 1 And CNonMed.Value = 1 Then
sql = "select plastname,pfirstname,pmi,page,padiagnose,patientno,idno,datec,TIMEIN from Patient where FORMAT(datec,'MMMM') = '" & combomonth.Text & "' AND DAY(datec) = '" & combodate.Text & "'  AND ISNULL(dateD) = TRUE  order by DATEC ,TIMEIN ,plastname,pfirstname,pmi;"


'***
ElseIf combomonth.Text <> "None" And combodate.Text = "None" And txtyear.Text <> "____" And CMED.Value = 1 And CNonMed.Value = 0 Then
sql = "select plastname,pfirstname,pmi,page,padiagnose,patientno,idno,datec,TIMEIN from Patient where FORMAT(datec,'MMMM') = '" & combomonth.Text & "' AND   YEAR(datec) = '" & txtyear.Text & "' AND ISNULL(dateD) = TRUE AND ISNULL(IDNO)=FALSE order by DATEC ,TIMEIN ,plastname,pfirstname,pmi;"
ElseIf combomonth.Text <> "None" And combodate.Text = "None" And txtyear.Text <> "____" And CMED.Value = 0 And CNonMed.Value = 1 Then
sql = "select plastname,pfirstname,pmi,page,padiagnose,patientno,idno,datec,TIMEIN from Patient where FORMAT(datec,'MMMM') = '" & combomonth.Text & "' AND   YEAR(datec) = '" & txtyear.Text & "' AND ISNULL(dateD) = TRUE AND ISNULL(IDNO)=TRUE order by DATEC ,TIMEIN ,plastname,pfirstname,pmi;"
ElseIf combomonth.Text <> "None" And combodate.Text = "None" And txtyear.Text <> "____" And CMED.Value = 0 And CNonMed.Value = 0 Then
sql = "select plastname,pfirstname,pmi,page,padiagnose,patientno,idno,datec,TIMEIN from Patient where FORMAT(datec,'MMMM') = '" & combomonth.Text & "' AND   YEAR(datec) = '" & txtyear.Text & "' AND ISNULL(dateD) = TRUE  order by DATEC ,TIMEIN ,plastname,pfirstname,pmi;"
ElseIf combomonth.Text <> "None" And combodate.Text = "None" And txtyear.Text <> "____" And CMED.Value = 1 And CNonMed.Value = 1 Then
sql = "select plastname,pfirstname,pmi,page,padiagnose,patientno,idno,datec,TIMEIN from Patient where FORMAT(datec,'MMMM') = '" & combomonth.Text & "' AND   YEAR(datec) = '" & txtyear.Text & "' AND ISNULL(dateD) = TRUE  order by DATEC ,TIMEIN ,plastname,pfirstname,pmi;"

'***
ElseIf combomonth.Text = "None" And combodate.Text <> "None" And txtyear.Text <> "____" And CMED.Value = 1 And CNonMed.Value = 0 Then
sql = "select plastname,pfirstname,pmi,page,padiagnose,patientno,idno,datec,TIMEIN from Patient where DAY(datec) = '" & combodate.Text & "'  AND   YEAR(datec) = '" & txtyear.Text & "' AND ISNULL(dateD) = TRUE AND ISNULL(IDNO)=FALSE order by DATEC ,TIMEIN ,plastname,pfirstname,pmi;"
ElseIf combomonth.Text = "None" And combodate.Text <> "None" And txtyear.Text <> "____" And CMED.Value = 0 And CNonMed.Value = 1 Then
sql = "select plastname,pfirstname,pmi,page,padiagnose,patientno,idno,datec,TIMEIN from Patient where DAY(datec) = '" & combodate.Text & "'  AND   YEAR(datec) = '" & txtyear.Text & "' AND ISNULL(dateD) = TRUE AND ISNULL(IDNO)=TRUE order by DATEC ,TIMEIN ,plastname,pfirstname,pmi;"
ElseIf combomonth.Text = "None" And combodate.Text <> "None" And txtyear.Text <> "____" And CMED.Value = 0 And CNonMed.Value = 0 Then
sql = "select plastname,pfirstname,pmi,page,padiagnose,patientno,idno,datec,TIMEIN from Patient where DAY(datec) = '" & combodate.Text & "'  AND   YEAR(datec) = '" & txtyear.Text & "' AND ISNULL(dateD) = TRUE  order by DATEC ,TIMEIN ,plastname,pfirstname,pmi;"
ElseIf combomonth.Text = "None" And combodate.Text <> "None" And txtyear.Text <> "____" And CMED.Value = 1 And CNonMed.Value = 1 Then
sql = "select plastname,pfirstname,pmi,page,padiagnose,patientno,idno,datec,TIMEIN from Patient where DAY(datec) = '" & combodate.Text & "'  AND   YEAR(datec) = '" & txtyear.Text & "' AND ISNULL(dateD) = TRUE  order by DATEC ,TIMEIN ,plastname,pfirstname,pmi;"

'***
ElseIf combomonth.Text = "None" And combodate.Text = "None" And txtyear.Text = "____" And CMED.Value = 1 And CNonMed.Value = 0 Then
sql = "select plastname,pfirstname,pmi,page,padiagnose,patientno,idno,datec,TIMEIN from Patient where ISNULL(dateD) = TRUE AND ISNULL(IDNO)=FALSE order by DATEC ,TIMEIN ,plastname,pfirstname,pmi;"
ElseIf combomonth.Text = "None" And combodate.Text = "None" And txtyear.Text = "____" And CMED.Value = 0 And CNonMed.Value = 1 Then
sql = "select plastname,pfirstname,pmi,page,padiagnose,patientno,idno,datec,TIMEIN from Patient where ISNULL(dateD) = TRUE AND ISNULL(IDNO)=TRUE order by DATEC ,TIMEIN ,plastname,pfirstname,pmi;"
ElseIf combomonth.Text = "None" And combodate.Text = "None" And txtyear.Text = "____" And CMED.Value = 0 And CNonMed.Value = 0 Then
sql = "select plastname,pfirstname,pmi,page,padiagnose,patientno,idno,datec,TIMEIN from Patient where ISNULL(dateD) = TRUE  order by DATEC ,TIMEIN ,plastname,pfirstname,pmi;"
ElseIf combomonth.Text = "None" And combodate.Text = "None" And txtyear.Text = "____" And CMED.Value = 1 And CNonMed.Value = 1 Then
sql = "select plastname,pfirstname,pmi,page,padiagnose,patientno,idno,datec,TIMEIN from Patient where ISNULL(dateD) = TRUE  order by DATEC ,TIMEIN ,plastname,pfirstname,pmi;"
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
                        y.SubItems(7) = Format(rs.Fields(7), "mm/dd/yyyy")
                        y.SubItems(8) = Format(rs.Fields(8), "medium time")
                        rs.MoveNext
                 Loop
              End If
           '*************************
TRAPPER:

End Sub

Private Sub cmdRep_Click()
Dim TRAPPER As Variant

On Error GoTo TRAPPER
DBa.Open "DRIVER={Microsoft Access Driver (*.mdb)};dbq=" & App.Path & "\hospital.mdb", , "student"

'***
If combomonth.Text <> "None" And combodate.Text = "None" And txtyear.Text = "____" And CMED.Value = 1 And CNonMed.Value = 0 Then
sql = "select plastname,pfirstname,pmi,page,padiagnose,patientno,idno,datec,TIMEIN,NOID from Patient where FORMAT(datec,'MMMM') = '" & combomonth.Text & "' AND ISNULL(dateD) = TRUE AND ISNULL(IDNO)=FALSE order by DATEC,TIMEIN ,plastname,pfirstname,pmi;"
ElseIf combomonth.Text <> "None" And combodate.Text = "None" And txtyear.Text = "____" And CMED.Value = 0 And CNonMed.Value = 1 Then
sql = "select plastname,pfirstname,pmi,page,padiagnose,patientno,idno,datec,TIMEIN,NOID from Patient where FORMAT(datec,'MMMM') = '" & combomonth.Text & "' AND ISNULL(dateD) = TRUE AND ISNULL(IDNO)=TRUE order by DATEC,TIMEIN ,plastname,pfirstname,pmi;"
ElseIf combomonth.Text <> "None" And combodate.Text = "None" And txtyear.Text = "____" And CMED.Value = 0 And CNonMed.Value = 0 Then
sql = "select plastname,pfirstname,pmi,page,padiagnose,patientno,idno,datec,TIMEIN,NOID from Patient where FORMAT(datec,'MMMM') = '" & combomonth.Text & "' AND ISNULL(dateD) = TRUE order by DATEC,TIMEIN ,plastname,pfirstname,pmi;"
ElseIf combomonth.Text <> "None" And combodate.Text = "None" And txtyear.Text = "____" And CMED.Value = 1 And CNonMed.Value = 1 Then
sql = "select plastname,pfirstname,pmi,page,padiagnose,patientno,idno,datec,TIMEIN,NOID from Patient where FORMAT(datec,'MMMM') = '" & combomonth.Text & "' AND ISNULL(dateD) = TRUE order by DATEC,TIMEIN ,plastname,pfirstname,pmi;"

'***
ElseIf combomonth.Text = "None" And combodate.Text <> "None" And txtyear.Text = "____" And CMED.Value = 1 And CNonMed.Value = 0 Then
sql = "select plastname,pfirstname,pmi,page,padiagnose,patientno,idno,datec,TIMEIN,NOID from Patient where DAY(datec) = '" & combodate.Text & "' AND ISNULL(dateD) = TRUE AND ISNULL(IDNO)=FALSE order by DATEC ,TIMEIN ,plastname,pfirstname,pmi;"
ElseIf combomonth.Text = "None" And combodate.Text <> "None" And txtyear.Text = "____" And CMED.Value = 0 And CNonMed.Value = 1 Then
sql = "select plastname,pfirstname,pmi,page,padiagnose,patientno,idno,datec,TIMEIN,NOID from Patient where DAY(datec) = '" & combodate.Text & "' AND ISNULL(dateD) = TRUE AND ISNULL(IDNO)=TRUE order by DATEC ,TIMEIN ,plastname,pfirstname,pmi;"
ElseIf combomonth.Text = "None" And combodate.Text <> "None" And txtyear.Text = "____" And CMED.Value = 0 And CNonMed.Value = 0 Then
sql = "select plastname,pfirstname,pmi,page,padiagnose,patientno,idno,datec,TIMEIN,NOID from Patient where DAY(datec) = '" & combodate.Text & "' AND ISNULL(dateD) = TRUE  order by DATEC ,TIMEIN ,plastname,pfirstname,pmi;"
ElseIf combomonth.Text = "None" And combodate.Text <> "None" And txtyear.Text = "____" And CMED.Value = 1 And CNonMed.Value = 1 Then
sql = "select plastname,pfirstname,pmi,page,padiagnose,patientno,idno,datec,TIMEIN,NOID from Patient where DAY(datec) = '" & combodate.Text & "' AND ISNULL(dateD) = TRUE  order by DATEC ,TIMEIN ,plastname,pfirstname,pmi;"

'***
ElseIf combomonth.Text = "None" And combodate.Text = "None" And txtyear.Text <> "____" And CMED.Value = 1 And CNonMed.Value = 0 Then
sql = "select plastname,pfirstname,pmi,page,padiagnose,patientno,idno,datec,TIMEIN,NOID from Patient where YEAR(datec) = '" & txtyear.Text & "' AND ISNULL(dateD) = TRUE AND ISNULL(IDNO)=FALSE order by DATEC ,TIMEIN ,plastname,pfirstname,pmi;"
ElseIf combomonth.Text = "None" And combodate.Text = "None" And txtyear.Text <> "____" And CMED.Value = 0 And CNonMed.Value = 1 Then
sql = "select plastname,pfirstname,pmi,page,padiagnose,patientno,idno,datec,TIMEIN,NOID from Patient where YEAR(datec) = '" & txtyear.Text & "' AND ISNULL(dateD) = TRUE AND ISNULL(IDNO)=TRUE order by DATEC ,TIMEIN ,plastname,pfirstname,pmi;"
ElseIf combomonth.Text = "None" And combodate.Text = "None" And txtyear.Text <> "____" And CMED.Value = 0 And CNonMed.Value = 0 Then
sql = "select plastname,pfirstname,pmi,page,padiagnose,patientno,idno,datec,TIMEIN,NOID from Patient where YEAR(datec) = '" & txtyear.Text & "' AND ISNULL(dateD) = TRUE  order by DATEC ,TIMEIN ,plastname,pfirstname,pmi;"
ElseIf combomonth.Text = "None" And combodate.Text = "None" And txtyear.Text <> "____" And CMED.Value = 1 And CNonMed.Value = 1 Then
sql = "select plastname,pfirstname,pmi,page,padiagnose,patientno,idno,datec,TIMEIN,NOID from Patient where YEAR(datec) = '" & txtyear.Text & "' AND ISNULL(dateD) = TRUE  order by DATEC ,TIMEIN ,plastname,pfirstname,pmi;"

'***
ElseIf combomonth.Text <> "None" And combodate.Text <> "None" And txtyear.Text <> "____" And CMED.Value = 1 And CNonMed.Value = 0 Then
sql = "select plastname,pfirstname,pmi,page,padiagnose,patientno,idno,datec,TIMEIN,NOID from Patient where FORMAT(datec,'MMMM') = '" & combomonth.Text & "' AND DAY(datec) = '" & combodate.Text & "' AND  YEAR(datec) = '" & txtyear.Text & "' AND ISNULL(dateD) = TRUE AND ISNULL(IDNO)=FALSE order by DATEC ,TIMEIN ,plastname,pfirstname,pmi;"
ElseIf combomonth.Text <> "None" And combodate.Text <> "None" And txtyear.Text <> "____" And CMED.Value = 0 And CNonMed.Value = 1 Then
sql = "select plastname,pfirstname,pmi,page,padiagnose,patientno,idno,datec,TIMEIN,NOID from Patient where FORMAT(datec,'MMMM') = '" & combomonth.Text & "' AND DAY(datec) = '" & combodate.Text & "' AND  YEAR(datec) = '" & txtyear.Text & "' AND ISNULL(dateD) = TRUE AND ISNULL(IDNO)=TRUE order by DATEC ,TIMEIN ,plastname,pfirstname,pmi;"
ElseIf combomonth.Text <> "None" And combodate.Text <> "None" And txtyear.Text <> "____" And CMED.Value = 0 And CNonMed.Value = 0 Then
sql = "select plastname,pfirstname,pmi,page,padiagnose,patientno,idno,datec,TIMEIN,NOID from Patient where FORMAT(datec,'MMMM') = '" & combomonth.Text & "' AND DAY(datec) = '" & combodate.Text & "' AND  YEAR(datec) = '" & txtyear.Text & "' AND ISNULL(dateD) = TRUE  order by DATEC ,TIMEIN ,plastname,pfirstname,pmi;"
ElseIf combomonth.Text <> "None" And combodate.Text <> "None" And txtyear.Text <> "____" And CMED.Value = 1 And CNonMed.Value = 1 Then
sql = "select plastname,pfirstname,pmi,page,padiagnose,patientno,idno,datec,TIMEIN,NOID from Patient where FORMAT(datec,'MMMM') = '" & combomonth.Text & "' AND DAY(datec) = '" & combodate.Text & "' AND  YEAR(datec) = '" & txtyear.Text & "' AND ISNULL(dateD) = TRUE  order by DATEC ,TIMEIN ,plastname,pfirstname,pmi;"

'***
ElseIf combomonth.Text <> "None" And combodate.Text <> "None" And txtyear.Text = "____" And CMED.Value = 1 And CNonMed.Value = 0 Then
sql = "select plastname,pfirstname,pmi,page,padiagnose,patientno,idno,datec,TIMEIN,NOID from Patient where FORMAT(datec,'MMMM') = '" & combomonth.Text & "' AND DAY(datec) = '" & combodate.Text & "'  AND ISNULL(dateD) = TRUE AND ISNULL(IDNO)=FALSE order by DATEC ,TIMEIN ,plastname,pfirstname,pmi;"
ElseIf combomonth.Text <> "None" And combodate.Text <> "None" And txtyear.Text = "____" And CMED.Value = 0 And CNonMed.Value = 1 Then
sql = "select plastname,pfirstname,pmi,page,padiagnose,patientno,idno,datec,TIMEIN,NOID from Patient where FORMAT(datec,'MMMM') = '" & combomonth.Text & "' AND DAY(datec) = '" & combodate.Text & "'  AND ISNULL(dateD) = TRUE AND ISNULL(IDNO)=TRUE order by DATEC ,TIMEIN ,plastname,pfirstname,pmi;"
ElseIf combomonth.Text <> "None" And combodate.Text <> "None" And txtyear.Text = "____" And CMED.Value = 0 And CNonMed.Value = 0 Then
sql = "select plastname,pfirstname,pmi,page,padiagnose,patientno,idno,datec,TIMEIN,NOID from Patient where FORMAT(datec,'MMMM') = '" & combomonth.Text & "' AND DAY(datec) = '" & combodate.Text & "'  AND ISNULL(dateD) = TRUE  order by DATEC ,TIMEIN ,plastname,pfirstname,pmi;"
ElseIf combomonth.Text <> "None" And combodate.Text <> "None" And txtyear.Text = "____" And CMED.Value = 1 And CNonMed.Value = 1 Then
sql = "select plastname,pfirstname,pmi,page,padiagnose,patientno,idno,datec,TIMEIN,NOID from Patient where FORMAT(datec,'MMMM') = '" & combomonth.Text & "' AND DAY(datec) = '" & combodate.Text & "'  AND ISNULL(dateD) = TRUE  order by DATEC ,TIMEIN ,plastname,pfirstname,pmi;"


'***
ElseIf combomonth.Text <> "None" And combodate.Text = "None" And txtyear.Text <> "____" And CMED.Value = 1 And CNonMed.Value = 0 Then
sql = "select plastname,pfirstname,pmi,page,padiagnose,patientno,idno,datec,TIMEIN,NOID from Patient where FORMAT(datec,'MMMM') = '" & combomonth.Text & "' AND   YEAR(datec) = '" & txtyear.Text & "' AND ISNULL(dateD) = TRUE AND ISNULL(IDNO)=FALSE order by DATEC ,TIMEIN ,plastname,pfirstname,pmi;"
ElseIf combomonth.Text <> "None" And combodate.Text = "None" And txtyear.Text <> "____" And CMED.Value = 0 And CNonMed.Value = 1 Then
sql = "select plastname,pfirstname,pmi,page,padiagnose,patientno,idno,datec,TIMEIN,NOID from Patient where FORMAT(datec,'MMMM') = '" & combomonth.Text & "' AND   YEAR(datec) = '" & txtyear.Text & "' AND ISNULL(dateD) = TRUE AND ISNULL(IDNO)=TRUE order by DATEC ,TIMEIN ,plastname,pfirstname,pmi;"
ElseIf combomonth.Text <> "None" And combodate.Text = "None" And txtyear.Text <> "____" And CMED.Value = 0 And CNonMed.Value = 0 Then
sql = "select plastname,pfirstname,pmi,page,padiagnose,patientno,idno,datec,TIMEIN,NOID from Patient where FORMAT(datec,'MMMM') = '" & combomonth.Text & "' AND   YEAR(datec) = '" & txtyear.Text & "' AND ISNULL(dateD) = TRUE order by DATEC ,TIMEIN ,plastname,pfirstname,pmi;"
ElseIf combomonth.Text <> "None" And combodate.Text = "None" And txtyear.Text <> "____" And CMED.Value = 1 And CNonMed.Value = 1 Then
sql = "select plastname,pfirstname,pmi,page,padiagnose,patientno,idno,datec,TIMEIN,NOID from Patient where FORMAT(datec,'MMMM') = '" & combomonth.Text & "' AND   YEAR(datec) = '" & txtyear.Text & "' AND ISNULL(dateD) = TRUE order by DATEC ,TIMEIN ,plastname,pfirstname,pmi;"

'***
ElseIf combomonth.Text = "None" And combodate.Text <> "None" And txtyear.Text <> "____" And CMED.Value = 1 And CNonMed.Value = 0 Then
sql = "select plastname,pfirstname,pmi,page,padiagnose,patientno,idno,datec,TIMEIN,NOID from Patient where DAY(datec) = '" & combodate.Text & "'  AND   YEAR(datec) = '" & txtyear.Text & "' AND ISNULL(dateD) = TRUE AND ISNULL(IDNO)=FALSE order by DATEC ,TIMEIN ,plastname,pfirstname,pmi;"
ElseIf combomonth.Text = "None" And combodate.Text <> "None" And txtyear.Text <> "____" And CMED.Value = 0 And CNonMed.Value = 1 Then
sql = "select plastname,pfirstname,pmi,page,padiagnose,patientno,idno,datec,TIMEIN,NOID from Patient where DAY(datec) = '" & combodate.Text & "'  AND   YEAR(datec) = '" & txtyear.Text & "' AND ISNULL(dateD) = TRUE AND ISNULL(IDNO)=TRUE order by DATEC ,TIMEIN ,plastname,pfirstname,pmi;"
ElseIf combomonth.Text = "None" And combodate.Text <> "None" And txtyear.Text <> "____" And CMED.Value = 0 And CNonMed.Value = 0 Then
sql = "select plastname,pfirstname,pmi,page,padiagnose,patientno,idno,datec,TIMEIN,NOID from Patient where DAY(datec) = '" & combodate.Text & "'  AND   YEAR(datec) = '" & txtyear.Text & "' AND ISNULL(dateD) = TRUE  order by DATEC ,TIMEIN ,plastname,pfirstname,pmi;"
ElseIf combomonth.Text = "None" And combodate.Text <> "None" And txtyear.Text <> "____" And CMED.Value = 1 And CNonMed.Value = 1 Then
sql = "select plastname,pfirstname,pmi,page,padiagnose,patientno,idno,datec,TIMEIN,NOID from Patient where DAY(datec) = '" & combodate.Text & "'  AND   YEAR(datec) = '" & txtyear.Text & "' AND ISNULL(dateD) = TRUE  order by DATEC ,TIMEIN ,plastname,pfirstname,pmi;"

'***
ElseIf combomonth.Text = "None" And combodate.Text = "None" And txtyear.Text = "____" And CMED.Value = 1 And CNonMed.Value = 0 Then
sql = "select plastname,pfirstname,pmi,page,padiagnose,patientno,idno,datec,TIMEIN,NOID from Patient where ISNULL(dateD) = TRUE AND ISNULL(IDNO)=FALSE order by DATEC ,TIMEIN ,plastname,pfirstname,pmi;"
ElseIf combomonth.Text = "None" And combodate.Text = "None" And txtyear.Text = "____" And CMED.Value = 0 And CNonMed.Value = 1 Then
sql = "select plastname,pfirstname,pmi,page,padiagnose,patientno,idno,datec,TIMEIN,NOID from Patient where ISNULL(dateD) = TRUE AND ISNULL(IDNO)=TRUE order by DATEC ,TIMEIN ,plastname,pfirstname,pmi;"
ElseIf combomonth.Text = "None" And combodate.Text = "None" And txtyear.Text = "____" And CMED.Value = 0 And CNonMed.Value = 0 Then
sql = "select plastname,pfirstname,pmi,page,padiagnose,patientno,idno,datec,TIMEIN,NOID from Patient where ISNULL(dateD) = TRUE  order by DATEC ,TIMEIN ,plastname,pfirstname,pmi;"
ElseIf combomonth.Text = "None" And combodate.Text = "None" And txtyear.Text = "____" And CMED.Value = 1 And CNonMed.Value = 1 Then
sql = "select plastname,pfirstname,pmi,page,padiagnose,patientno,idno,datec,TIMEIN,NOID from Patient where ISNULL(dateD) = TRUE  order by DATEC ,TIMEIN ,plastname,pfirstname,pmi;"
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
                        y.SubItems(7) = Format(rs.Fields(7), "mm/dd/yyyy")
                        y.SubItems(8) = Format(rs.Fields(8), "medium time")
                        rs.MoveNext
                 Loop
                 Set PHconRep.DataSource = rsa
                 PHconRep.Show vbModal
                
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
'**** FX SCANHEAD
Add3DBorder Me
Add3DBorder ListView1
'*********************
combomonth.Text = MonthName(Month(now))
combodate.Text = Day(now)
txtyear.Text = Year(now)
CNonMed.Value = 1
CMED.Value = 0
sql = "select plastname,pfirstname,pmi,page,padiagnose,patientno,idno,datec,TIMEIN from Patient where FORMAT(datec,'MMMM') = '" & combomonth.Text & "' AND DAY(datec) = '" & combodate.Text & "' AND  YEAR(datec) = '" & txtyear.Text & "' AND ISNULL(dateD) = TRUE AND ISNULL(IDNO)=TRUE order by DATEC,TIMEIN,plastname,pfirstname,pmi;"
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
                     y.SubItems(7) = Format(rs.Fields(7), "mm/dd/yyyy")
                        y.SubItems(8) = Format(rs.Fields(8), "medium time")
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

