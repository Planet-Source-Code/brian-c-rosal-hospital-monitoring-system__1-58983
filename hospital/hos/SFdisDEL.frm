VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form SFdisDEL 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search Discharged Patient(s)"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4365
   Icon            =   "SFdisDEL.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   4365
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5505
      Left            =   75
      TabIndex        =   11
      Top             =   75
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   9710
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Search Area"
      TabPicture(0)   =   "SFdisDEL.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame Frame1 
         Height          =   4875
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   3930
         Begin VB.Frame Frame2 
            Caption         =   "Search By:"
            ForeColor       =   &H00FF0000&
            Height          =   975
            Left            =   330
            TabIndex        =   17
            Top             =   3240
            Width           =   3345
            Begin VB.CheckBox CNonMed 
               Caption         =   "Check1"
               Height          =   375
               Left            =   195
               TabIndex        =   5
               Top             =   330
               Width           =   255
            End
            Begin VB.CheckBox CMed 
               Caption         =   "Check1"
               Height          =   375
               Left            =   1800
               TabIndex        =   6
               Top             =   330
               Width           =   255
            End
            Begin VB.Label Label1 
               Caption         =   "NON-MED only"
               Height          =   255
               Index           =   5
               Left            =   465
               TabIndex        =   19
               Top             =   405
               Width           =   1095
            End
            Begin VB.Label Label1 
               Caption         =   "PH MED only"
               Height          =   255
               Index           =   4
               Left            =   2070
               TabIndex        =   18
               Top             =   420
               Width           =   990
            End
         End
         Begin MSMask.MaskEdBox txtyear 
            Height          =   390
            Left            =   360
            TabIndex        =   3
            Top             =   2745
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   688
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
         Begin VB.CheckBox Checkyear 
            Caption         =   "Check1"
            Height          =   375
            Left            =   3465
            TabIndex        =   10
            Top             =   2760
            Width           =   255
         End
         Begin VB.TextBox txtlast 
            Height          =   375
            Left            =   345
            MaxLength       =   35
            TabIndex        =   1
            Top             =   1305
            Width           =   2895
         End
         Begin VB.CheckBox Checklast 
            Height          =   375
            Left            =   3480
            TabIndex        =   8
            Top             =   1320
            Width           =   255
         End
         Begin VB.CheckBox Checkid 
            Height          =   375
            Left            =   3480
            TabIndex        =   7
            Top             =   540
            Width           =   285
         End
         Begin VB.TextBox txtfirst 
            Height          =   375
            Left            =   345
            TabIndex        =   2
            Top             =   2040
            Width           =   2895
         End
         Begin VB.CheckBox CheckFirst 
            Caption         =   "Check1"
            Height          =   375
            Left            =   3480
            TabIndex        =   9
            Top             =   2040
            Width           =   255
         End
         Begin MSMask.MaskEdBox txtId 
            Height          =   390
            Left            =   360
            TabIndex        =   0
            Top             =   615
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   688
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "P-########"
            PromptChar      =   "_"
         End
         Begin Project1.chameleonButton srchbut2 
            Height          =   405
            Left            =   1500
            TabIndex        =   4
            Top             =   4350
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   714
            BTYPE           =   5
            TX              =   "&Perform"
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
            MICON           =   "SFdisDEL.frx":045E
            PICN            =   "SFdisDEL.frx":047A
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
            Caption         =   "Search by Year"
            Height          =   255
            Index           =   3
            Left            =   375
            TabIndex        =   16
            Top             =   2520
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Last Name"
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   15
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Patient No."
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   14
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "First Name"
            Height          =   255
            Index           =   2
            Left            =   345
            TabIndex        =   13
            Top             =   1815
            Width           =   1575
         End
      End
   End
End
Attribute VB_Name = "SFdisDEL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' the scanhead of windsor ** david brian rosal ***'


Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Add3DBorder Me
Add3DBorder txtId
Add3DBorder txtlast
Add3DBorder txtfirst
Add3DBorder txtyear

End Sub

Private Sub srchbut2_Click()
  Dim y As Variant
  Dim TMP_KEY As String
  Dim TmpFN As String
  Dim TmpLN, TmpYr, tmpid, testdate, NOID As String
  Dim Madz, Mbry As Boolean
  Dim X As Byte
  On Error Resume Next
  NOID = "Non-Med"
  testdate = ""
  search_sql = ""
  Public_Sql = ""
  TmpFN = ""
  TmpLN = ""
  tmpid = ""
  TmpYr = ""
  id = False
  Madz = False
  Mbry = False
  ln = False
  fn = False
  y = False
  X = False
  yr = False
     
 'Store Last Name in Temporary Variables
  tmpid = Trim(txtId.Text)
  TmpFN = Trim(txtfirst.Text)
  TmpLN = Trim(txtlast.Text)
  TmpYr = Trim(txtyear.Text)
 'ID no.
 If Checkid.Value = 1 Then
     If Len(tmpid) < 1 Then
         MsgBox "You must provide a value for the Patient No.", vbInformation, "Attention"
         id = False
         Exit Sub
     Else
       
        id = True
        Mbry = True
     End If
  End If
    
  'lastname :dbrymadz
If Checklast.Value = 1 Then
      If Len(TmpLN) < 1 Then
         MsgBox "You need to enter a value for Last Name.", vbInformation, "Attention"
         ln = False
         Exit Sub
      Else
         ln = True
      End If
End If
  'firstname :dbrymadz
If CheckFirst.Value = 1 Then
    If Len(TmpFN) < 1 Then
          MsgBox "You need to enter a value for first Name.", vbInformation, "Attention"
          fn = False
          Exit Sub
    Else
          fn = True
    End If
End If
        
 'year :dbrymadz
If Checkyear.Value = 1 Then
    If Len(TmpYr) < 1 Then
          MsgBox "You need to enter a value for Year.", vbInformation, "Attention"
          yr = False
          Exit Sub
    Else
          yr = True
    End If
End If
  
 If CNonMed.Value = 1 Then
          NONMED = True
 End If
  
 If CMED.Value = 1 Then
          PHMED = True
 End If
 
 If id = True And ln = False And fn = False And NONMED = False And PHMED = False Then
   search_sql = "SELECT patientno,dateD,Fdiagnose,rmbrd,labfee,Tmeds,Pfee,Philpay,Diff,idno,Plast,Pfirst,Pmi,Page,Pdiag,pdatec,Mlast,Mfirst,Mmi,timein,timeout,rmpay,labpay,medpay,pfpay FROM Pdischarged WHERE  patientno = '" & tmpid & "' ;"
   Madz = True
 ElseIf id = True And ln = False And fn = True And NONMED = False And PHMED = False Then
   search_sql = "SELECT patientno,dateD,Fdiagnose,rmbrd,labfee,Tmeds,Pfee,Philpay,Diff,idno,Plast,Pfirst,Pmi,Page,Pdiag,pdatec,Mlast,Mfirst,Mmi,timein,timeout,rmpay,labpay,medpay,pfpay FROM Pdischarged WHERE  patientno = '" & tmpid & "' and  Pfirst = '" & TmpFN & "';"
   Madz = True
 ElseIf id = True And ln = True And fn = False And NONMED = False And PHMED = False Then
   search_sql = "SELECT patientno,dateD,Fdiagnose,rmbrd,labfee,Tmeds,Pfee,Philpay,Diff,idno,Plast,Pfirst,Pmi,Page,Pdiag,pdatec,Mlast,Mfirst,Mmi,timein,timeout,rmpay,labpay,medpay,pfpay FROM Pdischarged WHERE  patientno = '" & tmpid & "' and PLast = '" & TmpLN & "' ;"
   Madz = True
 ElseIf id = True And ln = True And fn = True And NONMED = False And PHMED = False Then
     search_sql = "SELECT patientno,dateD,Fdiagnose,rmbrd,labfee,Tmeds,Pfee,Philpay,Diff,idno,Plast,Pfirst,Pmi,Page,Pdiag,pdatec,Mlast,Mfirst,Mmi,timein,timeout,rmpay,labpay,medpay,pfpay  FROM Pdischarged WHERE  patientno = '" & tmpid & "' and PLast = '" & TmpLN & "' and Pfirst = '" & TmpFN & "' ;"
     Madz = True
 ElseIf id = False And ln = True And fn = True And NONMED = False And PHMED = False Then
     search_sql = "SELECT patientno,dateD,Fdiagnose,rmbrd,labfee,Tmeds,Pfee,Philpay,Diff,idno,Plast,Pfirst,Pmi,Page,Pdiag,pdatec,Mlast,Mfirst,Mmi,timein,timeout,rmpay,labpay,medpay,pfpay FROM Pdischarged WHERE  PLast = '" & TmpLN & "' and Pfirst = '" & TmpFN & "' ;"
     Madz = True
 ElseIf id = False And ln = False And fn = True And NONMED = False And PHMED = False Then
    search_sql = "SELECT patientno,dateD,Fdiagnose,rmbrd,labfee,Tmeds,Pfee,Philpay,Diff,idno,Plast,Pfirst,Pmi,Page,Pdiag,pdatec,Mlast,Mfirst,Mmi,timein,timeout,rmpay,labpay,medpay,pfpay FROM Pdischarged WHERE Pfirst = '" & TmpFN & "' ;"
    Madz = True
 ElseIf id = False And ln = True And fn = False And NONMED = False And PHMED = False Then
    search_sql = "SELECT patientno,dateD,Fdiagnose,rmbrd,labfee,Tmeds,Pfee,Philpay,Diff,idno,Plast,Pfirst,Pmi,Page,Pdiag,pdatec,Mlast,Mfirst,Mmi,timein,timeout,rmpay,labpay,medpay,pfpay FROM Pdischarged WHERE PLAST = '" & TmpLN & "' ;"
    Madz = True
 '***** STAGE 2 :1
 ElseIf id = True And ln = False And fn = False And NONMED = True And PHMED = False Then
   search_sql = "SELECT patientno,dateD,Fdiagnose,rmbrd,labfee,Tmeds,Pfee,Philpay,Diff,idno,Plast,Pfirst,Pmi,Page,Pdiag,pdatec,Mlast,Mfirst,Mmi,timein,timeout,rmpay,labpay,medpay,pfpay FROM Pdischarged WHERE  patientno = '" & tmpid & "' AND TRIM(IDNO) = '" & NOID & "';"
   Madz = True
 ElseIf id = True And ln = False And fn = False And NONMED = False And PHMED = True Then
   search_sql = "SELECT patientno,dateD,Fdiagnose,rmbrd,labfee,Tmeds,Pfee,Philpay,Diff,idno,Plast,Pfirst,Pmi,Page,Pdiag,pdatec,Mlast,Mfirst,Mmi,timein,timeout,rmpay,labpay,medpay,pfpay FROM Pdischarged WHERE  patientno = '" & tmpid & "' AND TRIM(IDNO) <> '" & NOID & "';"
   Madz = True
'****2
ElseIf id = True And ln = False And fn = True And NONMED = True And PHMED = False Then
 search_sql = "SELECT patientno,dateD,Fdiagnose,rmbrd,labfee,Tmeds,Pfee,Philpay,Diff,idno,Plast,Pfirst,Pmi,Page,Pdiag,pdatec,Mlast,Mfirst,Mmi,timein,timeout,rmpay,labpay,medpay,pfpay FROM Pdischarged WHERE  patientno = '" & tmpid & "' and  Pfirst = '" & TmpFN & "' AND TRIM(IDNO) = '" & NOID & "';"
 Madz = True
ElseIf id = True And ln = False And fn = True And NONMED = False And PHMED = True Then
   search_sql = "SELECT patientno,dateD,Fdiagnose,rmbrd,labfee,Tmeds,Pfee,Philpay,Diff,idno,Plast,Pfirst,Pmi,Page,Pdiag,pdatec,Mlast,Mfirst,Mmi,timein,timeout,rmpay,labpay,medpay,pfpay FROM Pdischarged WHERE  patientno = '" & tmpid & "' and  Pfirst = '" & TmpFN & "'AND TRIM(IDNO) <> '" & NOID & "';"
   Madz = True
'*3
 ElseIf id = True And ln = True And fn = False And NONMED = True And PHMED = False Then
    search_sql = "SELECT patientno,dateD,Fdiagnose,rmbrd,labfee,Tmeds,Pfee,Philpay,Diff,idno,Plast,Pfirst,Pmi,Page,Pdiag,pdatec,Mlast,Mfirst,Mmi,timein,timeout,rmpay,labpay,medpay,pfpay FROM Pdischarged WHERE  patientno = '" & tmpid & "' and PLast = '" & TmpLN & "' AND TRIM(IDNO) = '" & NOID & "';"
 Madz = True
 ElseIf id = True And ln = True And fn = False And NONMED = False And PHMED = True Then
     search_sql = "SELECT patientno,dateD,Fdiagnose,rmbrd,labfee,Tmeds,Pfee,Philpay,Diff,idno,Plast,Pfirst,Pmi,Page,Pdiag,pdatec,Mlast,Mfirst,Mmi,timein,timeout,rmpay,labpay,medpay,pfpay FROM Pdischarged WHERE  patientno = '" & tmpid & "' and PLast = '" & TmpLN & "' AND TRIM(IDNO) <> '" & NOID & "';"
 Madz = True
'*4
 ElseIf id = True And ln = True And fn = True And NONMED = True And PHMED = False Then
      search_sql = "SELECT patientno,dateD,Fdiagnose,rmbrd,labfee,Tmeds,Pfee,Philpay,Diff,idno,Plast,Pfirst,Pmi,Page,Pdiag,pdatec,Mlast,Mfirst,Mmi,timein,timeout,rmpay,labpay,medpay,pfpay FROM Pdischarged WHERE  patientno = '" & tmpid & "' and PLast = '" & TmpLN & "' and Pfirst = '" & TmpFN & "' AND TRIM(IDNO) = '" & NOID & "';"
    Madz = True
 ElseIf id = True And ln = True And fn = True And NONMED = False And PHMED = True Then
      search_sql = "SELECT patientno,dateD,Fdiagnose,rmbrd,labfee,Tmeds,Pfee,Philpay,Diff,idno,Plast,Pfirst,Pmi,Page,Pdiag,pdatec,Mlast,Mfirst,Mmi,timein,timeout,rmpay,labpay,medpay,pfpay FROM Pdischarged WHERE  patientno = '" & tmpid & "' and PLast = '" & TmpLN & "' and Pfirst = '" & TmpFN & "' AND TRIM(IDNO) <> '" & NOID & "';"
   Madz = True
 '*5
  ElseIf id = False And ln = True And fn = True And NONMED = True And PHMED = False Then
     search_sql = "SELECT patientno,dateD,Fdiagnose,rmbrd,labfee,Tmeds,Pfee,Philpay,Diff,idno,Plast,Pfirst,Pmi,Page,Pdiag,pdatec,Mlast,Mfirst,Mmi,timein,timeout,rmpay,labpay,medpay,pfpay FROM Pdischarged WHERE  PLast = '" & TmpLN & "' and Pfirst = '" & TmpFN & "' AND TRIM(IDNO) = '" & NOID & "' ;"
     Madz = True
 
  ElseIf id = False And ln = True And fn = True And NONMED = False And PHMED = True Then
     search_sql = "SELECT patientno,dateD,Fdiagnose,rmbrd,labfee,Tmeds,Pfee,Philpay,Diff,idno,Plast,Pfirst,Pmi,Page,Pdiag,pdatec,Mlast,Mfirst,Mmi,timein,timeout,rmpay,labpay,medpay,pfpay  FROM Pdischarged WHERE  PLast = '" & TmpLN & "' and Pfirst = '" & TmpFN & "' AND TRIM(IDNO) <> '" & NOID & "' ;"
     Madz = True
 '*6
  ElseIf id = False And ln = False And fn = True And NONMED = True And PHMED = False Then
    search_sql = "SELECT patientno,dateD,Fdiagnose,rmbrd,labfee,Tmeds,Pfee,Philpay,Diff,idno,Plast,Pfirst,Pmi,Page,Pdiag,pdatec,Mlast,Mfirst,Mmi,timein,timeout,rmpay,labpay,medpay,pfpay FROM Pdischarged WHERE Pfirst = '" & TmpFN & "' and TRIM(IDNO) = '" & NOID & "' ;"
    Madz = True
  ElseIf id = False And ln = False And fn = True And NONMED = False And PHMED = True Then
    search_sql = "SELECT patientno,dateD,Fdiagnose,rmbrd,labfee,Tmeds,Pfee,Philpay,Diff,idno,Plast,Pfirst,Pmi,Page,Pdiag,pdatec,Mlast,Mfirst,Mmi,timein,timeout,rmpay,labpay,medpay,pfpay FROM Pdischarged WHERE Pfirst = '" & TmpFN & "' AND TRIM(IDNO) <> '" & NOID & "' ;"
    Madz = True
 '*7
 ElseIf id = False And ln = True And fn = False And NONMED = True And PHMED = False Then
      search_sql = "SELECT patientno,dateD,Fdiagnose,rmbrd,labfee,Tmeds,Pfee,Philpay,Diff,idno,Plast,Pfirst,Pmi,Page,Pdiag,pdatec,Mlast,Mfirst,Mmi,timein,timeout,rmpay,labpay,medpay,pfpay FROM Pdischarged WHERE PLAST = '" & TmpLN & "' and TRIM(IDNO) = '" & NOID & "';"
  Madz = True

 ElseIf id = False And ln = True And fn = False And NONMED = False And PHMED = True Then
   search_sql = "SELECT patientno,dateD,Fdiagnose,rmbrd,labfee,Tmeds,Pfee,Philpay,Diff,idno,Plast,Pfirst,Pmi,Page,Pdiag,pdatec,Mlast,Mfirst,Mmi,timein,timeout,rmpay,labpay,medpay,pfpay FROM Pdischarged WHERE PLAST = '" & TmpLN & "' AND TRIM(IDNO) <> '" & NOID & "';"
  Madz = True
 '*8
 ElseIf id = False And ln = False And fn = False And NONMED = True And PHMED = False Then
  search_sql = "SELECT patientno,dateD,Fdiagnose,rmbrd,labfee,Tmeds,Pfee,Philpay,Diff,idno,Plast,Pfirst,Pmi,Page,Pdiag,pdatec,Mlast,Mfirst,Mmi,timein,timeout,rmpay,labpay,medpay,pfpay FROM Pdischarged WHERE TRIM(IDNO) = '" & NOID & "';"
  Madz = True
 ElseIf id = False And ln = False And fn = False And NONMED = False And PHMED = True Then
  search_sql = "SELECT patientno,dateD,Fdiagnose,rmbrd,labfee,Tmeds,Pfee,Philpay,Diff,idno,Plast,Pfirst,Pmi,Page,Pdiag,pdatec,Mlast,Mfirst,Mmi,timein,timeout,rmpay,labpay,medpay,pfpay FROM Pdischarged WHERE  TRIM(IDNO) <> '" & NOID & "';"
    Madz = True
 '*9
 ElseIf id = False And ln = False And fn = False And NONMED = True And PHMED = True Then
  search_sql = "SELECT patientno,dateD,Fdiagnose,rmbrd,labfee,Tmeds,Pfee,Philpay,Diff,idno,Plast,Pfirst,Pmi,Page,Pdiag,pdatec,Mlast,Mfirst,Mmi,timein,timeout,rmpay,labpay,medpay,pfpay FROM Pdischarged ;"
    Madz = True
 ElseIf id = False And ln = False And fn = False And NONMED = False And PHMED = False Then
    Madz = False
 End If





If yr = True Then
    search_sql = "SELECT patientno,dateD,Fdiagnose,rmbrd,labfee,Tmeds,Pfee,Philpay,Diff,idno,Plast,Pfirst,Pmi,Page,Pdiag,pdatec,Mlast,Mfirst,Mmi,timein,timeout,rmpay,labpay,medpay,pfpay  FROM Pdischarged WHERE format$(dateD,'yyyy') like '" & TmpYr & "';"
    Madz = True
End If
   
 '/********** displaying of values to the form: scanhead *******
  Set TRS = DBMain.OpenRecordset(search_sql)
  'yaH im searchin?
 PHdisDelete.txtDISdate.Text = Format(TRS.Fields(1), "mm/dd/yyyy")
 PHdisDelete.txtDISdiag.Text = TRS.Fields(2)
 PHdisDelete.txtDISrm.Text = TRS.Fields(3)
 PHdisDelete.txtDISlab.Text = TRS.Fields(4)
 PHdisDelete.txtDISmed.Text = TRS.Fields(5)
 PHdisDelete.txtDISpf.Text = TRS.Fields(6)
 PHdisDelete.txtDISpaid.Text = TRS.Fields(7)
 PHdisDelete.txtmidno.Text = TRS.Fields(9)
 PHdisDelete.txtpno.Text = TRS.Fields(0)
 PHdisDelete.txtplast.Text = TRS.Fields(10)
 PHdisDelete.txtpfirst.Text = TRS.Fields(11)
 PHdisDelete.txtpmi.Text = TRS.Fields(12)
 PHdisDelete.txtpage.Text = TRS.Fields(13)
 PHdisDelete.txtpAd.Text = TRS.Fields(14)
 PHdisDelete.txtdate.Text = Format(TRS.Fields(15), "mm/dd/yyyy")
 PHdisDelete.txtctime.Text = Format(TRS.Fields(19), "medium time")
 PHdisDelete.txttime.Text = Format(TRS.Fields(20), "medium time")
 PHdisDelete.txtPAYrm.Text = TRS.Fields(21)
 PHdisDelete.txtPAYlab.Text = TRS.Fields(22)
 PHdisDelete.txtPAYmed.Text = TRS.Fields(23)
 PHdisDelete.txtPAYpf.Text = TRS.Fields(24)
 
 
 PHdisDelete.txtDIStot.Text = Format$(Val(TRS.Fields(3)) + Val(TRS.Fields(4)) + Val(TRS.Fields(5)) + Val(TRS.Fields(6)), "###,###,###.00")
 PHdisDelete.lblDISdiff.Caption = "Php " + Format$(Val(TRS.Fields(8)), "###,###,###.00")
        
  '***************************
     If Madz = True Then
        Set TRS = DBMain.OpenRecordset(search_sql)
         PHdisDelete.ListView2.ListItems.Clear
         If TRS.RecordCount > 0 Then
             TRS.Fields.Refresh
             Do While Not TRS.EOF
                  Set List = PHdisDelete.ListView2.ListItems.Add(, , TRS.Fields(0))
                        List.SubItems(1) = TRS.Fields(9)
                        List.SubItems(2) = Format(TRS.Fields(1), "mm/dd/yyyy")
                        List.SubItems(3) = TRS.Fields(10)
                        List.SubItems(4) = TRS.Fields(11)
                        List.SubItems(5) = TRS.Fields(12)
                        List.SubItems(6) = TRS.Fields(13)
                        List.SubItems(7) = TRS.Fields(2)
                        TRS.MoveNext
                     Loop
           'the scanhead of windsor
           Else
           MsgBox "Search Data does not exist!, Click OK to Reset!", vbCritical, "Attention"
           clearDdel
           PHdisDelete.cmdreset.Enabled = True
           PHdisDelete.DisplayLst
           End If
       
     Else
         MsgBox "Nothing was Performed!, Click OK to Reset!", vbCritical, "Attention"
         clearDdel
         PHdisDelete.DisplayLst
         PHdisDelete.cmdreset.Enabled = True
         End If
 Unload Me
End Sub
' the scanhead of windsor ** david brian rosal ***'
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

Private Sub txtfirst_GotFocus()
Checkyear.Value = 0
txtyear.Text = "____"
End Sub

Private Sub txtId_GotFocus()
With txtId
.SelStart = 0
.SelLength = Len(txtId.Text)
End With
Checkyear.Value = 0
txtyear.Text = "____"
End Sub

Private Sub txtid_KeyPress(KeyAscii As Integer)
Dim madzbry As String
 madzbry = "P-0123456789"
          If KeyAscii > 26 Then
            If InStr(madzbry, Chr(KeyAscii)) = 0 Then
              KeyAscii = 0
            End If
          End If
      
Checkid.Value = 1
End Sub

Private Sub txtlast_GotFocus()
Checkyear.Value = 0
txtyear.Text = "____"
End Sub

Private Sub txtlast_KeyPress(KeyAscii As Integer)
Dim madzbry As String
  madzbry = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ ."
          If KeyAscii > 26 Then
            If InStr(madzbry, Chr(KeyAscii)) = 0 Then
              KeyAscii = 0
            End If
          End If
Checklast.Value = 1
End Sub

Private Sub txtfirst_KeyPress(KeyAscii As Integer)
Dim madzbry As String
  madzbry = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ ."
          If KeyAscii > 26 Then
            If InStr(madzbry, Chr(KeyAscii)) = 0 Then
              KeyAscii = 0
            End If
          End If
CheckFirst.Value = 1
End Sub

Private Sub txtyear_GotFocus()
With txtyear
.SelStart = 0
.SelLength = Len(txtyear.Text)
End With
End Sub

Private Sub txtyear_KeyPress(KeyAscii As Integer)
Checkyear.Value = 1
CheckFirst.Value = 0
Checklast.Value = 0
Checkid.Value = 0
txtfirst.Text = ""
txtlast.Text = ""
txtId.Text = "P-________"
End Sub
