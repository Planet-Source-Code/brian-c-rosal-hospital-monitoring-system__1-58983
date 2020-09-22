VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form SFDisP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search Discharged Patient(s)"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4395
   Icon            =   "SFDisP.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   4395
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5055
      Left            =   90
      TabIndex        =   9
      Top             =   75
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   8916
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Search Area"
      TabPicture(0)   =   "SFDisP.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame Frame1 
         Height          =   4455
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   3930
         Begin VB.Frame Frame2 
            Caption         =   "Search By:"
            ForeColor       =   &H00FF0000&
            Height          =   975
            Left            =   285
            TabIndex        =   14
            Top             =   2640
            Width           =   3345
            Begin VB.CheckBox CMed 
               Caption         =   "Check1"
               Height          =   375
               Left            =   1770
               TabIndex        =   8
               Top             =   330
               Width           =   255
            End
            Begin VB.CheckBox CNonMed 
               Caption         =   "Check1"
               Height          =   375
               Left            =   195
               TabIndex        =   7
               Top             =   330
               Width           =   255
            End
            Begin VB.Label Label1 
               Caption         =   "PH MED only"
               Height          =   255
               Index           =   4
               Left            =   2070
               TabIndex        =   16
               Top             =   420
               Width           =   990
            End
            Begin VB.Label Label1 
               Caption         =   "NON-MED only"
               Height          =   255
               Index           =   3
               Left            =   480
               TabIndex        =   15
               Top             =   405
               Width           =   1095
            End
         End
         Begin VB.CheckBox CheckFirst 
            Caption         =   "Check1"
            Height          =   375
            Left            =   3480
            TabIndex        =   6
            Top             =   2040
            Width           =   255
         End
         Begin VB.TextBox txtfirst 
            Height          =   375
            Left            =   345
            TabIndex        =   2
            Top             =   2040
            Width           =   2895
         End
         Begin VB.CheckBox Checkid 
            Height          =   375
            Left            =   3480
            TabIndex        =   4
            Top             =   540
            Width           =   285
         End
         Begin VB.CheckBox Checklast 
            Height          =   375
            Left            =   3480
            TabIndex        =   5
            Top             =   1320
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
            Left            =   1440
            TabIndex        =   3
            Top             =   3840
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
            MICON           =   "SFDisP.frx":045E
            PICN            =   "SFDisP.frx":047A
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
            Caption         =   "First Name"
            Height          =   255
            Index           =   2
            Left            =   345
            TabIndex        =   13
            Top             =   1815
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Patient No."
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   12
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Last Name"
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   11
            Top             =   1080
            Width           =   1575
         End
      End
   End
End
Attribute VB_Name = "SFDisP"
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


End Sub

Private Sub srchbut2_Click()
  Dim y As Variant
  Dim TMP_KEY As String
  Dim TmpFN As String
  Dim TmpLN, TmpYr, tmpid, testdate, PHMED, NONMED, NOID As String
  Dim Madz, Mbry As Boolean
  Dim X As Byte
  On Error Resume Next
  NOID = "Non-Med"
  search_sql = ""
  Public_Sql = ""
  TmpFN = ""
  TmpLN = ""
  tmpid = ""
  id = False
  Madz = False
  Mbry = False
  ln = False
  fn = False
  y = False
  X = False
  NONMED = False
  PHMED = False
   
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
     search_sql = "SELECT patientno,dateD,Fdiagnose,rmbrd,labfee,Tmeds,Pfee,Philpay,Diff,idno,Plast,Pfirst,Pmi,Page,Pdiag,pdatec,Mlast,Mfirst,Mmi,timein,timeout,rmpay,labpay,medpay,pfpay FROM Pdischarged WHERE  patientno = '" & tmpid & "' and PLast = '" & TmpLN & "' and Pfirst = '" & TmpFN & "' ;"
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
     search_sql = "SELECT patientno,dateD,Fdiagnose,rmbrd,labfee,Tmeds,Pfee,Philpay,Diff,idno,Plast,Pfirst,Pmi,Page,Pdiag,pdatec,Mlast,Mfirst,Mmi,timein,timeout,rmpay,labpay,medpay,pfpay FROM Pdischarged WHERE  PLast = '" & TmpLN & "' and Pfirst = '" & TmpFN & "' AND TRIM(IDNO) <> '" & NOID & "' ;"
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













 
   
 '/********** displaying of values to the form: scanhead *******
  Set TRS = DBMain.OpenRecordset(search_sql)
  'yaH im searchin?
 PHdischarged.txtDISdate.Text = Format(TRS.Fields(1), "mm/dd/yyyy")
 PHdischarged.txtDISdiag.Text = TRS.Fields(2)
 PHdischarged.txtDISrm.Text = TRS.Fields(3)
 PHdischarged.txtDISlab.Text = TRS.Fields(4)
 PHdischarged.txtDISmed.Text = TRS.Fields(5)
 PHdischarged.txtDISpf.Text = TRS.Fields(6)
 PHdischarged.txtDISpaid.Text = TRS.Fields(7)
 PHdischarged.txtmidno.Text = TRS.Fields(9)
 PHdischarged.txtpno.Text = TRS.Fields(0)
 PHdischarged.txtplast.Text = TRS.Fields(10)
 PHdischarged.txtpfirst.Text = TRS.Fields(11)
 PHdischarged.txtpmi.Text = TRS.Fields(12)
 PHdischarged.txtpage.Text = TRS.Fields(13)
 PHdischarged.txtpAd.Text = TRS.Fields(14)
 PHdischarged.txtdate.Text = Format(TRS.Fields(15), "mm/dd/yyyy")
 PHdischarged.txtMlast.Text = TRS.Fields(16)
 PHdischarged.txtMfirst.Text = TRS.Fields(17)
 PHdischarged.txtMmi.Text = TRS.Fields(18)
 PHdischarged.txtctime.Text = Format(TRS.Fields(19), "medium time")
 PHdischarged.txttime.Text = Format(TRS.Fields(20), "medium time")
 PHdischarged.txtPAYrm.Text = TRS.Fields(21)
 PHdischarged.txtPAYlab.Text = TRS.Fields(22)
 PHdischarged.txtPAYmed.Text = TRS.Fields(23)
 PHdischarged.txtPAYpf.Text = TRS.Fields(24)
  
 PHdischarged.txtDIStot.Text = Format$(Val(TRS.Fields(3)) + Val(TRS.Fields(4)) + Val(TRS.Fields(5)) + Val(TRS.Fields(6)), "###,###,###.00")
 PHdischarged.lblDISdiff.Caption = "Php " + Format$(Val(TRS.Fields(8)), "###,###,###.00")
        
  '***************************
     If Madz = True Then
        Set TRS = DBMain.OpenRecordset(search_sql)
         PHdischarged.ListView2.ListItems.Clear
         If TRS.RecordCount > 0 Then
             TRS.Fields.Refresh
             Do While Not TRS.EOF
                  Set List = PHdischarged.ListView2.ListItems.Add(, , TRS.Fields(0))
                        List.SubItems(1) = TRS.Fields(9)
                        List.SubItems(2) = Format(TRS.Fields(1), "mm/dd/yyyy")
                        List.SubItems(3) = TRS.Fields(10)
                        List.SubItems(4) = TRS.Fields(11)
                        List.SubItems(5) = TRS.Fields(12)
                        List.SubItems(6) = TRS.Fields(13)
                        List.SubItems(7) = TRS.Fields(2)
                        TRS.MoveNext
                     Loop
           PHdischarged.cmdmod.Enabled = True
           'the scanhead of windsor
           Else
           MsgBox "Search Data does not exist, Click OK to Reset!", vbCritical, "Attention"
           PHdischarged.cmdreset.Enabled = True
           '***************
           PHdischarged.DisplayLstx
           PHdischarged.DisplayLstx2
           clearD
           PHdischarged.ListView1.Enabled = True
           PHdischarged.ListView2.Enabled = True
           PHdischarged.LBLDIS.Caption = "  "
           PHdischarged.cmdexit.SetFocus
           PHdischarged.cmdmod.Enabled = False
           '***************
         End If
       
     Else
         MsgBox "Nothing was Performed, Click OK to Reset!", vbCritical, "Attention"
         PHdischarged.cmdreset.Enabled = True
         '***************
           PHdischarged.DisplayLstx
           PHdischarged.DisplayLstx2
            clearD
           PHdischarged.ListView1.Enabled = True
           PHdischarged.ListView2.Enabled = True
           PHdischarged.LBLDIS.Caption = "  "
           PHdischarged.cmdexit.SetFocus
           PHdischarged.cmdmod.Enabled = False
           '***************
             
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


Private Sub txtId_GotFocus()
With txtId
.SelStart = 0
.SelLength = Len(txtId.Text)
End With
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


