VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form SrchFormC 
   Caption         =   "Search Confined Patient"
   ClientHeight    =   5235
   ClientLeft      =   4110
   ClientTop       =   2295
   ClientWidth     =   4440
   Icon            =   "SrchFormC.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5235
   ScaleWidth      =   4440
   Begin TabDlg.SSTab SSTab1 
      Height          =   4995
      Left            =   105
      TabIndex        =   8
      Top             =   60
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   8811
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Search Area"
      TabPicture(0)   =   "SrchFormC.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame Frame1 
         Height          =   4320
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   3930
         Begin VB.Frame Frame2 
            Caption         =   "Search By:"
            ForeColor       =   &H00FF0000&
            Height          =   975
            Left            =   360
            TabIndex        =   12
            Top             =   2610
            Width           =   3345
            Begin VB.CheckBox CNonMed 
               Caption         =   "Check1"
               Height          =   375
               Left            =   195
               TabIndex        =   14
               Top             =   330
               Width           =   255
            End
            Begin VB.CheckBox CMed 
               Caption         =   "Check1"
               Height          =   375
               Left            =   1770
               TabIndex        =   13
               Top             =   330
               Width           =   255
            End
            Begin VB.Label Label1 
               Caption         =   "NON-MED only"
               Height          =   255
               Index           =   3
               Left            =   465
               TabIndex        =   16
               Top             =   405
               Width           =   1095
            End
            Begin VB.Label Label1 
               Caption         =   "PH MED only"
               Height          =   255
               Index           =   4
               Left            =   2070
               TabIndex        =   15
               Top             =   420
               Width           =   990
            End
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
            TabIndex        =   6
            Top             =   1320
            Width           =   255
         End
         Begin VB.CheckBox Checkid 
            Height          =   375
            Left            =   3480
            TabIndex        =   5
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
            TabIndex        =   7
            Top             =   2040
            Width           =   255
         End
         Begin Project1.chameleonButton srchbut2 
            Height          =   405
            Left            =   1440
            TabIndex        =   3
            Top             =   3765
            Width           =   2295
            _extentx        =   4048
            _extenty        =   714
            btype           =   5
            tx              =   "&Perform"
            enab            =   -1  'True
            font            =   "SrchFormC.frx":045E
            coltype         =   2
            focusr          =   -1  'True
            bcol            =   12632256
            bcolo           =   12632256
            fcol            =   0
            fcolo           =   255
            mcol            =   12632256
            mptr            =   1
            micon           =   "SrchFormC.frx":0482
            picn            =   "SrchFormC.frx":04A0
            umcol           =   -1  'True
            soft            =   0   'False
            picpos          =   0
            ngrey           =   0   'False
            fx              =   1
            hand            =   0   'False
            check           =   0   'False
            value           =   0   'False
         End
         Begin VB.Label Label1 
            Caption         =   "Last Name"
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   4
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Patient No."
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   11
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "First Name"
            Height          =   255
            Index           =   2
            Left            =   345
            TabIndex        =   10
            Top             =   1815
            Width           =   1575
         End
      End
   End
End
Attribute VB_Name = "SrchFormC"
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
  Dim TmpLN, tmpid, testdate, NOID, NONMED, PHMED As String
  Dim Madz, Mbry As Boolean
  Dim X As Byte
  On Error Resume Next
  testdate = ""
  NOID = ""
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
  NOMED = False
  PHMED = False
 'Store Last Name in Temporary Variables
  tmpid = Trim(txtId.Text)
  TmpFN = Trim(txtfirst.Text)
  TmpLN = Trim(txtlast.Text)
   
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
         MsgBox "You need to enter a value for Last Name", vbInformation, "Attention"
         ln = False
         Exit Sub
      Else
         ln = True
      End If
End If
  'firstname :dbrymadz
If CheckFirst.Value = 1 Then
    If Len(TmpFN) < 1 Then
          MsgBox "You need to enter a value for first Name", vbInformation, "Attention"
          fn = False
          Exit Sub
    Else
          fn = True
    End If
End If
        
  
 If CNonMed.Value = 1 Then
          NONMED = True
 End If
  
 If CMed.Value = 1 Then
          PHMED = True
 End If
  
  
 If id = True And ln = False And fn = False And NONMED = False And PHMED = False Then
   search_sql = "SELECT patientno,idno,datec,plastname,pfirstname,pmi,page,padiagnose,mlast,mfirst,mmi,TIMEIN FROM Patient WHERE  patientno Like '*" & tmpid & "*' AND ISNULL(dateD) = TRUE;"
   Madz = True
 ElseIf id = True And ln = False And fn = True And NONMED = False And PHMED = False Then
   search_sql = "SELECT patientno,idno,datec,plastname,pfirstname,pmi,page,padiagnose,mlast,mfirst,mmi,TIMEIN FROM Patient WHERE  patientno Like '*" & tmpid & "*' and  Pfirstname like '" & TmpFN & "' AND ISNULL(dateD) = TRUE;"
   Madz = True
 ElseIf id = True And ln = True And fn = False And NONMED = False And PHMED = False Then
   search_sql = "SELECT patientno,idno,datec,plastname,pfirstname,pmi,page,padiagnose,mlast,mfirst,mmi,TIMEIN FROM Patient WHERE  patientno Like '*" & tmpid & "*' and PLastname Like '*" & TmpLN & "*' AND ISNULL(dateD) = TRUE;"
   Madz = True
 ElseIf id = True And ln = True And fn = True And NONMED = False And PHMED = False Then
     search_sql = "SELECT patientno,idno,datec,plastname,pfirstname,pmi,page,padiagnose,mlast,mfirst,mmi,TIMEIN FROM Patient WHERE  patientno Like '*" & tmpid & "*' and PLastname Like '*" & TmpLN & "*' and Pfirstname like '" & TmpFN & "' AND ISNULL(dateD) = TRUE;"
     Madz = True
 ElseIf id = False And ln = True And fn = True And NONMED = False And PHMED = False Then
     search_sql = "SELECT patientno,idno,datec,plastname,pfirstname,pmi,page,padiagnose,mlast,mfirst,mmi,TIMEIN FROM Patient WHERE  PLastname Like '*" & TmpLN & "*' and Pfirstname like '" & TmpFN & "' AND ISNULL(dateD) = TRUE;"
     Madz = True
 ElseIf id = False And ln = False And fn = True And NONMED = False And PHMED = False Then
    search_sql = "SELECT patientno,idno,datec,plastname,pfirstname,pmi,page,padiagnose,mlast,mfirst,mmi,TIMEIN FROM Patient WHERE Pfirstname like '" & TmpFN & "' AND ISNULL(dateD) = TRUE;"
    Madz = True
 ElseIf id = False And ln = True And fn = False And NONMED = False And PHMED = False Then
    search_sql = "SELECT patientno,idno,datec,plastname,pfirstname,pmi,page,padiagnose,mlast,mfirst,mmi,TIMEIN FROM Patient WHERE PLASTname like '" & TmpLN & "' AND ISNULL(dateD) = TRUE;"
    Madz = True
 '************* STAGE 2
 '*
 ElseIf id = True And ln = False And fn = False And NONMED = True And PHMED = False Then
   search_sql = "SELECT patientno,idno,datec,plastname,pfirstname,pmi,page,padiagnose,mlast,mfirst,mmi,TIMEIN FROM Patient WHERE  patientno Like '*" & tmpid & "*' AND ISNULL(dateD) = TRUE AND ISNULL(IDNO)=TRUE ;"
   Madz = True
 ElseIf id = True And ln = False And fn = False And NONMED = False And PHMED = True Then
   search_sql = "SELECT patientno,idno,datec,plastname,pfirstname,pmi,page,padiagnose,mlast,mfirst,mmi,TIMEIN FROM Patient WHERE  patientno Like '*" & tmpid & "*' AND ISNULL(dateD) = TRUE AND ISNULL(IDNO)=FALSE ;"
   Madz = True
 '*
 ElseIf id = True And ln = False And fn = True And NONMED = True And PHMED = False Then
   search_sql = "SELECT patientno,idno,datec,plastname,pfirstname,pmi,page,padiagnose,mlast,mfirst,mmi,TIMEIN FROM Patient WHERE  patientno Like '*" & tmpid & "*' and  Pfirstname like '" & TmpFN & "' AND ISNULL(dateD) = TRUE AND ISNULL(IDNO)=TRUE;"
   Madz = True
 ElseIf id = True And ln = False And fn = True And NONMED = False And PHMED = True Then
   search_sql = "SELECT patientno,idno,datec,plastname,pfirstname,pmi,page,padiagnose,mlast,mfirst,mmi,TIMEIN FROM Patient WHERE  patientno Like '*" & tmpid & "*' and  Pfirstname like '" & TmpFN & "' AND ISNULL(dateD) = TRUE AND ISNULL(IDNO)=FALSE;"
   Madz = True
 '*
 ElseIf id = True And ln = True And fn = False And NONMED = True And PHMED = False Then
   search_sql = "SELECT patientno,idno,datec,plastname,pfirstname,pmi,page,padiagnose,mlast,mfirst,mmi,TIMEIN FROM Patient WHERE  patientno Like '*" & tmpid & "*' and PLastname Like '*" & TmpLN & "*' AND ISNULL(dateD) = TRUE AND ISNULL(IDNO)=TRUE;"
   Madz = True
 ElseIf id = True And ln = True And fn = False And NONMED = False And PHMED = True Then
   search_sql = "SELECT patientno,idno,datec,plastname,pfirstname,pmi,page,padiagnose,mlast,mfirst,mmi,TIMEIN FROM Patient WHERE  patientno Like '*" & tmpid & "*' and PLastname Like '*" & TmpLN & "*' AND ISNULL(dateD) = TRUE AND ISNULL(IDNO)=FALSE;"
   Madz = True
 '*
 ElseIf id = True And ln = True And fn = True And NONMED = True And PHMED = False Then
     search_sql = "SELECT patientno,idno,datec,plastname,pfirstname,pmi,page,padiagnose,mlast,mfirst,mmi,TIMEIN FROM Patient WHERE  patientno Like '*" & tmpid & "*' and PLastname Like '*" & TmpLN & "*' and Pfirstname like '" & TmpFN & "' AND ISNULL(dateD) = TRUE AND ISNULL(IDNO)=TRUE;"
     Madz = True
 ElseIf id = True And ln = True And fn = True And NONMED = False And PHMED = True Then
     search_sql = "SELECT patientno,idno,datec,plastname,pfirstname,pmi,page,padiagnose,mlast,mfirst,mmi,TIMEIN FROM Patient WHERE  patientno Like '*" & tmpid & "*' and PLastname Like '*" & TmpLN & "*' and Pfirstname like '" & TmpFN & "' AND ISNULL(dateD) = TRUE AND ISNULL(IDNO)=FALSE;"
     Madz = True
 '*
 ElseIf id = False And ln = True And fn = True And NONMED = True And PHMED = False Then
     search_sql = "SELECT patientno,idno,datec,plastname,pfirstname,pmi,page,padiagnose,mlast,mfirst,mmi,TIMEIN FROM Patient WHERE  PLastname Like '*" & TmpLN & "*' and Pfirstname like '" & TmpFN & "' AND ISNULL(dateD) = TRUE AND ISNULL(IDNO)=TRUE;"
     Madz = True
ElseIf id = False And ln = True And fn = True And NONMED = False And PHMED = True Then
     search_sql = "SELECT patientno,idno,datec,plastname,pfirstname,pmi,page,padiagnose,mlast,mfirst,mmi,TIMEIN FROM Patient WHERE  PLastname Like '*" & TmpLN & "*' and Pfirstname like '" & TmpFN & "' AND ISNULL(dateD) = TRUE AND ISNULL(IDNO)=FALSE;"
     Madz = True
'*
 ElseIf id = False And ln = False And fn = True And NONMED = True And PHMED = False Then
    search_sql = "SELECT patientno,idno,datec,plastname,pfirstname,pmi,page,padiagnose,mlast,mfirst,mmi,TIMEIN FROM Patient WHERE Pfirstname like '" & TmpFN & "' AND ISNULL(dateD) = TRUE AND ISNULL(IDNO)=TRUE;"
    Madz = True
 ElseIf id = False And ln = False And fn = True And NONMED = False And PHMED = True Then
    search_sql = "SELECT patientno,idno,datec,plastname,pfirstname,pmi,page,padiagnose,mlast,mfirst,mmi,TIMEIN FROM Patient WHERE Pfirstname like '" & TmpFN & "' AND ISNULL(dateD) = TRUE AND ISNULL(IDNO)=FALSE;"
    Madz = True
 '*
 ElseIf id = False And ln = True And fn = False And NONMED = True And PHMED = False Then
    search_sql = "SELECT patientno,idno,datec,plastname,pfirstname,pmi,page,padiagnose,mlast,mfirst,mmi,TIMEIN FROM Patient WHERE PLASTname like '" & TmpLN & "' AND ISNULL(dateD) = TRUE AND ISNULL(IDNO)=TRUE;"
    Madz = True

 ElseIf id = False And ln = True And fn = False And NONMED = False And PHMED = True Then
    search_sql = "SELECT patientno,idno,datec,plastname,pfirstname,pmi,page,padiagnose,mlast,mfirst,mmi,TIMEIN FROM Patient WHERE PLASTname like '" & TmpLN & "' AND ISNULL(dateD) = TRUE AND ISNULL(IDNO)=FALSE;"
    Madz = True
 '*
 ElseIf id = False And ln = False And fn = False And NONMED = True And PHMED = False Then
    search_sql = "SELECT patientno,idno,datec,plastname,pfirstname,pmi,page,padiagnose,mlast,mfirst,mmi,TIMEIN FROM Patient WHERE  ISNULL(dateD) = TRUE AND ISNULL(IDNO)=TRUE;"
    Madz = True
 ElseIf id = False And ln = False And fn = False And NONMED = False And PHMED = True Then
    search_sql = "SELECT patientno,idno,datec,plastname,pfirstname,pmi,page,padiagnose,mlast,mfirst,mmi,TIMEIN FROM Patient WHERE  ISNULL(dateD) = TRUE AND ISNULL(IDNO)=FALSE;"
    Madz = True
 '*
 ElseIf id = False And ln = False And fn = False And NONMED = True And PHMED = True Then
    search_sql = "SELECT patientno,idno,datec,plastname,pfirstname,pmi,page,padiagnose,mlast,mfirst,mmi,TIMEIN FROM Patient WHERE  ISNULL(dateD) = TRUE ;"
    Madz = True
 ElseIf id = False And ln = False And fn = False And NONMED = False And PHMED = False Then
    Madz = False
 End If


   
 '/********** displaying of values to the form: scanhead *******
  Set TRS = DBMain.OpenRecordset(search_sql)
   If IsNull(TRS.Fields(1)) = True Then
       PHdischarged.txtMidno.Text = "Non-Med"
       PHdischarged.txtMlast.Text = "--------------------"
       PHdischarged.txtMfirst.Text = "--------------------"
       PHdischarged.txtMmi.Text = "----"
   Else
      PHdischarged.txtMidno.Text = TRS.Fields(1)
      PHdischarged.txtMlast.Text = TRS.Fields(8)
      PHdischarged.txtMfirst.Text = TRS.Fields(9)
      PHdischarged.txtMmi.Text = TRS.Fields(10)
   End If
  PHdischarged.txtpno.Text = TRS.Fields(0)
  PHdischarged.txtplast.Text = TRS.Fields(3)
  PHdischarged.txtpfirst.Text = TRS.Fields(4)
  PHdischarged.txtpmi.Text = TRS.Fields(5)
  PHdischarged.txtpage.Text = TRS.Fields(6)
  PHdischarged.txtpAd.Text = TRS.Fields(7)
  PHdischarged.txtdate.Text = Format(TRS.Fields(2), "mm/dd/yyyy")
  PHdischarged.txtctime.Text = Format(TRS.Fields(11), "mEDIUM TIME")
 'yaH im searchin?
        
     If Madz = True Then
        Set TRS = DBMain.OpenRecordset(search_sql)
         PHdischarged.ListView1.ListItems.Clear
         If TRS.RecordCount > 0 Then
            TRS.Fields.Refresh
             Do While Not TRS.EOF
                  Set List = PHdischarged.ListView1.ListItems.Add(, , TRS.Fields(0))
                       If IsNull(TRS.Fields(1)) = True Then
                          List.SubItems(1) = "Non-Med"
                        Else
                          List.SubItems(1) = TRS.Fields(1)
                        End If
                        List.SubItems(2) = Format(TRS.Fields(2), "mm/dd/yyyy") + " " + Format(TRS.Fields(11), "MEDIUM TIME")
                        List.SubItems(3) = TRS.Fields(3)
                        List.SubItems(4) = TRS.Fields(4)
                        List.SubItems(5) = TRS.Fields(5)
                        List.SubItems(6) = TRS.Fields(6)
                        List.SubItems(7) = TRS.Fields(7)
                        TRS.MoveNext
                     Loop
           '&H80000000&]
           Else
           MsgBox "Search Data does not exist, Click OK to Reset!", vbCritical, "Attention"
           clearD
           PHdischarged.cmdreset.Enabled = True
           '**********
           PHdischarged.DisplayLstx
           PHdischarged.DisplayLstx2
           PHdischarged.ListView1.Enabled = True
           PHdischarged.ListView2.Enabled = True
           PHdischarged.LBLDIS.Caption = "  "
           PHdischarged.cmdExit.SetFocus
           PHdischarged.cmdmod.Enabled = False
           '**********
           End If
       
     Else
         MsgBox "Nothing was Performed, Click OK to Reset!", vbCritical, "Attention"
         clearD
           PHdischarged.cmdreset.Enabled = True
           '**********
           PHdischarged.DisplayLstx
           PHdischarged.DisplayLstx2
           PHdischarged.ListView1.Enabled = True
           PHdischarged.ListView2.Enabled = True
           PHdischarged.LBLDIS.Caption = "  "
           PHdischarged.cmdExit.SetFocus
           PHdischarged.cmdmod.Enabled = False
           '**********
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



