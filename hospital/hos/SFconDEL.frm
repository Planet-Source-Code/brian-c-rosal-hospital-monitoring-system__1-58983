VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form SFconDEL 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search Patient"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4320
   Icon            =   "SFconDEL.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   4320
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   4935
      Left            =   45
      TabIndex        =   4
      Top             =   45
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   8705
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Search Area"
      TabPicture(0)   =   "SFconDEL.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame Frame1 
         Height          =   4320
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   3960
         Begin VB.Frame Frame2 
            Caption         =   "Search By:"
            ForeColor       =   &H00FF0000&
            Height          =   975
            Left            =   315
            TabIndex        =   14
            Top             =   2565
            Width           =   3345
            Begin VB.CheckBox CMed 
               Caption         =   "Check1"
               Height          =   375
               Left            =   1800
               TabIndex        =   10
               Top             =   330
               Width           =   255
            End
            Begin VB.CheckBox CNonMed 
               Caption         =   "Check1"
               Height          =   375
               Left            =   195
               TabIndex        =   9
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
               Left            =   465
               TabIndex        =   15
               Top             =   405
               Width           =   1095
            End
         End
         Begin VB.CheckBox CheckFirst 
            Caption         =   "Check1"
            Height          =   375
            Left            =   3480
            TabIndex        =   8
            Top             =   2040
            Width           =   255
         End
         Begin VB.TextBox txtfirst 
            Height          =   375
            Left            =   360
            TabIndex        =   2
            Top             =   2040
            Width           =   2895
         End
         Begin VB.CheckBox Checkid 
            Height          =   375
            Left            =   3480
            TabIndex        =   7
            Top             =   540
            Width           =   285
         End
         Begin VB.CheckBox Checklast 
            Height          =   375
            Left            =   3480
            TabIndex        =   6
            Top             =   1320
            Width           =   255
         End
         Begin VB.TextBox txtlast 
            Height          =   375
            Left            =   360
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
            Left            =   1515
            TabIndex        =   3
            Top             =   3825
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
            MICON           =   "SFconDEL.frx":045E
            PICN            =   "SFconDEL.frx":047A
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
Attribute VB_Name = "SFconDEL"
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
  Dim TmpLN, tmpid, testdate As String
  Dim Madz, Mbry As Boolean
  Dim X As Byte
  On Error Resume Next
  testdate = ""
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
  
 If CMED.Value = 1 Then
          PHMED = True
 End If
  
  
 If id = True And ln = False And fn = False And NONMED = False And PHMED = False Then
   search_sql = "SELECT patientno,idno,datec,plastname,pfirstname,pmi,page,padiagnose,mlast,mfirst,mmi FROM Patient WHERE  patientno Like '*" & tmpid & "*' AND format(dateD,'mm/dd/yyyy') = '" & testdate & "';"
   Madz = True
 ElseIf id = True And ln = False And fn = True And NONMED = False And PHMED = False Then
   search_sql = "SELECT patientno,idno,datec,plastname,pfirstname,pmi,page,padiagnose,mlast,mfirst,mmi FROM Patient WHERE  patientno Like '*" & tmpid & "*' and  Pfirstname like '" & TmpFN & "' AND format(dateD,'mm/dd/yyyy') = '" & testdate & "';"
   Madz = True
 ElseIf id = True And ln = True And fn = False And NONMED = False And PHMED = False Then
   search_sql = "SELECT patientno,idno,datec,plastname,pfirstname,pmi,page,padiagnose,mlast,mfirst,mmi FROM Patient WHERE  patientno Like '*" & tmpid & "*' and PLastname Like '*" & TmpLN & "*' AND format(dateD,'mm/dd/yyyy') = '" & testdate & "';"
   Madz = True
 ElseIf id = True And ln = True And fn = True And NONMED = False And PHMED = False Then
     search_sql = "SELECT patientno,idno,datec,plastname,pfirstname,pmi,page,padiagnose,mlast,mfirst,mmi FROM Patient WHERE  patientno Like '*" & tmpid & "*' and PLastname Like '*" & TmpLN & "*' and Pfirstname like '" & TmpFN & "' AND format(dateD,'mm/dd/yyyy') = '" & testdate & "';"
     Madz = True
 ElseIf id = False And ln = True And fn = True And NONMED = False And PHMED = False Then
     search_sql = "SELECT patientno,idno,datec,plastname,pfirstname,pmi,page,padiagnose,mlast,mfirst,mmi FROM Patient WHERE  PLastname Like '*" & TmpLN & "*' and Pfirstname like '" & TmpFN & "' AND format(dateD,'mm/dd/yyyy') = '" & testdate & "';"
     Madz = True
 ElseIf id = False And ln = False And fn = True And NONMED = False And PHMED = False Then
    search_sql = "SELECT patientno,idno,datec,plastname,pfirstname,pmi,page,padiagnose,mlast,mfirst,mmi FROM Patient WHERE Pfirstname like '" & TmpFN & "' AND format(dateD,'mm/dd/yyyy') = '" & testdate & "';"
    Madz = True
 ElseIf id = False And ln = True And fn = False And NONMED = False And PHMED = False Then
    search_sql = "SELECT patientno,idno,datec,plastname,pfirstname,pmi,page,padiagnose,mlast,mfirst,mmi FROM Patient WHERE PLASTname like '" & TmpLN & "' AND format(dateD,'mm/dd/yyyy') = '" & testdate & "';"
    Madz = True
 '************* STAGE 2
 '*
 ElseIf id = True And ln = False And fn = False And NONMED = True And PHMED = False Then
   search_sql = "SELECT patientno,idno,datec,plastname,pfirstname,pmi,page,padiagnose,mlast,mfirst,mmi FROM Patient WHERE  patientno Like '*" & tmpid & "*' AND format(dateD,'mm/dd/yyyy') = '" & testdate & "' AND ISNULL(IDNO)=TRUE ;"
   Madz = True
 ElseIf id = True And ln = False And fn = False And NONMED = False And PHMED = True Then
   search_sql = "SELECT patientno,idno,datec,plastname,pfirstname,pmi,page,padiagnose,mlast,mfirst,mmi FROM Patient WHERE  patientno Like '*" & tmpid & "*' AND format(dateD,'mm/dd/yyyy') = '" & testdate & "' AND ISNULL(IDNO)=FALSE ;"
   Madz = True
 '*
 ElseIf id = True And ln = False And fn = True And NONMED = True And PHMED = False Then
   search_sql = "SELECT patientno,idno,datec,plastname,pfirstname,pmi,page,padiagnose,mlast,mfirst,mmi FROM Patient WHERE  patientno Like '*" & tmpid & "*' and  Pfirstname like '" & TmpFN & "' AND format(dateD,'mm/dd/yyyy') = '" & testdate & "' AND ISNULL(IDNO)=TRUE;"
   Madz = True
 ElseIf id = True And ln = False And fn = True And NONMED = False And PHMED = True Then
   search_sql = "SELECT patientno,idno,datec,plastname,pfirstname,pmi,page,padiagnose,mlast,mfirst,mmi FROM Patient WHERE  patientno Like '*" & tmpid & "*' and  Pfirstname like '" & TmpFN & "' AND format(dateD,'mm/dd/yyyy') = '" & testdate & "' AND ISNULL(IDNO)=FALSE;"
   Madz = True
 '*
 ElseIf id = True And ln = True And fn = False And NONMED = True And PHMED = False Then
   search_sql = "SELECT patientno,idno,datec,plastname,pfirstname,pmi,page,padiagnose,mlast,mfirst,mmi FROM Patient WHERE  patientno Like '*" & tmpid & "*' and PLastname Like '*" & TmpLN & "*' AND format(dateD,'mm/dd/yyyy') = '" & testdate & "' AND ISNULL(IDNO)=TRUE;"
   Madz = True
 ElseIf id = True And ln = True And fn = False And NONMED = False And PHMED = True Then
   search_sql = "SELECT patientno,idno,datec,plastname,pfirstname,pmi,page,padiagnose,mlast,mfirst,mmi FROM Patient WHERE  patientno Like '*" & tmpid & "*' and PLastname Like '*" & TmpLN & "*' AND format(dateD,'mm/dd/yyyy') = '" & testdate & "' AND ISNULL(IDNO)=FALSE;"
   Madz = True
 '*
 ElseIf id = True And ln = True And fn = True And NONMED = True And PHMED = False Then
     search_sql = "SELECT patientno,idno,datec,plastname,pfirstname,pmi,page,padiagnose,mlast,mfirst,mmi FROM Patient WHERE  patientno Like '*" & tmpid & "*' and PLastname Like '*" & TmpLN & "*' and Pfirstname like '" & TmpFN & "' AND format(dateD,'mm/dd/yyyy') = '" & testdate & "' AND ISNULL(IDNO)=TRUE;"
     Madz = True
 ElseIf id = True And ln = True And fn = True And NONMED = False And PHMED = True Then
     search_sql = "SELECT patientno,idno,datec,plastname,pfirstname,pmi,page,padiagnose,mlast,mfirst,mmi FROM Patient WHERE  patientno Like '*" & tmpid & "*' and PLastname Like '*" & TmpLN & "*' and Pfirstname like '" & TmpFN & "' AND format(dateD,'mm/dd/yyyy') = '" & testdate & "' AND ISNULL(IDNO)=FALSE;"
     Madz = True
 '*
 ElseIf id = False And ln = True And fn = True And NONMED = True And PHMED = False Then
     search_sql = "SELECT patientno,idno,datec,plastname,pfirstname,pmi,page,padiagnose,mlast,mfirst,mmi FROM Patient WHERE  PLastname Like '*" & TmpLN & "*' and Pfirstname like '" & TmpFN & "' AND format(dateD,'mm/dd/yyyy') = '" & testdate & "' AND ISNULL(IDNO)=TRUE;"
     Madz = True
ElseIf id = False And ln = True And fn = True And NONMED = False And PHMED = True Then
     search_sql = "SELECT patientno,idno,datec,plastname,pfirstname,pmi,page,padiagnose,mlast,mfirst,mmi FROM Patient WHERE  PLastname Like '*" & TmpLN & "*' and Pfirstname like '" & TmpFN & "' AND format(dateD,'mm/dd/yyyy') = '" & testdate & "' AND ISNULL(IDNO)=FALSE;"
     Madz = True
'*
 ElseIf id = False And ln = False And fn = True And NONMED = True And PHMED = False Then
    search_sql = "SELECT patientno,idno,datec,plastname,pfirstname,pmi,page,padiagnose,mlast,mfirst,mmi FROM Patient WHERE Pfirstname like '" & TmpFN & "' AND format(dateD,'mm/dd/yyyy') = '" & testdate & "' AND ISNULL(IDNO)=TRUE;"
    Madz = True
 ElseIf id = False And ln = False And fn = True And NONMED = False And PHMED = True Then
    search_sql = "SELECT patientno,idno,datec,plastname,pfirstname,pmi,page,padiagnose,mlast,mfirst,mmi FROM Patient WHERE Pfirstname like '" & TmpFN & "' AND format(dateD,'mm/dd/yyyy') = '" & testdate & "' AND ISNULL(IDNO)=FALSE;"
    Madz = True
 '*
 ElseIf id = False And ln = True And fn = False And NONMED = True And PHMED = False Then
    search_sql = "SELECT patientno,idno,datec,plastname,pfirstname,pmi,page,padiagnose,mlast,mfirst,mmi FROM Patient WHERE PLASTname like '" & TmpLN & "' AND format(dateD,'mm/dd/yyyy') = '" & testdate & "' AND ISNULL(IDNO)=TRUE;"
    Madz = True

 ElseIf id = False And ln = True And fn = False And NONMED = False And PHMED = True Then
    search_sql = "SELECT patientno,idno,datec,plastname,pfirstname,pmi,page,padiagnose,mlast,mfirst,mmi FROM Patient WHERE PLASTname like '" & TmpLN & "' AND format(dateD,'mm/dd/yyyy') = '" & testdate & "' AND ISNULL(IDNO)=FALSE;"
    Madz = True
 '*
 ElseIf id = False And ln = False And fn = False And NONMED = True And PHMED = False Then
    search_sql = "SELECT patientno,idno,datec,plastname,pfirstname,pmi,page,padiagnose,mlast,mfirst,mmi FROM Patient WHERE  format(dateD,'mm/dd/yyyy') = '" & testdate & "' AND ISNULL(IDNO)=TRUE;"
    Madz = True
 ElseIf id = False And ln = False And fn = False And NONMED = False And PHMED = True Then
    search_sql = "SELECT patientno,idno,datec,plastname,pfirstname,pmi,page,padiagnose,mlast,mfirst,mmi FROM Patient WHERE  format(dateD,'mm/dd/yyyy') = '" & testdate & "' AND ISNULL(IDNO)=FALSE;"
    Madz = True
 '*
 ElseIf id = False And ln = False And fn = False And NONMED = True And PHMED = True Then
    search_sql = "SELECT patientno,idno,datec,plastname,pfirstname,pmi,page,padiagnose,mlast,mfirst,mmi FROM Patient WHERE  format(dateD,'mm/dd/yyyy') = '" & testdate & "' ;"
    Madz = True
 ElseIf id = False And ln = False And fn = False And NONMED = False And PHMED = False Then
    Madz = False
 End If

   
 '/********** displaying of values to the form: scanhead *******
  Set TRS = DBMain.OpenRecordset(search_sql)
  PHconDelete.txtcon(0).Text = TRS.Fields(0) 'pno
  If IsNull(TRS.Fields(1)) = True Then
      PHconDelete.txtconM.Text = "Non-Med"
  Else
      PHconDelete.txtconM.Text = TRS.Fields(1)
  End If
  PHconDelete.txtcon(2).Text = TRS.Fields(3) 'lastname
  PHconDelete.txtcon(3).Text = TRS.Fields(4) 'firstname
  PHconDelete.txtcon(4).Text = TRS.Fields(5) 'MI
  PHconDelete.txtcon(5).Text = TRS.Fields(6) 'age
  PHconDelete.txtcon(6).Text = TRS.Fields(7) 'diagnosis
  PHconDelete.txtdate.Text = Format(TRS.Fields(2), "mm/dd/yyyy")
 'yaH im searchin?
        
     If Madz = True Then
        Set TRS = DBMain.OpenRecordset(search_sql)
         PHconDelete.ListView1.ListItems.Clear
         If TRS.RecordCount > 0 Then
            TRS.Fields.Refresh
             Do While Not TRS.EOF
                  Set List = PHconDelete.ListView1.ListItems.Add(, , TRS.Fields(0))
                       If IsNull(TRS.Fields(1)) = True Then
                          List.SubItems(1) = "Non-Med"
                        Else
                          List.SubItems(1) = TRS.Fields(1)
                        End If
                        List.SubItems(2) = Format(TRS.Fields(2), "mm/dd/yyyy")
                        List.SubItems(3) = TRS.Fields(3)
                        List.SubItems(4) = TRS.Fields(4)
                        List.SubItems(5) = TRS.Fields(5)
                        List.SubItems(6) = TRS.Fields(6)
                        List.SubItems(7) = TRS.Fields(7)
                        TRS.MoveNext
                     Loop
           '&H80000000&
           Else
           MsgBox "Search Data does not exist!, Click OK to Reset!", vbCritical, "Attention"
           ClearText2x
           PHconDelete.DisplayLstx
           PHconDelete.cmdreset.Enabled = True
           End If
       
     Else
        MsgBox "Nothing was performed, click OK to Reset!", vbCritical, "Attention"
        ClearText2x
         PHconDelete.DisplayLstx
        PHconDelete.cmdreset.Enabled = True
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




