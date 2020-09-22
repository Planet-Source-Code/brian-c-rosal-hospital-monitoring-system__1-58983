VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form SFMemDel 
   Caption         =   "Search PHmember"
   ClientHeight    =   4020
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4365
   Icon            =   "SFMem.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4020
   ScaleWidth      =   4365
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   3855
      Left            =   75
      TabIndex        =   1
      Top             =   75
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   6800
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Search Area"
      TabPicture(0)   =   "SFMem.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame Frame1 
         Height          =   3135
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   3930
         Begin VB.TextBox txtlast 
            Height          =   375
            Left            =   345
            MaxLength       =   35
            TabIndex        =   4
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
         Begin VB.TextBox txtid 
            Height          =   375
            Left            =   360
            TabIndex        =   0
            Top             =   600
            Width           =   2895
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
            TabIndex        =   6
            Top             =   2040
            Width           =   255
         End
         Begin Project1.chameleonButton srchbut2 
            Height          =   405
            Left            =   1440
            TabIndex        =   3
            Top             =   2565
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
            MICON           =   "SFMem.frx":045E
            PICN            =   "SFMem.frx":047A
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
            Caption         =   "Last Name"
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   9
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "PH Member ID No."
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
Attribute VB_Name = "SFMemDel"
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
  Dim TMP_KEY As String
  Dim TmpFN As String
  Dim TmpLN, tmpid As String
  Dim Madz, Mbry As Boolean
  Dim X As Byte
  On Error Resume Next

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
     
 'Store Last Name in Temporary Variables
  tmpid = Trim(txtId.Text)
  TmpFN = Trim(txtfirst.Text)
  TmpLN = Trim(txtlast.Text)
   
 'ID no.
 If Checkid.Value = 1 Then
     If Len(tmpid) < 1 Then
         MsgBox "You must provide a value for the Identification No.", vbInformation, "Attention"
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
         MsgBox ("You need to enter a value for Last Name")
         ln = False
         Exit Sub
      Else
         ln = True
      End If
End If
  'firstname :dbrymadz
If CheckFirst.Value = 1 Then
    If Len(TmpFN) < 1 Then
          MsgBox ("You need to enter a value for first Name")
          fn = False
          Exit Sub
    Else
          fn = True
    End If
End If
        
  
 If id = True And ln = False And fn = False Then
   search_sql = "SELECT * FROM PHmember WHERE  IDno Like '*" & tmpid & "*';"
   Madz = True
 ElseIf id = True And ln = False And fn = True Then
   search_sql = "SELECT * FROM PHmember WHERE  IDno Like '*" & tmpid & "*' and  firstname like '" & TmpFN & "';"
   Madz = True
 ElseIf id = True And ln = True And fn = False Then
   search_sql = "SELECT * FROM PHmember WHERE  IDno Like '*" & tmpid & "*' and Lastname Like '*" & TmpLN & "*';"
   Madz = True
 ElseIf id = True And ln = True And fn = True Then
     search_sql = "SELECT * FROM PHmember WHERE  IDno Like '*" & tmpid & "*' and Lastname Like '*" & TmpLN & "*' and firstname like '" & TmpFN & "';"
     Madz = True
 ElseIf id = False And ln = True And fn = True Then
     search_sql = "SELECT * FROM PHmember WHERE  Lastname Like '*" & TmpLN & "*' and firstname like '" & TmpFN & "';"
     Madz = True
 ElseIf id = False And ln = False And fn = True Then
    search_sql = "SELECT * FROM PHmember WHERE firstname like '" & TmpFN & "';"
    Madz = True
 ElseIf id = False And ln = True And fn = False Then
    search_sql = "SELECT * FROM PHmember WHERE LASTname like '" & TmpLN & "';"
    Madz = True
 ElseIf id = False And ln = False And fn = False Then
    Madz = False
End If

 
 '/*******Displaying of records the scanhead : david bry ****/
  
  Set TRS = DBMain.OpenRecordset(search_sql)
  For X = 0 To 5
  PHmemDelete.txtmem(X).Text = TRS.Fields(X)
  Next X
        
 'yaH im searchin?
        
     If Madz = True Then
        Set TRS = DBMain.OpenRecordset(search_sql)
         PHmemDelete.ListView.ListItems.Clear
         If TRS.RecordCount > 0 Then
            TRS.Fields.Refresh
             Do While Not TRS.EOF
                Set List = PHmemDelete.ListView.ListItems.Add(, , TRS.Fields(0))
                With List
                  .SubItems(1) = TRS.Fields(1) 'barangay
                  .SubItems(2) = TRS.Fields(2) 'population
                  .SubItems(3) = TRS.Fields(3) 'area
                  .SubItems(4) = TRS.Fields(4) 'profile
                  .SubItems(5) = TRS.Fields(5) 'City_mun
                End With
               TRS.MoveNext
            Loop
            PHmemDelete.cmdreset.Enabled = True
          Else
         MsgBox "No Data Found, Click OK to Reset!", vbCritical, "Attention"
          ClearTextBri
          PHmemDelete.DisplayLst
          PHmemDelete.cmdreset.Enabled = True
          PHmemDelete.ListView.SetFocus
           End If
       
     Else
       MsgBox "Nothing was performed, click OK to Reset!", vbCritical, "Attention"
       ClearTextBri
       PHmemDelete.DisplayLst
       PHmemDelete.cmdreset.Enabled = True
       PHmemDelete.ListView.SetFocus
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



Private Sub txtid_KeyPress(KeyAscii As Integer)
Dim madzbry As String
 madzbry = "0123456789"
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


