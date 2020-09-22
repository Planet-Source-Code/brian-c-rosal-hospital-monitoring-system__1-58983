VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MainForm 
   BackColor       =   &H00000000&
   Caption         =   "PH-Med/Non-Med Billing System  Ver 1.0"
   ClientHeight    =   4395
   ClientLeft      =   1380
   ClientTop       =   1035
   ClientWidth     =   7695
   HelpContextID   =   1
   Icon            =   "MDIfrm.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIfrm.frx":0442
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3600
      Top             =   3240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIfrm.frx":A727
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIfrm.frx":B2FB
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIfrm.frx":B74F
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIfrm.frx":BA6F
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIfrm.frx":BBCB
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIfrm.frx":CA1F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIfrm.frx":D2FB
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIfrm.frx":D74F
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIfrm.frx":DBA3
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   688
      BandCount       =   2
      _CBWidth        =   7695
      _CBHeight       =   390
      _Version        =   "6.0.8169"
      BandBackColor1  =   -2147483638
      Child1          =   "Toolbar1"
      MinWidth1       =   150
      MinHeight1      =   330
      Width1          =   1605
      FixedBackground1=   0   'False
      UseCoolbarColors1=   0   'False
      UseCoolbarPicture1=   0   'False
      NewRow1         =   0   'False
      Child2          =   "Toolbar2"
      MinWidth2       =   105
      MinHeight2      =   330
      Width2          =   1545
      UseCoolbarColors2=   0   'False
      NewRow2         =   0   'False
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   330
         Left            =   1800
         TabIndex        =   3
         Top             =   30
         Width           =   5805
         _ExtentX        =   10239
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Icon1"
               Object.ToolTipText     =   "User Manager"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Icon2"
               Object.ToolTipText     =   "Information"
               ImageIndex      =   8
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   165
         TabIndex        =   2
         Top             =   30
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Icon1"
               Object.ToolTipText     =   "PH member(s) Section"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Icon2"
               Object.ToolTipText     =   "PH-Med Admittance Section"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Icon3"
               Object.ToolTipText     =   "Non-Med Admittance Section"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Icon4"
               Object.ToolTipText     =   "PH-Med/Non-Med Discharging Section"
               ImageIndex      =   7
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   4140
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   13044
            MinWidth        =   13053
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   2
            TextSave        =   "8:48 AM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   2
            TextSave        =   "2/23/03"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   2
            Enabled         =   0   'False
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Alignment       =   2
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   873
            MinWidth        =   882
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   2
            Enabled         =   0   'False
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "NUM"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu DataM 
      Caption         =   "&Data Maintenance"
      Begin VB.Menu mem 
         Caption         =   "&PH Member Section"
         Shortcut        =   {F3}
      End
      Begin VB.Menu x3 
         Caption         =   "-"
      End
      Begin VB.Menu ConP 
         Caption         =   "&Confine Patient"
         Begin VB.Menu Con 
            Caption         =   "&PH-Med "
            Shortcut        =   ^P
         End
         Begin VB.Menu x6 
            Caption         =   "-"
         End
         Begin VB.Menu Con2 
            Caption         =   "&NON-MED"
            Shortcut        =   ^N
         End
      End
      Begin VB.Menu x4 
         Caption         =   "-"
      End
      Begin VB.Menu Dis 
         Caption         =   "&Discharge Patient"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu Reports 
      Caption         =   "&Reports"
      Begin VB.Menu memRep 
         Caption         =   "&PH Member(s) "
         Shortcut        =   {F6}
      End
      Begin VB.Menu ConRep 
         Caption         =   "&Confined Patients"
         Shortcut        =   {F7}
      End
      Begin VB.Menu line 
         Caption         =   "-"
      End
      Begin VB.Menu DisRep 
         Caption         =   "&Discharged Patients"
         Shortcut        =   {F8}
      End
   End
   Begin VB.Menu Util 
      Caption         =   "&Utilities"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Content"
      End
      Begin VB.Menu Deletion 
         Caption         =   "Data De&letion"
         Begin VB.Menu MemDel 
            Caption         =   "&PHMember Section"
            Shortcut        =   ^M
         End
         Begin VB.Menu DelCon 
            Caption         =   "&Confined Patients"
            Shortcut        =   ^C
         End
         Begin VB.Menu DisDel 
            Caption         =   "&Discharged Patients"
            Shortcut        =   ^D
         End
      End
      Begin VB.Menu log 
         Caption         =   "Log&ged Users"
      End
      Begin VB.Menu Manager 
         Caption         =   "User &Manager"
      End
      Begin VB.Menu xx2 
         Caption         =   "-"
      End
      Begin VB.Menu AA 
         Caption         =   "&About"
      End
   End
   Begin VB.Menu Win 
      Caption         =   "&Window"
      Begin VB.Menu Cd 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu Th 
         Caption         =   "Tile Hor&izontal"
      End
      Begin VB.Menu TV 
         Caption         =   "Tile &Vertical"
      End
      Begin VB.Menu AI 
         Caption         =   "Arrange I&cons"
      End
   End
   Begin VB.Menu exit 
      Caption         =   "&Exit"
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hWnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)
'hapit nako mabuang!!!!!  scanhead

Private Sub AA_Click()
AA.Enabled = False
frmAbout.Show vbModal
End Sub


Private Sub AI_Click()
 Me.Arrange vbArrangeIcons
End Sub

Private Sub Back_Click()
Dim bry As String

On Error GoTo bry
Shell App.Path + "\backup.bat", vbHide
MsgBox "Back-up database successfully, read the content section for restoring the database file.", vbInformation, "Attention"

bry:

End Sub

Private Sub restore_Click()
Dim rs As DAO.Recordset
Dim bry As String
'On Error GoTo bry

Set WSMain = Nothing
Set DBMain = Nothing
Set rs = Nothing

Close Index
Close Table
Close Database

Shell App.Path + "\restore.bat", vbNormalFocus

MsgBox "Restoring records successfully, Exit the application to retrieve the data properly!"

'bry:
End Sub

Private Sub Cd_Click()
Me.Arrange vbCascade
End Sub

Private Sub CON_Click()
Con.Enabled = False
Toolbar1.Buttons.Item(2).Enabled = False
PHconfined.Show
End Sub

Private Sub Cord_Click()
GovCordReport.Show
End Sub

Private Sub Cord_E_Click()
InstCordReport_E.Show
End Sub

Private Sub Cord_P_Click()
InstCordReport_P.Show
End Sub

Private Sub Cord_S_Click()
InstCordReport_S.Show
End Sub

Private Sub Cord_T_Click()
InstCordReport_T.Show
End Sub

Private Sub CordHouse_Click()
HouseCordReport.Show
End Sub

Private Sub Dev_Click()
Dev.Enabled = False
Toolbar1.Buttons.Item(2).Enabled = False
frmdev.Show
End Sub

Private Sub DevDel_Click()
DevDel.Enabled = False
Load Dlogin_form
Dlogin_form.Show vbModal
End Sub

Private Sub Con2_Click()
Con2.Enabled = False
Toolbar1.Buttons.Item(3).Enabled = False
PHconNONMed.Show
End Sub

Private Sub ConRep_Click()
listcon4.Show vbModal
End Sub

Private Sub DelCon_Click()
Dim X As Variant
'*****************
On Error GoTo X
Unload PHmember
Unload PHconfined
Unload PHconNONMed
Unload PHdischarged
MainForm.mem.Enabled = True
MainForm.Toolbar1.Buttons.Item(1).Enabled = True
MainForm.Con.Enabled = True
MainForm.Toolbar1.Buttons.Item(2).Enabled = True
MainForm.Con2.Enabled = True
MainForm.Toolbar1.Buttons.Item(3).Enabled = True
MainForm.Dis.Enabled = True
MainForm.Toolbar1.Buttons.Item(4).Enabled = True

X:
DelCon.Enabled = False
Load PHconDelete
PHconDelete.Show
End Sub

Private Sub dis_Click()
Dis.Enabled = False
Toolbar1.Buttons.Item(4).Enabled = False
PHdischarged.Show
End Sub

Private Sub DisDel_Click()
Dim X As Variant
'*****************
On Error GoTo X
Unload PHmember
Unload PHconfined
Unload PHconNONMed
Unload PHdischarged
MainForm.mem.Enabled = True
MainForm.Toolbar1.Buttons.Item(1).Enabled = True
MainForm.Con.Enabled = True
MainForm.Toolbar1.Buttons.Item(2).Enabled = True
MainForm.Con2.Enabled = True
MainForm.Toolbar1.Buttons.Item(3).Enabled = True
MainForm.Dis.Enabled = True
MainForm.Toolbar1.Buttons.Item(4).Enabled = True
X:
DisDel.Enabled = False
PHdisDelete.Show
End Sub

Private Sub DisRep_Click()
listDis.Show vbModal
End Sub

Private Sub exit_Click()
'Confirm exiting
Dim Msg$
Msg$ = "Are you sure you want to quit?"
If MsgBox(Msg$, vbYesNo + vbQuestion) = vbYes Then
   Unload frmLogin
   Shell App.Path + "\deltemp.bat", vbHide
End
End If
End Sub



Private Sub GCord_Click()
AnnCord.Show
End Sub





Private Sub Gllc_Click()
AnnLLC.Show
End Sub

Private Sub HotelCord_Click()
HotelsCordReport.Show
End Sub

Private Sub HotelLLC_Click()
HotelsLLCReport.Show
End Sub
'hapit nako mabuang!!!!!  scanhead

Private Sub IndusCord_Click()
IndusCordReport.Show
End Sub

Private Sub IndusLLC_Click()
IndusLLCReport.Show
End Sub

Private Sub LLC_Click()
GovLLcReport.Show
End Sub

Private Sub LLC_E_Click()
InstLLCReport_E.Show
End Sub

Private Sub LLC_P_Click()
InstLLcReport_P.Show
End Sub

Private Sub LLC_S_Click()
InstLLcReport_S.Show
End Sub
 
Private Sub LLC_T_Click()
InstLLCReport_T.Show
End Sub

Private Sub LLCHouse_Click()
HouseLLCreport.Show
End Sub

Private Sub log_Click()
Viewlog.Show vbModal
End Sub

Private Sub Manager_Click()
Manager.Enabled = False
Toolbar2.Buttons.Item(1).Enabled = False
Load frmUsrMngr
frmUsrMngr.Show vbModal
End Sub






Private Sub MDIForm_Load()
Add3DBorder Me
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
    End If
    Close All
    Unload Me
    Shell App.Path + "\deltemp.bat", vbHide
    End Sub

Private Sub MEM_Click()
mem.Enabled = False
Toolbar1.Buttons.Item(1).Enabled = False
PHmember.Show
End Sub

Private Sub MemDel_Click()
Dim X As Variant
'*****************
On Error GoTo X
Unload PHmember
Unload PHconfined
Unload PHconNONMed
Unload PHdischarged
MainForm.mem.Enabled = True
MainForm.Toolbar1.Buttons.Item(1).Enabled = True
MainForm.Con.Enabled = True
MainForm.Toolbar1.Buttons.Item(2).Enabled = True
MainForm.Con2.Enabled = True
MainForm.Toolbar1.Buttons.Item(3).Enabled = True
MainForm.Dis.Enabled = True
MainForm.Toolbar1.Buttons.Item(4).Enabled = True
X:
MemDel.Enabled = False
Load PHmemDelete
PHmemDelete.Show

End Sub

Private Sub memRep_Click()
Dim TRAPPER As Variant
Dim sql As String
On Error GoTo TRAPPER

DBa.Open "DRIVER={Microsoft Access Driver (*.mdb)};dbq=" & App.Path & "\hospital.mdb", , "student"
sql = "Select * from Phmember  order by lastname,firstname,mi  ASC;"
Set rsa = DBa.Execute(sql)
Set PHmemRep.DataSource = rsa
PHmemRep.Show vbModal
TRAPPER:
DBa.Close
Set DBa = Nothing
End Sub

Private Sub mnuHelpcontents_Click()
    Toolbar2.Buttons.Item(2).Enabled = False
   Load frmhlp
   frmhlp.Show vbModal
  
End Sub

Private Sub PopCord_Click()
PopCordReport.Show
End Sub

Private Sub PopLLC_Click()
PopLLCReport.Show
End Sub

Private Sub Rd_Click()
Dim bry As String

On Error GoTo bry
Shell App.Path + "\restore.bat", vbNormalFocus
MsgBox "Back-up database successfully, read the content section for restoring the database file.", vbInformation, "Attention"

bry:

End Sub

Private Sub SCord_Click()
SpotsCord.Show

End Sub

Private Sub Sllc_Click()
SpotsLLC.Show
End Sub

Private Sub Th_Click()
  Me.Arrange vbTileHorizontal
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
 Select Case Button.Key
        Case "Icon1"
            mem.Enabled = False
            Toolbar1.Buttons.Item(1).Enabled = False
            PHmember.Show
        Case "Icon2"
            Con.Enabled = False
            Toolbar1.Buttons.Item(2).Enabled = False
            PHconfined.Show
        Case "Icon3"
             Con2.Enabled = False
             Toolbar1.Buttons.Item(3).Enabled = False
             PHconNONMed.Show
        Case "Icon4"
             Dis.Enabled = False
             Toolbar1.Buttons.Item(4).Enabled = False
             PHdischarged.Show
 End Select
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim OpenFixDB As Long
'-----haaaaaaaaaayyyyyyy
     Select Case Button.Key
             Case "Icon1"  'toolbar2
             Toolbar2.Buttons.Item(1).Enabled = False
             frmUsrMngr.Show vbModal
             Case "Icon2"  'toolbar2
             Toolbar2.Buttons.Item(2).Enabled = False
             frmhlp.Show vbModal
        End Select

End Sub
'Scanhead of windsor

Private Sub TV_Click()
 Me.Arrange vbTileVertical
End Sub

Private Sub WB_Click()
Toolbar2.Buttons.Item(1).Enabled = False
             frmWeb.StartingAddress = "http://localhost/mysite/"
             frmWeb.Show
End Sub

Private Sub WS_Click()
'Open WebSite Server
            OpenFixDB = Shell("c:\website\httpd32.exe", vbNormalFocus)
            MsgBox "WebSite Server is activated!", vbInformation
End Sub
