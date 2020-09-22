VERSION 5.00
Begin VB.Form frmhlp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Information"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7200
   Icon            =   "frmhlp.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   7200
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   90
      TabIndex        =   1
      Top             =   5205
      Width           =   7020
      Begin Project1.chameleonButton hlpClose 
         Height          =   465
         Left            =   5160
         TabIndex        =   3
         Top             =   180
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   820
         BTYPE           =   5
         TX              =   "E&xit"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
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
         MICON           =   "frmhlp.frx":0442
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
         Left            =   315
         TabIndex        =   2
         Top             =   315
         Width           =   4215
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000009&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000009&
         FillColor       =   &H00808000&
         FillStyle       =   0  'Solid
         Height          =   420
         Left            =   105
         Shape           =   4  'Rounded Rectangle
         Top             =   210
         Width           =   4695
      End
   End
   Begin VB.TextBox hlpText 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   60
      Width           =   6975
   End
End
Attribute VB_Name = "frmhlp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************
'To God Be The Glory
'**********************************************************************
Option Explicit


'**********************************************************************
'This Function is used to load a text file into a text box
'**********************************************************************
Function GetTextFromFile(txtFile, txtopen As TextBox)
    Dim sfile As String
    Dim nfile As Integer
    On Error Resume Next
    
    nfile = FreeFile
    sfile = txtFile
    Open sfile For Input As nfile
    txtopen = Input(LOF(nfile), nfile)
    Close nfile
End Function
'**********************************************************************
'**********************************************************************


'**********************************************************************
'**********************************************************************
Private Sub Form_Load()
  Add3DBorder Me
  Add3DBorder hlpText
  hlpText.Locked = False
 'Load Readme.txt into the text box
  Call GetTextFromFile(App.Path & "\README.txt", hlpText)
  hlpText.Locked = True
  hlpText.Enabled = True
End Sub
'**********************************************************************
'**********************************************************************


'**********************************************************************
'**********************************************************************
Private Sub hlpClose_Click()
  MainForm.Toolbar2.Buttons.Item(2).Enabled = True
  Unload Me
End Sub
'**********************************************************************
'**********************************************************************

