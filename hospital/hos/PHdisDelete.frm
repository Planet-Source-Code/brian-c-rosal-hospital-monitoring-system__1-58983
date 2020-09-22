VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PHdisDelete 
   BackColor       =   &H80000007&
   Caption         =   "Discharged Patients Data Deletion"
   ClientHeight    =   6465
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11820
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "PHdisDelete.frx":0000
   ScaleHeight     =   6465
   ScaleWidth      =   11820
   WindowState     =   2  'Maximized
   Begin Project1.chameleonButton butx 
      Height          =   285
      Left            =   105
      TabIndex        =   35
      Top             =   195
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
      MICON           =   "PHdisDelete.frx":A2E5
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
      Height          =   1680
      Left            =   480
      TabIndex        =   33
      Top             =   3465
      Width           =   3315
      Begin Project1.chameleonButton cmdDelete 
         Height          =   615
         Left            =   360
         TabIndex        =   34
         Top             =   240
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   1085
         BTYPE           =   5
         TX              =   "&Undo Discharged Patient(s)"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
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
         MICON           =   "PHdisDelete.frx":A301
         PICN            =   "PHdisDelete.frx":A31D
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   1
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Project1.chameleonButton cmdDel2 
         Height          =   615
         Left            =   360
         TabIndex        =   41
         Top             =   915
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   1085
         BTYPE           =   5
         TX              =   "Del&ete Discharged Patient(s)"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
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
         MICON           =   "PHdisDelete.frx":A76F
         PICN            =   "PHdisDelete.frx":A78B
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   1
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin Project1.chameleonButton cmdExit 
      Height          =   585
      Left            =   495
      TabIndex        =   29
      Top             =   5280
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   1032
      BTYPE           =   5
      TX              =   "E&xit"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
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
      MICON           =   "PHdisDelete.frx":ABDD
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame2 
      Caption         =   "Delete By Year"
      Height          =   885
      Left            =   480
      TabIndex        =   26
      Top             =   2520
      Width           =   3330
      Begin VB.CheckBox Check1 
         Height          =   345
         Left            =   2805
         TabIndex        =   32
         Top             =   330
         Width           =   300
      End
      Begin MSMask.MaskEdBox txtyear 
         Height          =   375
         Left            =   840
         TabIndex        =   31
         Top             =   315
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label9 
         Caption         =   "Year:"
         Height          =   495
         Left            =   240
         TabIndex        =   30
         Top             =   255
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Patient Section"
      ForeColor       =   &H00C00000&
      Height          =   2235
      Left            =   4680
      TabIndex        =   11
      Top             =   105
      Width           =   6615
      Begin VB.TextBox txtctime 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "d MMMM yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   3840
         TabIndex        =   37
         Top             =   1845
         Width           =   1215
      End
      Begin VB.TextBox txtmidno 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   270
         Left            =   2700
         MaxLength       =   15
         TabIndex        =   27
         Top             =   450
         Width           =   1605
      End
      Begin VB.TextBox txtpno 
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   255
         Left            =   180
         TabIndex        =   18
         Top             =   465
         Width           =   2415
      End
      Begin VB.TextBox txtplast 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   195
         TabIndex        =   17
         Top             =   990
         Width           =   2655
      End
      Begin VB.TextBox txtpfirst 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2940
         TabIndex        =   16
         Top             =   990
         Width           =   3090
      End
      Begin VB.TextBox txtpmi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6120
         TabIndex        =   15
         Top             =   990
         Width           =   345
      End
      Begin VB.TextBox txtpage 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   255
         Left            =   210
         TabIndex        =   14
         Top             =   1470
         Width           =   450
      End
      Begin VB.TextBox txtpAd 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   255
         Left            =   765
         TabIndex        =   13
         Top             =   1470
         Width           =   4290
      End
      Begin VB.TextBox txtdate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "d MMMM yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2400
         TabIndex        =   12
         Top             =   1845
         Width           =   1350
      End
      Begin Project1.chameleonButton cmdreset 
         Height          =   525
         Left            =   4425
         TabIndex        =   38
         Top             =   210
         Width           =   2100
         _extentx        =   2064
         _extenty        =   926
         btype           =   5
         tx              =   "&Reset"
         enab            =   -1  'True
         font            =   "PHdisDelete.frx":ABF9
         coltype         =   2
         focusr          =   -1  'True
         bcol            =   12632256
         bcolo           =   12632256
         fcol            =   0
         fcolo           =   255
         mcol            =   12632256
         mptr            =   1
         micon           =   "PHdisDelete.frx":AC1D
         umcol           =   -1  'True
         soft            =   0   'False
         picpos          =   0
         ngrey           =   0   'False
         fx              =   1
         hand            =   0   'False
         check           =   0   'False
         value           =   0   'False
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H80000003&
         Height          =   765
         Left            =   5130
         Top             =   1365
         Width           =   1365
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Note: Check the data carefully before deleting."
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
         Height          =   630
         Index           =   7
         Left            =   5250
         TabIndex        =   43
         Top             =   1410
         Width           =   1230
      End
      Begin VB.Label Label2 
         Caption         =   "PH Member ID No."
         Height          =   255
         Index           =   6
         Left            =   2700
         TabIndex        =   28
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Patient No.:"
         Height          =   255
         Index           =   3
         Left            =   180
         TabIndex        =   25
         Top             =   255
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "Age"
         Height          =   255
         Index           =   2
         Left            =   195
         TabIndex        =   24
         Top             =   1245
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Admission Diagnosed:"
         Height          =   255
         Index           =   1
         Left            =   795
         TabIndex        =   23
         Top             =   1245
         Width           =   2415
      End
      Begin VB.Label Label4 
         Caption         =   "Admission Date/Time:"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   195
         TabIndex        =   22
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "MI"
         Height          =   195
         Left            =   6135
         TabIndex        =   21
         Top             =   780
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "First Name"
         Height          =   210
         Index           =   0
         Left            =   2970
         TabIndex        =   20
         Top             =   780
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Last Name"
         Height          =   255
         Index           =   0
         Left            =   210
         TabIndex        =   19
         Top             =   780
         Width           =   975
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "List of Discharged Patients"
      ForeColor       =   &H000000FF&
      Height          =   2220
      Left            =   495
      TabIndex        =   9
      Top             =   105
      Width           =   4065
      Begin MSComctlLib.ListView ListView2 
         Height          =   1785
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   3795
         _ExtentX        =   6694
         _ExtentY        =   3149
         View            =   3
         Arrange         =   1
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
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Patient No."
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "PH Member ID No."
            Object.Width           =   9
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Date/Time Discharged"
            Object.Width           =   9
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Last Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "First Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "MI"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Age"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Final Diagnosed"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Patient's Discharged Transaction Section"
      ForeColor       =   &H00C00000&
      Height          =   3420
      Left            =   3900
      TabIndex        =   1
      Top             =   2445
      Width           =   7350
      Begin VB.Frame Frame8 
         Height          =   2685
         Left            =   3330
         TabIndex        =   44
         Top             =   150
         Width           =   3855
         Begin MSMask.MaskEdBox txtDISmed 
            Height          =   315
            Left            =   1230
            TabIndex        =   45
            Top             =   1005
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Enabled         =   0   'False
            Format          =   "#,##0.00;($#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtDISpf 
            Height          =   345
            Left            =   1230
            TabIndex        =   46
            Top             =   1395
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   609
            _Version        =   393216
            BackColor       =   16777215
            Enabled         =   0   'False
            Format          =   "#,##0.00;($#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtDISpaid 
            Height          =   375
            Left            =   1245
            TabIndex        =   47
            Top             =   2235
            Width           =   2505
            _ExtentX        =   4419
            _ExtentY        =   661
            _Version        =   393216
            BackColor       =   -2147483624
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;($#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtDISlab 
            Height          =   315
            Left            =   1245
            TabIndex        =   48
            Top             =   615
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Enabled         =   0   'False
            Format          =   "#,##0.00;($#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtDISrm 
            Height          =   315
            Left            =   1245
            TabIndex        =   49
            Top             =   240
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Enabled         =   0   'False
            Format          =   "#,##0.00;($#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtDIStot 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
            Height          =   345
            Left            =   1245
            TabIndex        =   50
            Top             =   1830
            Width           =   2520
            _ExtentX        =   4445
            _ExtentY        =   609
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;($#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtPAYrm 
            Height          =   315
            Left            =   2565
            TabIndex        =   51
            Top             =   240
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483624
            Enabled         =   0   'False
            Format          =   "#,##0.00;($#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtPAYlab 
            Height          =   315
            Left            =   2565
            TabIndex        =   52
            Top             =   630
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483624
            Enabled         =   0   'False
            Format          =   "#,##0.00;($#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtPAYmed 
            Height          =   315
            Left            =   2565
            TabIndex        =   53
            Top             =   1005
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483624
            Enabled         =   0   'False
            Format          =   "#,##0.00;($#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtPAYpf 
            Height          =   345
            Left            =   2565
            TabIndex        =   54
            Top             =   1395
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   609
            _Version        =   393216
            BackColor       =   -2147483624
            Enabled         =   0   'False
            Format          =   "#,##0.00;($#,##0.00)"
            PromptChar      =   "_"
         End
         Begin VB.Label Label6 
            Caption         =   "Rm/Brding:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   105
            TabIndex        =   60
            Top             =   255
            Width           =   1500
         End
         Begin VB.Label Label6 
            Caption         =   "Lab. Fee:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   59
            Top             =   645
            Width           =   1455
         End
         Begin VB.Label Label6 
            Caption         =   "Ttl Med.(s):"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   90
            TabIndex        =   58
            Top             =   1065
            Width           =   1335
         End
         Begin VB.Label Label6 
            Caption         =   "Prof. Fee:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   57
            Top             =   1440
            Width           =   1740
         End
         Begin VB.Label Label6 
            Caption         =   "PH/Amt Paid:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Index           =   4
            Left            =   90
            TabIndex        =   56
            Top             =   2310
            Width           =   1485
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            BackColor       =   &H80000016&
            Caption         =   "Ttl Fee(s)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   165
            TabIndex        =   55
            Top             =   1890
            Width           =   1005
         End
         Begin VB.Shape Shape2 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H80000010&
            Height          =   345
            Left            =   105
            Top             =   1830
            Width           =   1230
         End
      End
      Begin VB.TextBox txttime 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "d MMMM yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   1470
         TabIndex        =   42
         Top             =   555
         Width           =   1245
      End
      Begin VB.Frame Frame7 
         Height          =   1245
         Left            =   120
         TabIndex        =   36
         Top             =   1575
         Width           =   3120
         Begin Project1.chameleonButton cmdsrch 
            Height          =   840
            Left            =   240
            TabIndex        =   39
            Top             =   240
            Width           =   1350
            _extentx        =   2381
            _extenty        =   1482
            btype           =   5
            tx              =   "Search D&ischarged Patient(s)"
            enab            =   -1  'True
            font            =   "PHdisDelete.frx":AC3B
            coltype         =   2
            focusr          =   -1  'True
            bcol            =   12632256
            bcolo           =   12632256
            fcol            =   0
            fcolo           =   255
            mcol            =   12632256
            mptr            =   1
            micon           =   "PHdisDelete.frx":AC67
            umcol           =   -1  'True
            soft            =   0   'False
            picpos          =   0
            ngrey           =   0   'False
            fx              =   1
            hand            =   0   'False
            check           =   0   'False
            value           =   0   'False
         End
         Begin Project1.chameleonButton cmdVdis 
            Height          =   840
            Left            =   1665
            TabIndex        =   40
            Top             =   255
            Width           =   1350
            _extentx        =   2381
            _extenty        =   1482
            btype           =   5
            tx              =   "&Discharged Patient(s)"
            enab            =   -1  'True
            font            =   "PHdisDelete.frx":AC85
            coltype         =   2
            focusr          =   -1  'True
            bcol            =   12632256
            bcolo           =   12632256
            fcol            =   0
            fcolo           =   255
            mcol            =   12632256
            mptr            =   1
            micon           =   "PHdisDelete.frx":ACB1
            umcol           =   -1  'True
            soft            =   0   'False
            picpos          =   0
            ngrey           =   0   'False
            fx              =   1
            hand            =   0   'False
            check           =   0   'False
            value           =   0   'False
         End
      End
      Begin VB.TextBox txtDISdiag 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Left            =   150
         TabIndex        =   3
         Top             =   1125
         Width           =   2895
      End
      Begin MSMask.MaskEdBox txtDISdate 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/d/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   330
         Left            =   150
         TabIndex        =   2
         Top             =   540
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         _Version        =   393216
         BackColor       =   -2147483624
         Enabled         =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblDISdiff 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   3570
         TabIndex        =   6
         Top             =   2925
         Width           =   3495
      End
      Begin VB.Label Label2 
         Caption         =   "Date/Time Discharged:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   210
         Index           =   4
         Left            =   150
         TabIndex        =   8
         Top             =   300
         Width           =   2295
      End
      Begin VB.Label Label2 
         Caption         =   "Final Diagnosed:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   5
         Left            =   150
         TabIndex        =   7
         Top             =   900
         Width           =   2295
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H008080FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000018&
         FillColor       =   &H80000016&
         FillStyle       =   0  'Solid
         Height          =   420
         Left            =   3345
         Top             =   2910
         Width           =   3840
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Balance:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   465
         Left            =   1110
         TabIndex        =   4
         Top             =   2940
         Width           =   2040
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Balance:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   465
         Left            =   1140
         TabIndex        =   5
         Top             =   2970
         Width           =   2295
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   6075
      Index           =   2
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11475
      _ExtentX        =   20241
      _ExtentY        =   10716
      TabWidthStyle   =   2
      MultiRow        =   -1  'True
      HotTracking     =   -1  'True
      Placement       =   2
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
      EndProperty
      Enabled         =   0   'False
   End
End
Attribute VB_Name = "PHdisDelete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sql, rs, y As Variant, ctr As Byte
Dim sql2 As Variant
Dim MyType, passString As String
Dim formod As String
Dim sum, Dis As Double



Private Sub butx_Click()
Unload Me
MainForm.DisDel.Enabled = True
End Sub


Private Sub chameleonButton1_Click()

End Sub

Private Sub cmdDel2_Click()
Dim X, TRAPPER, passString, testdate As String
testdate = ""
X = Format$("01/01/" + txtyear.Text, "MM/DD/YYYY")
If Check1.Value = 1 Then
     If IsDate(X) = True Then
       If MsgBox("Are you sure you want to Delete Permanently All Discharged Patients belong to Year: " + txtyear.Text + "? Note: No restoration for this.", vbQuestion + vbYesNo, "Confirmation") = vbYes Then
           
           '****** Search for the year *********
           sql = "select * from Pdischarged where format(dateD,'yyyy') like '" + txtyear.Text + "';"
           Set rs = DBMain.OpenRecordset(sql)
            If rs.RecordCount < 0 Or rs.RecordCount = 0 Then
                       MsgBox "No Patients discharged at the year: " + txtyear.Text + " .", vbInformation, "Sorry"
                       GoTo TRAPPER
            End If
           
           '***Pdischarged table:scanhead ******
           DBMain.Execute "DELETE * FROM Pdischarged WHERE  format(dateD,'yyyy') like '" + txtyear.Text + "' ;"
           '***Patient table:scanhead ******
           DBMain.Execute "DELETE * FROM Patient WHERE format(dateD,'yyyy') like '" + txtyear.Text + "' ;"
           ListView2.ListItems.Clear
           disp
           txtyear.Text = "____"
           Check1.Value = 0
           cmdDelete.Enabled = False
           ListView2.SetFocus
           clearDdel
       Else
           txtyear.Text = "____"
           Check1.Value = 0
           ListView2.SetFocus
           clearDdel
           GoTo TRAPPER
        End If
     Else
          MsgBox "Invalid Year", vbOKOnly + vbCritical, "Attention"
     End If
 Else
         If txtpno.Text = "" Then
             MsgBox "Please select the Discharged Patient to DELETE permanently.", vbOKOnly + vbCritical, "Attention"
             GoTo TRAPPER
         Else
              If MsgBox("Are you sure you want to DELETE " + txtpfirst.Text + " ." + txtpmi.Text + " " + txtplast.Text + " permanently?", vbExclamation + vbYesNo, "Confirm Deletion") = vbYes Then
                  Dim Dellist As ListItem
                  '***** discharged table ***********'
                   DBMain.Execute "DELETE * FROM Pdischarged WHERE PatientNo = '" + txtpno.Text + "';"
                  '***Patient table the restore section :scanhead ******
                   passString = "DELETE * FROM Patient WHERE patient.patientno='" & txtpno.Text & "' AND ISNULL(PATIENT.DATED)=FALSE;"
                   DBMain.Execute passString
            
                  Set Dellist = ListView2.FindItem(txtpno.Text, , , lvwPartial)
                  ListView2.ListItems.Remove Dellist.Index
                  txtyear.Text = "____"
                  Check1.Value = 0
                  cmdDel2.Enabled = False
                  ListView2.SetFocus
                  clearDdel
              End If
        End If
  End If

TRAPPER:

End Sub

Private Sub cmdDelete_Click()
Dim X, TRAPPER, passString, testdate As String
testdate = ""
X = Format$("01/01/" + txtyear.Text, "MM/DD/YYYY")
If Check1.Value = 1 Then
     If IsDate(X) = True Then
       If MsgBox("Are you sure you want to DELETE Permanently All Discharged Patients belong to Year: " + txtyear.Text + "? Note: No restoration for this.", vbQuestion + vbYesNo, "Confirmation") = vbYes Then
           
           '****** Search for the year *********
           sql = "select * from Pdischarged where format(dateD,'yyyy') like '" + txtyear.Text + "';"
           Set rs = DBMain.OpenRecordset(sql)
            If rs.RecordCount < 0 Or rs.RecordCount = 0 Then
                       MsgBox "No Patients discharged at the year: " + txtyear.Text + " .", vbInformation, "Sorry"
                       GoTo TRAPPER
            End If
           
           '***Pdischarged table:scanhead ******
           DBMain.Execute "DELETE * FROM Pdischarged WHERE  format(dateD,'yyyy') like '" + txtyear.Text + "' ;"
           '***Patient table:scanhead ******
           DBMain.Execute "DELETE * FROM Patient WHERE format(dateD,'yyyy') like '" + txtyear.Text + "' ;"
           ListView2.ListItems.Clear
           disp
           txtyear.Text = "____"
           Check1.Value = 0
           cmdDelete.Enabled = False
           ListView2.SetFocus
           clearDdel
       Else
           txtyear.Text = "____"
           Check1.Value = 0
           ListView2.SetFocus
           clearDdel
           GoTo TRAPPER
        End If
     Else
          MsgBox "Invalid Year", vbOKOnly + vbCritical, "Attention"
     End If
 Else
         If txtpno.Text = "" Then
             MsgBox "Please select the Discharged Patient to RESTORE as Confined Patient.", vbOKOnly + vbCritical, "Attention"
             GoTo TRAPPER
         Else
              If MsgBox("Are you sure you want to RESTORE " + txtpfirst.Text + " ." + txtpmi.Text + " " + txtplast.Text + " as Confined Patient?", vbExclamation + vbYesNo, "Confirm Deletion") = vbYes Then
                  Dim Dellist As ListItem
                  '***** discharged table ***********'
                   DBMain.Execute "DELETE * FROM Pdischarged WHERE PatientNo = '" + txtpno.Text + "';"
                  '***Patient table the restore section :scanhead ******
                   passString = "UPDATE Patient SET [DateD]='" & testdate & "' WHERE patient.patientno='" & txtpno.Text & "' ;"
                   DBMain.Execute passString
            
                  Set Dellist = ListView2.FindItem(txtpno.Text, , , lvwPartial)
                  ListView2.ListItems.Remove Dellist.Index
                  txtyear.Text = "____"
                  Check1.Value = 0
                  cmdDelete.Enabled = False
                  ListView2.SetFocus
                  clearDdel
              End If
        End If
  End If

TRAPPER:

End Sub

Private Sub cmdExit_Click()
Unload Me
MainForm.DisDel.Enabled = True
End Sub

Private Sub cmdreset_Click()
clearDdel
DisplayLst
End Sub

Private Sub cmdsrch_Click()
Load SFdisDEL
cmdreset.Enabled = True
SFdisDEL.Show vbModal
End Sub

Private Sub cmdVdis_Click()
listDis.Show vbModal
End Sub

Private Sub Form_Load()
 Dim testdate As String
  testdate = ""
  add3d
  Set WSMain = DBEngine.Workspaces(0)
  Set DBMain = WSMain.OpenDatabase(App.Path + "\hospital.mdb", False, False, ";pwd=scanhead")
   '/*****************Discharged Patients ***********************
 sql = "select Pdischarged.patientno,Pdischarged.idno,Pdischarged.dateD,Patient.plastname,Patient.pfirstname,Patient.pmi,Patient.page,Pdischarged.fdiagnose,Pdischarged.Timeout FROM Patient INNER JOIN Pdischarged ON  PDISCHARGED.PATIENTNO = PATIENT.PATIENTNO  AND PDISCHARGED.dateD = PATIENT.dateD order by patient.plastname,patient.pfirstname,patient.pmi;"
           Set rs = DBMain.OpenRecordset(sql)
                Do Until rs.EOF
                    Set y = ListView2.ListItems.Add(, , rs.Fields(0))
                        y.SubItems(1) = rs.Fields(1)
                        y.SubItems(2) = Format(rs.Fields(2), "mm/dd/yyyy") + " " + Format(rs.Fields(8), "medium time")
                        y.SubItems(3) = rs.Fields(3)
                        y.SubItems(4) = rs.Fields(4)
                        y.SubItems(5) = rs.Fields(5)
                        y.SubItems(6) = rs.Fields(6)
                        y.SubItems(7) = rs.Fields(7)
                        rs.MoveNext
                 Loop
cmdreset.Enabled = True
End Sub

Private Sub ListView2_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim TQR As DAO.QueryDef
Dim TRS As DAO.Recordset
 Set TQR = DBMain.CreateQueryDef("", "SELECT Pdischarged.dateD,Pdischarged.Fdiagnose,Pdischarged.rmbrd,Pdischarged.labfee,Pdischarged.Tmeds,Pdischarged.Pfee,Pdischarged.Philpay,Pdischarged.Diff,Pdischarged.idno,Pdischarged.Mlast,Pdischarged.Mfirst,Pdischarged.Mmi,Pdischarged.Plast,Pdischarged.Pfirst,Pdischarged.Pmi,Pdischarged.Page,Pdischarged.Pdiag,Pdischarged.patientno,pdischarged.pdatec,Pdischarged.Timein,Pdischarged.Timeout,Pdischarged.rmpay , Pdischarged.labpay, Pdischarged.medpay, Pdischarged.pfpay FROM Pdischarged,Patient WHERE Pdischarged.patientno ='" & ListView2.SelectedItem.Text & "' AND Patient.dateD = Pdischarged.DateD ORDER BY pdischarged.mlast,pdischarged.mfirst,pdischarged.mmi")
 Set TRS = TQR.OpenRecordset()
 txtDISdate.Text = Format(TRS.Fields(0), "mm/dd/yyyy")
 txtDISdiag.Text = TRS.Fields(1)
 txtDISrm.Text = TRS.Fields(2)
 txtDISlab.Text = TRS.Fields(3)
 txtDISmed.Text = TRS.Fields(4)
 txtDISpf.Text = TRS.Fields(5)
 txtDISpaid.Text = TRS.Fields(6)
 txtmidno.Text = TRS.Fields(8)
 txtpno.Text = TRS.Fields(17)
 txtplast.Text = TRS.Fields(12)
 txtpfirst.Text = TRS.Fields(13)
 txtpmi.Text = TRS.Fields(14)
 txtpage.Text = TRS.Fields(15)
 txtpAd.Text = TRS.Fields(16)
 txtdate.Text = Format(TRS.Fields(18), "mm/dd/yyyy")
 txtctime.Text = Format(TRS.Fields(19), "medium time")
 txttime.Text = Format(TRS.Fields(20), "medium time")
 txtPAYrm.Text = TRS.Fields(21)
 txtPAYlab.Text = TRS.Fields(22)
 txtPAYmed.Text = TRS.Fields(23)
 txtPAYpf.Text = TRS.Fields(24)

 txtDIStot.Text = Format$(Val(TRS.Fields(2)) + Val(TRS.Fields(3)) + Val(TRS.Fields(4)) + Val(TRS.Fields(5)), "###,###,###.00")
 lblDISdiff.Caption = "Php " + Format$(Val(TRS.Fields(7)), "###,###,###.00")
 cmdDelete.Enabled = True
 cmdDel2.Enabled = True
End Sub

Private Sub txtyear_GotFocus()
With txtyear
.SelStart = 0
.SelLength = Len(txtyear.Text)
End With
cmdDelete.Enabled = True
End Sub

Private Sub txtyear_KeyPress(KeyAscii As Integer)
Check1.Value = 1
End Sub

Public Sub disp()
'/*****************Discharged Patients ***********************
 sql = "select Pdischarged.patientno,Pdischarged.idno,Pdischarged.dateD,Patient.plastname,Patient.pfirstname,Patient.pmi,Patient.page,Pdischarged.fdiagnose FROM Patient INNER JOIN Pdischarged ON  PDISCHARGED.PATIENTNO = PATIENT.PATIENTNO  AND PDISCHARGED.dateD = PATIENT.dateD order by patient.plastname,patient.pfirstname,patient.pmi;"
           Set rs = DBMain.OpenRecordset(sql)
                Do Until rs.EOF
                    Set y = ListView2.ListItems.Add(, , rs.Fields(0))
                        y.SubItems(1) = rs.Fields(1)
                        y.SubItems(2) = Format(rs.Fields(2), "mm/dd/yyyy")
                        y.SubItems(3) = rs.Fields(3)
                        y.SubItems(4) = rs.Fields(4)
                        y.SubItems(5) = rs.Fields(5)
                        y.SubItems(6) = rs.Fields(6)
                        y.SubItems(7) = rs.Fields(7)
                        rs.MoveNext
                 Loop

 
End Sub

Private Function add3d()
  Add3DBorder Me
  Add3DBorder ListView2
  Add3DBorder txtmidno
  Add3DBorder txtpno
  Add3DBorder txtplast
  Add3DBorder txtpfirst
  Add3DBorder txtpmi
  Add3DBorder txtpage
  Add3DBorder txtpAd
  Add3DBorder txtyear
  Add3DBorder txtctime
  Add3DBorder txttime
  Add3DBorder txtdate
  Add3DBorder txtDISrm
  Add3DBorder txtDISdate
  Add3DBorder txtDISlab
  Add3DBorder txtDISdiag
  Add3DBorder txtDISmed
  Add3DBorder txtDISpaid
  Add3DBorder txtDISpf
  
  Add3DBorder txtPAYrm
  Add3DBorder txtPAYmed
  Add3DBorder txtPAYlab
  Add3DBorder txtPAYpf
  End Function

Public Sub DisplayLst() 'Reset control display: gotcha the scanhedbri
On Error GoTo X
    Dim TQR As DAO.QueryDef
    Dim rs As DAO.Recordset
    Dim y As ListItem
     Dim sql As String
    Dim X As Long
   
  sql = "select Pdischarged.patientno,Pdischarged.idno,Pdischarged.dateD,Patient.plastname,Patient.pfirstname,Patient.pmi,Patient.page,Pdischarged.fdiagnose FROM Patient INNER JOIN Pdischarged ON  PDISCHARGED.PATIENTNO = PATIENT.PATIENTNO  AND PDISCHARGED.dateD = PATIENT.dateD order by patient.plastname,patient.pfirstname,patient.pmi;"
  Set rs = DBMain.OpenRecordset(sql)
     ListView2.ListItems.Clear
    Do Until rs.EOF
        Set y = ListView2.ListItems.Add(, , rs.Fields(0))
            y.SubItems(1) = rs.Fields(1)
            y.SubItems(2) = Format(rs.Fields(2), "mm/dd/yyyy")
            y.SubItems(3) = rs.Fields(3)
            y.SubItems(4) = rs.Fields(4)
            y.SubItems(5) = rs.Fields(5)
            y.SubItems(6) = rs.Fields(6)
            y.SubItems(7) = rs.Fields(7)
            rs.MoveNext
    Loop
X:
End Sub
