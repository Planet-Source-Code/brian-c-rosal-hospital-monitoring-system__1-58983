VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PHdischarged 
   BackColor       =   &H00000000&
   Caption         =   "Patient(s) Discharging Section "
   ClientHeight    =   8415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11790
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   ForeColor       =   &H00000000&
   Icon            =   "PHdischarged.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "PHdischarged.frx":0442
   ScaleHeight     =   8415
   ScaleWidth      =   11790
   WindowState     =   2  'Maximized
   Begin Project1.chameleonButton butx 
      Height          =   285
      Left            =   90
      TabIndex        =   24
      Top             =   150
      Width           =   285
      _extentx        =   503
      _extenty        =   503
      btype           =   5
      tx              =   "X"
      enab            =   -1  'True
      font            =   "PHdischarged.frx":A727
      coltype         =   2
      focusr          =   -1  'True
      bcol            =   12632256
      bcolo           =   12582912
      fcol            =   0
      fcolo           =   16777215
      mcol            =   12632256
      mptr            =   1
      micon           =   "PHdischarged.frx":A753
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   0
      ngrey           =   0   'False
      fx              =   1
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin Project1.chameleonButton cmdExit 
      Height          =   480
      Left            =   465
      TabIndex        =   15
      Top             =   6765
      Width           =   1815
      _extentx        =   3201
      _extenty        =   847
      btype           =   5
      tx              =   "E&xit"
      enab            =   -1  'True
      font            =   "PHdischarged.frx":A771
      coltype         =   2
      focusr          =   -1  'True
      bcol            =   12632256
      bcolo           =   12632256
      fcol            =   0
      fcolo           =   255
      mcol            =   12632256
      mptr            =   1
      micon           =   "PHdischarged.frx":A795
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   0
      ngrey           =   0   'False
      fx              =   1
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin Project1.chameleonButton cmdcancel 
      Height          =   480
      Left            =   2370
      TabIndex        =   13
      Top             =   6765
      Width           =   1845
      _extentx        =   3254
      _extenty        =   847
      btype           =   5
      tx              =   "&Cancel"
      enab            =   -1  'True
      font            =   "PHdischarged.frx":A7B3
      coltype         =   2
      focusr          =   -1  'True
      bcol            =   12632256
      bcolo           =   12632256
      fcol            =   0
      fcolo           =   255
      mcol            =   12632256
      mptr            =   1
      micon           =   "PHdischarged.frx":A7D7
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   0
      ngrey           =   0   'False
      fx              =   1
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin VB.Frame Frame7 
      Height          =   735
      Left            =   450
      TabIndex        =   54
      Top             =   5940
      Width           =   3765
      Begin Project1.chameleonButton cmdsave 
         Height          =   480
         Left            =   1935
         TabIndex        =   12
         Top             =   165
         Width           =   1665
         _extentx        =   2937
         _extenty        =   847
         btype           =   8
         tx              =   "&Save"
         enab            =   -1  'True
         font            =   "PHdischarged.frx":A7F5
         coltype         =   2
         focusr          =   -1  'True
         bcol            =   12632256
         bcolo           =   12632256
         fcol            =   0
         fcolo           =   255
         mcol            =   12632256
         mptr            =   1
         micon           =   "PHdischarged.frx":A821
         picn            =   "PHdischarged.frx":A83F
         umcol           =   -1  'True
         soft            =   0   'False
         picpos          =   0
         ngrey           =   0   'False
         fx              =   1
         hand            =   0   'False
         check           =   0   'False
         value           =   0   'False
      End
      Begin Project1.chameleonButton cmdmod 
         Height          =   480
         Left            =   120
         TabIndex        =   14
         Top             =   165
         Width           =   1725
         _extentx        =   3043
         _extenty        =   847
         btype           =   8
         tx              =   "&Modify"
         enab            =   -1  'True
         font            =   "PHdischarged.frx":AC93
         coltype         =   2
         focusr          =   -1  'True
         bcol            =   12632256
         bcolo           =   12632256
         fcol            =   0
         fcolo           =   255
         mcol            =   12632256
         mptr            =   1
         micon           =   "PHdischarged.frx":ACBF
         picn            =   "PHdischarged.frx":ACDD
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
   Begin VB.Frame Frame6 
      Caption         =   "Patient's Discharged Transaction Section"
      ForeColor       =   &H00C00000&
      Height          =   3420
      Left            =   4320
      TabIndex        =   53
      Top             =   3825
      Width           =   7185
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
         TabIndex        =   1
         Top             =   540
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtDISdiag 
         Height          =   315
         Left            =   150
         TabIndex        =   3
         Top             =   1125
         Width           =   2895
      End
      Begin VB.Frame Frame8 
         Height          =   2685
         Left            =   3180
         TabIndex        =   55
         Top             =   150
         Width           =   3855
         Begin MSMask.MaskEdBox txtDISmed 
            Height          =   315
            Left            =   1230
            TabIndex        =   6
            Top             =   1005
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Format          =   "#,##0.00;($#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtDISpf 
            Height          =   345
            Left            =   1230
            TabIndex        =   7
            Top             =   1395
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   609
            _Version        =   393216
            BackColor       =   16777215
            Format          =   "#,##0.00;($#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtDISpaid 
            Height          =   375
            Left            =   1245
            TabIndex        =   19
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
            TabIndex        =   5
            Top             =   615
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Format          =   "#,##0.00;($#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtDISrm 
            Height          =   315
            Left            =   1245
            TabIndex        =   4
            Top             =   240
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
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
            TabIndex        =   25
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
            TabIndex        =   8
            Top             =   240
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483624
            Format          =   "#,##0.00;($#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtPAYlab 
            Height          =   315
            Left            =   2565
            TabIndex        =   9
            Top             =   630
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483624
            Format          =   "#,##0.00;($#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtPAYmed 
            Height          =   315
            Left            =   2565
            TabIndex        =   10
            Top             =   1005
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483624
            Format          =   "#,##0.00;($#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtPAYpf 
            Height          =   345
            Left            =   2565
            TabIndex        =   11
            Top             =   1395
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   609
            _Version        =   393216
            BackColor       =   -2147483624
            Format          =   "#,##0.00;($#,##0.00)"
            PromptChar      =   "_"
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
            TabIndex        =   64
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
            TabIndex        =   62
            Top             =   2310
            Width           =   1485
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
            TabIndex        =   61
            Top             =   1440
            Width           =   1740
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
            TabIndex        =   60
            Top             =   1065
            Width           =   1335
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
            TabIndex        =   58
            Top             =   255
            Width           =   1500
         End
      End
      Begin Project1.chameleonButton cmdSrchDis 
         Height          =   735
         Left            =   240
         TabIndex        =   22
         Top             =   2040
         Width           =   1335
         _extentx        =   2302
         _extenty        =   1296
         btype           =   5
         tx              =   "Search D&ischarged  Patient"
         enab            =   -1
         font            =   "PHdischarged.frx":B131
         coltype         =   2
         focusr          =   -1
         bcol            =   12632256
         bcolo           =   12632256
         fcol            =   0
         fcolo           =   255
         mcol            =   12632256
         mptr            =   1
         micon           =   "PHdischarged.frx":B155
         umcol           =   -1
         soft            =   0
         picpos          =   0
         ngrey           =   0
         fx              =   1
         hand            =   0
         check           =   0
         value           =   0
      End
      Begin Project1.chameleonButton chameleonButton1 
         Height          =   735
         Left            =   1650
         TabIndex        =   23
         Top             =   2040
         Width           =   1335
         _extentx        =   2302
         _extenty        =   1296
         btype           =   5
         tx              =   "Discharged  P&atient(s)"
         enab            =   -1
         font            =   "PHdischarged.frx":B173
         coltype         =   2
         focusr          =   -1
         bcol            =   12632256
         bcolo           =   12632256
         fcol            =   0
         fcolo           =   255
         mcol            =   12632256
         mptr            =   1
         micon           =   "PHdischarged.frx":B197
         umcol           =   -1
         soft            =   0
         picpos          =   0
         ngrey           =   0
         fx              =   1
         hand            =   0
         check           =   0
         value           =   0
      End
      Begin MSMask.MaskEdBox txttime 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "h:nn AM/PM"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   4
         EndProperty
         Height          =   330
         Left            =   1470
         TabIndex        =   2
         Top             =   540
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   582
         _Version        =   393216
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##:## ?M"
         PromptChar      =   "_"
      End
      Begin VB.Label Label10 
         Caption         =   "::Press Enter Key in every Entry::"
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   405
         TabIndex        =   69
         Top             =   1605
         Width           =   2370
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H8000000C&
         Height          =   945
         Left            =   150
         Top             =   1920
         Width           =   2910
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Balance :"
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
         Left            =   1560
         TabIndex        =   66
         Top             =   2880
         Width           =   2040
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Balance :"
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
         Left            =   1560
         TabIndex        =   65
         Top             =   2880
         Width           =   2295
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
         Left            =   3435
         TabIndex        =   63
         Top             =   2925
         Width           =   3495
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H008080FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000018&
         FillColor       =   &H80000016&
         FillStyle       =   0  'Solid
         Height          =   420
         Left            =   3180
         Top             =   2910
         Width           =   3840
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
         TabIndex        =   57
         Top             =   900
         Width           =   2295
      End
      Begin VB.Label Label2 
         Caption         =   "Date/Time Discharged:"
         BeginProperty Font 
            Name            =   "Arial"
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
         TabIndex        =   56
         Top             =   300
         Width           =   2415
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "List of Discharged Patients"
      ForeColor       =   &H000000FF&
      Height          =   2115
      Left            =   480
      TabIndex        =   52
      Top             =   3825
      Width           =   3735
      Begin MSComctlLib.ListView ListView2 
         Height          =   1740
         Left            =   120
         TabIndex        =   68
         Top             =   240
         Width           =   3465
         _ExtentX        =   6112
         _ExtentY        =   3069
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
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Age"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Final Diagnosed"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame Frame3 
      Height          =   885
      Left            =   480
      TabIndex        =   51
      Top             =   2925
      Width           =   4470
      Begin Project1.chameleonButton cmdDIS 
         Height          =   585
         Left            =   120
         TabIndex        =   21
         Top             =   195
         Width           =   4245
         _extentx        =   7488
         _extenty        =   1032
         btype           =   3
         tx              =   "&Discharge"
         enab            =   -1
         font            =   "PHdischarged.frx":B1B5
         coltype         =   2
         focusr          =   -1
         bcol            =   12632256
         bcolo           =   12632256
         fcol            =   0
         fcolo           =   255
         mcol            =   12632256
         mptr            =   1
         micon           =   "PHdischarged.frx":B1D9
         picn            =   "PHdischarged.frx":B1F7
         umcol           =   -1
         soft            =   0
         picpos          =   0
         ngrey           =   0
         fx              =   1
         hand            =   0
         check           =   0
         value           =   0
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "List of Confined Patients"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2880
      Left            =   495
      TabIndex        =   43
      Top             =   30
      Width           =   4470
      Begin MSComctlLib.ListView ListView1 
         Height          =   2505
         Left            =   150
         TabIndex        =   20
         Top             =   240
         Width           =   4185
         _ExtentX        =   7382
         _ExtentY        =   4419
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483624
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
            Text            =   "Admission Date/Time"
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
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Age"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Admission  Diagnosed"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Patient Section"
      ForeColor       =   &H00C00000&
      Height          =   2235
      Left            =   5040
      TabIndex        =   35
      Top             =   1575
      Width           =   6435
      Begin Project1.chameleonButton cmdcalc 
         Height          =   330
         Left            =   5355
         TabIndex        =   71
         Top             =   1800
         Width           =   915
         _extentx        =   1614
         _extenty        =   582
         btype           =   5
         tx              =   "Ca&lc"
         enab            =   -1
         font            =   "PHdischarged.frx":B64B
         coltype         =   2
         focusr          =   -1
         bcol            =   12632256
         bcolo           =   12632256
         fcol            =   0
         fcolo           =   255
         mcol            =   12632256
         mptr            =   1
         micon           =   "PHdischarged.frx":B66F
         umcol           =   -1
         soft            =   0
         picpos          =   0
         ngrey           =   0
         fx              =   1
         hand            =   0
         check           =   0
         value           =   0
      End
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
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   3960
         TabIndex        =   70
         Top             =   1800
         Width           =   1245
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
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   2400
         TabIndex        =   50
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox txtpAd 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   255
         Left            =   795
         TabIndex        =   49
         Top             =   1425
         Width           =   5475
      End
      Begin VB.TextBox txtpage 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   255
         Left            =   210
         TabIndex        =   48
         Top             =   1425
         Width           =   495
      End
      Begin VB.TextBox txtpmi 
         Appearance      =   0  'Flat
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
         Left            =   5925
         TabIndex        =   47
         Top             =   915
         Width           =   345
      End
      Begin VB.TextBox txtpfirst 
         Appearance      =   0  'Flat
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
         Left            =   2700
         TabIndex        =   46
         Top             =   915
         Width           =   3150
      End
      Begin VB.TextBox txtplast 
         Appearance      =   0  'Flat
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
         TabIndex        =   45
         Top             =   915
         Width           =   2415
      End
      Begin VB.TextBox txtpno 
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   255
         Left            =   195
         TabIndex        =   44
         Top             =   450
         Width           =   2415
      End
      Begin VB.Shape Shape6 
         BorderColor     =   &H80000003&
         Height          =   360
         Left            =   4140
         Top             =   210
         Width           =   2130
      End
      Begin VB.Shape Shape5 
         BorderColor     =   &H80000005&
         Height          =   360
         Left            =   4155
         Top             =   225
         Width           =   2130
      End
      Begin VB.Label LBLDIS 
         BackColor       =   &H80000004&
         BackStyle       =   0  'Transparent
         Caption         =   ":: DISCHARGED ::"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000000&
         Height          =   255
         Left            =   4305
         TabIndex        =   67
         Top             =   270
         Width           =   1845
      End
      Begin VB.Label Label1 
         Caption         =   "Last Name"
         Height          =   255
         Index           =   0
         Left            =   210
         TabIndex        =   42
         Top             =   705
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "First Name"
         Height          =   210
         Index           =   0
         Left            =   2700
         TabIndex        =   41
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "MI"
         Height          =   195
         Left            =   5895
         TabIndex        =   40
         Top             =   705
         Width           =   255
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
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   210
         TabIndex        =   39
         Top             =   1740
         Width           =   2355
      End
      Begin VB.Label Label2 
         Caption         =   "Admission Diagnosed:"
         Height          =   255
         Index           =   1
         Left            =   795
         TabIndex        =   38
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "Age"
         Height          =   255
         Index           =   2
         Left            =   195
         TabIndex        =   37
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Patient No.:"
         Height          =   255
         Index           =   3
         Left            =   180
         TabIndex        =   36
         Top             =   255
         Width           =   2415
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Member Section"
      ForeColor       =   &H00C00000&
      Height          =   1500
      Left            =   5040
      TabIndex        =   26
      Top             =   60
      Width           =   6405
      Begin VB.TextBox txtMidno 
         BackColor       =   &H80000014&
         Enabled         =   0   'False
         Height          =   285
         Left            =   195
         MaxLength       =   15
         TabIndex        =   30
         Top             =   570
         Width           =   2520
      End
      Begin VB.TextBox txtMmi 
         Enabled         =   0   'False
         Height          =   300
         Left            =   5895
         TabIndex        =   29
         Top             =   1110
         Width           =   360
      End
      Begin VB.TextBox txtMfirst 
         Enabled         =   0   'False
         Height          =   300
         Left            =   2790
         TabIndex        =   28
         Top             =   1110
         Width           =   3030
      End
      Begin VB.TextBox txtMlast 
         Enabled         =   0   'False
         Height          =   285
         Left            =   195
         TabIndex        =   27
         Top             =   1110
         Width           =   2505
      End
      Begin Project1.chameleonButton SrchP 
         Height          =   570
         Left            =   2835
         TabIndex        =   16
         Top             =   225
         Width           =   1620
         _extentx        =   2910
         _extenty        =   1005
         btype           =   5
         tx              =   "Searc&h Confined Patient"
         enab            =   -1
         font            =   "PHdischarged.frx":B68D
         coltype         =   2
         focusr          =   -1
         bcol            =   12632256
         bcolo           =   12632256
         fcol            =   0
         fcolo           =   255
         mcol            =   12632256
         mptr            =   1
         micon           =   "PHdischarged.frx":B6B1
         umcol           =   -1
         soft            =   0
         picpos          =   0
         ngrey           =   0
         fx              =   1
         hand            =   0
         check           =   0
         value           =   0
      End
      Begin Project1.chameleonButton cmdreset 
         Height          =   570
         Left            =   5490
         TabIndex        =   18
         Top             =   225
         Width           =   810
         _extentx        =   1429
         _extenty        =   1005
         btype           =   5
         tx              =   "&Reset"
         enab            =   -1
         font            =   "PHdischarged.frx":B6CF
         coltype         =   2
         focusr          =   -1
         bcol            =   12632256
         bcolo           =   12632256
         fcol            =   0
         fcolo           =   255
         mcol            =   12632256
         mptr            =   1
         micon           =   "PHdischarged.frx":B6F3
         umcol           =   -1
         soft            =   0
         picpos          =   0
         ngrey           =   0
         fx              =   1
         hand            =   0
         check           =   0
         value           =   0
      End
      Begin Project1.chameleonButton cmdConP 
         Height          =   570
         Left            =   4500
         TabIndex        =   17
         Top             =   225
         Width           =   945
         _extentx        =   1667
         _extenty        =   1005
         btype           =   5
         tx              =   "Con&fined Patient(s)"
         enab            =   -1
         font            =   "PHdischarged.frx":B711
         coltype         =   2
         focusr          =   -1
         bcol            =   12632256
         bcolo           =   12632256
         fcol            =   0
         fcolo           =   255
         mcol            =   12632256
         mptr            =   1
         micon           =   "PHdischarged.frx":B735
         umcol           =   -1
         soft            =   0
         picpos          =   0
         ngrey           =   0
         fx              =   1
         hand            =   0
         check           =   0
         value           =   0
      End
      Begin VB.Label Label1 
         Caption         =   "MI"
         Height          =   255
         Index           =   4
         Left            =   5895
         TabIndex        =   34
         Top             =   870
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "First Name"
         Height          =   255
         Index           =   3
         Left            =   2790
         TabIndex        =   33
         Top             =   885
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Last Name"
         Height          =   255
         Index           =   2
         Left            =   195
         TabIndex        =   32
         Top             =   885
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "PH Member ID No."
         Height          =   255
         Index           =   1
         Left            =   210
         TabIndex        =   31
         Top             =   330
         Width           =   2775
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   7440
      Left            =   0
      TabIndex        =   0
      Top             =   -45
      Width           =   11625
      _ExtentX        =   20505
      _ExtentY        =   13123
      MultiRow        =   -1  'True
      TabFixedHeight  =   5
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
Attribute VB_Name = "PHdischarged"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sql, rs, y As Variant, ctr As Byte
Dim sql2 As Variant
Dim MyType, Typex, passString As String
Dim formod As String
Dim sum, SUMpay, Dis As Double



Private Sub butx_Click()
Unload Me
MainForm.Dis.Enabled = True
MainForm.Toolbar1.Buttons.Item(4).Enabled = True
End Sub

Private Sub chameleonButton1_Click()
listDis.Show vbModal
End Sub

Private Sub cmdcalc_Click()
Dim calc As Variant
On Error GoTo TRAPPER
calc = Shell("c:\windows\calc.exe")
AppActivate calc
TRAPPER:
End Sub

Private Sub cmdCancel_Click()
ListView1.Enabled = True
ListView2.Enabled = True
SetText3 False
Color3 False
cmdcancel.Enabled = False
cleartext3
ListView1.SetFocus
cmdDIS.Enabled = True
cmdsave.Enabled = False

End Sub

Private Sub cmdConP_Click()
listcon4.Show vbModal
End Sub

Private Sub cmdDIS_Click()
 If txtMidno.Text = "" Then
        MsgBox "You must select a patient to discharge!!", vbExclamation, ""
        Exit Sub
 End If
sum = 0
Dis = 0
SetText3 True
Color3 True
ListView1.Enabled = False
ListView2.Enabled = False

cmdDIS.Enabled = False
cmdcancel.Enabled = True
txtDISdate.Text = Format(now, "mm/dd/yyyy")
txttime.Text = Format(now, "medium time")
txtDISdate.SetFocus

MyType = "ADD"
End Sub




Private Sub cmdExit_Click()
Unload Me
MainForm.Dis.Enabled = True
MainForm.Toolbar1.Buttons.Item(4).Enabled = True
End Sub

Private Sub cmdmod_Click()
 SetText3 True
 Color3 True
 ListView1.Enabled = False
 ListView2.Enabled = False
 cmdmod.Enabled = False
 cmdcancel.Enabled = True
 MyType = "EDIT"
End Sub

Private Sub cmdreset_Click()
DisplayLstx
DisplayLstx2
clearD
ListView1.Enabled = True
ListView2.Enabled = True
PHdischarged.LBLDIS.Caption = "  "
cmdExit.SetFocus
cmdmod.Enabled = False
End Sub

Private Sub CMDSAVE_Click()
On Error GoTo TRAPPER
    Dim TRS As DAO.Recordset
    Dim TQR As DAO.QueryDef
    Dim P, madzsrch As String
    Dim Query, tb, td As String
    Dim List, delist As ListItem
    Dim X As Long
    Dim Flag As Boolean
    
    '******* for assurance **********************
    If IsDate(txtDISdate.Text) = False Then
     MsgBox "Please Input the proper Discharged Date!", vbCritical, "Attention"
     txtDISdate.SetFocus
     End If
     
    If txtMidno.Text = "" Then
       MsgBox "Invalid Entries!", vbCritical, "Attention"
       GoTo TRAPPER
    End If
    
     If txtDISdiag.Text = "" Then
        MsgBox "It is important to input the Final Diagnosis!", vbCritical, "Check This out"
        txtDISdiag.SetFocus
        GoTo TRAPPER
     End If
     
     If txtDISlab.Text = "" Then
        PHdischarged.txtDISlab.Text = "0.00"
     End If
     
     If txtPAYlab.Text = "" Then
        PHdischarged.txtPAYlab.Text = "0.00"
     End If
     
     If txtDISmed.Text = "" Then
        PHdischarged.txtDISmed.Text = "0.00"
     End If
     
     If txtPAYmed.Text = "" Then
        PHdischarged.txtPAYmed.Text = "0.00"
     End If
           
     If txtDISpf.Text = "" Then
        PHdischarged.txtDISpf.Text = "0.00"
     End If
       
     If txtPAYpf.Text = "" Then
        PHdischarged.txtPAYpf.Text = "0.00"
     End If
       
     If txtDISrm.Text = "" Then
       PHdischarged.txtDISrm.Text = "0.00"
     End If
     
     If txtPAYrm.Text = "" Then
       PHdischarged.txtPAYrm.Text = "0.00"
     End If
     txtDIStot.Text = Val(txtDISrm.Text) + Val(txtDISlab.Text) + Val(txtDISpf.Text) + Val(txtDISmed.Text)
     txtDISpaid.Text = Val(txtPAYrm.Text) + Val(txtPAYlab.Text) + Val(txtPAYpf.Text) + Val(txtPAYmed.Text)
     lblDISdiff.Caption = "Php    " + Format$(Val(txtDIStot.Text) - Val(txtDISpaid.Text), "###,###,###.00")
    '***************************************
    
    If MyType = "ADD" Then
    
      If MsgBox("Are you sure of what you are doing?", vbInformation + vbYesNo, "Confirmation") = vbYes Then
          '/**********Discharged table******************************
          P = "INSERT INTO Pdischarged (patientno,idno,[dateD],Fdiagnose,RmBrd,LabFee,Tmeds,Pfee,PhilPay,DIFF,mlast,mfirst,mMi,plast,pfirst,pmi,page,pdiag,[PdateC],TIMEOUT,TIMEIN,rmpay,labpay,medpay,pfpay,DISTOT) VALUES ('" & txtpno.Text & "','" & txtMidno.Text & "','" & txtDISdate.Text & "','" & txtDISdiag.Text & "','" & txtDISrm.Text & "','" & txtDISlab.Text & "','" & txtDISmed.Text & "','" & txtDISpf.Text & "','" & txtDISpaid.Text & "','" & Dis & "','" & txtMlast.Text & "','" & txtMfirst.Text & "','" & txtMmi.Text & "','" & txtplast.Text & "','" & txtpfirst.Text & "','" & txtpmi.Text & "','" & txtpage.Text & "','" & txtpAd.Text & "','" & txtdate.Text & "','" & txttime.Text & "','" & txtctime.Text & "','" & txtPAYrm.Text & "','" & txtPAYlab.Text & "','" & txtPAYmed.Text & "','" & txtPAYpf.Text & "','" & txtDIStot.Text & "') ;"
          Set TQR = DBMain.CreateQueryDef("", P)
          TQR.Execute
          '**********Removing patient from confined table********************
          passString = "UPDATE Patient SET [DateD]='" & txtDISdate.Text & "' WHERE patient.patientno='" & txtpno.Text & "' ;"
          DBMain.Execute passString
          '/*********Removing patient from listbox*******************
          Set delist = ListView1.FindItem(txtpno.Text, , , lvwPartial)
          ListView1.ListItems.Remove delist.Index
          '/****************************************************************
          Set List = ListView2.ListItems.Add(, , txtpno.Text)
          With List
            .SubItems(1) = txtMidno.Text
            .SubItems(2) = Format(txtDISdate.Text, "mm/dd/yyyy") + " " + Format(txttime.Text, "medium time")
            .SubItems(3) = txtplast.Text
            .SubItems(4) = txtpfirst.Text
            .SubItems(5) = txtpmi.Text
            .SubItems(6) = txtpage.Text
            .SubItems(7) = txtDISdiag.Text
        End With
          
          DisplayLstx
       End If
       
     ElseIf MyType = "EDIT" Then
          If Trim(txtpno.Text) = "" Then
            MsgBox "No record to Modify!", vbCritical, "Attention"
            GoTo TRAPPER
          End If
          '******** SAVE MODIFIED DISCHARGED PATIENTS ******
          P = "UPDATE  Pdischarged SET patientno= '" & txtpno.Text & "',idno='" & txtMidno.Text & "',[dateD]='" & txtDISdate.Text & "',Fdiagnose='" & txtDISdiag.Text & "',RmBrd='" & txtDISrm.Text & "',Rmpay='" & txtPAYrm.Text & "',LabFee='" & txtDISlab.Text & "',Labpay='" & txtPAYlab.Text & "',Tmeds='" & txtDISmed.Text & "',medpay='" & txtPAYmed.Text & "',Pfee='" & txtDISpf.Text & "',Pfpay='" & txtPAYpf.Text & "',PhilPay='" & txtDISpaid.Text & "',DIFF='" & Dis & "',mlast='" & txtMlast.Text & "',mfirst='" & txtMfirst.Text & "',mMi='" & txtMmi.Text & "',plast='" & txtplast.Text & "',pfirst='" & txtpfirst.Text & "',pmi='" & txtpmi.Text & "',page='" & txtpage.Text & "',pdiag='" & txtpAd.Text & "',[PdateC]='" & txtdate.Text & "',TIMEOUT ='" & txttime.Text & "',TIMEIN ='" & txtctime.Text & "',DISTOT='" & txtDIStot.Text & "';"
          Set TQR = DBMain.CreateQueryDef("", P)
          TQR.Execute
          '**********MODIFY patient from confined table********************
          passString = "UPDATE Patient SET [DateD]='" & txtDISdate.Text & "' WHERE patient.patientno='" & txtpno.Text & "' ;"
          DBMain.Execute passString
          DisplayLstx2
          '**************************************************
    End If
    
    '/*********** POST DISCHARGED *************
    PHdischarged.LBLDIS.Caption = ""
    cmdsave.Enabled = False
    ListView1.Enabled = True
    ListView2.Enabled = True
    SetText3 False
    Color3 False
    cmdcancel.Enabled = False
    cleartext3
    ListView1.SetFocus
    cmdDIS.Enabled = True
    Dis = 0
    sum = 0
    '/****************************************
TRAPPER:
End Sub



Private Sub cmdsave_MouseOver()
 If txtDISlab.Text = "" Then
        PHdischarged.txtDISlab.Text = "0.00"
     End If
     
     If txtDISmed.Text = "" Then
        PHdischarged.txtDISmed.Text = "0.00"
     End If
     
     If txtDISpaid.Text = "" Then
        PHdischarged.txtDISpaid.Text = "0.00"
     End If
       
     If txtDISpf.Text = "" Then
        PHdischarged.txtDISpf.Text = "0.00"
     End If
       
     If txtDISrm.Text = "" Then
       PHdischarged.txtDISrm.Text = "0.00"
     End If
     
     txtDIStot.Text = Val(txtDISrm.Text) + Val(txtDISlab.Text) + Val(txtDISpf.Text) + Val(txtDISmed.Text)
     lblDISdiff.Caption = "Php    " + Format$(Val(txtDIStot.Text) - Val(txtDISpaid.Text), "###,###,###.00")
    '***************************************
  
End Sub

Private Sub cmdSrchDis_Click()
cmdmod.Enabled = False
cmdreset.Enabled = True
SFDisP.Show vbModal
Typex = "Dis"
End Sub

Private Sub Form_Load()
  Dim testdate As String
  testdate = ""
  add3d
  Set WSMain = DBEngine.Workspaces(0)
  Set DBMain = WSMain.OpenDatabase(App.Path + "\hospital.mdb", False, False, ";pwd=scanhead")
  '/*************** Confined patients ACCEPTS NOID***************************
  sql = "select patientno,idno,datec,plastname,pfirstname,pmi,page,padiagnose,timein from Patient where ISNULL(dateD) = TRUE order by plastname,pfirstname,pmi;"
    
           Set rs = DBMain.OpenRecordset(sql)
           If rs.RecordCount = 0 Then
               cmdDIS.Enabled = False
           End If
                
                Do Until rs.EOF
                    Set y = ListView1.ListItems.Add(, , rs.Fields(0))
                        If IsNull(rs.Fields(1)) = True Then
                          y.SubItems(1) = "NON-Med"
                        Else
                          y.SubItems(1) = rs.Fields(1)
                        End If
                        y.SubItems(2) = Format(rs.Fields(2), "mm/dd/yyyy") + " " + Format(rs.Fields(8), "medium time")
                        y.SubItems(3) = rs.Fields(3)
                        y.SubItems(4) = rs.Fields(4)
                        y.SubItems(5) = rs.Fields(5)
                        y.SubItems(6) = rs.Fields(6)
                        y.SubItems(7) = rs.Fields(7)
                        rs.MoveNext
                 Loop
 '/*****************Discharged Patients NOID ACCEPTED ***********************
 sql = "select Pdischarged.patientno,Pdischarged.idno,Pdischarged.dateD,Patient.plastname,Patient.pfirstname,Patient.pmi,Patient.page,Pdischarged.fdiagnose,pdischarged.timeout FROM Patient INNER JOIN Pdischarged ON  PDISCHARGED.PATIENTNO = PATIENT.PATIENTNO  AND PDISCHARGED.dateD = PATIENT.dateD order by patient.plastname,patient.pfirstname,patient.pmi;"
           Set rs = DBMain.OpenRecordset(sql)
                Do Until rs.EOF
                    Set y = ListView2.ListItems.Add(, , rs.Fields(0))
                        If IsNull(rs.Fields(1)) = True Then
                          y.SubItems(1) = "NON-Med"
                        Else
                          y.SubItems(1) = rs.Fields(1)
                        End If
                        y.SubItems(2) = Format(rs.Fields(2), "mm/dd/yyyy") + " " + Format(rs.Fields(8), "medium time")
                        y.SubItems(3) = rs.Fields(3)
                        y.SubItems(4) = rs.Fields(4)
                        y.SubItems(5) = rs.Fields(5)
                        y.SubItems(6) = rs.Fields(6)
                        y.SubItems(7) = rs.Fields(7)
                        rs.MoveNext
                 Loop

 
 
 
 PHdischarged.LBLDIS.ForeColor = &H8000000F
 cleartext3
 SetText3 False
 Color3 False
 cmdsave.Enabled = False
 cmdmod.Enabled = False
 cmdcancel.Enabled = False
 
End Sub



Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo TRAPPER
    Dim testdate As String
    Dim X As Long
    Dim TQR As DAO.QueryDef
    Dim TRS As DAO.Recordset
   'Scanhead rules
   Color3 False
   testdate = ""
   cmdDIS.Enabled = True
   cmdmod.Enabled = False
   cleartext3
   Set TQR = DBMain.CreateQueryDef("", "select patientno,idno,datec,plastname,pfirstname,pmi,page,padiagnose,mlast,mfirst,Mmi,timein from Patient where patientno ='" & ListView1.SelectedItem.Text & "' AND ISNULL(dateD) = TRUE order by plastname,pfirstname,pmi")
   Set TRS = TQR.OpenRecordset()
   If IsNull(TRS.Fields(1)) = True Then
      
      txtMidno.Text = "Non-Med"
      txtMlast.Text = "-----------"
      txtMfirst.Text = "-----------"
      txtMmi.Text = "-----------"
      txtpno.Text = TRS.Fields(0)
      txtplast.Text = TRS.Fields(3)
      txtpfirst.Text = TRS.Fields(4)
      txtpmi.Text = TRS.Fields(5)
      txtdate.Text = Format(TRS.Fields(2), "mm/dd/yyyy")
      txtctime.Text = Format(TRS.Fields(11), "medium time")
      txtpage.Text = TRS.Fields(6)
      txtpAd.Text = TRS.Fields(7)
    Else
      txtMidno.Text = TRS.Fields(1)
      txtMlast.Text = TRS.Fields(8)
      txtMfirst.Text = TRS.Fields(9)
      txtMmi.Text = TRS.Fields(10)
      txtpno.Text = TRS.Fields(0)
      txtplast.Text = TRS.Fields(3)
      txtpfirst.Text = TRS.Fields(4)
      txtpmi.Text = TRS.Fields(5)
      txtdate.Text = Format(TRS.Fields(2), "mm/dd/yyyy")
      txtctime.Text = Format(TRS.Fields(11), "medium time")
      txtpage.Text = TRS.Fields(6)
      txtpAd.Text = TRS.Fields(7)
    End If
    PHdischarged.LBLDIS.Caption = " :: CONFINED :: "
    PHdischarged.LBLDIS.ForeColor = &HFF0000

TRAPPER:
End Sub




Private Sub ListView2_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim TQR As DAO.QueryDef
Dim TRS As DAO.Recordset
cmdDIS.Enabled = False
cmdmod.Enabled = True
Color3 True
 Set TQR = DBMain.CreateQueryDef("", "SELECT Pdischarged.dateD,Pdischarged.Fdiagnose,Pdischarged.rmbrd,Pdischarged.labfee,Pdischarged.Tmeds,Pdischarged.Pfee,Pdischarged.Philpay,Pdischarged.Diff,Pdischarged.idno,Pdischarged.Mlast,Pdischarged.Mfirst,Pdischarged.Mmi,Pdischarged.Plast,Pdischarged.Pfirst,Pdischarged.Pmi,Pdischarged.Page,Pdischarged.Pdiag,Pdischarged.patientno,Pdischarged.pdateC,Pdischarged.TIMEOUT,Pdischarged.TIMEIN,Pdischarged.rmpay,Pdischarged.labpay,Pdischarged.medpay,Pdischarged.pfpay FROM Pdischarged,Patient WHERE Pdischarged.patientno ='" & ListView2.SelectedItem.Text & "' AND ISNULL(Patient.dateD) = FALSE  ORDER BY pdischarged.mlast,pdischarged.mfirst,pdischarged.mmi")
 Set TRS = TQR.OpenRecordset()
 txtDISdate.Text = Format(TRS.Fields(0), "mm/dd/yyyy")
 txtDISdiag.Text = TRS.Fields(1)
 txtDISrm.Text = TRS.Fields(2)
 txtDISlab.Text = TRS.Fields(3)
 txtDISmed.Text = TRS.Fields(4)
 txtDISpf.Text = TRS.Fields(5)
 txtDISpaid.Text = TRS.Fields(6)
 If IsNull(TRS.Fields(8)) = True Then
    txtMidno.Text = "Non-Med"
 Else
    txtMidno.Text = TRS.Fields(8)
 End If
    
 txtMlast.Text = TRS.Fields(9)
 txtMfirst.Text = TRS.Fields(10)
 txtMmi.Text = TRS.Fields(11)
 txtpno.Text = TRS.Fields(17)
 txtplast.Text = TRS.Fields(12)
 txtpfirst.Text = TRS.Fields(13)
 txtpmi.Text = TRS.Fields(14)
 txtpage.Text = TRS.Fields(15)
 txtpAd.Text = TRS.Fields(16)
 txtdate.Text = Format(TRS.Fields(18), "mm/dd/yyyy")
 txttime.Text = Format(TRS.Fields(19), "medium time")
 txtctime.Text = Format(TRS.Fields(20), "medium time")
 txtPAYrm.Text = TRS.Fields(21)
 txtPAYlab.Text = TRS.Fields(22)
 txtPAYmed.Text = TRS.Fields(23)
 txtPAYpf.Text = TRS.Fields(24)
 
 txtDIStot.Text = Format$(Val(TRS.Fields(2)) + Val(TRS.Fields(3)) + Val(TRS.Fields(4)) + Val(TRS.Fields(5)), "###,###,###.00")
 txtDISpaid.Text = Format$(Val(TRS.Fields(21)) + Val(TRS.Fields(22)) + Val(TRS.Fields(23)) + Val(TRS.Fields(24)), "###,###,###.00")

 lblDISdiff.Caption = "Php " + Format$(Val(TRS.Fields(7)), "###,###,###.00")
 
 PHdischarged.LBLDIS.Caption = " :: DISCHARGED :: "
 PHdischarged.LBLDIS.ForeColor = &HFF&
End Sub

Private Sub SrchP_Click()
cmdreset.Enabled = True
SrchFormC.Show vbModal
Typex = "Con"
End Sub

Private Sub txtDISdate_GotFocus()
With txtDISdate
.SelStart = 0
.SelLength = Len(txtDISdate.Text)
End With
End Sub









Private Sub txtDISdiag_LostFocus()
cmdsave.Enabled = True
End Sub

Private Sub txtDISlab_GotFocus()
With txtDISlab
.SelStart = 0
.SelLength = Len(txtDISlab.Text)
End With
End Sub

Private Sub txtDISlab_LOSTFocus()
sum = Val(txtDISrm.Text) + Val(txtDISlab.Text) + Val(txtDISpf.Text) + Val(txtDISmed.Text)
SUMpay = Val(txtPAYrm.Text) + Val(txtPAYlab.Text) + Val(txtPAYpf.Text) + Val(txtPAYmed.Text)
txtDIStot.Text = sum
txtDISpaid.Text = SUMpay
Dis = Val(txtDIStot.Text) - Val(txtDISpaid.Text)
lblDISdiff.Caption = "Php    " + Format$(Val(txtDIStot.Text) - Val(txtDISpaid.Text), "###,###,###.00")
cmdsave.Enabled = True
End Sub

Private Sub txtDISlab_KeyPress(KeyAscii As Integer)
Dim madzbry As String
madzbry = "0123456789."
          If KeyAscii > 26 Then
            If InStr(madzbry, Chr(KeyAscii)) = 0 Then
              KeyAscii = 0
            End If
         End If '


End Sub

Private Sub txtDISmed_GotFocus()
With txtDISmed
.SelStart = 0
.SelLength = Len(txtDISmed.Text)
End With
End Sub

Private Sub txtDISmed_LOSTFocus()
sum = Val(txtDISrm.Text) + Val(txtDISlab.Text) + Val(txtDISpf.Text) + Val(txtDISmed.Text)
SUMpay = Val(txtPAYrm.Text) + Val(txtPAYlab.Text) + Val(txtPAYpf.Text) + Val(txtPAYmed.Text)
txtDIStot.Text = sum
txtDISpaid.Text = SUMpay
Dis = Val(txtDIStot.Text) - Val(txtDISpaid.Text)
lblDISdiff.Caption = "Php    " + Format$(Val(txtDIStot.Text) - Val(txtDISpaid.Text), "###,###,###.00")
cmdsave.Enabled = True
End Sub

Private Sub txtDISmed_KeyPress(KeyAscii As Integer)
Dim madzbry As String
madzbry = "0123456789."
          If KeyAscii > 26 Then
            If InStr(madzbry, Chr(KeyAscii)) = 0 Then
              KeyAscii = 0
            End If
         End If


End Sub

Private Sub txtDISpaid_GotFocus()
With txtDISpaid
.SelStart = 0
.SelLength = Len(txtDISpaid.Text)
End With
End Sub





Private Sub txtDISpf_GotFocus()
With txtDISpf
.SelStart = 0
.SelLength = Len(txtDISpf.Text)
End With
End Sub
Private Sub txtDISpf_LOSTFocus()
sum = Val(txtDISrm.Text) + Val(txtDISlab.Text) + Val(txtDISpf.Text) + Val(txtDISmed.Text)
SUMpay = Val(txtPAYrm.Text) + Val(txtPAYlab.Text) + Val(txtPAYpf.Text) + Val(txtPAYmed.Text)
txtDIStot.Text = sum
txtDISpaid.Text = SUMpay
Dis = Val(txtDIStot.Text) - Val(txtDISpaid.Text)
lblDISdiff.Caption = "Php    " + Format$(Val(txtDIStot.Text) - Val(txtDISpaid.Text), "###,###,###.00")
cmdsave.Enabled = True

End Sub

Private Sub txtDISpf_KeyPress(KeyAscii As Integer)
Dim madzbry As String
madzbry = "0123456789."
          If KeyAscii > 26 Then
            If InStr(madzbry, Chr(KeyAscii)) = 0 Then
              KeyAscii = 0
            End If
         End If
End Sub


Private Sub txtDISrm_GotFocus()
With txtDISrm
.SelStart = 0
.SelLength = Len(txtDISrm.Text)
End With
End Sub



Private Sub txtDISrm_LOSTFocus()
sum = Val(txtDISrm.Text) + Val(txtDISlab.Text) + Val(txtDISpf.Text) + Val(txtDISmed.Text)
SUMpay = Val(txtPAYrm.Text) + Val(txtPAYlab.Text) + Val(txtPAYpf.Text) + Val(txtPAYmed.Text)
txtDIStot.Text = sum
txtDISpaid.Text = SUMpay
Dis = Val(txtDIStot.Text) - Val(txtDISpaid.Text)
lblDISdiff.Caption = "Php    " + Format$(Val(txtDIStot.Text) - Val(txtDISpaid.Text), "###,###,###.00")
cmdsave.Enabled = True
End Sub


Private Sub txtDISrm_KeyPress(KeyAscii As Integer)
Dim madzbry As String
madzbry = "0123456789."
          If KeyAscii > 26 Then
            If InStr(madzbry, Chr(KeyAscii)) = 0 Then
              KeyAscii = 0
            End If
         End If


End Sub

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

Private Function add3d()
  Add3DBorder Me
  Add3DBorder ListView1
  Add3DBorder ListView2
  Add3DBorder txtMidno
  Add3DBorder txtMlast
  Add3DBorder txtMfirst
  Add3DBorder txtMmi
  Add3DBorder txtpno
  Add3DBorder txtplast
  Add3DBorder txtpfirst
  Add3DBorder txtpmi
  Add3DBorder txtpage
  Add3DBorder txtpAd
  Add3DBorder txtdate
  Add3DBorder txtDISrm
  Add3DBorder txtDISdate
  Add3DBorder txtDISlab
  Add3DBorder txtDISdiag
  Add3DBorder txtDISmed
  Add3DBorder txtDISpaid
  Add3DBorder txtDISpf
  Add3DBorder txttime
  Add3DBorder txtctime
  
  Add3DBorder txtPAYrm
  Add3DBorder txtPAYpf
  Add3DBorder txtPAYlab
  Add3DBorder txtPAYmed
  
End Function

Public Sub DisplayLstx() 'Reset control display: gotcha the scanhedbri
On Error GoTo X
    
    Dim TQR As DAO.QueryDef
    Dim rs As DAO.Recordset
    Dim y As ListItem
     Dim sql As String
    Dim X As Long
    Dim testdate As String
    
    testdate = ""
   
    '/*************** Confined patients ***************************
   sql = "select patientno,idno,datec,plastname,pfirstname,pmi,page,padiagnose,timein from Patient where ISNULL(dateD) = TRUE order by plastname,pfirstname,pmi;"
          Set rs = DBMain.OpenRecordset(sql)
           ListView1.ListItems.Clear
                Do Until rs.EOF
                    Set y = ListView1.ListItems.Add(, , rs.Fields(0))
                        If IsNull(rs.Fields(1)) = True Then
                          y.SubItems(1) = "NON-Med"
                        Else
                          y.SubItems(1) = rs.Fields(1)
                        End If
                        y.SubItems(2) = Format(rs.Fields(2), "mm/dd/yyyy") + " " + Format(rs.Fields(8), "medium time")
                        y.SubItems(3) = rs.Fields(3)
                        y.SubItems(4) = rs.Fields(4)
                        y.SubItems(5) = rs.Fields(5)
                        y.SubItems(6) = rs.Fields(6)
                        y.SubItems(7) = rs.Fields(7)
                        rs.MoveNext
                 Loop
   clearD
  SetText3 False
X:
End Sub

'**** for discharged patients *******
Public Sub DisplayLstx2() 'Reset control display: gotcha the scanhedbri
On Error GoTo X
    
    Dim TQR As DAO.QueryDef
    Dim rs As DAO.Recordset
    Dim y As ListItem
     Dim sql As String
    Dim X As Long
    Dim testdate As String
    '/*************** Discharged patients NOID ACCEPTED ***************************
  sql = "select Pdischarged.patientno,Pdischarged.idno,Pdischarged.dateD,Patient.plastname,Patient.pfirstname,Patient.pmi,Patient.page,Pdischarged.fdiagnose,pdischarged.timeout FROM Patient INNER JOIN Pdischarged ON  PDISCHARGED.PATIENTNO = PATIENT.PATIENTNO  AND PDISCHARGED.dateD = PATIENT.dateD  order by patient.plastname,patient.pfirstname,patient.pmi;"
           Set rs = DBMain.OpenRecordset(sql)
            ListView2.ListItems.Clear
                Do Until rs.EOF
                    Set y = ListView2.ListItems.Add(, , rs.Fields(0))
                        If IsNull(rs.Fields(1)) = True Then
                          y.SubItems(1) = "NON-Med"
                        Else
                          y.SubItems(1) = rs.Fields(1)
                        End If
                        
                        y.SubItems(2) = Format(rs.Fields(2), "mm/dd/yyyy") + " " + Format(rs.Fields(8), "medium time")
                        y.SubItems(3) = rs.Fields(3)
                        y.SubItems(4) = rs.Fields(4)
                        y.SubItems(5) = rs.Fields(5)
                        y.SubItems(6) = rs.Fields(6)
                        y.SubItems(7) = rs.Fields(7)
                        
                        rs.MoveNext
                 Loop

   clearD
   SetText3 False
X:
End Sub




Private Sub txtPAYlab_gotfocus()
With txtPAYlab
.SelStart = 0
.SelLength = Len(txtPAYlab.Text)
End With
End Sub
Private Sub txtpaylab_KeyPress(KeyAscii As Integer)
Dim madzbry As String
madzbry = "0123456789."
          If KeyAscii > 26 Then
            If InStr(madzbry, Chr(KeyAscii)) = 0 Then
              KeyAscii = 0
            End If
         End If
End Sub

Private Sub txtPAYlab_LostFocus()
If Val(txtPAYlab.Text) > Val(txtDISlab.Text) Then
     MsgBox "PH Paid for LABORATORY FEE is Greater than the indicated amount.", vbCritical, "Attention"
     txtPAYlab.SetFocus
Else
sum = Val(txtDISrm.Text) + Val(txtDISlab.Text) + Val(txtDISpf.Text) + Val(txtDISmed.Text)
SUMpay = Val(txtPAYrm.Text) + Val(txtPAYlab.Text) + Val(txtPAYpf.Text) + Val(txtPAYmed.Text)
txtDIStot.Text = sum
txtDISpaid.Text = SUMpay
Dis = Val(txtDIStot.Text) - Val(txtDISpaid.Text)
lblDISdiff.Caption = "Php    " + Format$(Val(txtDIStot.Text) - Val(txtDISpaid.Text), "###,###,###.00")
cmdsave.Enabled = True
End If

End Sub

Private Sub txtPAYmed_gotfocus()
With txtPAYmed
.SelStart = 0
.SelLength = Len(txtPAYmed.Text)
End With
End Sub
Private Sub txtpaymed_KeyPress(KeyAscii As Integer)
Dim madzbry As String
madzbry = "0123456789."
          If KeyAscii > 26 Then
            If InStr(madzbry, Chr(KeyAscii)) = 0 Then
              KeyAscii = 0
            End If
         End If
End Sub
Private Sub txtPAYmed_LostFocus()
If Val(txtPAYmed.Text) > Val(txtDISmed.Text) Then
     MsgBox "PH Paid for TOTAL MED. is Greater than the indicated amount.", vbCritical, "Attention"
     txtPAYmed.SetFocus
Else
sum = Val(txtDISrm.Text) + Val(txtDISlab.Text) + Val(txtDISpf.Text) + Val(txtDISmed.Text)
SUMpay = Val(txtPAYrm.Text) + Val(txtPAYlab.Text) + Val(txtPAYpf.Text) + Val(txtPAYmed.Text)
txtDIStot.Text = sum
txtDISpaid.Text = SUMpay
Dis = Val(txtDIStot.Text) - Val(txtDISpaid.Text)
lblDISdiff.Caption = "Php    " + Format$(Val(txtDIStot.Text) - Val(txtDISpaid.Text), "###,###,###.00")
cmdsave.Enabled = True
End If
End Sub

Private Sub txtPAYpf_GotFocus()
With txtPAYpf
.SelStart = 0
.SelLength = Len(txtPAYpf.Text)
End With
End Sub
Private Sub txtpaypf_KeyPress(KeyAscii As Integer)
Dim madzbry As String
madzbry = "0123456789."
          If KeyAscii > 26 Then
            If InStr(madzbry, Chr(KeyAscii)) = 0 Then
              KeyAscii = 0
            End If
         End If
End Sub

Private Sub txtPAYpf_LostFocus()
If Val(txtPAYpf.Text) > Val(txtDISpf.Text) Then
     MsgBox "PH Paid for PROF. FEE is Greater than the indicated amount.", vbCritical, "Attention"
     txtPAYpf.SetFocus
Else
sum = Val(txtDISrm.Text) + Val(txtDISlab.Text) + Val(txtDISpf.Text) + Val(txtDISmed.Text)
SUMpay = Val(txtPAYrm.Text) + Val(txtPAYlab.Text) + Val(txtPAYpf.Text) + Val(txtPAYmed.Text)
txtDIStot.Text = sum
txtDISpaid.Text = SUMpay
Dis = Val(txtDIStot.Text) - Val(txtDISpaid.Text)
lblDISdiff.Caption = "Php    " + Format$(Val(txtDIStot.Text) - Val(txtDISpaid.Text), "###,###,###.00")
cmdsave.Enabled = True
End If
End Sub

Private Sub txtPAYrm_GotFocus()
With txtPAYrm
.SelStart = 0
.SelLength = Len(txtPAYrm.Text)
End With
End Sub

Private Sub txtpayrm_KeyPress(KeyAscii As Integer)
Dim madzbry As String
madzbry = "0123456789."
          If KeyAscii > 26 Then
            If InStr(madzbry, Chr(KeyAscii)) = 0 Then
              KeyAscii = 0
            End If
         End If
End Sub



Private Sub txtPAYrm_LostFocus()
If Val(txtPAYrm.Text) > Val(txtDISrm.Text) Then
     MsgBox "PH Paid for RM/BOARDING is Greater than the indicated amount.", vbCritical, "Attention"
     txtPAYrm.SetFocus
     
Else
sum = Val(txtDISrm.Text) + Val(txtDISlab.Text) + Val(txtDISpf.Text) + Val(txtDISmed.Text)
SUMpay = Val(txtPAYrm.Text) + Val(txtPAYlab.Text) + Val(txtPAYpf.Text) + Val(txtPAYmed.Text)
txtDIStot.Text = sum
txtDISpaid.Text = SUMpay
Dis = Val(txtDIStot.Text) - Val(txtDISpaid.Text)
lblDISdiff.Caption = "Php    " + Format$(Val(txtDIStot.Text) - Val(txtDISpaid.Text), "###,###,###.00")
cmdsave.Enabled = True
End If
TRAPPER:
End Sub

Private Sub txtTIME_GotFocus()
With txttime
.SelStart = 0
.SelLength = Len(txttime.Text)
End With
End Sub

