VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PHconfined 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "Patient(s) Confinement Section "
   ClientHeight    =   7680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10680
   ControlBox      =   0   'False
   DrawStyle       =   1  'Dash
   Icon            =   "PHconfined.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   Picture         =   "PHconfined.frx":08CA
   ScaleHeight     =   7680
   ScaleWidth      =   10680
   WindowState     =   2  'Maximized
   Begin Project1.chameleonButton butx 
      Height          =   285
      Left            =   150
      TabIndex        =   41
      Top             =   240
      Width           =   285
      _extentx        =   503
      _extenty        =   503
      btype           =   5
      tx              =   "X"
      enab            =   -1  'True
      font            =   "PHconfined.frx":ABAF
      coltype         =   2
      focusr          =   -1  'True
      bcol            =   12632256
      bcolo           =   12582912
      fcol            =   0
      fcolo           =   16777215
      mcol            =   12632256
      mptr            =   1
      micon           =   "PHconfined.frx":ABDB
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
      Height          =   825
      Left            =   8910
      TabIndex        =   15
      Top             =   5835
      Width           =   1425
      _extentx        =   2514
      _extenty        =   1455
      btype           =   5
      tx              =   "E&xit"
      enab            =   -1  'True
      font            =   "PHconfined.frx":ABF9
      coltype         =   2
      focusr          =   -1  'True
      bcol            =   12632256
      bcolo           =   12632256
      fcol            =   0
      fcolo           =   255
      mcol            =   12632256
      mptr            =   1
      micon           =   "PHconfined.frx":AC1D
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
      Height          =   825
      Left            =   7380
      TabIndex        =   12
      Top             =   5835
      Width           =   1425
      _extentx        =   2514
      _extenty        =   1455
      btype           =   5
      tx              =   "&Cancel"
      enab            =   -1  'True
      font            =   "PHconfined.frx":AC3B
      coltype         =   2
      focusr          =   -1  'True
      bcol            =   12632256
      bcolo           =   12632256
      fcol            =   0
      fcolo           =   255
      mcol            =   12632256
      mptr            =   1
      micon           =   "PHconfined.frx":AC5F
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   0
      ngrey           =   0   'False
      fx              =   1
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin VB.Frame Frame5 
      Height          =   915
      Left            =   525
      TabIndex        =   39
      Top             =   5745
      Width           =   5715
      Begin Project1.chameleonButton cmdreset 
         Height          =   600
         Left            =   4635
         TabIndex        =   14
         Top             =   225
         Width           =   960
         _extentx        =   1693
         _extenty        =   1058
         btype           =   5
         tx              =   "&Reset"
         enab            =   -1  'True
         font            =   "PHconfined.frx":AC7D
         coltype         =   2
         focusr          =   -1  'True
         bcol            =   12632256
         bcolo           =   12632256
         fcol            =   0
         fcolo           =   255
         mcol            =   12632256
         mptr            =   1
         micon           =   "PHconfined.frx":ACA9
         umcol           =   -1  'True
         soft            =   0   'False
         picpos          =   0
         ngrey           =   0   'False
         fx              =   1
         hand            =   0   'False
         check           =   0   'False
         value           =   0   'False
      End
      Begin Project1.chameleonButton cmdsave 
         Height          =   555
         Left            =   2880
         TabIndex        =   11
         Top             =   240
         Width           =   1260
         _extentx        =   2223
         _extenty        =   979
         btype           =   8
         tx              =   "&Save"
         enab            =   -1  'True
         font            =   "PHconfined.frx":ACC7
         coltype         =   2
         focusr          =   -1  'True
         bcol            =   12632256
         bcolo           =   12632256
         fcol            =   0
         fcolo           =   255
         mcol            =   12632256
         mptr            =   1
         micon           =   "PHconfined.frx":ACF3
         picn            =   "PHconfined.frx":AD11
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
         Height          =   555
         Left            =   1515
         TabIndex        =   13
         Top             =   210
         Width           =   1260
         _extentx        =   2223
         _extenty        =   979
         btype           =   8
         tx              =   "&Modify"
         enab            =   -1  'True
         font            =   "PHconfined.frx":B165
         coltype         =   2
         focusr          =   -1  'True
         bcol            =   12632256
         bcolo           =   12632256
         fcol            =   0
         fcolo           =   255
         mcol            =   12632256
         mptr            =   1
         micon           =   "PHconfined.frx":B191
         picn            =   "PHconfined.frx":B1AF
         umcol           =   -1  'True
         soft            =   0   'False
         picpos          =   0
         ngrey           =   0   'False
         fx              =   1
         hand            =   0   'False
         check           =   0   'False
         value           =   0   'False
      End
      Begin Project1.chameleonButton cmdadd 
         Height          =   555
         Left            =   150
         TabIndex        =   2
         Top             =   225
         Width           =   1260
         _extentx        =   2223
         _extenty        =   979
         btype           =   8
         tx              =   "&New Confine"
         enab            =   -1  'True
         font            =   "PHconfined.frx":B603
         coltype         =   2
         focusr          =   -1  'True
         bcol            =   12632256
         bcolo           =   12632256
         fcol            =   0
         fcolo           =   255
         mcol            =   12632256
         mptr            =   1
         micon           =   "PHconfined.frx":B62F
         picn            =   "PHconfined.frx":B64D
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
   Begin VB.Frame Frame4 
      Caption         =   "Confined Patient(s) List"
      ForeColor       =   &H000000FF&
      Height          =   2880
      Left            =   540
      TabIndex        =   32
      Top             =   2790
      Width           =   3735
      Begin MSComctlLib.ListView ListView1 
         Height          =   2505
         Left            =   150
         TabIndex        =   20
         Top             =   240
         Width           =   3435
         _ExtentX        =   6059
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
            Text            =   "Date Confined"
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
            Text            =   "Admission  Diagnosed"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Patient's Confinement Data"
      ForeColor       =   &H00C00000&
      Height          =   2880
      Left            =   4455
      TabIndex        =   25
      Top             =   2790
      Width           =   5940
      Begin VB.TextBox txtcon 
         Height          =   300
         Index           =   6
         Left            =   900
         TabIndex        =   8
         Top             =   1785
         Width           =   4890
      End
      Begin VB.TextBox txtcon 
         Height          =   270
         Index           =   5
         Left            =   225
         MaxLength       =   3
         TabIndex        =   7
         Top             =   1785
         Width           =   450
      End
      Begin VB.TextBox txtcon 
         Height          =   285
         Index           =   4
         Left            =   5400
         MaxLength       =   1
         TabIndex        =   6
         Top             =   1140
         Width           =   390
      End
      Begin VB.TextBox txtcon 
         Height          =   285
         Index           =   3
         Left            =   3000
         TabIndex        =   5
         Top             =   1140
         Width           =   2310
      End
      Begin VB.TextBox txtcon 
         Height          =   285
         Index           =   2
         Left            =   225
         TabIndex        =   4
         Top             =   1140
         Width           =   2655
      End
      Begin MSMask.MaskEdBox txtdate 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "M/d/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   375
         Left            =   450
         TabIndex        =   9
         Top             =   2415
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   661
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
      Begin VB.TextBox txtcon 
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   225
         TabIndex        =   3
         Top             =   570
         Width           =   1590
      End
      Begin Project1.chameleonButton cmddiag 
         Height          =   690
         Left            =   4575
         TabIndex        =   18
         Top             =   195
         Width           =   1215
         _extentx        =   1905
         _extenty        =   529
         btype           =   5
         tx              =   "Con&fined Patient(s)"
         enab            =   -1  'True
         font            =   "PHconfined.frx":BF29
         coltype         =   2
         focusr          =   -1  'True
         bcol            =   12632256
         bcolo           =   12632256
         fcol            =   0
         fcolo           =   255
         mcol            =   12632256
         mptr            =   1
         micon           =   "PHconfined.frx":BF4D
         umcol           =   -1  'True
         soft            =   0   'False
         picpos          =   0
         ngrey           =   0   'False
         fx              =   1
         hand            =   0   'False
         check           =   0   'False
         value           =   0   'False
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
         Height          =   375
         Left            =   1935
         TabIndex        =   10
         Top             =   2415
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   661
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
      Begin Project1.chameleonButton cmdsrch2 
         Height          =   690
         Left            =   3240
         TabIndex        =   17
         Top             =   195
         Width           =   1245
         _extentx        =   1905
         _extenty        =   529
         btype           =   5
         tx              =   "Searc&h Patient(s)"
         enab            =   -1  'True
         font            =   "PHconfined.frx":BF6B
         coltype         =   2
         focusr          =   -1  'True
         bcol            =   12632256
         bcolo           =   12632256
         fcol            =   0
         fcolo           =   255
         mcol            =   12632256
         mptr            =   1
         micon           =   "PHconfined.frx":BF8F
         umcol           =   -1  'True
         soft            =   0   'False
         picpos          =   0
         ngrey           =   0   'False
         fx              =   1
         hand            =   0   'False
         check           =   0   'False
         value           =   0   'False
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "PH-Med"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   0
         Left            =   2070
         TabIndex        =   43
         Top             =   585
         Width           =   1020
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Note: This Form is exclusived for PH-Med Entries only."
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
         Height          =   420
         Index           =   1
         Left            =   3615
         TabIndex        =   42
         Top             =   2280
         Width           =   2085
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H80000003&
         Height          =   555
         Left            =   3435
         Top             =   2220
         Width           =   2385
      End
      Begin VB.Label Label1 
         Caption         =   "Patient No.:"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   40
         Top             =   330
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Admission Date/Time:"
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
         Height          =   240
         Index           =   11
         Left            =   240
         TabIndex        =   38
         Top             =   2160
         Width           =   2130
      End
      Begin VB.Label Label1 
         Caption         =   "Admission Diagnosis:"
         Height          =   255
         Index           =   10
         Left            =   840
         TabIndex        =   37
         Top             =   1545
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Age:"
         Height          =   255
         Index           =   9
         Left            =   240
         TabIndex        =   36
         Top             =   1530
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "MI:"
         Height          =   255
         Index           =   8
         Left            =   5415
         TabIndex        =   35
         Top             =   900
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "First Name:"
         Height          =   255
         Index           =   7
         Left            =   3000
         TabIndex        =   34
         Top             =   900
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Last Name:"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   33
         Top             =   900
         Width           =   1575
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000005&
         FillColor       =   &H00C00000&
         FillStyle       =   0  'Solid
         Height          =   315
         Left            =   1920
         Top             =   555
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Member Identification Section"
      ForeColor       =   &H00C00000&
      Height          =   2535
      Left            =   5670
      TabIndex        =   23
      Top             =   150
      Width           =   4650
      Begin VB.TextBox memtext2 
         BackColor       =   &H80000016&
         Enabled         =   0   'False
         Height          =   300
         Left            =   4050
         TabIndex        =   30
         Top             =   1215
         Width           =   435
      End
      Begin VB.TextBox memtext3 
         BackColor       =   &H80000016&
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   27
         Top             =   1875
         Width           =   4365
      End
      Begin VB.TextBox memtext1 
         BackColor       =   &H80000016&
         Enabled         =   0   'False
         Height          =   300
         Left            =   120
         TabIndex        =   26
         Top             =   1215
         Width           =   3825
      End
      Begin VB.TextBox txtcon 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   1
         Left            =   120
         MaxLength       =   15
         TabIndex        =   0
         Top             =   600
         Width           =   2175
      End
      Begin Project1.chameleonButton se 
         Height          =   645
         Left            =   2400
         TabIndex        =   1
         Top             =   255
         Width           =   1005
         _extentx        =   1773
         _extenty        =   1138
         btype           =   5
         tx              =   "&Process"
         enab            =   -1  'True
         font            =   "PHconfined.frx":BFAD
         coltype         =   2
         focusr          =   -1  'True
         bcol            =   12632256
         bcolo           =   12632256
         fcol            =   0
         fcolo           =   255
         mcol            =   12632256
         mptr            =   1
         micon           =   "PHconfined.frx":BFD1
         umcol           =   -1  'True
         soft            =   0   'False
         picpos          =   0
         ngrey           =   0   'False
         fx              =   1
         hand            =   0   'False
         check           =   0   'False
         value           =   0   'False
      End
      Begin Project1.chameleonButton cmdsrch 
         Height          =   645
         Left            =   3465
         TabIndex        =   16
         Top             =   255
         Width           =   1035
         _extentx        =   1826
         _extenty        =   1138
         btype           =   5
         tx              =   "&Search PHMbr"
         enab            =   -1  'True
         font            =   "PHconfined.frx":BFEF
         coltype         =   2
         focusr          =   -1  'True
         bcol            =   12632256
         bcolo           =   12632256
         fcol            =   0
         fcolo           =   255
         mcol            =   12632256
         mptr            =   1
         micon           =   "PHconfined.frx":C013
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
         Caption         =   "MI"
         Height          =   255
         Index           =   3
         Left            =   4050
         TabIndex        =   31
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "First Name:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   29
         Top             =   1635
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Last Name:"
         Height          =   255
         Index           =   1
         Left            =   135
         TabIndex        =   28
         Top             =   975
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "PH Member ID No."
         Height          =   255
         Index           =   0
         Left            =   135
         TabIndex        =   24
         Top             =   330
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Viewing Section: PHmember(s)"
      ForeColor       =   &H000000FF&
      Height          =   2565
      Left            =   540
      TabIndex        =   21
      Top             =   135
      Width           =   5025
      Begin MSComctlLib.ListView ListView 
         Height          =   2160
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   4770
         _ExtentX        =   8414
         _ExtentY        =   3810
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         OLEDragMode     =   1
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483624
         BorderStyle     =   1
         Appearance      =   0
         OLEDragMode     =   1
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Recno"
            Object.Width           =   4
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "PH Member ID No."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Last Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "First Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "MI"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Address"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   6840
      Index           =   2
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   12065
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
Attribute VB_Name = "PHconfined"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sql, rs, y As Variant, ctr As Byte
Dim MyType, passString As String
Dim formod As String



Private Sub cmdprocess_Click()
SetText_T True

End Sub

Private Sub butx_Click()
Unload Me
MainForm.Con.Enabled = True
MainForm.Toolbar1.Buttons.Item(2).Enabled = True
End Sub

Private Sub cmdAdd_Click()
SetText2 True
SetText_T True
ClearText2
txtcon(0).Text = AutoRecordNumber2
txtdate.Text = Format(now, "mm/dd/yyyy")
txttime.Text = Format(now, "medium time")
txtcon(2).SetFocus
ListView1.Enabled = False
cmdmod.Enabled = False
cmdsave.Enabled = True
cmdadd.Enabled = False
cmdcancel.Enabled = True
MyType = "ADD"
End Sub

Private Sub cmdCancel_Click()
 ListView1.ListItems.Clear
 ListView1.Enabled = True
 txtcon(1).Enabled = True
 SetText2 False
 ClearText2
 se.Enabled = True
 cmdsrch.Enabled = True
 ListView.Enabled = True
 cmdadd.Enabled = False
 cmdsave.Enabled = False
 cmdmod.Enabled = False
 cmdcancel.Enabled = False
 cmdreset.Enabled = True
 cmdsrch2.Enabled = True
End Sub


Private Sub cmdDIS_Click()
PHdischarged.Show vbModal
End Sub

Private Sub cmddiag_Click()
ListView.Enabled = True
listcon3.Show vbModal
End Sub

Private Sub cmdExit_Click()
Unload Me
MainForm.Con.Enabled = True
MainForm.Toolbar1.Buttons.Item(2).Enabled = True

End Sub

Private Sub cmdmod_Click()
cmdcancel.Enabled = True
If txtcon(0).Text = "" Then
   MsgBox "You must select a Record to Modify!", vbCritical, "Warning"
   ListView1.ListItems.Clear
   ListView1.Enabled = True
   txtcon(1).Enabled = True
   SetText2 False
   ClearText2
   se.Enabled = True
   cmdsrch.Enabled = True
   ListView.Enabled = True
   ListView.SetFocus
   cmdadd.Enabled = False
   cmdsave.Enabled = False
   cmdmod.Enabled = False
   cmdcancel.Enabled = False
   cmdreset.Enabled = False
Else
   SetText2 True
  SetText_T True
   txtcon(2).SetFocus
   ListView1.Enabled = False
   cmdadd.Enabled = False
   cmdsave.Enabled = True
   cmdmod.Enabled = False
   MyType = "EDIT"
End If
End Sub

Private Sub cmdreset_Click()
ClearP
ClearText2
ListView.Enabled = True
ListView1.ListItems.Clear
DisplayLst
txtcon(1).Enabled = True
se.Enabled = True
cmdsrch.Enabled = True
cmdsrch2.Enabled = True
End Sub

Private Sub CMDSAVE_Click()
On Error GoTo TRAPPER
    Dim TRS As DAO.Recordset
    Dim TQR As DAO.QueryDef
    Dim P, madzsrch, testdate As String
    Dim Query, tb, td As String
    Dim List As ListItem
    Dim X As Long
    Dim Flag As Boolean
    Flag = False
    For X = 1 To 6
        If txtcon(X).Text = "" Then Flag = True
    Next X
    If txtdate.Text = "__/__/____" Then
       Flag = True
    End If
     
     
    If Flag Then
        MsgBox "Please Enter all information to Continue ?", vbInformation, "Confirmation"
        GoTo TRAPPER
     End If
   
   If MyType = "ADD" Then
   
   '******** search for 3 months patients *********
    testdate = ""
    tb = Trim(txtcon(1).Text)
   
    sql = "select patientno,idno,datec,plastname,pfirstname,pmi,page,padiagnose from Patient Where plastname = '" & txtcon(2).Text & "' and pfirstname ='" & txtcon(3).Text & "' and pmi = '" & txtcon(4).Text & "' and ISNULL(dateD) = TRUE order by plastname,pfirstname,pmi;"
           Set rs = DBMain.OpenRecordset(sql)
            If rs.RecordCount <> 0 Or rs.RecordCount > 0 Then
                MsgBox "The Patient exist in the admission list! [ Note: If not as PH-Med as Non-Med ] Check the Confined Patients List for confirmation.", vbCritical, "Warning"
           GoTo TRAPPER
            End If
   
    
    
    sql = "select patientno,idno,datec,plastname,pfirstname,pmi,page,padiagnose,DateD from Patient Where plastname = '" & txtcon(2).Text & "' and pfirstname ='" & txtcon(3).Text & "' and pmi = '" & txtcon(4).Text & "' and trim(idno) = '" & tb & "' and ISNULL(dateD) = FALSE and isnull(idno) = False order by plastname,pfirstname,pmi;"
            Dim Scan As String
            Set rs = DBMain.OpenRecordset(sql)
            If rs.RecordCount <> 0 Or rs.RecordCount > 0 Then
                Scan = DateDiff("m", rs.Fields(8), now)
                 
                If Val(Scan) > 3 Then
                   If MsgBox("RETURNEE Discharged Date: " + Format(rs.Fields(8), "mm/dd/yyyy") + ". Months Ellapsed since Dischaged: " + Scan + Space(20) + "[ Note: Before clicking OK, be sure the RETURNEE is Cleared.]", vbCritical + vbOKCancel, "Warning") = vbCancel Then
                      GoTo TRAPPER
                   End If
                Else
                   MsgBox "RETURNEE Discharged Date: " + Format(rs.Fields(8), "mm/dd/yyyy") + "  Current Date: " + Format(now, "mm/dd/yyyy") + Space(27) + ". Months Ellapsed since Discharged: " + Scan + Space(80) + "Sorry cannot be Admitted as PHILHEALTH Patient [ For Non-Med Only. ]", vbCritical, "Warning"
                   GoTo TRAPPER
                End If
            End If
   
   
   
   '**********************************************
   P = "INSERT INTO Patient (patientno,idno,plastname,pfirstname,PMI,Page,PAdiagnose,[datec],TIMEIN,mlast,mfirst,mmi) VALUES ('" & txtcon(0).Text & "','" & txtcon(1).Text & "','" & txtcon(2).Text & "','" & txtcon(3).Text & "','" & txtcon(4).Text & "','" & txtcon(5).Text & "','" & txtcon(6).Text & "','" & txtdate.Text & "','" & txttime.Text & "','" & memtext1.Text & "','" & memtext3.Text & "','" & memtext2.Text & "') ;"
   Set TQR = DBMain.CreateQueryDef("", P)
   TQR.Execute
   Set List = ListView1.ListItems.Add(, , txtcon(0).Text)
        With List
            .SubItems(1) = txtcon(1).Text
            .SubItems(2) = Format(txtdate.Text, "mm/dd/yyyy")
            .SubItems(3) = txtcon(2).Text
            .SubItems(4) = txtcon(3).Text
            .SubItems(5) = txtcon(4).Text
            .SubItems(6) = txtcon(5).Text
            .SubItems(7) = txtcon(6).Text
        End With
   
   ElseIf MyType = "EDIT" Then
                    
            
            passString = "UPDATE Patient SET  idno='" & txtcon(1).Text & "', plastname='" & txtcon(2).Text & "',pfirstname='" & txtcon(3).Text & "',PMI='" & txtcon(4).Text & "',Page='" & txtcon(5).Text & "',PAdiagnose='" & txtcon(6).Text & "',[DateC]='" & txtdate.Text & "',TimeIn = '" & txttime.Text & "',mlast = '" & memtext1.Text & "',mfirst = '" & memtext3.Text & "',mmi = '" & memtext2.Text & "'  WHERE patientno='" & txtcon(0).Text & "';"
            DBMain.Execute passString
            ListView1.Enabled = True
            MyType = ""
            Set List = ListView1.FindItem(txtcon(0).Text, , , lvwPartial)
              With List
                .SubItems(1) = txtcon(1).Text
                .SubItems(2) = Format(txtdate.Text, "mm/dd/yyyy")
                .SubItems(3) = txtcon(2).Text
                .SubItems(4) = txtcon(3).Text
                .SubItems(5) = txtcon(4).Text
                .SubItems(6) = txtcon(5).Text
                .SubItems(7) = txtcon(6).Text
              End With
    
   
   End If
   
   '************************************************
    cmdsave.Enabled = False
    cmdadd.Enabled = True
    cmdmod.Enabled = True
    SetText_T False
    ListView1.Enabled = True
     
TRAPPER:
   Exit Sub
    
End Sub


Private Sub cmdsrch_Click()
 Load srchFormP
 ClearText2
 cmdreset.Enabled = True
 cmdsrch2.Enabled = True
 srchFormP.Show vbModal
End Sub



Private Sub cmdsrch2_Click()
ListView.Enabled = False
txtcon(1).Enabled = False
SetText_T False
se.Enabled = False
cmdsrch.Enabled = False
cmdreset.Enabled = True
cmdcancel.Enabled = True
cmdmod.Enabled = True
cmdadd.Enabled = True
SFconPH.Show vbModal
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

Private Sub Form_Load()
 Dim TRAPPER As Variant
 On Error GoTo TRAPPER
  add3d
  MadzB = False
  Set WSMain = DBEngine.Workspaces(0)
  Set DBMain = WSMain.OpenDatabase(App.Path + "\hospital.mdb", False, False, ";pwd=scanhead")
    
    sql = "select * from PHmember order by 2;"
    Set rs = DBMain.OpenRecordset(sql)
    Do Until rs.EOF
        Set y = ListView.ListItems.Add(, , rs.Fields(0))
        y.SubItems(1) = rs.Fields(1)
        y.SubItems(2) = rs.Fields(2)
        y.SubItems(3) = rs.Fields(3)
        y.SubItems(4) = rs.Fields(4)
        y.SubItems(5) = rs.Fields(5)
        rs.MoveNext
    Loop
    
           SetText2 False
           ClearText2
           cmdadd.Enabled = False
           cmdsave.Enabled = False
           cmdmod.Enabled = False
           cmdcancel.Enabled = False
           cmdreset.Enabled = True
           se.Enabled = True
           cmdsrch.Enabled = True
           cmdsrch2.Enabled = True
           
TRAPPER:
End Sub

Private Sub ListView_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo TRAPPER
    Dim X As Long
    Dim TQR As DAO.QueryDef
    Dim TRS As DAO.Recordset
    'Scanhead rules
    '*********
  
   Set TQR = DBMain.CreateQueryDef("", "SELECT * FROM PHmember WHERE recno ='" & ListView.SelectedItem.Text & "' ORDER BY IDNO")
   Set TRS = TQR.OpenRecordset()
   PHconfined.txtcon(1).Text = TRS.Fields(1)
   PHconfined.memtext1.Text = TRS.Fields(2)
   PHconfined.memtext2.Text = TRS.Fields(4)
   PHconfined.memtext3.Text = TRS.Fields(3)
   PHconfined.txtcon(1).Enabled = True
   cmdsrch2.Enabled = True
TRAPPER:
End Sub





Private Sub ListView1_GotFocus()
On Error GoTo TRAPPER
    Dim X As Long
    Dim testdate As String
    Dim TQR As DAO.QueryDef
    Dim TRS As DAO.Recordset
    testdate = ""
    'Scanhead rules
    Set TQR = DBMain.CreateQueryDef("", "SELECT Patientno,idno,plastname,pfirstname,PMI,page,padiagnose,datec,TIMEIN,mlast,mfirst,mmi FROM Patient WHERE patientno ='" & ListView1.SelectedItem.Text & "' and format(dateD,'mm/dd/yyyy') = '" & testdate & "' and isnull(idno) = False ORDER BY plastname,pfirstname,pmi")
   Set TRS = TQR.OpenRecordset()
   PHconfined.txtcon(0).Text = TRS.Fields(0)
   PHconfined.txtcon(1).Text = TRS.Fields(1)
   PHconfined.txtcon(2).Text = TRS.Fields(2)
   PHconfined.txtcon(3).Text = TRS.Fields(3)
   PHconfined.txtcon(4).Text = TRS.Fields(4)
   PHconfined.txtcon(5).Text = TRS.Fields(5)
   PHconfined.txtcon(6).Text = TRS.Fields(6)
   PHconfined.txtdate.Text = Format(TRS.Fields(7), "mm/dd/yyyy")
   PHconfined.txttime.Text = Format(TRS.Fields(8), "medium time")
   PHconfined.memtext1.Text = TRS.Fields(9)
   PHconfined.memtext2.Text = TRS.Fields(10)
   PHconfined.memtext3.Text = TRS.Fields(11)
   cmdsrch2.Enabled = True
TRAPPER:
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo TRAPPER
    Dim X As Long
    Dim testdate As String
    Dim TQR As DAO.QueryDef
    Dim TRS As DAO.Recordset
    'Scanhead rules
    '*********
    testdate = ""
   Set TQR = DBMain.CreateQueryDef("", "SELECT Patientno,idno,plastname,pfirstname,PMI,page,padiagnose,datec,TIMEIN,mlast,mfirst,mmi FROM Patient WHERE patientno ='" & ListView1.SelectedItem.Text & "' and format(dateD,'mm/dd/yyyy') = '" & testdate & "' and isnull(idno) = False ORDER BY plastname,pfirstname,pmi")
   Set TRS = TQR.OpenRecordset()
   PHconfined.txtcon(0).Text = TRS.Fields(0)
   PHconfined.txtcon(2).Text = TRS.Fields(2)
   PHconfined.txtcon(3).Text = TRS.Fields(3)
   PHconfined.txtcon(4).Text = TRS.Fields(4)
   PHconfined.txtcon(5).Text = TRS.Fields(5)
   PHconfined.txtcon(6).Text = TRS.Fields(6)
   PHconfined.txtdate.Text = Format(TRS.Fields(7), "mm/dd/yyyy")
   PHconfined.txttime.Text = Format(TRS.Fields(8), "medium time")
   PHconfined.txtcon(1).Text = TRS.Fields(1)
   PHconfined.memtext1.Text = TRS.Fields(9)
   PHconfined.memtext2.Text = TRS.Fields(10)
   PHconfined.memtext3.Text = TRS.Fields(11)
TRAPPER:
End Sub

Private Sub Psrch_Click()
cmdreset.Enabled = False
SrchFormC.Show vbModal
End Sub

Private Sub MaskEdBox1_Change()

End Sub

Private Sub se_Click()
Dim TRAPPER As String
Dim TRS As DAO.Recordset
Dim TQR As DAO.QueryDef
Dim P, madzsrch As String
Dim Query, tb, testdate, td As String
Dim List As ListItem
Dim X As Long

cmdsrch2.Enabled = False
ListView1.Enabled = True
'//this area is for searchin the inputted record
    tb = Trim(txtcon(1).Text)
    testdate = ""
    If tb = "" Then
      MsgBox "Null entries are not acceptable!!", vbCritical, "Warning"
      PHconfined.txtcon(1).Enabled = True
      PHconfined.txtcon(1).SetFocus
    Else
      madzsrch = "SELECT * FROM PHmember WHERE idno like '*" & tb & "*';"
      Set TRS = DBMain.OpenRecordset(madzsrch)
        If TRS.RecordCount > 0 Then
           SetText_T False
           PHconfined.memtext1.Text = TRS.Fields(2)
           PHconfined.memtext2.Text = TRS.Fields(4)
           PHconfined.memtext3.Text = TRS.Fields(3)
           PHconfined.txtcon(1).Enabled = False
           se.Enabled = False
           cmdsrch.Enabled = False
           ListView.Enabled = False
           cmdcancel.Enabled = True
           cmdadd.Enabled = True
           cmdsave.Enabled = False
           '************************************
           sql = "select patientno,idno,datec,plastname,pfirstname,pmi,page,padiagnose from Patient where trim(idno) = trim('" & tb & "') and ISNULL(dateD) = TRUE and isnull(idno) = False order by plastname,pfirstname,pmi;"
           Set rs = DBMain.OpenRecordset(sql)
            If rs.RecordCount = 0 Or rs.RecordCount < 0 Then
                cmdmod.Enabled = False
            Else
                Do Until rs.EOF
                    Set y = ListView1.ListItems.Add(, , rs.Fields(0))
                        y.SubItems(1) = rs.Fields(1)
                        y.SubItems(2) = Format(rs.Fields(2), "mm/dd/yyyy")
                        y.SubItems(3) = rs.Fields(3)
                        y.SubItems(4) = rs.Fields(4)
                        y.SubItems(5) = rs.Fields(5)
                        y.SubItems(6) = rs.Fields(6)
                        y.SubItems(7) = rs.Fields(7)
                        rs.MoveNext
                 Loop
                 cmdmod.Enabled = True
            End If
            
           '*******************************
         Else
           If MsgBox("PhilHealth Member ID does not exist, Click OK to Register the New ID!", vbExclamation + vbOKCancel, "Attention") = vbOK Then
              PHmember.Show
              PHmember.cmdadd.SetFocus
              SendKeys "{Enter}"
              PHmember.txtmem(1).Text = tb
              MadzB = True
           Else
              PHconfined.txtcon(1).Enabled = True
              PHconfined.txtcon(1).SetFocus
           End If
        End If
    End If
          '/****************************************
End Sub

Private Sub txtcon_GotFocus(Index As Integer)
If Index = 1 Then
  txtcon(1).Text = ""
  memtext1.Text = ""
  memtext2.Text = ""
  memtext3.Text = ""
End If

End Sub

Private Sub txtcon_KeyPress(Index As Integer, KeyAscii As Integer)
'the madzbry txt validation
Dim madzbry As String
 Select Case Index
        Case 1, 5
         madzbry = "0123456789"
          If KeyAscii > 26 Then
            If InStr(madzbry, Chr(KeyAscii)) = 0 Then
              KeyAscii = 0
            End If
         End If
        Case 2, 3, 4
         madzbry = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ ."
          If KeyAscii > 26 Then
            If InStr(madzbry, Chr(KeyAscii)) = 0 Then
              KeyAscii = 0
            End If
          End If
End Select
End Sub

Public Function AutoRecordNumber2() As String
    Dim TQR As DAO.QueryDef
    Dim TRS As DAO.Recordset
    Dim Start As String
    Start = "00000000"
    Set TQR = DBMain.CreateQueryDef("", "SELECT PatientNo FROM Patient")
    Set TRS = TQR.OpenRecordset()
    Do While Not TRS.EOF
        TRS.FindFirst "PatientNo ='P-" + Start + "'"
        If Not TRS.NoMatch Then
            Start = Format(Str(Val(Mid$(Start, 3)) + 1), "00000000")
        Else
            AutoRecordNumber2 = "P-" + Start
            Exit Function
        End If
    Loop
    AutoRecordNumber2 = "P-" + Start
End Function

Public Function add3d()
 Dim X As Integer
 
 Add3DBorder Me
 Add3DBorder ListView
 Add3DBorder ListView1
 For X = 0 To 6
   Add3DBorder txtcon(X)
 Next X
 Add3DBorder memtext1
 Add3DBorder memtext2
 Add3DBorder memtext3
 Add3DBorder txtdate
 Add3DBorder txttime
 
End Function



Private Sub txtdate_GotFocus()
With txtdate
.SelStart = 0
.SelLength = Len(txtdate.Text)
End With
End Sub

Public Sub DisplayLst() 'Reset control display: gotcha the scanhedbri
On Error GoTo X
    
    Dim TQR As DAO.QueryDef
    Dim rs As DAO.Recordset
    Dim y As ListItem
     Dim sql As String
    Dim X As Long
   
    sql = "select * from PHmember order by 2;"
    Set rs = DBMain.OpenRecordset(sql)
     ListView.ListItems.Clear
    Do Until rs.EOF
        Set y = ListView.ListItems.Add(, , rs.Fields(0))
        y.SubItems(1) = rs.Fields(1)
        y.SubItems(2) = rs.Fields(2)
        y.SubItems(3) = rs.Fields(3)
        y.SubItems(4) = rs.Fields(4)
        y.SubItems(5) = rs.Fields(5)
        rs.MoveNext
    Loop
   
    SetText2 False
X:
End Sub

Private Sub txtTIME_GotFocus()
With txttime
.SelStart = 0
.SelLength = Len(txttime.Text)
End With
End Sub
