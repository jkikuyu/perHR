VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{A22D979F-2684-11D2-8E21-10B404C10000}#1.4#0"; "CPOPMENU.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Personnel Manager HR"
   ClientHeight    =   5445
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6990
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   2400
      Left            =   0
      TabIndex        =   2
      Top             =   780
      Width           =   6990
      _ExtentX        =   12330
      _ExtentY        =   4233
      ButtonWidth     =   1720
      ButtonHeight    =   1376
      AllowCustomize  =   0   'False
      ImageList       =   "imgold"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   12
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Login"
            Key             =   "Login"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Logout"
            Key             =   "Logout"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Employees"
            Key             =   "Emp"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Employment"
            Key             =   "Employment"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Pay"
            Key             =   "Pay"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Banking"
            Key             =   "Banking"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Dependants"
            Key             =   "Dependants"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Applications"
            Key             =   "Applications"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Appraisals"
            Key             =   "Appraisals"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Leave"
            Key             =   "Leave"
            Object.Tag             =   ""
            ImageIndex      =   10
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img 
      Left            =   9435
      Top             =   4125
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0896
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0CEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":113E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1592
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":19E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1E3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":228E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":26E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2A06
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2E5A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Pic1 
      Align           =   2  'Align Bottom
      Height          =   735
      Left            =   0
      Picture         =   "frmMain.frx":32AE
      ScaleHeight     =   675
      ScaleWidth      =   6930
      TabIndex        =   1
      Top             =   4440
      Visible         =   0   'False
      Width           =   6990
      Begin PerHR.TransTBWrapper TWr 
         Height          =   345
         Left            =   15330
         TabIndex        =   3
         Top             =   165
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   609
      End
      Begin cPopMenu.PopMenu pMenu 
         Left            =   6075
         Top             =   0
         _ExtentX        =   1058
         _ExtentY        =   1058
         HighlightCheckedItems=   0   'False
         TickIconIndex   =   0
      End
      Begin VB.Label lblPass 
         Caption         =   "Label1"
         Height          =   330
         Left            =   7605
         TabIndex        =   4
         Top             =   210
         Visible         =   0   'False
         Width           =   1665
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1875
      Top             =   2355
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5C4D0
            Key             =   ""
            Object.Tag             =   "Employees"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5C924
            Key             =   ""
            Object.Tag             =   "Pay"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5CC40
            Key             =   ""
            Object.Tag             =   "Applications"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5CF5C
            Key             =   ""
            Object.Tag             =   "Dependants"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5D3B0
            Key             =   ""
            Object.Tag             =   "Employment File"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5D6CC
            Key             =   ""
            Object.Tag             =   "Banking"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5DB20
            Key             =   ""
            Object.Tag             =   "Lay-offs"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5DE3C
            Key             =   ""
            Object.Tag             =   "Leave"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5E290
            Key             =   ""
            Object.Tag             =   "Apparaisals"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5E5AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5EA0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5EE6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5F2C0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   5175
      Width           =   6990
      _ExtentX        =   12330
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6694
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "05/Oct/2001"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "7:57 AM"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbl 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   6990
      _ExtentX        =   12330
      _ExtentY        =   1376
      Picture         =   "frmMain.frx":5F5DC
      MouseIcon       =   "frmMain.frx":65724
      EmbossPicture   =   -1  'True
      _CBWidth        =   6990
      _CBHeight       =   780
      _Version        =   "6.7.8988"
      MinHeight1      =   360
      Width1          =   9285
      Key1            =   "NavBand"
      NewRow1         =   0   'False
      Caption2        =   "Selected Company"
      Child2          =   "txtCo"
      MinHeight2      =   330
      Width2          =   11745
      Key2            =   "Company"
      NewRow2         =   -1  'True
      AllowVertical2  =   0   'False
      Caption3        =   "Year of Operation"
      Child3          =   "txtYear"
      MinHeight3      =   330
      Width3          =   9750
      Key3            =   "Periods"
      NewRow3         =   0   'False
      AllowVertical3  =   0   'False
      Begin VB.TextBox txtYear 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   330
         Left            =   6960
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   420
         Width           =   150
      End
      Begin VB.TextBox txtCo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   330
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   420
         Width           =   3870
      End
   End
   Begin ComctlLib.ImageList imgold 
      Left            =   5625
      Top             =   4965
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   10
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":65740
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":65A5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":65D74
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":6608E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":663A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":666C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":669DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":66CF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":67010
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":6732A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open Company..."
      End
      Begin VB.Menu sep10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIncEr 
         Caption         =   "&Income Earner..."
      End
      Begin VB.Menu sep29 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLogin 
         Caption         =   "Log &In"
         Enabled         =   0   'False
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuLock 
         Caption         =   "Log &Out"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuChangePass 
         Caption         =   "Change P&assword..."
      End
      Begin VB.Menu sep30 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuData 
      Caption         =   "Curr&ent"
      Enabled         =   0   'False
      Begin VB.Menu mnuPrData 
         Caption         =   "Employee Pe&rsonal Information"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuEmpData 
         Caption         =   "&Employment Information"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuPData 
         Caption         =   "&Payment Information"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuBanking 
         Caption         =   "&Banking Information"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuRecords 
      Caption         =   "Re&cords"
      Enabled         =   0   'False
      Begin VB.Menu mnuDependants 
         Caption         =   "&Dependants"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuEmpHist 
         Caption         =   "E&mployment History"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuEduHist 
         Caption         =   "E&ducation History"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuTHist 
         Caption         =   "Employee &Training History"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuDevt 
         Caption         =   "Professional De&velopment"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuPMember 
         Caption         =   "Professional &Membership"
      End
      Begin VB.Menu mnuSocial 
         Caption         =   "&Social/Public Responsibilities"
      End
      Begin VB.Menu mnuAbsent 
         Caption         =   "A&bsenteeism"
         Enabled         =   0   'False
      End
      Begin VB.Menu sep24 
         Caption         =   "-"
      End
      Begin VB.Menu mnuApp 
         Caption         =   "&Applicants"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuCriteria 
         Caption         =   "&Selection Criteria"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuOperations 
      Caption         =   "&Operations"
      Enabled         =   0   'False
      Begin VB.Menu mnuEmpLeave 
         Caption         =   "&Leave Requests"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuLent 
         Caption         =   "Leave Entitlements"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuLBal 
         Caption         =   "Leave Balances"
         Enabled         =   0   'False
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMedEnt 
         Caption         =   "Medical Entitlements"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuEmpMedSch 
         Caption         =   "Medical Scheme Assignments"
         Enabled         =   0   'False
      End
      Begin VB.Menu sep28 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAppraisals 
         Caption         =   "Appraisals"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuUtilities 
      Caption         =   "&Utilities"
      Enabled         =   0   'False
      Begin VB.Menu mnuqlf 
         Caption         =   "&Qualifications"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuLeave 
         Caption         =   "&Leave Types"
         Enabled         =   0   'False
      End
      Begin VB.Menu sep25 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTTypes 
         Caption         =   "&Training types"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuExtTrn 
         Caption         =   "&External Trainers"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnucoTraining 
         Caption         =   "Company Training &Scheduling"
         Enabled         =   0   'False
      End
      Begin VB.Menu sep27 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMedSch 
         Caption         =   "&Medical Scheme Types"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnucodocs 
         Caption         =   "Company &Doctors/Hospitals"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuChemists 
         Caption         =   "Company Chemists/&Pharmacies"
      End
      Begin VB.Menu sep26 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAppr 
         Caption         =   "&Appraisals"
         Enabled         =   0   'False
         Begin VB.Menu mnuAppParam 
            Caption         =   "&Appraisal Parameters"
         End
         Begin VB.Menu mnuRApp 
            Caption         =   "&Reconcile Appraisals"
         End
      End
      Begin VB.Menu mnuYOp 
         Caption         =   "&Year of Operation"
         Enabled         =   0   'False
         Begin VB.Menu mnuCYr 
            Caption         =   "&Close Current Year"
         End
         Begin VB.Menu mnuNewY 
            Caption         =   "&Open New Year"
         End
         Begin VB.Menu sep20 
            Caption         =   "-"
         End
         Begin VB.Menu mnuYHist 
            Caption         =   "Operation Year &History"
         End
      End
      Begin VB.Menu sep21 
         Caption         =   "-"
      End
      Begin VB.Menu mnuListBuilder 
         Caption         =   "&List Builder..."
         Enabled         =   0   'False
      End
      Begin VB.Menu sep22 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWHours 
         Caption         =   "&Working Hours Definitions..."
      End
      Begin VB.Menu mnuCalendar 
         Caption         =   "&Calendar..."
         Enabled         =   0   'False
      End
      Begin VB.Menu sep23 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCoReg 
         Caption         =   "Company &Registers"
      End
      Begin VB.Menu sep39 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "&Options..."
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "&Reports"
      Enabled         =   0   'False
      Begin VB.Menu mnuViewer 
         Caption         =   "&Report Viewer..."
         Enabled         =   0   'False
      End
      Begin VB.Menu sep32 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRptmngr 
         Caption         =   "Report &Manager..."
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowNewWindow 
         Caption         =   "&New Window"
      End
      Begin VB.Menu mnuWindowBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnuWindowTileHorizontal 
         Caption         =   "Tile &Horizontal"
      End
      Begin VB.Menu mnuWindowTileVertical 
         Caption         =   "Tile &Vertical"
      End
      Begin VB.Menu mnuWindowArrangeIcons 
         Caption         =   "&Arrange Icons"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents"
      End
      Begin VB.Menu mnuHelpSearchForHelpOn 
         Caption         =   "&Search For Help On..."
      End
      Begin VB.Menu mnuHelpBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About "
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Const EM_UNDO = &HC7
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hWnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)


Public Sub EvaluateObjects()
    Dim mnu As Menu
    Dim ctl As Control
    Dim datSubObj As Recordset
    'This procedure evaluates the object assignment and loads the desirable picture
       'Find the form in the objects list
       Set datExtra = New Recordset
       datExtra.Open "SELECT * FROM Objects WHERE s_name ='" & Me.Name & _
       "' AND obj_cat = 2 AND App_code = 2", cn, adOpenStatic, adLockOptimistic
       If datExtra.RecordCount <> 0 Then
          scrid = datExtra!objid
          'Get all objects registered for this form
          Set datExtra = New Recordset
          datExtra.Open "SELECT * FROM Objects WHERE Parent =" & CLng(scrid), cn, adOpenStatic, adLockOptimistic
          While Not datExtra.EOF
                'Check if this object has any sub objects e.g. in menus
                Set datSubObj = New Recordset
                datSubObj.Open "SELECT * FROM Objects WHERE Parent =" & CLng(datExtra!objid), cn, adOpenStatic, adLockOptimistic
                While Not datSubObj.EOF
                    'Check if user group is assigned this object
                    Set datExtra2 = New Recordset
                    datExtra2.Open "SELECT * FROM grp_objs WHERE grp_id = " & grpID & " AND objid =" & _
                    datSubObj!objid, cn, adOpenStatic, adLockOptimistic
                    If datExtra2.RecordCount <> 0 Then
                       'Enable this object
                        For Each ctl In Me
                            If (TypeOf ctl Is Menu) Then
                                If ctl.Name = Trim(datSubObj!s_name) Then
                                   ctl.Enabled = True
                                Else
                                End If
                            End If
                        Next
                    End If
                    datSubObj.MoveNext
                Wend
                'Now assign the parent object
                'Check if employee is assigned this object
                Set datExtra2 = New Recordset
                datExtra2.Open "SELECT * FROM grp_objs WHERE grp_id = " & grpID & " AND objid =" & _
                datExtra!objid, cn, adOpenStatic, adLockOptimistic
                If datExtra2.RecordCount <> 0 Then
                   'Enable this object
                    For Each Menu In Me
                        If Menu.Name = Trim(datExtra!s_name) Then
                           Menu.Enabled = True
                        Else
                        End If
                    Next Menu
                End If
                datExtra.MoveNext
          Wend
       End If
    'this portion is only added for frmMain to control the toolbar
    If mnuPrData.Enabled = True Then
       Toolbar1.Buttons(4).Enabled = True
    Else
        Toolbar1.Buttons(4).Enabled = False
    End If
    If mnuEmpData.Enabled = True Then
       Toolbar1.Buttons(5).Enabled = True
    Else
        Toolbar1.Buttons(5).Enabled = False
    End If
    If mnuPData.Enabled = True Then
       Toolbar1.Buttons(6).Enabled = True
    Else
        Toolbar1.Buttons(6).Enabled = False
    End If
    If mnuBanking.Enabled = True Then
       Toolbar1.Buttons(7).Enabled = True
    Else
        Toolbar1.Buttons(7).Enabled = False
    End If
    
    If mnuDependants.Enabled = True Then
       Toolbar1.Buttons(8).Enabled = True
    Else
        Toolbar1.Buttons(8).Enabled = False
    End If
    If mnuApp.Enabled = True Then
       Toolbar1.Buttons(10).Enabled = True
    Else
        Toolbar1.Buttons(10).Enabled = False
    End If
    If mnuAppraisals.Enabled = True Then
       Toolbar1.Buttons(11).Enabled = True
    Else
        Toolbar1.Buttons(11).Enabled = False
    End If
    If mnuEmpLeave.Enabled = True Then
       Toolbar1.Buttons(12).Enabled = True
    Else
        Toolbar1.Buttons(12).Enabled = False
    End If
End Sub

Private Sub MDIForm_Load()
    Dim ctl As Control
    
    EvaluateObjects
    'Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    'Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    'Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
    'Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
    
    If Operation_Year = 0 Then
        Me.Caption = "Human Resources System"
        
        txtCo = uCo
        txtYear = "Not Opened"
    Else
        Me.Caption = "Human Resources System"
        
        txtCo = uCo
        txtYear = Operation_Year
    End If
    frmMain.StatusBar1.Panels(3).Text = uUser
    lblPass = uPass
    
    'Set the menu Background
    With pMenu
         .ImageList = img
         
         .SubClassMenu Me
         .HighlightStyle = cspHighlightButton
         Set .BackgroundPicture = Pic1.Picture
                
         .MenuDefault("mnuOpen") = -1
         '.MenuDefault("mnuIncEr") = -1
         
         .ItemIcon("mnuPrData") = 0
         .ItemIcon("mnuLogin") = 1
         .ItemIcon("mnuLock") = 2
         .ItemIcon("mnuEmpData") = 3
         .ItemIcon("mnuAppraisals") = 4
         .ItemIcon("mnuBanking") = 5
         .ItemIcon("mnuDependants") = 6
         .ItemIcon("mnuApp") = 7
         .ItemIcon("mnuMedEnt") = 8
         .ItemIcon("mnuEmpLeave") = 9
         .ItemIcon("mnuOpen") = 10
    End With
    TWr.Transparent = True
    
    Set TWr.Container = cbl
    Set cbl.Bands.item("NavBand").Child = TWr
    
    
    Set Toolbar1.Container = cbl
    TWr.Transparent = True
    Set TWr.Toolbar = Toolbar1
    Toolbar1.Buttons(1).Enabled = False
End Sub


Private Sub MDIForm_Unload(Cancel As Integer)
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
    End If
    End
End Sub

Private Sub mnuAbsent_Click()
    frmAbsentee.Show vbModal
    
End Sub

Private Sub mnuApp_Click()
    frmApplicants.Show vbModal
End Sub

Private Sub mnuAppParam_Click()
    frmAppParams.Show vbModal
    
End Sub

Private Sub mnuAppraisals_Click()
    frmEmpApp.Show vbModal
    
End Sub

Private Sub mnuBanking_Click()
    frmBanking_data.Show vbModal
    
End Sub

Private Sub mnuCalendar_Click()
    frmCalendar.Show vbModal
    
End Sub

Private Sub mnuChangePass_Click()
    vbPass = 3
    frmPass.Show vbModal
    
    'frmChangePass.Show vbModal
End Sub

Private Sub mnuChemists_Click()
    frmChemists.Show vbModal
    
End Sub

Private Sub mnucodocs_Click()
    frmco_doctors.Show vbModal
    
End Sub

Private Sub mnuCoReg_Click()
    CoReg.Show vbModal
End Sub

Private Sub mnucoTraining_Click()
    frmcoTraining.Show vbModal
    
End Sub

Private Sub mnuCriteria_Click()
    frmAppSel.Show vbModal
    
End Sub

Private Sub mnuCYr_Click()
    frmCloseYear.Show vbModal
    
End Sub

Private Sub mnuDependants_Click()
    frmDependants.Show vbModal
    
End Sub

Private Sub mnuDevt_Click()
    frmEmpDevt.Show vbModal
End Sub

Private Sub mnuEduHist_Click()
    frmEduHist.Show vbModal
End Sub

Private Sub mnuEmpData_Click()
    frmEmployment.Show vbModal
    
End Sub

Private Sub mnuEmpHist_Click()
    frmEmpHist.Show vbModal
End Sub

Private Sub mnuEmpLeave_Click()
    If Operation_Year = 0 Then
       MsgBox ("This facility cannot be accessed without a valid operation year." & Chr(13) & _
       "Please assign an operation year for your company."), vbInformation
       Exit Sub
    End If
    frmEmpLeave.Show vbModal
    
End Sub

Private Sub mnuEmpMedSch_Click()
    frmempmedsch.Show vbModal
    
End Sub

Private Sub mnuExtTrn_Click()
    frmExtTrainers.Show vbModal
    
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal
End Sub

Private Sub mnuHelpSearchForHelpOn_Click()
    Dim nRet As Integer


    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hWnd, App.HelpFile, 261, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If

End Sub

Private Sub mnuHelpContents_Click()
    Dim nRet As Integer


    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hWnd, App.HelpFile, 3, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If

End Sub


Private Sub mnuIncEr_Click()
On Error GoTo ShellErr
    'Find the executable for the human resource
    Dim IncFile As String
    Dim ShellHr As Double
    If Right(App.Path, 1) = "\" Then
       IncFile = Dir(App.Path + "PerMan.exe")
    Else
       IncFile = Dir(App.Path + "\PerMan.exe")
    End If
    If IncFile = "" Then
       MsgBox ("Cannot find the executable for the Income Earner System." & Chr(13) & _
       "Possibly the Income Earner system is not well installed." & Chr(13) & _
       "Please re-install the Income Earner system."), vbInformation
    Else
        If Right(App.Path, 1) = "\" Then
           ShellHr = Shell(App.Path + "PerMan.exe", vbNormalFocus)
        Else
           ShellHr = Shell(App.Path + "\PerMan.exe", vbNormalFocus)
        End If
    End If
    Exit Sub
ShellErr:
MsgBox (Err.Description & Chr(13) & "Cannot launch application."), vbInformation
End Sub

Private Sub mnuLBal_Click()
    If Operation_Year = 0 Then
       MsgBox ("This facility cannot be accessed without a valid operation year." & Chr(13) & _
       "Please assign an operation year for your company."), vbInformation
       Exit Sub
    End If
    frmLeaveBal.Show vbModal
    
End Sub

Private Sub mnuLeave_Click()
    frmLeaves.Show vbModal
    
End Sub

Private Sub mnuLent_Click()
    If Operation_Year = 0 Then
       MsgBox ("This facility cannot be accessed without a valid operation year." & Chr(13) & _
       "Please assign an operation year for your company."), vbInformation
       Exit Sub
    End If
    frmlent.Show vbModal
End Sub

Private Sub mnuListBuilder_Click()
    frmListBuilder.Show vbModal
    
End Sub

Private Sub mnuLock_Click()
        Dim i
        For i = 2 To 12
            Toolbar1.Buttons(i).Enabled = False
        Next i
        Toolbar1.Buttons(1).Enabled = True
        ID = ""
        Me.mnuData.Enabled = False
        Me.mnuLogin.Enabled = True
        Me.mnuLock.Enabled = False
        Me.mnuIncEr.Enabled = False
        Me.mnuOpen.Enabled = False
        Me.mnuChangePass.Enabled = False
        Me.mnuRecords.Enabled = False
        Me.mnuOperations.Enabled = False
        Me.mnuUtilities.Enabled = False
        '*****Temporarily disabled******
        'Set datExtra = New Recordset
        'datExtra.Open "SELECT * FROM ULogins WHERE user_name ='" & _
        UCase(username) & "'", cn, adOpenStatic, adLockOptimistic
        'datExtra.Delete
        
        UserName = ""
        MsgBar "", False
End Sub

Private Sub mnuLogin_Click()
    frmEntry.Show vbModal
    
    If UserName <> "" Then
        Me.mnuData.Enabled = True
        Me.mnuLogin.Enabled = False
        Me.mnuLock.Enabled = True
        Me.mnuIncEr.Enabled = True
        Me.mnuOpen.Enabled = True
        Me.mnuChangePass.Enabled = True
        Me.mnuRecords.Enabled = True
        Me.mnuOperations.Enabled = True
        Me.mnuUtilities.Enabled = True
        
        Dim i
        For i = 2 To 12
            Toolbar1.Buttons(i).Enabled = True
        Next i
        Toolbar1.Buttons(1).Enabled = False
    End If
End Sub

Private Sub mnuMedEnt_Click()
    If Operation_Year = 0 Then
       MsgBox ("This facility cannot be accessed without a valid operation year." & Chr(13) & _
       "Please assign an operation year for your company."), vbInformation
       Exit Sub
    End If
    frmempmed.Show vbModal
    
End Sub

Private Sub mnuMedSch_Click()
    frmmedschemes.Show vbModal
    
End Sub

Private Sub mnuNewY_Click()
    frmOpenYear.Show vbModal
    
End Sub

Private Sub mnuOpen_Click()
    CompID.Show vbModal
    If Operation_Year = 0 Then
        Me.Caption = "Human Resources System     " & uCo
    Else
        Me.Caption = "Human Resources System     " & uCo & " - " & Operation_Year
    End If
End Sub

Private Sub mnuOptions_Click()
    frmOptions.Show vbModal
    
End Sub

Private Sub mnuPData_Click()
    frmPayroll_data.Show vbModal
    
End Sub

Private Sub mnuPMember_Click()
    frmPMembership.Show vbModal
    
End Sub

Private Sub mnuPrData_Click()
    frmEmp.Show 'vbModal
    
End Sub

Private Sub mnuqlf_Click()
    frmQlf.Show vbModal
End Sub

Private Sub mnuRApp_Click()
    frmRAppr.Show vbModal
    
End Sub

Private Sub mnuRptmngr_Click()
    frmReportManager.Show vbModal
    
End Sub

Private Sub mnuSocial_Click()
    frmSResponsibilities.Show vbModal
    
End Sub

Private Sub mnuTHist_Click()
    frmEmpTraining.Show vbModal
    
End Sub

Private Sub mnuTTypes_Click()
    frmTrainings.Show vbModal
    
End Sub

Private Sub mnuViewer_Click()
    frmReportViewer.Show vbModal
    
End Sub

Private Sub mnuWHours_Click()
    frmWHours.Show vbModal
End Sub

Private Sub mnuWindowArrangeIcons_Click()
    Me.Arrange vbArrangeIcons
End Sub

Private Sub mnuWindowTileVertical_Click()
    Me.Arrange vbTileVertical
End Sub

Private Sub mnuWindowTileHorizontal_Click()
    Me.Arrange vbTileHorizontal
End Sub

Private Sub mnuWindowCascade_Click()
    Me.Arrange vbCascade
End Sub

Private Sub mnuFileExit_Click()
    'unload the form
    End

End Sub

Private Sub mnuYHist_Click()
    frmYearsHistory.Show vbModal
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
    Select Case Button.key
        Case "Emp"
            mnuPrData_Click
        Case "Employment"
            mnuEmpData_Click
        Case "Pay"
            mnuPData_Click
        Case "Banking"
            mnuBanking_Click
        Case "Dependants"
            mnuDependants_Click
        Case "Applications"
            mnuApp_Click
        Case "LayOffs"
        
        Case "Appraisals"
            mnuAppraisals_Click
        Case "Leave"
            mnuEmpLeave_Click
        Case "Login"
            mnuLogin_Click
        Case "Logout"
            mnuLock_Click
    End Select
End Sub
