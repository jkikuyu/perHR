VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmEmpLeave 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Employee Leave Requests"
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7800
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEmpLeave.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   7800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtleave 
      DataField       =   "leave_code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   9195
      TabIndex        =   12
      Top             =   1890
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CommandButton cmdLeave 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   9705
      Picture         =   "frmEmpLeave.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2205
      Width           =   285
   End
   Begin VB.CommandButton cmdFind 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5595
      Picture         =   "frmEmpLeave.frx":053C
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   270
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   7800
      TabIndex        =   1
      Top             =   6900
      Width           =   7800
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   5595
         TabIndex        =   56
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdGoTo 
         Caption         =   "Go to..."
         Height          =   350
         Left            =   75
         TabIndex        =   55
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   1275
         TabIndex        =   6
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "C&lose"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   6795
         TabIndex        =   5
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   4515
         TabIndex        =   4
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   3435
         TabIndex        =   3
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   2355
         TabIndex        =   2
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.TextBox txtleave 
      Height          =   330
      Index           =   1
      Left            =   8595
      TabIndex        =   0
      Top             =   2205
      Width           =   1080
   End
   Begin Threed.SSPanel pnlother_names 
      Height          =   375
      Left            =   2910
      TabIndex        =   13
      Top             =   270
      Width           =   2655
      _Version        =   65536
      _ExtentX        =   4683
      _ExtentY        =   661
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
   End
   Begin Threed.SSPanel pnlLast_name 
      Height          =   375
      Left            =   975
      TabIndex        =   14
      Top             =   270
      Width           =   1920
      _Version        =   65536
      _ExtentX        =   3387
      _ExtentY        =   661
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSPanel pnlStaff_code 
      Height          =   375
      Left            =   0
      TabIndex        =   15
      Top             =   270
      Width           =   975
      _Version        =   65536
      _ExtentX        =   1720
      _ExtentY        =   661
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame3 
      Height          =   5865
      Left            =   15
      TabIndex        =   17
      Top             =   975
      Width           =   7770
      Begin VB.CommandButton cmdCalc 
         Caption         =   "&Calculate Days"
         Height          =   345
         Left            =   5835
         TabIndex        =   53
         Top             =   2475
         Width           =   1830
      End
      Begin VB.CommandButton fnext 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6855
         Picture         =   "frmEmpLeave.frx":06BE
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   5355
         Width           =   375
      End
      Begin VB.CommandButton fprev 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6435
         Picture         =   "frmEmpLeave.frx":0808
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   5355
         Width           =   375
      End
      Begin VB.CommandButton ffirst 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6030
         Picture         =   "frmEmpLeave.frx":0952
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   5355
         Width           =   375
      End
      Begin VB.CommandButton flast 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7260
         Picture         =   "frmEmpLeave.frx":0A9C
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   5355
         Width           =   375
      End
      Begin VB.CommandButton cmdEntitlement 
         Caption         =   "Leave Entitlement....."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5835
         TabIndex        =   48
         Top             =   1470
         Width           =   1830
      End
      Begin VB.CommandButton cmdBal 
         Caption         =   "Leave Balances......"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5835
         TabIndex        =   47
         Top             =   1875
         Width           =   1830
      End
      Begin VB.CommandButton cmdAnalyze 
         Caption         =   "A&nalyze"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5835
         TabIndex        =   46
         Top             =   270
         Width           =   1830
      End
      Begin VB.CommandButton cmdApprove 
         Caption         =   "A&pprove"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5835
         TabIndex        =   45
         Top             =   675
         Width           =   1830
      End
      Begin VB.CommandButton cmdDisapprove 
         Caption         =   "&Dis- Approve"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5835
         TabIndex        =   44
         Top             =   1065
         Width           =   1830
      End
      Begin VB.Frame Frame2 
         Caption         =   "Official Use"
         Height          =   1020
         Left            =   135
         TabIndex        =   37
         Top             =   4770
         Width           =   5535
         Begin VB.TextBox txtleave 
            DataField       =   "adays"
            Height          =   330
            Index           =   7
            Left            =   2295
            TabIndex        =   41
            Top             =   555
            Width           =   1200
         End
         Begin VB.TextBox txtleave 
            DataField       =   "adate"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MMM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   330
            Index           =   8
            Left            =   225
            TabIndex        =   40
            Top             =   555
            Width           =   1560
         End
         Begin VB.CommandButton cmdPost 
            Caption         =   "Post Leave"
            Height          =   345
            Left            =   3915
            TabIndex        =   38
            Top             =   555
            Width           =   1545
         End
         Begin MSComCtl2.DTPicker DTPicker3 
            Height          =   330
            Left            =   1770
            TabIndex        =   39
            Top             =   555
            Width           =   270
            _ExtentX        =   476
            _ExtentY        =   582
            _Version        =   393216
            Format          =   24510465
            CurrentDate     =   36893
         End
         Begin VB.Label Label11 
            Caption         =   "Actual Number of Days Taken"
            Height          =   270
            Left            =   2325
            TabIndex        =   43
            Top             =   300
            Width           =   1920
         End
         Begin VB.Label Label12 
            Caption         =   "Actual End Date"
            Height          =   180
            Left            =   270
            TabIndex        =   42
            Top             =   315
            Width           =   1455
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Leave Options"
         Height          =   750
         Left            =   150
         TabIndex        =   34
         Top             =   3945
         Width           =   5505
         Begin VB.CheckBox chkPaid 
            Caption         =   "Paid Leave"
            DataField       =   "paid_leave"
            Height          =   255
            Left            =   225
            TabIndex        =   36
            Top             =   360
            Width           =   1575
         End
         Begin VB.CheckBox chkunpaid 
            Caption         =   "Un-Paid Leave"
            DataField       =   "unpaid_leave"
            Height          =   255
            Left            =   2850
            TabIndex        =   35
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.TextBox txtleave 
         DataField       =   "notes"
         Height          =   990
         Index           =   5
         Left            =   135
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   26
         Top             =   2880
         Width           =   5505
      End
      Begin VB.TextBox txtleave 
         DataField       =   "start_date"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MMM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   150
         TabIndex        =   25
         Top             =   1755
         Width           =   1905
      End
      Begin VB.TextBox txtleave 
         DataField       =   "end_date"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MMM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   330
         Index           =   3
         Left            =   3165
         TabIndex        =   24
         Top             =   1755
         Width           =   1905
      End
      Begin VB.TextBox txtleave 
         DataField       =   "rdays"
         Height          =   330
         Index           =   4
         Left            =   1500
         TabIndex        =   23
         Top             =   2205
         Width           =   540
      End
      Begin VB.TextBox txtleave 
         DataField       =   "req_date"
         Height          =   285
         Index           =   6
         Left            =   165
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   465
         Width           =   1560
      End
      Begin MSDataListLib.DataCombo cboLeave 
         DataField       =   "leave_code"
         Height          =   345
         Left            =   150
         TabIndex        =   18
         Top             =   1035
         Width           =   5520
         _ExtentX        =   9737
         _ExtentY        =   609
         _Version        =   393216
         Style           =   2
         ListField       =   "leave_name"
         BoundColumn     =   "leave_code"
         Text            =   ""
      End
      Begin MSComCtl2.DTPicker rqdate 
         DataField       =   "req_date"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MMM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   300
         Left            =   1740
         TabIndex        =   20
         Top             =   465
         Width           =   270
         _ExtentX        =   476
         _ExtentY        =   529
         _Version        =   393216
         Format          =   24444929
         CurrentDate     =   36893.4265393519
         MinDate         =   36526
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   330
         Left            =   5070
         TabIndex        =   21
         Top             =   1755
         Width           =   270
         _ExtentX        =   476
         _ExtentY        =   582
         _Version        =   393216
         Format          =   24444929
         CurrentDate     =   36873
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   330
         Left            =   2055
         TabIndex        =   22
         Top             =   1755
         Width           =   270
         _ExtentX        =   476
         _ExtentY        =   582
         _Version        =   393216
         Format          =   24444929
         CurrentDate     =   36873
      End
      Begin VB.Label lblDays 
         Caption         =   "Calculating Number of Days, Please wait..."
         ForeColor       =   &H000000FF&
         Height          =   600
         Left            =   5850
         TabIndex        =   54
         Top             =   2985
         Visible         =   0   'False
         Width           =   1800
      End
      Begin VB.Label Label1 
         Caption         =   "Leave Requested:"
         Height          =   255
         Left            =   165
         TabIndex        =   33
         Top             =   795
         Width           =   2535
      End
      Begin VB.Label Label2 
         Caption         =   "Start Date"
         Height          =   255
         Left            =   150
         TabIndex        =   32
         Top             =   1515
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Resumption Date"
         Height          =   255
         Left            =   3180
         TabIndex        =   31
         Top             =   1515
         Width           =   1860
      End
      Begin VB.Label Label4 
         Caption         =   "Notes"
         Height          =   255
         Left            =   135
         TabIndex        =   30
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Leave Duration"
         Height          =   255
         Left            =   150
         TabIndex        =   29
         Top             =   2250
         Width           =   1275
      End
      Begin VB.Label Label9 
         Caption         =   "working days"
         Height          =   240
         Left            =   2085
         TabIndex        =   28
         Top             =   2235
         Width           =   1080
      End
      Begin VB.Label Label10 
         Caption         =   "Request Date"
         Height          =   225
         Left            =   165
         TabIndex        =   27
         Top             =   225
         Width           =   1890
      End
   End
   Begin VB.Label lblstatus 
      Caption         =   "Approved"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   630
      Width           =   2295
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000001&
      Caption         =   " Other Names"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   3120
      TabIndex        =   11
      Top             =   0
      Width           =   2415
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000001&
      Caption         =   " Last Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   960
      TabIndex        =   10
      Top             =   0
      Width           =   2175
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000001&
      Caption         =   " Staff Code"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   975
   End
End
Attribute VB_Name = "frmEmpLeave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents datLeave As Recordset
Attribute datLeave.VB_VarHelpID = -1
Dim datLeaveTypes As Recordset
Dim LAdd As Boolean

Private Sub cmdAnalyze_Click()
    'This analyzes this leave and tells us whether this leave
    'should be approved or not.
    
    'First check if this employee is entitled to this leave
    Set datExtra = New Recordset
    datExtra.Open "SELECT * FROM leave_entitlement WHERE co_code ='" & _
    ID & "' AND staff_code ='" & Trim(datLeave!staff_code) & "' AND leave_code=" & _
    CLng(txtleave(0)), cn, adOpenStatic, adLockOptimistic
    
    If datExtra.RecordCount = 0 Then
       'Employee not entitled to this leave, no need to continue analysis
       Open App.Path + "\Leave.txt" For Append Access Read Write As #1
       Print #1,
       Print #1, Spc(5); "Action:         DISAPPROVE"
       Print #1,
       Print #1, Spc(5); "Explanation:  Employee not entitled to this leave type"
       Print #1,
       Print #1, Spc(5); "Alternative:    Assign this leave type to this employee under entitlements"
       Close #1
    Else
        'Entitlement Found Continue Analysis
        'Check the balances against the number of days requested
        cmdCalc_Click
        Set datExtra = New Recordset
        datExtra.Open "SELECT * FROM leave_bal WHERE co_code ='" & _
        ID & "' AND staff_code = '" & Trim(datLeave!staff_code) & _
        "' AND leave_code =" & cboLeave.BoundText & _
        " AND year = " & Operation_Year, cn, adOpenStatic, adLockOptimistic
        
        If datExtra!bal < CDbl(txtleave(4)) Then
            'Leave balance less than days required
            Open App.Path + "\Leave.txt" For Append Access Read Write As #1
            Print #1,
            Print #1, Spc(5); "Action:         DISAPPROVE"
            Print #1,
            Print #1, Spc(5); "Explanation:  Balance Less than number of days requested"
            Print #1,
            Print #1, Spc(5); "Balance:          " & datExtra!bal
            Print #1, Spc(5); "Requested Days:   " & CDbl(txtleave(4))
            Close #1
        Else
            'This leave has met all requirements so approve it
            Open App.Path + "\Leave.txt" For Append Access Read Write As #1
            Print #1,
            Print #1, Spc(5); "Action:         APPROVE"
            Print #1,
            Print #1, Spc(5); "Explanation:  Leave request has met all valid criteria"
            Print #1,
            Print #1, Spc(5); "Entitlement:      Valid"
            Print #1, Spc(5); "Balance:          " & datExtra!bal
            Print #1, Spc(5); "Requested Days:   " & CDbl(txtleave(4))
            Close #1
        End If
    End If
    frmLAna.Show vbModal
    
End Sub

Private Sub cmdApprove_Click()
    If MsgBox("Approve this leave request?", vbYesNo + vbQuestion, "Leave Requests") = vbYes Then
       'First check if this employee is entitled to this leave
       Set datExtra = New Recordset
       datExtra.Open "SELECT * FROM leave_entitlement WHERE co_code ='" & _
       ID & "' AND staff_code ='" & Trim(datLeave!staff_code) & "' AND leave_code=" & _
       CLng(txtleave(0)), cn, adOpenStatic, adLockOptimistic
       If datExtra.RecordCount = 0 Then
          MsgBox ("This leave cannot be approved!" & Chr(13) & _
          "This employee is not entitled to this leave."), vbInformation, "Leave Requests"
          Exit Sub
       Else
            'Entitlement Found Continue Analysis
            'Check the balances against the number of days requested
            cmdCalc_Click
            Set datExtra = New Recordset
            datExtra.Open "SELECT * FROM leave_bal WHERE co_code ='" & _
            ID & "' AND staff_code = '" & Trim(datLeave!staff_code) & _
            "' AND leave_code =" & cboLeave.BoundText & _
            " AND year = " & Operation_Year, cn, adOpenStatic, adLockOptimistic
            If datExtra!bal < CDbl(txtleave(4)) Then
               'Leave Balance less than the number of days required by the employee
               MsgBox ("This leave cannot be approved." & Chr(13) & _
               "The employee is requesting for more days than are available in the leave balance."), vbInformation, "Leave Requests"
               Exit Sub
            Else
                'It is okay to approve this leave
                datLeave!approved = True
                datLeave.Update
                bk = datLeave.Bookmark
                datLeave.Bookmark = bk
                MsgBox ("Leave approved!" & Chr(13) & _
                "Please remember to post the number of days taken when the employee resumes, so as to update the leave balances."), vbInformation, "Leave Requests"
                
            End If
       End If
       
    End If
End Sub

Private Sub cmdCalc_Click()
    lblDays.Visible = True
    DoEvents
    If Trim(txtleave(2)) = "" Or Trim(txtleave(2)) = "" Then
       MsgBox ("Unable to calculate the leave duration"), vbInformation, "Leave Request"
       lblDays.Visible = False
       Exit Sub
    ElseIf Not IsDate(txtleave(2)) Then
       MsgBox ("Invalid start date"), vbInformation, "Leave Request"
       txtleave(2).SetFocus
       SendKeys "{HOME}"
       SendKeys "+{END}"
       Exit Sub
    ElseIf Not IsDate(txtleave(3)) Then
       MsgBox ("Invalid start date"), vbInformation, "Leave Request"
       txtleave(3).SetFocus
       SendKeys "{HOME}"
       SendKeys "+{END}"
       Exit Sub
    End If
    '*****16th/02/2001
    'Modification made so that the last day specified should be the day when the user resumes work
    
    txtleave(4) = CalculateDays(CVDate(txtleave(2)), CVDate(txtleave(3)) - 1)
    txtleave(7) = txtleave(4)
    lblDays.Visible = False
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdDelete_Click()
    If datLeave.RecordCount = 0 Then
       MsgBox "No record to delete", vbInformation, "Leave Entitlements"
       Exit Sub
    End If
    If MsgBox("Delete this request entry?", vbInformation + vbYesNo) = vbYes Then
       datLeave.Delete
       datLeave.Requery
       MsgBox ("Record Deleted"), vbInformation
    Else
       MsgBox ("Deletion aborted"), vbInformation
    End If
End Sub

Private Sub cmdDisapprove_Click()
    If MsgBox("Disapprove this leave request?", vbYesNo + vbQuestion) = vbYes Then
       datLeave!approved = False
       datLeave.Update
       bk = datLeave.Bookmark
       datLeave.Bookmark = bk
    End If
End Sub

Private Sub cmdEdit_Click()
    If datLeave.RecordCount = 0 Then
       MsgBox ("No Record to Edit"), vbInformation, "Leave Requests"
       Exit Sub
    End If
    Reset False
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdGoTo_Click()
    If datEmp.RecordCount = 0 Then
       MsgBox ("File is empty."), vbInformation, "Find"
       Exit Sub
    End If
    Dim Response
    Response = InputBox("Go To Employee number...", "Go To")
    If Response <> "" Then
        With datEmp
            .Find "Staff_code ='" & Trim(Response) & "'", 0, adSearchForward, 1
            If datEmp.EOF Then
               MsgBox ("Employee Record not found"), vbExclamation, "Record Missing"
               datEmp.Requery
            Else
               pnlStaff_code = Trim(!staff_code)
               pnlLast_name = Trim(!last_name)
               If Not IsNull(!mid_name) Then
                  pnlother_names = Trim(!first_name) & ", " & Trim(!mid_name)
               Else
                  pnlother_names = Trim(!first_name)
               End If
               datLeave.Filter = "Staff_code ='" & Trim(pnlStaff_code) & "'"
            End If
        End With
    
    End If
End Sub

Private Sub cmdPost_Click()
    'Post this leave, reduce the balance of the number of days
    'remaining
    If MsgBox("Posting this leave as complete will adjust the leave balance " & _
        Chr(13) & "and lock any editing to be done to this leave." & _
        Chr(13) & "Would you like to proceed with the posting?", vbYesNo + vbQuestion, "Leave Posting") = vbYes Then
        
        'Validate the actual day of resuming
        If CVDate(txtleave(8)) <= CVDate(txtleave(2)) Then
           MsgBox ("Actual date of leave end cannot be earlier than or same as days of starting the leave." & Chr(13) & _
           "Leave cannot be posted."), vbCritical, "Leave Error"
           Exit Sub
        End If
        'Calculate the number of days actually taken
        txtleave(7) = CalculateDays(CVDate(txtleave(2)), CVDate(txtleave(8)) - 1)
        
        
        Set datExtra = New Recordset
        datExtra.Open "SELECT * FROM leave_bal WHERE co_code ='" & ID & _
        "' AND Staff_code ='" & datLeave!staff_code & "' AND leave_code =" & _
        datLeave!Leave_code & " AND year =" & Operation_Year, cn, adOpenStatic, adLockOptimistic
        
        If datExtra.RecordCount <> 0 Then
           datExtra!days_taken = datExtra!days_taken + CDbl(txtleave(7))
           datExtra!bal = datExtra!bal - CDbl(txtleave(7))
           datExtra.Update
           
           'Mark this leave as posted
           datLeave!posted = True
           datLeave.Update
           cmdPost.Enabled = False
        End If
        MsgBox ("Posting Complete!"), vbInformation
    Else
        MsgBox ("Posting Aborted!"), vbInformation
    End If
End Sub

Private Sub DTPicker1_CloseUp()
    txtleave(2) = Format(DTPicker1.Value, "dd/mmm/yyyy")
End Sub

Private Sub DTPicker2_CloseUp()
    txtleave(3) = Format(DTPicker2.Value, "dd/mmm/yyyy")
    txtleave(8) = txtleave(3)
End Sub

Private Sub DTPicker3_CloseUp()
    txtleave(8) = Format(DTPicker3.Value, "dd/mmm/yyyy")
End Sub

Private Sub ffirst_Click()
On Error GoTo PrevErr
    With datLeave
         .MoveFirst
    End With
    Exit Sub
PrevErr:
MsgBox (Err.Description), vbInformation
Err.Clear
End Sub

Private Sub flast_Click()
On Error GoTo PrevErr
    With datLeave
         .MoveLast
    End With
    Exit Sub
PrevErr:
MsgBox (Err.Description), vbInformation
Err.Clear
End Sub
Private Sub fnext_Click()
On Error GoTo PrevErr
    With datLeave
         .MoveNext
         If .EOF Then
            .MoveLast
            MsgBox ("Last Record encountered"), vbInformation
         End If
    End With
    Exit Sub
PrevErr:
MsgBox (Err.Description), vbInformation
Err.Clear
End Sub
Private Sub fprev_Click()
On Error GoTo PrevErr
    With datLeave
         .MovePrevious
         If .BOF Then
            .MoveFirst
            MsgBox ("First Record encountered"), vbInformation
         End If
    End With
    Exit Sub
PrevErr:
MsgBox (Err.Description), vbInformation
Err.Clear
End Sub
Private Sub cmdAdd_Click()
    If Trim(pnlStaff_code) = "" Then
       MsgBox ("Employee not selected."), vbInformation, "Leave Requests"
       Exit Sub
    End If
    
    LAdd = True
    datLeave.AddNew
    LAdd = False
    chkPaid.Value = 0
    chkunpaid.Value = 0
    txtleave(6) = Format(Date, "dd/mmm/yyyy")
    txtleave(2) = Format(Date, "dd/mmm/yyyy")
    txtleave(3) = Format(Date, "dd/mmm/yyyy")
    txtleave(8) = Format(Date, "dd/mmm/yyyy")
    Reset False
    cmdApprove.Enabled = False
    cmdDisapprove.Enabled = False
End Sub

Private Sub cmdBal_Click()
    frmLeaveBal.Show vbModal
    
End Sub

Private Sub cmdCancel_Click()
    datLeave.CancelUpdate
    Reset True
End Sub

Private Sub cmdClose_Click()
    Unload Me
    
End Sub

Private Sub cmdEntitlement_Click()
    frmlent.Show vbModal
    
End Sub

Private Sub cmdFind_Click()
    frmSelectEmp.Left = Me.Left + 40
    frmSelectEmp.Top = Me.Top + pnlStaff_code.Top + pnlStaff_code.Height
    frmSelectEmp.Show vbModal
    
    If pnlStaff_code <> "" Then
       datLeave.Filter = "Staff_code ='" & Trim(pnlStaff_code) & "'"
    End If
End Sub

Private Sub cmdLeave_Click()
    frmSelLeave.Top = Me.Top + (Label2.Top - Label6.Top) + 100
    frmSelLeave.Left = Me.Left + 60
    frmSelLeave.Show vbModal
    
End Sub

Private Sub cmdSave_Click()
    'Validate for the entries made by the user
    If txtleave(2) <> "" And txtleave(3) <> "" Then
        If CVDate(txtleave(3)) <= CVDate(txtleave(2)) Then
           MsgBox ("End date cannot be earlier or same as than start date."), vbInformation, "Leave"
           Exit Sub
        End If
    End If
    If Me.cboLeave.BoundText = "" Then
       MsgBox ("You must select the leave requested before attempting a save."), vbInformation, "Leave Requests"
       Exit Sub
    End If
    cmdCalc_Click
    datLeave!staff_code = pnlStaff_code
    datLeave!co_code = ID
    datLeave!posted = False
    datLeave.Update
    Reset True
    'lAdd = False
End Sub

Private Sub datLeave_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error Resume Next
    
    If IsNull(datLeave!approved) Then
       lblstatus = "Pending"
       lblstatus.ForeColor = &HFF
       cmdPost.Enabled = False
    ElseIf datLeave!approved = True Then
       lblstatus = "Approved"
       lblstatus.ForeColor = &HFF0000
       cmdPost.Enabled = True
    ElseIf datLeave!approved = False Then
       lblstatus = "Dis-Approved"
       lblstatus.ForeColor = &HFF
       cmdPost.Enabled = False
    Else
       cmdPost.Enabled = False
       lblstatus.Visible = False
    End If
    
    If datLeave!posted = True Then
       cmdEdit.Enabled = False
       cmdPost.Enabled = False
    Else
       cmdEdit.Enabled = True
    End If
    If datLeave.RecordCount = 0 Then
       cmdAnalyze.Enabled = False
    Else
       cmdAnalyze.Enabled = True
    End If
    
    If datLeave.RecordCount = 0 Then
       lblstatus.Visible = False
    Else
        lblstatus.Visible = True
    End If
End Sub

Private Sub DTPicker2_LostFocus()
    'If txtleave(2) <> "" And txtleave(3) <> "" Then
    '    If txtleave(3) < txtleave(2) Then
    '       MsgBox ("End date cannot be earlier than start date."), vbInformation, "Leave"
    '       txtleave(2) = ""
    '       txtleave(3) = ""
    '    End If
    'End If
End Sub

Private Sub Form_Load()
    PLeave = True
    Dim txt As TextBox
    Set datLeave = New Recordset
    datLeave.Open "SELECT * FROM emp_leaves WHERE co_code ='" & ID & _
    "'", cn, adOpenStatic, adLockOptimistic
    
    For Each txt In Me.txtleave
        Set txt.DataSource = datLeave
    Next
    Set chkPaid.DataSource = datLeave
    Set chkunpaid.DataSource = datLeave
    Set cboLeave.DataSource = datLeave
    
    Set datLeaveTypes = New Recordset
    datLeaveTypes.Open "SELECT * FROM leave_mast WHERE co_code ='" & _
    ID & "'", cn, adOpenStatic, adLockOptimistic
    Set cboLeave.RowSource = datLeaveTypes
    
    Set datEmp = New Recordset
    With datEmp
        datEmp.Open "SELECT * FROM Personal_data WHERE co_code ='" & ID & "'", cn, adOpenStatic, adLockOptimistic
        If .RecordCount <> 0 Then
            pnlStaff_code = Trim(!staff_code)
            pnlLast_name = Trim(!last_name)
            If Not IsNull(!mid_name) Then
                pnlother_names = Trim(!first_name) & ", " & Trim(!mid_name)
            Else
                pnlother_names = Trim(!first_name)
            End If
            datLeave.Filter = "Staff_code ='" & Trim(pnlStaff_code) & "'"
        Else
            'cmdModify.Enabled = False
        End If
    End With
    DTPicker1.Value = Date
    DTPicker2.Value = Date
    Reset True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PLeave = False
End Sub

Private Sub rqdate_Change()
    txtleave(6) = Format(rqdate.Value, "dd/mmm/yyyy")
End Sub

Private Sub txtleave_Change(Index As Integer)
    If Index = 0 Then
        If Trim(txtleave(0)) <> "" Then
            'Find this leave from the leaves master file
            Set datExtra = New Recordset
            datExtra.Open "SELECT * FROM leave_mast WHERE leave_code =" & _
            Trim(txtleave(0)), cn, adOpenStatic, adLockOptimistic
            If datExtra.RecordCount <> 0 Then
               txtleave(1) = datExtra!leave_name
            Else
                txtleave(1) = ""
            End If
        Else
            txtleave(1) = ""
        End If
    ElseIf Index = 2 Or Index = 3 Then
        If txtleave(2) <> "" And txtleave(3) <> "" Then
           If LAdd = True Then
              Exit Sub
           End If
           Screen.MousePointer = vbDefault
        Else
            txtleave(4) = ""
            txtleave(7) = ""
        End If
    End If
               
    
End Sub

Public Sub Reset(bval As Boolean)
    Dim txt As TextBox
    cmdAdd.Enabled = bval
    cmdEdit.Enabled = bval
    cmdSave.Enabled = Not bval
    cmdCancel.Enabled = Not bval
    cmdLeave.Enabled = Not bval
    cmdGoTo.Enabled = bval
    
    
    Me.cmdApprove.Enabled = Not bval
    Me.cmdDisapprove.Enabled = Not bval
    ffirst.Enabled = bval
    fprev.Enabled = bval
    fnext.Enabled = bval
    flast.Enabled = bval
    
    DTPicker1.Enabled = Not bval
    DTPicker2.Enabled = Not bval
    DTPicker3.Enabled = Not bval
    rqdate.Enabled = Not bval
    Frame1.Enabled = Not bval
    
    For Each txt In Me.txtleave
        txt.Locked = bval
    Next
    txtleave(1).Locked = True
    txtleave(2).Locked = True
    txtleave(3).Locked = True
    txtleave(4).Locked = True
    txtleave(6).Locked = True
    cboLeave.Enabled = Not bval
End Sub
