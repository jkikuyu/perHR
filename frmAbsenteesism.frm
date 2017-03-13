VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmAbsentee 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Employee Absenteeism"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5985
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbsenteesism.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   5985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPost 
      Caption         =   "&Post this Data to Adjust the Leave Balance"
      Height          =   345
      Left            =   45
      TabIndex        =   47
      Top             =   3735
      Width           =   5865
   End
   Begin VB.TextBox txtDays 
      DataField       =   "days"
      Height          =   330
      Left            =   3225
      Locked          =   -1  'True
      TabIndex        =   46
      Top             =   2925
      Width           =   2265
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "Calc&ulate then Hours and Duration of Absenteeism"
      Height          =   345
      Left            =   45
      TabIndex        =   43
      Top             =   3345
      Width           =   5865
   End
   Begin VB.TextBox txtDur 
      DataField       =   "hrs"
      Height          =   330
      Left            =   45
      Locked          =   -1  'True
      TabIndex        =   42
      Top             =   2940
      Width           =   2250
   End
   Begin VB.CommandButton cmdGoTo 
      Caption         =   "&Go To..."
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
      Left            =   120
      TabIndex        =   40
      Top             =   6510
      Width           =   930
   End
   Begin VB.Frame Frame2 
      Caption         =   "Time Adjustment"
      Height          =   1140
      Left            =   4275
      TabIndex        =   37
      Top             =   900
      Width           =   1635
      Begin VB.OptionButton optHr 
         Caption         =   "By the Hour"
         Height          =   255
         Left            =   135
         TabIndex        =   39
         Top             =   795
         Width           =   1305
      End
      Begin VB.OptionButton optMin 
         Caption         =   "By the Minute"
         Height          =   270
         Left            =   120
         TabIndex        =   38
         Top             =   375
         Value           =   -1  'True
         Width           =   1410
      End
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
      Left            =   5040
      Picture         =   "frmAbsenteesism.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Next Record"
      Top             =   6495
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
      Left            =   4560
      Picture         =   "frmAbsenteesism.frx":058C
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Previous Record"
      Top             =   6495
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
      Left            =   4095
      Picture         =   "frmAbsenteesism.frx":06D6
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "First Record"
      Top             =   6495
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
      Left            =   5520
      Picture         =   "frmAbsenteesism.frx":0820
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Last Record"
      Top             =   6495
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
      ScaleWidth      =   5985
      TabIndex        =   22
      Top             =   6960
      Width           =   5985
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
         Left            =   4905
         TabIndex        =   28
         Top             =   0
         Width           =   930
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
         Left            =   2025
         TabIndex        =   27
         Top             =   0
         Width           =   930
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "&New"
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
         Left            =   105
         TabIndex        =   26
         Top             =   0
         Width           =   930
      End
      Begin VB.CommandButton cmdModify 
         Caption         =   "Modify"
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
         Left            =   1065
         TabIndex        =   25
         Top             =   0
         Width           =   930
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
         Left            =   2985
         TabIndex        =   24
         Top             =   0
         Width           =   930
      End
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
         Left            =   3945
         TabIndex        =   23
         Top             =   0
         Width           =   930
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   750
      Left            =   0
      TabIndex        =   19
      Top             =   4155
      Width           =   5925
      Begin VB.OptionButton optabs 
         Caption         =   "Not to be Paid"
         Height          =   270
         Index           =   1
         Left            =   2055
         TabIndex        =   21
         Top             =   270
         Width           =   2115
      End
      Begin VB.OptionButton optabs 
         Caption         =   "To be Paid"
         Height          =   285
         Index           =   0
         Left            =   165
         TabIndex        =   20
         Top             =   285
         Value           =   -1  'True
         Width           =   1230
      End
   End
   Begin VB.TextBox txtabs 
      DataField       =   "notes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1125
      Index           =   4
      Left            =   30
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   17
      Top             =   5280
      Width           =   5910
   End
   Begin VB.TextBox txtabs 
      DataField       =   "time_to"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "h:nn AM/PM"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   4
      EndProperty
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
      Index           =   3
      Left            =   2835
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   1770
      Width           =   1170
   End
   Begin VB.TextBox txtabs 
      DataField       =   "date_to"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MMM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
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
      Index           =   2
      Left            =   30
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   1785
      Width           =   1800
   End
   Begin VB.TextBox txtabs 
      DataField       =   "time_from"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "h:nn AM/PM"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   4
      EndProperty
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
      Index           =   1
      Left            =   2820
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1035
      Width           =   1170
   End
   Begin MSComCtl2.DTPicker dtfrom 
      Height          =   300
      Left            =   1845
      TabIndex        =   10
      Top             =   1035
      Width           =   270
      _ExtentX        =   476
      _ExtentY        =   529
      _Version        =   393216
      Format          =   24707073
      CurrentDate     =   36878
   End
   Begin VB.TextBox txtabs 
      DataField       =   "date_from"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MMM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
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
      Left            =   30
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1050
      Width           =   1800
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
      Picture         =   "frmAbsenteesism.frx":096A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   270
      Width           =   375
   End
   Begin Threed.SSPanel pnlother_names 
      Height          =   375
      Left            =   2910
      TabIndex        =   1
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
      TabIndex        =   2
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
      TabIndex        =   3
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
   Begin MSComCtl2.DTPicker dtto 
      Height          =   300
      Left            =   1845
      TabIndex        =   12
      Top             =   1770
      Width           =   270
      _ExtentX        =   476
      _ExtentY        =   529
      _Version        =   393216
      Format          =   24707073
      CurrentDate     =   36878
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   330
      Left            =   4005
      TabIndex        =   33
      Top             =   1020
      Width           =   195
      _ExtentX        =   423
      _ExtentY        =   582
      _Version        =   393216
      Value           =   1
      BuddyControl    =   "Text1"
      BuddyDispid     =   196632
      OrigLeft        =   2130
      OrigTop         =   1140
      OrigRight       =   2325
      OrigBottom      =   1605
      Max             =   60
      Min             =   1
      SyncBuddy       =   -1  'True
      Wrap            =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown UpDown2 
      Height          =   330
      Left            =   4005
      TabIndex        =   35
      Top             =   1740
      Width           =   195
      _ExtentX        =   423
      _ExtentY        =   582
      _Version        =   393216
      Value           =   1
      BuddyControl    =   "Text2"
      BuddyDispid     =   196633
      OrigLeft        =   2130
      OrigTop         =   1140
      OrigRight       =   2325
      OrigBottom      =   1605
      Max             =   60
      Min             =   1
      SyncBuddy       =   -1  'True
      Wrap            =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.TextBox Text1 
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
      Left            =   3735
      TabIndex        =   34
      Text            =   "1"
      Top             =   1035
      Width           =   150
   End
   Begin VB.TextBox Text2 
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
      Left            =   3615
      TabIndex        =   36
      Text            =   "1"
      Top             =   1770
      Width           =   150
   End
   Begin VB.Label Label11 
      Caption         =   "Duration as Work Days"
      ForeColor       =   &H00800000&
      Height          =   270
      Left            =   3210
      TabIndex        =   45
      Top             =   2715
      Width           =   2085
   End
   Begin VB.Label Label10 
      Caption         =   "Duration as Number of Hrs"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   75
      TabIndex        =   44
      Top             =   2715
      Width           =   2355
   End
   Begin VB.Label Label9 
      Caption         =   "Duration of Absenteeism (Computed on Basis of Working hours as at time of absenteeism)"
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   30
      TabIndex        =   41
      Top             =   2190
      Width           =   5880
   End
   Begin VB.Label Label5 
      Caption         =   "Comments"
      Height          =   270
      Left            =   60
      TabIndex        =   18
      Top             =   5025
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Time"
      Height          =   240
      Left            =   2835
      TabIndex        =   16
      Top             =   750
      Width           =   675
   End
   Begin VB.Label Label3 
      Caption         =   "Time"
      Height          =   255
      Left            =   2880
      TabIndex        =   15
      Top             =   1545
      Width           =   585
   End
   Begin VB.Label Label2 
      Caption         =   "To Date"
      Height          =   270
      Left            =   60
      TabIndex        =   8
      Top             =   1545
      Width           =   1860
   End
   Begin VB.Label Label1 
      Caption         =   "Absent From(Date)"
      Height          =   255
      Left            =   75
      TabIndex        =   7
      Top             =   735
      Width           =   2070
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
      TabIndex        =   6
      Top             =   0
      Width           =   975
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
      TabIndex        =   5
      Top             =   0
      Width           =   2175
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
      TabIndex        =   4
      Top             =   0
      Width           =   2415
   End
End
Attribute VB_Name = "frmAbsentee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents datAbsent As Recordset
Attribute datAbsent.VB_VarHelpID = -1
Dim txt As TextBox
Dim TAdd As Boolean
Dim x As Integer      'This varaible is declared to hold the
                      'Previous value of the change in the start time
Dim y As Integer      'This varaible is declared to hold the

Private Sub cmdCalc_Click()
    Dim abhrs As Double
    Dim TimeDiff
    Dim StartTime As Date
    Dim EndTime As Date
    Dim bStart As Date
    Dim bEnd As Date
    Dim brkduration
    Dim dbreak As Long
    Dim ddur As Long
    Dim duralg As Boolean
    Dim withinbreak As Boolean
    Dim afterbreak As Long
    
    'These variables are used in the second part of this procedure where the absenteeism record spans
    'over several days
    Dim NumberOfDays As Long
    Dim EntireDuration As Long
    Dim firstday
    Dim lastday
    
    'Get the start time and the end time of the day
    Set datExtra = New Recordset
    datExtra.Open "SELECT * FROM med_opts WHERE co_code ='" & ID & _
    "'", cn, adOpenStatic, adLockOptimistic
    StartTime = Format(datExtra!day_start, "hh:mm AM/PM")
    EndTime = Format(datExtra!day_end, "hh:mm AM/PM")
    'Get the days duration
    ddur = DateDiff("s", Format(datExtra!day_start, "hh:mm AM/PM"), Format(datExtra!day_end, "hh:mm AM/PM"))
    
    'Get the breaks within the day so as to get the lenght of the working day
    Set datExtra = New Recordset
    'Get the duration of the breaks within the day
    Set datExtra = New Recordset
    datExtra.Open "SELECT * FROM d_breaks WHERE co_code ='" & ID & _
    "'", cn, adOpenStatic, adLockOptimistic
    If datExtra.RecordCount <> 0 Then
       While Not datExtra.EOF
            dbreak = dbreak + DateDiff("s", Format(datExtra!start, "hh:mm AM/PM"), Format(datExtra!End, "hh:mm AM/PM"))
            datExtra.MoveNext
       Wend
    End If
    'Day duration without the breaks
    ddur = ddur - dbreak
    
    If CVDate(txtabs(0)) = CVDate(txtabs(2)) Then
       'This is the same day so calculate the number of working hours in this day
        If Format(txtabs(1), "hh:mm AM/PM") < StartTime Then 'format(txtabs(1),"hh:mm AM/PM")
           MsgBox "The Absent From Time you have put in is earlier than the reporting time of " & StartTime, vbInformation, "Absenteeism"
        Else
            'Get the duration of the breaks within the day
            Set datExtra = New Recordset
            datExtra.Open "SELECT * FROM d_breaks WHERE co_code ='" & ID & _
            "'", cn, adOpenStatic, adLockOptimistic
            With datExtra
                If .RecordCount <> 0 Then
                   While Not .EOF
                       If CVDate(Format(txtabs(1), "hh:mm AM/PM")) < CVDate(Format(!start, "hh:mm AM/PM")) Then
                          If CVDate(Format(txtabs(3), "hh:mm AM/PM")) >= CVDate(Format(!End, "hh:mm AM/PM")) Then
                             If duralg = False Then
                                brkduration = Format(CVDate(Format(!start, "hh:mm AM/PM")) - CVDate(Format(!End, "hh:mm AM/PM")), "hh:mm")
                                'brkduration = DateDiff("s", CVDate(Format(!start, "hh:mm AM/PM")), CVDate(Format(!End, "hh:mm AM/PM")))
                                duralg = True
                             Else
                                brkduration = Format(CVDate(!start) - CVDate(!End) - CVDate(brkduration), "hh:mm")
                                'MsgBox Round(DateDiff("s", Format(CVDate("8:00"), "hh:mm"), Format(CVDate("14:00"), "hh:mm")) / (3600 * 8), 3)
                            End If
                          ElseIf DateDiff("s", Format(CVDate(!start), "hh:mm AM/PM"), Format(CVDate(!End), "hh:mm AM/PM")) > DateDiff("s", Format(CVDate(!start), "hh:mm AM/PM"), Format(CVDate(txtabs(3)), "hh:mm AM/PM")) And DateDiff("s", Format(CVDate(!start), "hh:mm AM/PM"), Format(CVDate(txtabs(3)), "hh:mm AM/PM")) >= 0 Then
                          'Format(CVDate(Format(!start, "hh:mm AM/PM")) - CVDate(Format(!End, "hh:mm AM/PM")), "hh:mm") > Format(CVDate(Format(!start, "hh:mm AM/PM")) - CVDate(Format(txtabs(3), "hh:mm AM/PM")), "hh:mm") Then
                                'The employee came back during the break
                                brkduration = Format(CVDate(Format(!start, "hh:mm AM/PM")) - CVDate(Format(txtabs(3), "hh:mm AM/PM")), "hh:mm")
                          End If
                       'Else
                       '    'Employee could have go out during the break
                       '
                       '    afterbreak = DateDiff("s", CVDate(Format(txtabs(1), "hh:mm AM/PM")), CVDate(Format(!start, "hh:mm AM/PM")))
                       '    If afterbreak <= 0 Then
                       '       brkduration = Format(CVDate(Format(txtabs(1), "hh:mm AM/PM")) - CVDate(Format(!start, "hh:mm AM/PM")), "hh:mm")
                       '       withinbreak = True
                       '    End If
                       End If
                       .MoveNext
                   Wend
                End If
            End With
            If brkduration <> "" Then
               'txtDur = Format((CVDate(txtabs(1)) - CVDate(txtabs(3))), "hh:mm") - Format(CVDate(brkduration), "hh:mm")
               'If withinbreak = True Then
               '    txtDur = Format((CVDate(txtabs(1)) - CVDate(txtabs(3))) + CVDate(brkduration), "hh:mm")
               '    txtDur = ConvertToTime(ConvertToSeconds(txtDur) + afterbreak)
               '
               'Else
                   txtDur = Format((CVDate(txtabs(1)) - CVDate(txtabs(3))) + CVDate(brkduration), "hh:mm")
               'End If
            Else
                txtDur = Format((CVDate(txtabs(1)) - CVDate(txtabs(3))), "hh:mm")
            End If
            txtDays = Round(ConvertToSeconds(txtDur) / ddur, 3)
        End If
    Else
        'The absenteeism record spans over several days
        'Get the number of days in between
        NumberOfDays = (DateDiff("d", CVDate(txtabs(0)), CVDate(txtabs(2))) + 1)
        
        'Get the first day duration of absenteeism
        'Since the span is over several days, then we shall only check for the start of the absenteeism record
        'and check for any breaks within this day
        'The code used is a duplication from above
        '********************************
            'Get the duration of the first day
            Set datExtra = New Recordset
            datExtra.Open "SELECT * FROM d_breaks WHERE co_code ='" & ID & _
            "'", cn, adOpenStatic, adLockOptimistic
            With datExtra
                If .RecordCount <> 0 Then
                   While Not .EOF
                       If CVDate(Format(txtabs(1), "hh:mm AM/PM")) <= CVDate(Format(!End, "hh:mm AM/PM")) Then
                          'This emloyee returned after the end of the break so get the duration of the break
                          brkduration = DateDiff("s", Format(CVDate(!start), "hh:mm"), Format(CVDate(!End), "hh:mm"))
                          
                       End If
                       .MoveNext
                   Wend
                End If
            End With
            
            'Now get the time from the start of the day to when this person returned
            firstday = DateDiff("s", CVDate(Format(txtabs(1), "hh:mm AM/PM")), CVDate(Format(EndTime, "hh:mm AM/PM")))
            firstday = firstday - brkduration
        
            'Get the duration of hours for the days in between
            If NumberOfDays - 2 <> 0 Then
               EntireDuration = EntireDuration + (ddur * (NumberOfDays - 2))
               
            End If
        
        
            'Get the duration of the last day
            Set datExtra = New Recordset
            datExtra.Open "SELECT * FROM d_breaks WHERE co_code ='" & ID & _
            "'", cn, adOpenStatic, adLockOptimistic
            With datExtra
                If .RecordCount <> 0 Then
                   While Not .EOF
                       If CVDate(Format(txtabs(3), "hh:mm AM/PM")) >= CVDate(Format(!End, "hh:mm AM/PM")) Then
                          'This emloyee returned after the end of the break so get the duration of the break
                          brkduration = DateDiff("s", Format(CVDate(!start), "hh:mm"), Format(CVDate(!End), "hh:mm"))
                          
                       End If
                       .MoveNext
                   Wend
                End If
            End With
            
            'Now get the time from the start of the day to when this person returned
            lastday = DateDiff("s", CVDate(Format(StartTime, "hh:mm AM/PM")), CVDate(Format(txtabs(3), "hh:mm AM/PM")))
            lastday = lastday - brkduration
                        
            EntireDuration = EntireDuration + CLng(firstday) + CLng(lastday)
            txtDur = ConvertToTime(EntireDuration)
            txtDays = Round(ConvertToSeconds(txtDur) / ddur, 3)
            
            
        'Note that any consequent days before the last day will have a duration of absenteeism of a whole day
        
        
    End If
    
    
        
    'This procedure reads the working hours from the database and calculates the number of hours of absenteeism.
    
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
            End If
        End With
    
       datAbsent.Find "Staff_code ='" & Trim(pnlStaff_code) & "'", 0, adSearchForward, 1
    End If
End Sub

Private Sub cmdPost_Click()
'On Error GoTo PostErr
    If datAbsent.RecordCount = 0 Then
       MsgBox "No record to post", vbInformation, "Absenteeism"
       Exit Sub
    End If
    If datAbsent!posted = True Then
       MsgBox "This leave value has already been posted." & Chr(13) & _
       "Please await an update from the developer to unpost a leave value", vbInformation
       Exit Sub
    End If
    Dim lcode As Long
    'Get the leave that has been assigned to the absenteeism
    Set datExtra = New Recordset
    datExtra.Open "SELECT * FROM med_opts WHERE co_code ='" & ID & _
    "'", cn, adOpenStatic, adLockOptimistic
    If datExtra.RecordCount <> 0 Then
       If Not IsNull(datExtra!LEAVE) Then
          lcode = datExtra!LEAVE
       End If
    End If
    If lcode = 0 Then
       MsgBox "A leave has not been assigned for this operation." & Chr(13) & _
       "Please assign a leave type under options", vbInformation, "Absenteeism"
       Exit Sub
    End If
    'Check for the employee's entitlement to this leave type
    Set datExtra = New Recordset
    datExtra.Open "SELECT * FROM leave_entitlement WHERE staff_code ='" & Trim(datAbsent!staff_code) & _
    "' AND leave_code =" & lcode, cn, adOpenStatic, adLockOptimistic
    If datExtra.RecordCount = 0 Then
       MsgBox "This employee is not entitled to the linked leave type." & Chr(13) & _
       "Please assign the leave type to the employee before attempting a post.", vbInformation, "Absenteeism"
       Exit Sub
    End If
    
    cn.BeginTrans
    'Post this entry and set it as a leave entry
    Set datExtra = New Recordset
    datExtra.Open "SELECT * FROM Emp_Leaves", cn, adOpenStatic, adLockOptimistic
    With datExtra
        .AddNew
        !co_code = ID
        !staff_code = Trim(datAbsent!staff_code)
        !req_date = Date
        !Leave_code = lcode
        !Start_Date = CVDate(datAbsent!date_from)
        !End_Date = CVDate(datAbsent!date_to)
        !notes = "Posted from the Absenteeism Records"
        !approved = 1
        !adays = CDbl(datAbsent!days)
        !adate = CVDate(datAbsent!date_to)
        !posted = 1
        !ext_s = 1
        .Update
    End With
    
    'Adjust the leave balance
    'Post this leave, reduce the balance of the number of days
    'remaining
    Set datExtra = New Recordset
    datExtra.Open "SELECT * FROM leave_bal WHERE co_code ='" & ID & _
    "' AND Staff_code ='" & datAbsent!staff_code & "' AND leave_code =" & _
    lcode & " AND year =" & Operation_Year, cn, adOpenStatic, adLockOptimistic
    
    If datExtra.RecordCount <> 0 Then
       datExtra!days_taken = datExtra!days_taken + CDbl(datAbsent!days)
       datExtra!bal = datExtra!bal - CDbl(datAbsent!days)
       datExtra.Update
       
    End If
    datAbsent!posted = 1
    datAbsent.Update
    bk = datAbsent.Bookmark
    datAbsent.Bookmark = bk
    cn.CommitTrans
    MsgBox ("Posting Complete!"), vbInformation
    Exit Sub
PostErr:
   cn.RollbackTrans
   MsgBox Err.Description & Chr(13) & _
   "Posting Aborted", vbInformation
   
End Sub

Private Sub datAbsent_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    If datAbsent.RecordCount = 0 Then
       cmdModify.Enabled = False
       cmdDelete.Enabled = False
    Else
       cmdModify.Enabled = True
       cmdDelete.Enabled = True
    End If
    
End Sub

Private Sub dtfrom_CloseUp()
    txtabs(0) = Format(dtfrom.Value, "dd/mmm/yyyy")
End Sub

Private Sub dtto_CloseUp()
    txtabs(2) = Format(dtto.Value, "dd/mmm/yyyy")
End Sub

                      'Previous value of the change in the start time
Private Sub ffirst_Click()
On Error GoTo PrevErr
    With datAbsent
         .MoveFirst
    End With
    Exit Sub
PrevErr:
MsgBox (Err.Description), vbInformation
Err.Clear
End Sub

Private Sub flast_Click()
On Error GoTo PrevErr
    With datAbsent
         .MoveLast
    End With
    Exit Sub
PrevErr:
MsgBox (Err.Description), vbInformation
Err.Clear
End Sub
Private Sub fnext_Click()
On Error GoTo PrevErr
    With datAbsent
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
    With datAbsent
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
Private Sub cmdCancel_Click()
'On Error GoTo CancelErr
    datAbsent.CancelUpdate
    Reset False
    Exit Sub
CancelErr:
datAbsent.CancelBatch
Reset False
'datAbsent.Requery
End Sub

Private Sub cmdClose_Click()
    Unload Me
    
End Sub

Private Sub cmdDelete_Click()
    If datAbsent.RecordCount = 0 Then
       MsgBox "No Record to Delete", vbInformation, "Modification aborted"
       Exit Sub
    End If
    If MsgBox("Delete this absenteeism record.", vbQuestion + vbYesNo) = vbYes Then
        datAbsent.Delete
        datAbsent.Requery
    Else
        MsgBox ("Deletion aborted."), vbInformation
    End If
End Sub

Private Sub cmdFind_Click()
    frmSelectEmp.Left = Me.Left + 40
    frmSelectEmp.Top = Me.Top + pnlStaff_code.Top + pnlStaff_code.Height
    frmSelectEmp.Show vbModal
    
    If pnlStaff_code <> "" Then
       datAbsent.Filter = "Staff_code ='" & Trim(pnlStaff_code) & "'"
        If datAbsent.RecordCount = 0 Then
           cmdModify.Enabled = False
           cmdDelete.Enabled = False
        Else
           cmdModify.Enabled = True
           cmdDelete.Enabled = True
        End If
    End If
End Sub

Private Sub cmdModify_Click()
    If datAbsent.RecordCount = 0 Then
       MsgBox "No Record to Modify", vbInformation, "Editing Aborted"
       Exit Sub
    End If
    Reset True
    
End Sub

Private Sub cmdNew_Click()
    If Trim(pnlStaff_code) = "" Then
       MsgBox ("No Employee selected."), vbInformation, "Absenteeism"
       Exit Sub
    End If
    
    datAbsent.AddNew
    TAdd = True
    'Show todays date in the two date fields
    txtabs(0) = Format(Now, "dd/mmm/yyyy")
    txtabs(2) = Format(Now, "dd/mmm/yyyy")
    
    'Set time to noon
    txtabs(1) = Format("12:00", "hh:mm AM/PM")
    txtabs(3) = Format("12:00", "hh:mm AM/PM")
    datAbsent!co_code = ID
    Reset True
End Sub

Private Sub cmdSave_Click()
    If optabs(0).Value = True Then
       datAbsent!pay_nopay = True
    Else
        datAbsent!pay_nopay = False
    End If
    datAbsent!staff_code = pnlStaff_code
    datAbsent!posted = False
    datAbsent.Update
    Reset False
    
End Sub

Private Sub Form_Load()
    PAbs = True
    Set datAbsent = New Recordset
    datAbsent.Open "SELECT * FROM emp_Absentee WHERE co_code ='" & ID & _
    "'", cn, adOpenStatic, adLockOptimistic
    For Each txt In Me.txtabs
        Set txt.DataSource = datAbsent
    Next
    Set txtDur.DataSource = datAbsent
    Set txtDays.DataSource = datAbsent
    
    Reset False
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
            datAbsent.Filter = "Staff_code ='" & Trim(pnlStaff_code) & "'"
            If datAbsent.RecordCount = 0 Then
               cmdModify.Enabled = False
               cmdDelete.Enabled = False
            End If
        Else
            cmdModify.Enabled = False
        End If
    End With
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PAbs = False
End Sub

Private Sub UpDown1_Change()
    If (x = 1 Or x = 0) And UpDown1.Value = 60 Then
        TAdd = False
    ElseIf x = 60 And UpDown1.Value = 1 Then
        TAdd = True
    Else
        If x > UpDown1.Value Then
           'This is a decrease
           TAdd = False
        Else
            TAdd = True
        End If
    End If
    x = UpDown1.Value
    If TAdd = True Then
        If optMin.Value = True Then
           txtabs(1) = Format(CVDate(Format(txtabs(1), "hh:mm AM/PM")) + CVDate(Format("00:01", "hh:mm AM/PM")), "hh:mm AM/PM")
        Else
           txtabs(1) = Format(CVDate(Format(txtabs(1), "hh:mm AM/PM")) + CVDate(Format("01:00", "hh:mm AM/PM")), "hh:mm AM/PM")
        End If
    Else
        If optMin.Value = True Then
            txtabs(1) = Format(CVDate(Format(txtabs(1), "hh:mm AM/PM")) - CVDate(Format("00:01", "hh:mm AM/PM")), "hh:mm AM/PM")
        Else
            txtabs(1) = Format(CVDate(Format(txtabs(1), "hh:mm AM/PM")) - CVDate(Format("01:00", "hh:mm AM/PM")), "hh:mm AM/PM")
        End If
    End If
End Sub

Private Sub tfrom_UpClick()
    'Increment the time shown
    txtabs(1) = Format(txtabs(1), "hh:mm") + Format(0.01, "hh:mm")
End Sub

Private Sub UpDown2_Change()
    If (y = 1 Or y = 0) And UpDown2.Value = 60 Then
        TAdd = False
    ElseIf y = 60 And UpDown2.Value = 1 Then
        TAdd = True
    Else
        If y > UpDown2.Value Then
           'This is a decrease
           TAdd = False
        Else
            TAdd = True
        End If
    End If
    y = UpDown2.Value
    If TAdd = True Then
        If optMin.Value = True Then
           txtabs(3) = Format(CVDate(Format(txtabs(3), "hh:mm AM/PM")) + CVDate(Format("00:01", "hh:mm AM/PM")), "hh:mm AM/PM")
        Else
           txtabs(3) = Format(CVDate(Format(txtabs(3), "hh:mm AM/PM")) + CVDate(Format("01:00", "hh:mm AM/PM")), "hh:mm AM/PM")
        End If
    Else
        If optMin.Value = True Then
           txtabs(3) = Format(CVDate(Format(txtabs(3), "hh:mm AM/PM")) - CVDate(Format("00:01", "hh:mm AM/PM")), "hh:mm AM/PM")
        Else
           txtabs(3) = Format(CVDate(Format(txtabs(3), "hh:mm AM/PM")) - CVDate(Format("01:00", "hh:mm AM/PM")), "hh:mm AM/PM")
        End If
    End If
End Sub

Public Sub Reset(bval As Boolean)
    UpDown1.Enabled = bval
    UpDown2.Enabled = bval
    cmdCalc.Enabled = bval
    cmdPost.Enabled = Not bval
    cmdNew.Enabled = Not bval
    cmdSave.Enabled = bval
    cmdModify.Enabled = Not bval
    cmdCancel.Enabled = bval
    cmdDelete.Enabled = Not bval
    dtfrom.Enabled = bval
    dtto.Enabled = bval
    txtabs(4).Locked = Not bval
    
    ffirst.Enabled = Not bval
    flast.Enabled = Not bval
    fnext.Enabled = Not bval
    fprev.Enabled = Not bval
    cmdFind.Enabled = Not bval
End Sub

Public Function ConvertToSeconds(Time As String)
    'Patrick Odawo
    'This function has been developed to convert a sent time string into seconds
    'The length of the sent string should be ideally 5 with a ":" at position 3
    Dim leftstr
    Dim rightstr
    Dim colonposn
    Dim Hrsec As Long
    Dim Minsec As Long
    
    colonposn = InStr(1, Time, ":")
    leftstr = Left(Time, colonposn - 1)
    rightstr = Right(Time, Len(Time) - colonposn)
    
    'Get the seconds in the hour portion of this time
    Hrsec = CLng(leftstr) * 3600
    Minsec = CLng(rightstr) * 60
    
    ConvertToSeconds = Hrsec + Minsec
    
End Function

Public Function ConvertToTime(seconds As Long)
    'Patrick Odawo
    'Function converts the applied seconds and gives a definite hour:minute format
    Dim Hrs
    Dim Min
    
    Hrs = "00"
    Min = "00"
    
    Hrs = Abs(seconds / 3600)
    If seconds - (Hrs * 3600) <> 0 Then
       Min = (seconds - (Hrs * 3600)) / 60
    End If
    ConvertToTime = Hrs & ":" & Min
    
End Function
