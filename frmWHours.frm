VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmWHours 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Working Hours"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5880
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdQuit 
      Caption         =   "&Quit"
      Height          =   360
      Left            =   4740
      TabIndex        =   26
      Top             =   5310
      Width           =   1080
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   360
      Left            =   4110
      TabIndex        =   25
      Top             =   360
      Width           =   840
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   360
      Left            =   4965
      TabIndex        =   24
      Top             =   360
      Width           =   840
   End
   Begin VB.CommandButton cmdModify 
      Caption         =   "&Modify"
      Height          =   360
      Left            =   3255
      TabIndex        =   23
      Top             =   360
      Width           =   840
   End
   Begin VB.Frame Frame2 
      Caption         =   "Time Adjustment Mode"
      Height          =   675
      Left            =   30
      TabIndex        =   18
      Top             =   4560
      Width           =   5805
      Begin VB.OptionButton optHr 
         Caption         =   "By the Hour"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   285
         Width           =   1305
      End
      Begin VB.OptionButton optMin 
         Caption         =   "By the Minute"
         Height          =   270
         Left            =   2895
         TabIndex        =   19
         Top             =   285
         Value           =   -1  'True
         Width           =   1410
      End
   End
   Begin VB.TextBox txtabs 
      DataField       =   "day_start"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "h:nn AM/PM"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   4
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   165
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   390
      Width           =   1170
   End
   Begin VB.TextBox txtabs 
      DataField       =   "day_end"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "h:nn AM/PM"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   4
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   390
      Width           =   1170
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   3060
      Left            =   15
      TabIndex        =   3
      Top             =   870
      Width           =   5865
      _Version        =   65536
      _ExtentX        =   10345
      _ExtentY        =   5397
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
         Left            =   4935
         Picture         =   "frmWHours.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Next Record"
         Top             =   660
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
         Left            =   4455
         Picture         =   "frmWHours.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Previous Record"
         Top             =   660
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
         Left            =   3990
         Picture         =   "frmWHours.frx":0294
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "First Record"
         Top             =   660
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
         Left            =   5415
         Picture         =   "frmWHours.frx":03DE
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Last Record"
         Top             =   660
         Width           =   375
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   5280
         Top             =   1845
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmWHours.frx":0528
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmWHours.frx":0980
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmWHours.frx":0DD8
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmWHours.frx":1230
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmWHours.frx":1688
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlb 
         Height          =   570
         Left            =   45
         TabIndex        =   22
         Top             =   45
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   1005
         ButtonWidth     =   1032
         ButtonHeight    =   1005
         Appearance      =   1
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   6
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Add"
               Object.ToolTipText     =   "Add"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Edit"
               Object.ToolTipText     =   "Edit"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Save"
               Object.ToolTipText     =   "Save"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Cancel"
               Object.ToolTipText     =   "Cancel"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Delete"
               Object.ToolTipText     =   "Delete"
               ImageIndex      =   5
            EndProperty
         EndProperty
      End
      Begin VB.TextBox txtDur 
         Height          =   330
         Left            =   1680
         TabIndex        =   21
         Text            =   "Text2"
         Top             =   2340
         Width           =   825
      End
      Begin VB.TextBox txtbrk 
         Height          =   330
         Left            =   1260
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   1215
         Width           =   4515
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
         Height          =   285
         Index           =   3
         Left            =   3330
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1935
         Width           =   1170
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
         Height          =   285
         Index           =   1
         Left            =   135
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1935
         Width           =   1170
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   330
         Left            =   1320
         TabIndex        =   10
         Top             =   1905
         Width           =   195
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   393216
         Value           =   1
         OrigLeft        =   2130
         OrigTop         =   1140
         OrigRight       =   2325
         OrigBottom      =   1605
         Max             =   60
         Min             =   1
         Wrap            =   -1  'True
         Enabled         =   0   'False
      End
      Begin MSComCtl2.UpDown UpDown2 
         Height          =   330
         Left            =   4500
         TabIndex        =   12
         Top             =   1905
         Width           =   195
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   393216
         Value           =   1
         OrigLeft        =   2130
         OrigTop         =   1140
         OrigRight       =   2325
         OrigBottom      =   1605
         Max             =   60
         Min             =   1
         Wrap            =   -1  'True
         Enabled         =   0   'False
      End
      Begin VB.Label Label8 
         Caption         =   "Duration of Break:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   2400
         Width           =   1740
      End
      Begin VB.Label Label7 
         Caption         =   "End Time:"
         Height          =   315
         Left            =   3315
         TabIndex        =   7
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label Label6 
         Caption         =   "Start Time:"
         Height          =   270
         Left            =   135
         TabIndex        =   6
         Top             =   1680
         Width           =   1470
      End
      Begin VB.Label Label5 
         Caption         =   "Break Name"
         Height          =   270
         Left            =   135
         TabIndex        =   5
         Top             =   1230
         Width           =   1260
      End
      Begin VB.Label Label4 
         Caption         =   "Defined Breaks During the Day"
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   150
         TabIndex        =   4
         Top             =   825
         Width           =   3795
      End
   End
   Begin MSComCtl2.UpDown UpDown3 
      Height          =   330
      Left            =   1350
      TabIndex        =   16
      Top             =   360
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
      _Version        =   393216
      Value           =   1
      OrigLeft        =   2130
      OrigTop         =   1140
      OrigRight       =   2325
      OrigBottom      =   1605
      Max             =   60
      Min             =   1
      Wrap            =   -1  'True
      Enabled         =   0   'False
   End
   Begin MSComCtl2.UpDown UpDown4 
      Height          =   330
      Left            =   2970
      TabIndex        =   17
      Top             =   360
      Width           =   195
      _ExtentX        =   423
      _ExtentY        =   582
      _Version        =   393216
      Value           =   1
      OrigLeft        =   2130
      OrigTop         =   1140
      OrigRight       =   2325
      OrigBottom      =   1605
      Max             =   60
      Min             =   1
      Wrap            =   -1  'True
      Enabled         =   0   'False
   End
   Begin VB.Label Label3 
      Caption         =   "Day ends at"
      Height          =   270
      Left            =   1800
      TabIndex        =   2
      Top             =   120
      Width           =   1350
   End
   Begin VB.Label Label2 
      Caption         =   "Day Starts at"
      Height          =   300
      Left            =   165
      TabIndex        =   1
      Top             =   135
      Width           =   1470
   End
   Begin VB.Label Label1 
      Caption         =   "Number of Working Hours in a day"
      Height          =   315
      Left            =   105
      TabIndex        =   0
      Top             =   4050
      Width           =   2985
   End
End
Attribute VB_Name = "frmWHours"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TAdd As Boolean
Dim x As Integer      'This varaible is declared to hold the
                      'Previous value of the change in the start time
Dim y As Integer      'This varaible is declared to hold the
Dim datHrs As Recordset
Dim datBreaks As Recordset
Dim bCount As Long


Private Sub ffirst_Click()
On Error GoTo PrevErr
    With datBreaks
         .MoveFirst
         ShowBreaks
    End With
    Exit Sub
PrevErr:
MsgBox (Err.Description), vbInformation
Err.Clear
End Sub

Private Sub flast_Click()
On Error GoTo PrevErr
    With datBreaks
         .MoveLast
         ShowBreaks
    End With
    Exit Sub
PrevErr:
MsgBox (Err.Description), vbInformation
Err.Clear
End Sub

Private Sub fnext_Click()
On Error GoTo PrevErr
    With datBreaks
         .MoveNext
         If .EOF Then
            .MoveLast
            MsgBox ("Last Record encountered"), vbInformation
         End If
         ShowBreaks
    End With
    Exit Sub
PrevErr:
MsgBox (Err.Description), vbInformation
Err.Clear
End Sub
Private Sub fprev_Click()
On Error GoTo PrevErr
    With datBreaks
         .MovePrevious
         If .BOF Then
            .MoveFirst
            MsgBox ("First Record encountered"), vbInformation
         End If
         ShowBreaks
    End With
    Exit Sub
PrevErr:
MsgBox (Err.Description), vbInformation
Err.Clear
End Sub

Private Sub cmdCancel_Click()
    datHrs.CancelUpdate
    cn.RollbackTrans
    'bk = datHrs.Bookmark
    'datHrs.Bookmark = bk
    datHrs.Requery
    txtabs(2) = Format(datHrs!day_start, "hh:mm AM/PM")
    txtabs(0) = Format(datHrs!day_end, "hh:mm AM/PM")
    
    Reset False
End Sub

Private Sub cmdModify_Click()
    cn.BeginTrans
    Reset True
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    datHrs!day_start = Format(txtabs(2), "hh:mm")
    datHrs!day_end = Format(txtabs(0), "hh:mm")
    datHrs.Update
    cn.CommitTrans
    datHrs.Requery
    Reset False
End Sub

Private Sub Form_Load()
    'Set time to noon
    txtabs(1) = Format("12:00", "hh:mm AM/PM")
    txtabs(3) = Format("12:00", "hh:mm AM/PM")
    
    Set datHrs = New Recordset
    datHrs.Open "SELECT * FROM Med_Opts WHERE Co_Code='" & ID & "'", cn, adOpenStatic, adLockOptimistic
    
    Set datBreaks = New Recordset
    datBreaks.Open "SELECT * FROM d_breaks WHERE co_Code ='" & ID & "'", cn, adOpenStatic, adLockOptimistic
    
    'Get the time from the options file and display it
    txtabs(2) = Format(datHrs!day_start, "hh:mm AM/PM")
    txtabs(0) = Format(datHrs!day_end, "hh:mm AM/PM")
    ShowBreaks
    Reset False
    
End Sub

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Add"
            Set datExtra = New Recordset
            datExtra.Open "SELECT * FROM d_breaks", cn, adOpenStatic, adLockOptimistic
            If datExtra.RecordCount = 0 Then
               bCount = 1
            Else
               datExtra.MoveLast
               bCount = datExtra!break_id + 1
            End If
            With datBreaks
                BrkReset True
                .AddNew
                txtbrk = ""
                txtabs(1) = Format("00:00", "hh:mm AM/PM")
                txtabs(3) = Format("00:00", "hh:mm AM/PM")
                txtDur = ""
            End With
        Case "Save"
            SaveBreaks
        Case "Cancel"
            datBreaks.CancelUpdate
            ShowBreaks
        Case "Delete"
            If MsgBox("Delete this Break?", vbYesNo + vbQuestion) = vbYes Then
               datBreaks.Delete
               datBreaks.Requery
               ShowBreaks
            End If
    End Select
End Sub

Private Sub txtabs_Change(Index As Integer)
    If Index = 3 Or Index = 1 Then
       If Trim(txtabs(1)) <> "" And Trim(txtabs(3)) <> "" Then
       'Get the time difference
       txtDur = Format(CVDate(Format(txtabs(1), "hh:mm AM/PM")) - CVDate(Format(txtabs(3), "hh:mm AM/PM")), "hh:mm")
       End If
    End If
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


Private Sub UpDown3_Change()
    If (x = 1 Or x = 0) And UpDown3.Value = 60 Then
        TAdd = False
    ElseIf x = 60 And UpDown3.Value = 1 Then
        TAdd = True
    Else
        If x > UpDown3.Value Then
           'This is a decrease
           TAdd = False
        Else
            TAdd = True
        End If
    End If
    x = UpDown3.Value
    If TAdd = True Then
        If optMin.Value = True Then
           txtabs(2) = Format(CVDate(Format(txtabs(2), "hh:mm AM/PM")) + CVDate(Format("00:01", "hh:mm AM/PM")), "hh:mm AM/PM")
        Else
           txtabs(2) = Format(CVDate(Format(txtabs(2), "hh:mm AM/PM")) + CVDate(Format("01:00", "hh:mm AM/PM")), "hh:mm AM/PM")
        End If
    Else
        If optMin.Value = True Then
            txtabs(2) = Format(CVDate(Format(txtabs(2), "hh:mm AM/PM")) - CVDate(Format("00:01", "hh:mm AM/PM")), "hh:mm AM/PM")
        Else
            txtabs(2) = Format(CVDate(Format(txtabs(2), "hh:mm AM/PM")) - CVDate(Format("01:00", "hh:mm AM/PM")), "hh:mm AM/PM")
        End If
    End If
End Sub

Private Sub UpDown4_Change()
    If (y = 1 Or y = 0) And UpDown4.Value = 60 Then
        TAdd = False
    ElseIf y = 60 And UpDown4.Value = 1 Then
        TAdd = True
    Else
        If y > UpDown4.Value Then
           'This is a decrease
           TAdd = False
        Else
            TAdd = True
        End If
    End If
    y = UpDown4.Value
    If TAdd = True Then
        If optMin.Value = True Then
           txtabs(0) = Format(CVDate(Format(txtabs(0), "hh:mm AM/PM")) + CVDate(Format("00:01", "hh:mm AM/PM")), "hh:mm AM/PM")
        Else
           txtabs(0) = Format(CVDate(Format(txtabs(0), "hh:mm AM/PM")) + CVDate(Format("01:00", "hh:mm AM/PM")), "hh:mm AM/PM")
        End If
    Else
        If optMin.Value = True Then
           txtabs(0) = Format(CVDate(Format(txtabs(0), "hh:mm AM/PM")) - CVDate(Format("00:01", "hh:mm AM/PM")), "hh:mm AM/PM")
        Else
           txtabs(0) = Format(CVDate(Format(txtabs(0), "hh:mm AM/PM")) - CVDate(Format("01:00", "hh:mm AM/PM")), "hh:mm AM/PM")
        End If
    End If
End Sub

Public Sub Reset(bval As Boolean)
    UpDown3.Enabled = bval
    UpDown4.Enabled = bval
    cmdSave.Enabled = bval
    cmdModify.Enabled = Not bval
    cmdCancel.Enabled = bval
End Sub

Public Sub ShowBreaks()
    If datBreaks.RecordCount <> 0 Then
        txtabs(1) = Format(datBreaks!start, "hh:mm AM/PM")
        txtabs(3) = Format(datBreaks!End, "hh:mm AM/PM")
        txtbrk = datBreaks!Name
    End If
End Sub

Public Sub SaveBreaks()
    datBreaks!co_code = ID
    datBreaks!Name = Trim(txtbrk)
    datBreaks!start = Format(txtabs(1), "hh:mm")
    datBreaks!End = Format(txtabs(3), "hh:mm")
    datBreaks!break_id = bCount
    datBreaks.Update
End Sub

Public Sub BrkReset(bval As Boolean)
    UpDown1.Enabled = bval
    UpDown2.Enabled = bval
    
End Sub
