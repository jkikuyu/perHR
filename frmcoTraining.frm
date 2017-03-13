VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmcoTraining 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Company Training Scheduling"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7380
   Icon            =   "frmcoTraining.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Period Span"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   0
      TabIndex        =   4
      Top             =   45
      Width           =   5250
      Begin VB.TextBox txttr 
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
         Height          =   300
         Index           =   3
         Left            =   2445
         TabIndex        =   25
         Top             =   525
         Width           =   1770
      End
      Begin VB.TextBox txttr 
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
         Height          =   300
         Index           =   2
         Left            =   60
         TabIndex        =   24
         Top             =   525
         Width           =   1770
      End
      Begin MSComCtl2.DTPicker dtEnd 
         Height          =   300
         Left            =   4200
         TabIndex        =   30
         Top             =   525
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   529
         _Version        =   393216
         Format          =   24379393
         CurrentDate     =   36965
      End
      Begin MSComCtl2.DTPicker dtStart 
         Height          =   300
         Left            =   1815
         TabIndex        =   29
         Top             =   525
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   529
         _Version        =   393216
         Format          =   24379393
         CurrentDate     =   36965
      End
      Begin VB.Label Label3 
         Caption         =   "End Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2460
         TabIndex        =   6
         Top             =   255
         Width           =   795
      End
      Begin VB.Label Label2 
         Caption         =   "Start Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   60
         TabIndex        =   5
         Top             =   270
         Width           =   855
      End
   End
   Begin VB.TextBox txttr 
      DataField       =   "training_code"
      Height          =   285
      Index           =   0
      Left            =   5370
      TabIndex        =   27
      Top             =   345
      Width           =   1680
   End
   Begin VB.CheckBox chktr 
      Caption         =   "Suspend Training"
      DataField       =   "suspend"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   60
      TabIndex        =   26
      Top             =   5610
      Width           =   1830
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   390
      Left            =   0
      ScaleHeight     =   390
      ScaleWidth      =   7380
      TabIndex        =   19
      Top             =   5955
      Width           =   7380
      Begin VB.CommandButton cmdModify 
         Caption         =   "&Modify"
         Height          =   350
         Left            =   2820
         TabIndex        =   23
         Top             =   30
         Width           =   1100
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   3960
         TabIndex        =   22
         Top             =   30
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   5115
         TabIndex        =   21
         Top             =   30
         Width           =   1100
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "C&lose"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   6255
         TabIndex        =   20
         Top             =   30
         Width           =   1100
      End
   End
   Begin VB.TextBox txttr 
      DataField       =   "notes"
      Height          =   1125
      Index           =   5
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   18
      Top             =   4350
      Width           =   7230
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   ">>"
      Height          =   390
      Left            =   3480
      TabIndex        =   17
      Top             =   3345
      Width           =   390
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "<<"
      Height          =   390
      Left            =   3495
      TabIndex        =   16
      Top             =   2775
      Width           =   390
   End
   Begin VB.ListBox lsttrainees 
      Height          =   1425
      Left            =   3960
      TabIndex        =   15
      Top             =   2565
      Width           =   3330
   End
   Begin VB.ListBox lstemps 
      Height          =   1425
      Left            =   60
      TabIndex        =   14
      Top             =   2565
      Width           =   3330
   End
   Begin VB.TextBox txttr 
      DataField       =   "trainer"
      Height          =   285
      Index           =   4
      Left            =   75
      TabIndex        =   13
      Top             =   1875
      Width           =   7245
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print Notifications"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2445
      Picture         =   "frmcoTraining.frx":030A
      TabIndex        =   12
      Top             =   5535
      Width           =   1635
   End
   Begin VB.CommandButton cmdTopics 
      Caption         =   "S&hedule Topics"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5775
      TabIndex        =   11
      Top             =   5535
      Width           =   1560
   End
   Begin VB.CommandButton cmdComplete 
      Caption         =   "&Mark Complete"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4140
      TabIndex        =   10
      Top             =   5535
      Width           =   1560
   End
   Begin VB.CommandButton cmdTraining 
      Height          =   285
      Left            =   7080
      Picture         =   "frmcoTraining.frx":047C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   345
      Width           =   285
   End
   Begin VB.TextBox txttr 
      DataField       =   "training_name"
      Height          =   285
      Index           =   1
      Left            =   45
      TabIndex        =   0
      Top             =   1260
      Width           =   7275
   End
   Begin VB.Label Label5 
      Caption         =   "Training Code"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5355
      TabIndex        =   28
      Top             =   120
      Width           =   1350
   End
   Begin VB.Label Label9 
      Caption         =   "Notes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   45
      TabIndex        =   9
      Top             =   4110
      Width           =   1440
   End
   Begin VB.Label Label8 
      Caption         =   "Selected Trainees"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3990
      TabIndex        =   8
      Top             =   2250
      Width           =   1740
   End
   Begin VB.Label Label7 
      Caption         =   "Company Employees"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   45
      TabIndex        =   7
      Top             =   2250
      Width           =   1965
   End
   Begin VB.Label Label4 
      Caption         =   "Trainer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   45
      TabIndex        =   3
      Top             =   1650
      Width           =   675
   End
   Begin VB.Label Label1 
      Caption         =   "Training Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   60
      TabIndex        =   1
      Top             =   1005
      Width           =   1485
   End
End
Attribute VB_Name = "frmcoTraining"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim datTraining As Recordset
Private Sub Check1_Click()

End Sub

Private Sub cmdCancel_Click()
    datTraining.CancelUpdate
    Reset True
End Sub

Private Sub cmdClose_Click()
    Unload Me
    
End Sub

Private Sub cmdModify_Click()
    If datTraining.RecordCount = 0 Then
       MsgBox ("No Record to Edit."), vbInformation, "Company Trainings"
       Exit Sub
    End If
    Reset False
    
End Sub

Private Sub cmdSave_Click()
    datTraining!co_code = ID
    datTraining.Update
    Reset True
End Sub

Private Sub cmdTopics_Click()
    If datTraining.RecordCount <> 0 Then
        Tp = datTraining!training_code
        frmTopicScheduling.Left = Screen.Width - frmTopicScheduling.Width
        frmTopicScheduling.Top = (Screen.Height * 0.85) / 2 - frmTopicScheduling.Height / 2
        frmTopicScheduling.Show vbModal
    Else: MsgBox ("No Training selected"), vbInformation, "Training Scheduling"
    End If
End Sub

Private Sub cmdTraining_Click()
    frmTheTrainings.Top = Me.Top + Frame1.Height - 50
    frmTheTrainings.Left = Me.Left + Me.Width - frmTheTrainings.Width
    frmTheTrainings.Show vbModal
    
    If Trsb = True Then
       'Calculate the new training code
       Set datExtra = New Recordset
       datExtra.Open "SELECT TOP 1 * FROM co_training WHERE co_code ='" & ID & _
       "' ORDER BY Training_code DESC", cn, adOpenStatic, adLockOptimistic
       If datExtra.RecordCount = 0 Then
          x = 1
       Else
          x = datExtra!training_code + 1
       End If
       
       datTraining.AddNew
       txttr(2) = Date
       txttr(3) = Date
       txttr(0) = x
       Trsb = False
       Reset False
       
    ElseIf Trs <> 0 Then
       datTraining.Find "training_code =" & Trs, 0, adSearchForward, 1
       Trs = 0
    End If
End Sub

Private Sub dtEnd_CloseUp()
    txttr(3) = Format(dtEnd.Value, "dd/mmm/yyyy")
End Sub

Private Sub dtStart_CloseUp()
    txttr(2) = Format(dtStart.Value, "dd/mmm/yyyy")
    
End Sub

Private Sub Form_Load()
    Dim txt As TextBox
    Set datTraining = New Recordset
    datTraining.Open "SELECT * FROM co_training WHERE co_code ='" & ID & _
    "'", cn, adOpenStatic, adLockOptimistic
    
    For Each txt In Me.txttr
        Set txt.DataSource = datTraining
    Next
    Set chktr.DataSource = datTraining
    
    'Add the employees of this company to the list box
    Set datExtra = New Recordset
    datExtra.Open "SELECT * FROM Personal_data WHERE co_code ='" & ID & _
    "'", cn, adOpenStatic, adLockOptimistic
    If datExtra.RecordCount <> 0 Then
       While Not datExtra.EOF
            lstemps.AddItem datExtra!staff_code & " " & datExtra!last_name & " " & datExtra!first_name
            datExtra.MoveNext
       Wend
    End If
    Reset True
    
End Sub

Public Sub Reset(bval As Boolean)
Attribute Reset.VB_Description = "Procedure for enabling and disabling text boxes and command buttons"
    Dim txt As TextBox
    cmdTraining.Enabled = bval
    cmdSave.Enabled = Not bval
    cmdCancel.Enabled = Not bval
    cmdModify.Enabled = bval
    cmdComplete.Enabled = bval
    cmdPrint.Enabled = bval
    cmdTopics.Enabled = bval
    cmdAdd.Enabled = Not bval
    cmdRemove.Enabled = Not bval
    dtStart.Enabled = Not bval
    dtEnd.Enabled = Not bval
    
    
    For Each txt In Me.txttr
        txt.Locked = bval
    Next
    Me.chktr.Enabled = Not bval
    
    txttr(0).Locked = True
    txttr(2).Locked = True
    txttr(3).Locked = True
    
End Sub

Private Sub txtTr_GotFocus(Index As Integer)
    txttr(Index) = Trim(txttr(Index))
    
End Sub
