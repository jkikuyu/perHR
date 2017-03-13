VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmReportManager 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Report Manager"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5070
   Icon            =   "frmReportManager.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Report Options"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   0
      TabIndex        =   15
      Top             =   3600
      Width           =   4935
      Begin VB.CheckBox chkDate 
         Caption         =   "Provide Date Range"
         DataField       =   "date_range"
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
         Left            =   240
         TabIndex        =   18
         Top             =   1080
         Width           =   3135
      End
      Begin VB.CheckBox chkWage 
         Caption         =   "Filter by Wage Code"
         DataField       =   "wage_filter"
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
         Left            =   240
         TabIndex        =   17
         Top             =   720
         Width           =   4575
      End
      Begin VB.CheckBox chkEmp 
         Caption         =   "Filter by Employee Staff Code"
         DataField       =   "emp_filter"
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
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   4575
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   5070
      TabIndex        =   10
      Top             =   5040
      Width           =   5070
      Begin VB.CommandButton cmdCancel 
         Caption         =   "C&ancel"
         Enabled         =   0   'False
         Height          =   350
         Left            =   2880
         TabIndex        =   14
         Top             =   135
         Width           =   975
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Update"
         Enabled         =   0   'False
         Height          =   350
         Left            =   1800
         TabIndex        =   13
         Top             =   135
         Width           =   975
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   350
         Left            =   3960
         TabIndex        =   12
         Top             =   135
         Width           =   975
      End
      Begin VB.CommandButton cmdModify 
         Caption         =   "&Modify"
         Height          =   350
         Left            =   720
         TabIndex        =   11
         Top             =   135
         Width           =   975
      End
   End
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   4200
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "B&rowse"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3960
      TabIndex        =   9
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton cmdFind 
      Height          =   300
      Left            =   960
      Picture         =   "frmReportManager.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   300
   End
   Begin VB.TextBox txtrpt 
      DataField       =   "report_desc"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   3
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   2400
      Width           =   4935
   End
   Begin VB.TextBox txtrpt 
      DataField       =   "report_path"
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
      Index           =   2
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1800
      Width           =   4935
   End
   Begin VB.TextBox txtrpt 
      DataField       =   "report_name"
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
      Index           =   1
      Left            =   0
      MaxLength       =   100
      TabIndex        =   6
      Top             =   1080
      Width           =   4935
   End
   Begin VB.TextBox txtrpt 
      DataField       =   "report_id"
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
      Index           =   0
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Description"
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
      Left            =   0
      TabIndex        =   5
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Report Path"
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
      Left            =   0
      TabIndex        =   4
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Report Name"
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
      Left            =   0
      TabIndex        =   3
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Report ID"
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
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmReportManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim datRpt As Recordset

Private Sub cmdBrowse_Click()
    cdlg.InitDir = App.Path
    cdlg.Filter = "Crystal Report Files(*.rpt)|*.rpt"
    cdlg.DialogTitle = "Select a Report File"
    cdlg.ShowOpen
    If cdlg.FileName <> "" Then
       txtRpt(2) = cdlg.FileName
    End If
    
End Sub

Private Sub cmdCancel_Click()
    datRpt.CancelUpdate
    
    
    ResetButtons True
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    frmShowRptHr.Left = Me.Left + 50
    frmShowRptHr.Top = Me.Top + (Label2.Top - Label1.Top) + txtRpt(0).Height
    frmShowRptHr.Show vbModal
    If NewRpt = True Then
       Set datExtra = New Recordset
       datExtra.Open "SELECT TOP 1 * FROM hr_reports ORDER BY report_id DESC", cn, adOpenStatic, adLockOptimistic
       
       datRpt.AddNew
       If datExtra.RecordCount = 0 Then
          txtRpt(0) = 1
       Else
          txtRpt(0) = datExtra!Report_id + 1
       End If
       chkEmp.Value = 0
       chkWage.Value = 0
       chkDate.Value = 0
       
       ResetButtons False
       NewRpt = False
    Else
        If SelRpt <> "" Then
            datRpt.Find "Report_id =" & CLng(SelRpt), 0, adSearchForward, 1
            SelRpt = ""
        End If
    End If
End Sub

Private Sub cmdModify_Click()
    ResetButtons False
    
End Sub

Private Sub cmdSave_Click()
'On Error GoTo SaveErr
    '
    datRpt.Update
    
    ResetButtons True
    Exit Sub
SaveErr:
MsgBox (Err.Description), vbInformation

End Sub

Private Sub Form_Load()
    Dim txt As TextBox
    Set datRpt = New Recordset
    datRpt.Open "SELECT * FROM hr_Reports", cn, adOpenStatic, adLockOptimistic
    
    For Each txt In Me.txtRpt
        Set txt.DataSource = datRpt
    Next
    Set chkEmp.DataSource = datRpt
    Set chkWage.DataSource = datRpt
    Set chkDate.DataSource = datRpt
    ResetButtons True
    If datRpt.RecordCount = 0 Then
       cmdModify.Enabled = False
    End If
End Sub

Public Sub ResetButtons(bval As Boolean)
    cmdBrowse.Enabled = bval
    cmdFind.Enabled = Not bval
    cmdSave.Enabled = Not bval
    cmdModify.Enabled = bval
    cmdCancel.Enabled = Not bval
    cmdBrowse.Enabled = Not bval
    cmdFind.Enabled = bval
    txtRpt(1).Locked = bval
    txtRpt(3).Locked = bval
    chkEmp.Enabled = Not bval
    chkWage.Enabled = Not bval
    chkDate.Enabled = Not bval
    
End Sub

Private Sub txtrpt_GotFocus(Index As Integer)
    If Index = 1 Then
       txtRpt(1) = Trim(txtRpt(1))
    End If
End Sub
