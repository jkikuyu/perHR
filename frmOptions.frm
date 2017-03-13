VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   3195
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   6495
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   0
      ScaleHeight     =   360
      ScaleWidth      =   6495
      TabIndex        =   7
      Top             =   2835
      Width           =   6495
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   345
         Left            =   5355
         TabIndex        =   11
         Top             =   15
         Width           =   1095
      End
      Begin VB.CommandButton cmdModify 
         Caption         =   "&Modify"
         Height          =   345
         Left            =   1785
         TabIndex        =   10
         Top             =   15
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "C&ancel"
         Height          =   345
         Left            =   2970
         TabIndex        =   9
         Top             =   15
         Width           =   1095
      End
      Begin VB.CommandButton cmdApply 
         Caption         =   "Apply"
         Height          =   345
         Left            =   4170
         TabIndex        =   8
         Top             =   15
         Width           =   1095
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2715
      Left            =   15
      TabIndex        =   6
      Top             =   15
      Width           =   6465
      _ExtentX        =   11404
      _ExtentY        =   4789
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Employee Data Options"
      TabPicture(0)   =   "frmOptions.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label5"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtOpt(2)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtOpt(3)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Links to Payroll"
      TabPicture(1)   =   "frmOptions.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label6"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cboOpt"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "txtOpt(1)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "txtOpt(0)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Absenteeism-Leave Link"
      TabPicture(2)   =   "frmOptions.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label4"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label7"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "cboLeave"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
      Begin VB.TextBox txtOpt 
         DataField       =   "link_wage"
         Height          =   330
         Index           =   0
         Left            =   -69435
         TabIndex        =   17
         Top             =   1425
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.TextBox txtOpt 
         Height          =   330
         Index           =   1
         Left            =   -74850
         TabIndex        =   16
         Top             =   1680
         Width           =   4350
      End
      Begin VB.TextBox txtOpt 
         DataField       =   "sf_prefix"
         Height          =   330
         Index           =   3
         Left            =   1935
         MaxLength       =   15
         TabIndex        =   15
         Top             =   1650
         Width           =   3510
      End
      Begin VB.TextBox txtOpt 
         DataField       =   "sc_prefix"
         Height          =   330
         Index           =   2
         Left            =   1935
         MaxLength       =   1
         TabIndex        =   14
         Top             =   1155
         Width           =   915
      End
      Begin MSDataListLib.DataCombo cboOpt 
         Height          =   330
         Left            =   -74805
         TabIndex        =   19
         Top             =   1680
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   582
         _Version        =   393216
         ListField       =   "wage_name"
         BoundColumn     =   "wage_code"
         Text            =   "DataCombo1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo cboLeave 
         DataField       =   "leave"
         Height          =   345
         Left            =   -74880
         TabIndex        =   20
         Top             =   1380
         Width           =   4710
         _ExtentX        =   8308
         _ExtentY        =   609
         _Version        =   393216
         Style           =   2
         ListField       =   "leave_name"
         BoundColumn     =   "leave_code"
         Text            =   "DataCombo1"
      End
      Begin VB.Label Label7 
         Caption         =   $"frmOptions.frx":0060
         Height          =   600
         Left            =   -74865
         TabIndex        =   24
         Top             =   405
         Width           =   6255
      End
      Begin VB.Label Label6 
         Caption         =   $"frmOptions.frx":00E7
         Height          =   780
         Left            =   -74865
         TabIndex        =   23
         Top             =   420
         Width           =   6225
      End
      Begin VB.Label Label5 
         Caption         =   "This Option determines the format of the data that is captured when registering an employee"
         Height          =   555
         Left            =   150
         TabIndex        =   22
         Top             =   510
         Width           =   6060
      End
      Begin VB.Label Label4 
         Caption         =   "Absenteeism Records Should Post Leave To:"
         Height          =   315
         Left            =   -74865
         TabIndex        =   21
         Top             =   1140
         Width           =   4620
      End
      Begin VB.Label Label1 
         Caption         =   "Medical Recoveries Should be Posted to:"
         Height          =   285
         Left            =   -74805
         TabIndex        =   18
         Top             =   1380
         Width           =   3540
      End
      Begin VB.Label Label3 
         Caption         =   "Staff File Prefix"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   1665
         Width           =   1515
      End
      Begin VB.Label Label2 
         Caption         =   "Staff Code Prefix"
         Height          =   285
         Left            =   270
         TabIndex        =   12
         Top             =   1170
         Width           =   1530
      End
   End
   Begin VB.PictureBox picOptions 
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
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1785
         Left            =   2100
         TabIndex        =   5
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
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
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1785
         Left            =   1545
         TabIndex        =   4
         Top             =   675
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
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
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Sample 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1785
         Left            =   645
         TabIndex        =   3
         Top             =   300
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bk As Variant
Dim datOpt As Recordset
Dim datLeave As Recordset

Private Sub cboOpt_Click(Area As Integer)
    If Area = 2 Then
       txtOpt(0) = Trim(cboOpt.BoundText)
    End If
End Sub

Private Sub cmdApply_Click()
    datOpt!co_code = ID
    datOpt.Update
    bk = datOpt.Bookmark
    datOpt.Bookmark = bk
    Reset False
    
End Sub

Private Sub cmdCancel_Click()
    Reset False
    
End Sub

Private Sub cmdOK_Click()
    'MsgBox "Place code here to set options and close dialog!"
    Unload Me
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdModify_Click()
    'cboOpt.Enabled = True
    'cmdCancel.Enabled = True
    'cmdApply.Enabled = True
    'cmdModify.Enabled = False
    If datOpt.RecordCount = 0 Then
       datOpt.AddNew
    End If
    Reset True
End Sub
Private Sub Reset(bval As Boolean)
    Dim txt As TextBox
    cboOpt.Enabled = bval
    
    For Each txt In Me.txtOpt
        txt.Locked = Not bval
    Next
    
    cmdCancel.Enabled = bval
    cmdModify.Enabled = Not bval
    cmdApply.Enabled = bval
    
End Sub

Private Sub Form_Load()
    'center the form
    'Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    Set datLeave = New Recordset
    datLeave.Open "SELECT * FROM Leave_mast WHERE co_code ='" & ID & _
    "'", cn, adOpenStatic, adLockOptimistic
    Set cboleave.RowSource = datLeave
    
    
    Set datExtra = New Recordset
    datExtra.Open "SELECT * FROM Wage_types WHERE system_code = 0" & _
    " AND loan = 0 AND slip_display =0", cn, adOpenStatic, adLockOptimistic
    Set cboOpt.DataSource = datExtra
    Set cboOpt.RowSource = datExtra
    
    Set datOpt = New Recordset
    datOpt.Open "SELECT * FROM med_opts WHERE co_code ='" & ID & _
    "'", cn, adOpenStatic, adLockOptimistic
    Dim txt As TextBox
    For Each txt In Me.txtOpt
        Set txt.DataSource = datOpt
    Next
    Set cboleave.DataSource = datOpt
    Reset False
    
End Sub

Private Sub txtOpt_Change(Index As Integer)
    If Index = 0 Then
       Set datExtra = New Recordset
       datExtra.Open "SELECT * FROM Wage_types WHERE wage_code ='" & _
       Trim(txtOpt(0)) & "' AND slip_display = 0", cn, adOpenStatic, adLockOptimistic
       If datExtra.RecordCount <> 0 Then
          txtOpt(1) = Trim(datExtra!wage_name)
       Else
          txtOpt(1) = ""
       End If
    End If
End Sub

Private Sub txtOpt_GotFocus(Index As Integer)
    txtOpt(Index) = Trim(txtOpt(Index))
End Sub

Private Sub txtOpt_KeyPress(Index As Integer, KeyAscii As Integer)
        Char = Chr(KeyAscii)
        KeyAscii = Asc(UCase(Char))
End Sub
