VERSION 5.00
Begin VB.Form frmAppParams 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Appraisal Parameters"
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5115
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAppParams.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   5115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkParams 
      Caption         =   "Parameter is Valid for Use"
      DataField       =   "Valid"
      Height          =   225
      Left            =   30
      TabIndex        =   20
      Top             =   3270
      Width           =   2910
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
      Left            =   4725
      Picture         =   "frmAppParams.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Last Record"
      Top             =   3210
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
      Left            =   3300
      Picture         =   "frmAppParams.frx":058C
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "First Record"
      Top             =   3210
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
      Left            =   3765
      Picture         =   "frmAppParams.frx":06D6
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Previous Record"
      Top             =   3210
      Width           =   375
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
      Left            =   4245
      Picture         =   "frmAppParams.frx":0820
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Next Record"
      Top             =   3210
      Width           =   375
   End
   Begin VB.TextBox txtparam 
      DataField       =   "param_code"
      Height          =   330
      Index           =   4
      Left            =   2895
      TabIndex        =   15
      Top             =   15
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   5115
      TabIndex        =   9
      Top             =   3630
      Width           =   5115
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   350
         Left            =   4200
         TabIndex        =   14
         Top             =   0
         Width           =   855
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "C&ancel"
         Height          =   350
         Left            =   2610
         TabIndex        =   13
         Top             =   0
         Width           =   855
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   350
         Left            =   1740
         TabIndex        =   12
         Top             =   0
         Width           =   855
      End
      Begin VB.CommandButton cmdModify 
         Caption         =   "&Modify"
         Height          =   350
         Left            =   870
         TabIndex        =   11
         Top             =   0
         Width           =   855
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "&New"
         Height          =   350
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   855
      End
   End
   Begin VB.TextBox txtparam 
      DataField       =   "notes"
      Height          =   1215
      Index           =   3
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   1920
      Width           =   5115
   End
   Begin VB.Frame Frame1 
      Caption         =   "Scoring"
      Height          =   855
      Left            =   0
      TabIndex        =   3
      Top             =   720
      Width           =   4695
      Begin VB.TextBox txtparam 
         DataField       =   "hscore"
         Height          =   285
         Index           =   2
         Left            =   2520
         TabIndex        =   7
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox txtparam 
         DataField       =   "lscore"
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Highest Points"
         Height          =   255
         Left            =   2520
         TabIndex        =   5
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Lowest Points"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.TextBox txtparam 
      DataField       =   "param_name"
      Height          =   330
      Index           =   0
      Left            =   15
      TabIndex        =   2
      Top             =   345
      Width           =   4680
   End
   Begin VB.Label Label4 
      Caption         =   "Notes"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Parameter Name"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmAppParams"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim datParam As Recordset

Private Sub cmdClose_Click()
    Unload Me
    
End Sub

Private Sub ffirst_Click()
On Error GoTo PrevErr
    With datParam
         .MoveFirst
    End With
    Exit Sub
PrevErr:
MsgBox (Err.Description), vbInformation
Err.Clear
End Sub

Private Sub flast_Click()
On Error GoTo PrevErr
    With datParam
         .MoveLast
    End With
    Exit Sub
PrevErr:
MsgBox (Err.Description), vbInformation
Err.Clear
End Sub
Private Sub fnext_Click()
On Error GoTo PrevErr
    With datParam
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
    With datParam
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
    datParam.CancelUpdate
    If datParam.RecordCount <> 0 Then
       datParam.MoveLast
    Else
       datParam.Requery
    End If
    
    reset True
End Sub

Private Sub cmdModify_Click()
    reset False
    
End Sub

Private Sub cmdNew_Click()
    Dim Ncount As Long
    'Calculate the new param_code
    Set datExtra = New Recordset
    datExtra.Open "SELECT * FROM appparams WHERE co_code ='" & ID & _
    "' ORDER BY param_code", cn, adOpenStatic, adLockOptimistic
    
    If datExtra.RecordCount = 0 Then
       Ncount = 1
    Else
        datExtra.MoveLast
        Ncount = datExtra!param_code + 1
    End If
    datParam.AddNew
    txtparam(4) = Ncount
    reset False
End Sub

Private Sub cmdUpdate_Click()
    'Validate the entries made
    If txtparam(0) = "" Or txtparam(1) = "" Or txtparam(2) = "" Or txtparam(3) = "" Then
       MsgBox ("Insufficient data given to save record."), vbInformation
       Exit Sub
    Else
        datParam!co_code = ID
        datParam.Update
        bk = datParam.Bookmark
        datParam.Bookmark = bk
    End If
    reset True
End Sub

Private Sub Form_Load()
    Dim txt As TextBox
    Set datParam = New Recordset
    datParam.Open "SELECT * FROM appparams WHERE co_code ='" & ID & _
    "' ORDER BY param_code", cn, adOpenStatic, adLockOptimistic
    
    For Each txt In Me.txtparam
        Set txt.DataSource = datParam
    Next
    Set Me.chkParams.DataSource = datParam
    reset True
    
End Sub

Public Sub reset(bval As Boolean)
    Dim txt As TextBox
    For Each txt In Me.txtparam
        txt.Locked = bval
    Next
    
    cmdNew.Enabled = bval
    cmdModify.Enabled = bval
    cmdUpdate.Enabled = Not bval
    cmdCancel.Enabled = Not bval
    chkParams.Enabled = Not bval
    ffirst.Enabled = bval
    fprev.Enabled = bval
    fnext.Enabled = bval
    flast.Enabled = bval
End Sub

Private Sub txtparam_GotFocus(Index As Integer)
    txtparam(Index) = Trim(txtparam(Index))
    
End Sub
