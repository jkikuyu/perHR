VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmTrainings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Training Types"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4950
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTrainings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   4560
      Picture         =   "frmTrainings.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Last Record"
      Top             =   1650
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
      Left            =   3135
      Picture         =   "frmTrainings.frx":058C
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "First Record"
      Top             =   1650
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
      Left            =   3600
      Picture         =   "frmTrainings.frx":06D6
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Previous Record"
      Top             =   1650
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
      Left            =   4080
      Picture         =   "frmTrainings.frx":0820
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Next Record"
      Top             =   1650
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   390
      Left            =   0
      ScaleHeight     =   390
      ScaleWidth      =   4950
      TabIndex        =   5
      Top             =   2085
      Width           =   4950
      Begin VB.CommandButton cmdNew 
         Caption         =   "&New"
         Height          =   360
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   960
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   360
         Left            =   975
         TabIndex        =   9
         Top             =   0
         Width           =   960
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   360
         Left            =   1950
         TabIndex        =   8
         Top             =   0
         Width           =   960
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   360
         Left            =   2925
         TabIndex        =   7
         Top             =   0
         Width           =   960
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "C&lose"
         Height          =   360
         Left            =   3915
         TabIndex        =   6
         Top             =   0
         Width           =   960
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   1575
      Left            =   15
      TabIndex        =   0
      Top             =   30
      Width           =   4920
      _ExtentX        =   8678
      _ExtentY        =   2778
      _Version        =   393216
      Style           =   1
      Tabs            =   1
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
      TabCaption(0)   =   "Training Types"
      TabPicture(0)   =   "frmTrainings.frx":096A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtTr(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtTr(1)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      Begin VB.TextBox txtTr 
         DataField       =   "t_name"
         Height          =   330
         Index           =   1
         Left            =   45
         MaxLength       =   50
         TabIndex        =   4
         Top             =   1170
         Width           =   4830
      End
      Begin VB.TextBox txtTr 
         DataField       =   "t_code"
         Height          =   330
         Index           =   0
         Left            =   1485
         TabIndex        =   3
         Top             =   495
         Width           =   1905
      End
      Begin VB.Label Label2 
         Caption         =   "Training Name"
         Height          =   300
         Left            =   60
         TabIndex        =   2
         Top             =   855
         Width           =   1260
      End
      Begin VB.Label Label1 
         Caption         =   "Training Code"
         Height          =   300
         Left            =   60
         TabIndex        =   1
         Top             =   510
         Width           =   1260
      End
   End
End
Attribute VB_Name = "frmTrainings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim datTraining As Recordset

Private Sub cmdCancel_Click()
    datTraining.CancelUpdate
    reset True
End Sub

Private Sub cmdEdit_Click()
    reset False
End Sub

Private Sub cmdNew_Click()
    'Calculate a new code for the leave type to be added
    Set datExtra = New Recordset
    datExtra.Open "SELECT TOP 1 * FROM trainings ORDER BY t_code DESC", cn, adOpenStatic, adLockOptimistic
    If datExtra.RecordCount <> 0 Then
       x = datExtra!t_code + 1
    Else
        x = 1
    End If
    datTraining.AddNew
    txtTr(0) = x
    reset False
End Sub

Private Sub cmdUpdate_Click()
    datTraining.Update
    If datTraining.Bookmark > 0 Then
        bk = datTraining.Bookmark
        datTraining.Bookmark = bk
    End If
    reset True
End Sub

Private Sub ffirst_Click()
On Error GoTo PrevErr
    With datTraining
         .MoveFirst
    End With
    Exit Sub
PrevErr:
MsgBox (Err.Description), vbInformation
Err.Clear
End Sub

Private Sub flast_Click()
On Error GoTo PrevErr
    With datTraining
         .MoveLast
    End With
    Exit Sub
PrevErr:
MsgBox (Err.Description), vbInformation
Err.Clear
End Sub

Private Sub fnext_Click()
On Error GoTo PrevErr
    With datTraining
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
    With datTraining
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
Private Sub cmdClose_Click()
    Unload Me
    
End Sub

Private Sub Form_Load()
    Dim txt As TextBox
    Set datTraining = New Recordset
    datTraining.Open "SELECT * FROM trainings", cn, adOpenStatic, adLockOptimistic
    For Each txt In Me.txtTr
        Set txt.DataSource = datTraining
    Next
    reset True
End Sub

Private Sub grdLeave_KeyPress(KeyAscii As Integer)
    Char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(Char))
End Sub

Public Sub reset(bval As Boolean)
    cmdNew.Enabled = bval
    cmdEdit.Enabled = bval
    cmdUpdate.Enabled = Not bval
    cmdCancel.Enabled = Not bval
    txtTr(0).Locked = bval
    txtTr(1).Locked = bval
    
End Sub

Private Sub txtTr_GotFocus(Index As Integer)
    txtTr(Index) = Trim(txtTr(Index))
End Sub
