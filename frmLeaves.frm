VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmLeaves 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Leave Types"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLeaves.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   3810
      Picture         =   "frmLeaves.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Next Record"
      Top             =   2340
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
      Left            =   3330
      Picture         =   "frmLeaves.frx":058C
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Previous Record"
      Top             =   2340
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
      Left            =   2865
      Picture         =   "frmLeaves.frx":06D6
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "First Record"
      Top             =   2340
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
      Left            =   4290
      Picture         =   "frmLeaves.frx":0820
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Last Record"
      Top             =   2340
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
      ScaleWidth      =   4695
      TabIndex        =   1
      Top             =   2850
      Width           =   4695
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   345
         Left            =   2760
         TabIndex        =   11
         Top             =   0
         Width           =   900
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   345
         Left            =   1830
         TabIndex        =   10
         Top             =   0
         Width           =   900
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   345
         Left            =   915
         TabIndex        =   9
         Top             =   0
         Width           =   900
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "&New"
         Height          =   345
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   900
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
         Height          =   345
         Left            =   3675
         TabIndex        =   2
         Top             =   0
         Width           =   900
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2235
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   4650
      _ExtentX        =   8202
      _ExtentY        =   3942
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Leave Types"
      TabPicture(0)   =   "frmLeaves.frx":096A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "chkLeave"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtLeave(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtLeave(1)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      Begin VB.TextBox txtLeave 
         BackColor       =   &H00FFFFFF&
         DataField       =   "leave_name"
         Height          =   330
         Index           =   1
         Left            =   1410
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   7
         Top             =   1110
         Width           =   3150
      End
      Begin VB.TextBox txtLeave 
         DataField       =   "leave_code"
         Height          =   330
         Index           =   0
         Left            =   1425
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   600
         Width           =   1965
      End
      Begin VB.CheckBox chkLeave 
         Caption         =   "Carry Balance Forward at end of Year"
         DataField       =   "forward"
         Height          =   330
         Left            =   135
         TabIndex        =   5
         Top             =   1710
         Width           =   4395
      End
      Begin VB.Label Label2 
         Caption         =   "Leave Name"
         Height          =   240
         Left            =   135
         TabIndex        =   4
         Top             =   1155
         Width           =   1080
      End
      Begin VB.Label Label1 
         Caption         =   "Leave Code"
         Height          =   315
         Left            =   165
         TabIndex        =   3
         Top             =   630
         Width           =   1125
      End
   End
End
Attribute VB_Name = "frmLeaves"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim datLeave As Recordset

Private Sub cmdCancel_Click()
    datLeave.CancelUpdate
    Reset True
End Sub

Private Sub cmdEdit_Click()
    If datLeave.RecordCount = 0 Then
       MsgBox ("No Record to Edit"), vbInformation, "Leave Types"
       Exit Sub
    End If
    Reset False
End Sub

Private Sub cmdNew_Click()
    'Calculate a new code for the leave type to be added
    Set datExtra = New Recordset
    datExtra.Open "SELECT TOP 1 * FROM leave_mast ORDER BY leave_code DESC", cn, adOpenStatic, adLockOptimistic
    If datExtra.RecordCount <> 0 Then
       x = datExtra!Leave_code + 1
    Else
        x = 1
    End If
    datLeave.AddNew
    chkLeave.Value = 1
    txtleave(0) = x
    Reset False
End Sub

Private Sub cmdUpdate_Click()
    datLeave!co_code = ID
    datLeave.Update
    If datLeave.Bookmark > 0 Then
        bk = datLeave.Bookmark
        datLeave.Bookmark = bk
    End If
    Reset True
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
Public Sub FormatGrid()
    'Format the grid and give the desired dimensions
    With grdLeave
        .Columns(0).Locked = True
        .Columns(0).Width = 900
        .Columns(1).Width = 3020
        .Width = Me.Width - 30
    End With
End Sub
Private Sub cmdClose_Click()
    Unload Me
    
End Sub

Private Sub Form_Load()
    Dim txt As TextBox
    Set datLeave = New Recordset
    datLeave.Open "SELECT * FROM leave_mast", cn, adOpenStatic, adLockOptimistic
    For Each txt In Me.txtleave
        Set txt.DataSource = datLeave
    Next
    Set chkLeave.DataSource = datLeave
    Reset True
End Sub

Private Sub grdLeave_KeyPress(KeyAscii As Integer)
    Char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(Char))
End Sub

Private Sub grdLeave_OnAddNew()
    'Calculate a new code for the leave type to be added
    Set datExtra = New Recordset
    datExtra.Open "SELECT TOP 1 * FROM leave_mast ORDER BY leave_code DESC", cn, adOpenStatic, adLockOptimistic
    If datExtra.RecordCount <> 0 Then
       grdLeave.Columns(0) = datExtra!Leave_code + 1
    Else
        grdLeave.Columns(0) = 1
    End If
End Sub

Public Sub Reset(bval As Boolean)
    cmdNew.Enabled = bval
    cmdEdit.Enabled = bval
    cmdUpdate.Enabled = Not bval
    cmdCancel.Enabled = Not bval
    txtleave(1).Locked = bval
    chkLeave.Enabled = Not bval
    
    ffirst.Enabled = bval
    flast.Enabled = bval
    fnext.Enabled = bval
    fprev.Enabled = bval
    
End Sub

Private Sub txtLeave_GotFocus(Index As Integer)
    txtleave(Index) = Trim(txtleave(Index))
End Sub
