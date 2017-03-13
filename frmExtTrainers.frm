VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmExtTrainers 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "External Trainers"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6030
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmExtTrainers.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   6030
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
      Left            =   5145
      Picture         =   "frmExtTrainers.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Next Record"
      Top             =   5685
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
      Left            =   4665
      Picture         =   "frmExtTrainers.frx":058C
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Previous Record"
      Top             =   5685
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
      Left            =   4200
      Picture         =   "frmExtTrainers.frx":06D6
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "First Record"
      Top             =   5685
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
      Left            =   5625
      Picture         =   "frmExtTrainers.frx":0820
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Last Record"
      Top             =   5685
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   390
      Left            =   0
      ScaleHeight     =   390
      ScaleWidth      =   6030
      TabIndex        =   14
      Top             =   6120
      Width           =   6030
      Begin VB.CommandButton cmdClose 
         Caption         =   "C&lose"
         Height          =   360
         Left            =   5010
         TabIndex        =   19
         Top             =   0
         Width           =   960
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   360
         Left            =   2925
         TabIndex        =   18
         Top             =   0
         Width           =   960
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   360
         Left            =   1950
         TabIndex        =   17
         Top             =   0
         Width           =   960
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   360
         Left            =   975
         TabIndex        =   16
         Top             =   0
         Width           =   960
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "&New"
         Height          =   360
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   960
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5715
      Left            =   0
      TabIndex        =   0
      Top             =   -90
      Width           =   6000
      Begin VB.TextBox txtTrn 
         DataField       =   "trn_code"
         Height          =   330
         Index           =   6
         Left            =   1425
         TabIndex        =   23
         Top             =   270
         Width           =   1485
      End
      Begin VB.TextBox txtTrn 
         DataField       =   "notes"
         Height          =   930
         Index           =   5
         Left            =   105
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   20
         Top             =   4725
         Width           =   5835
      End
      Begin VB.CommandButton cmdTrTypes 
         Caption         =   "..."
         Height          =   300
         Left            =   5355
         TabIndex        =   13
         ToolTipText     =   "Add a new Training Type"
         Top             =   3990
         Width           =   300
      End
      Begin MSDataListLib.DataCombo cboTrn 
         DataField       =   "t_code"
         Height          =   345
         Left            =   150
         TabIndex        =   12
         Top             =   3975
         Width           =   5190
         _ExtentX        =   9155
         _ExtentY        =   609
         _Version        =   393216
         Style           =   2
         ListField       =   "t_name"
         BoundColumn     =   "t_code"
         Text            =   ""
      End
      Begin VB.TextBox txtTrn 
         DataField       =   "phone2"
         Height          =   330
         Index           =   3
         Left            =   2970
         TabIndex        =   11
         Top             =   2400
         Width           =   2805
      End
      Begin VB.TextBox txtTrn 
         DataField       =   "phone1"
         Height          =   330
         Index           =   2
         Left            =   2970
         TabIndex        =   10
         Top             =   1725
         Width           =   2820
      End
      Begin VB.TextBox txtTrn 
         DataField       =   "email"
         Height          =   330
         Index           =   4
         Left            =   120
         TabIndex        =   9
         Top             =   3240
         Width           =   5670
      End
      Begin VB.TextBox txtTrn 
         DataField       =   "address"
         Height          =   1185
         Index           =   1
         Left            =   135
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   1725
         Width           =   2475
      End
      Begin VB.TextBox txtTrn 
         DataField       =   "trn_name"
         Height          =   330
         Index           =   0
         Left            =   135
         TabIndex        =   7
         Top             =   945
         Width           =   5670
      End
      Begin VB.Label Label8 
         Caption         =   "Trainers Code"
         Height          =   225
         Left            =   135
         TabIndex        =   22
         Top             =   300
         Width           =   1275
      End
      Begin VB.Label Label7 
         Caption         =   "Notes"
         Height          =   270
         Left            =   165
         TabIndex        =   21
         Top             =   4455
         Width           =   1545
      End
      Begin VB.Label Label6 
         Caption         =   "Training Offered"
         Height          =   225
         Left            =   135
         TabIndex        =   6
         Top             =   3720
         Width           =   1650
      End
      Begin VB.Label Label5 
         Caption         =   "E-mail"
         Height          =   255
         Left            =   135
         TabIndex        =   5
         Top             =   3015
         Width           =   1185
      End
      Begin VB.Label Label4 
         Caption         =   "Phone 2"
         Height          =   255
         Left            =   2985
         TabIndex        =   4
         Top             =   2145
         Width           =   825
      End
      Begin VB.Label Label3 
         Caption         =   "Phone 1"
         Height          =   300
         Left            =   3000
         TabIndex        =   3
         Top             =   1470
         Width           =   840
      End
      Begin VB.Label Label2 
         Caption         =   "Address"
         Height          =   285
         Left            =   135
         TabIndex        =   2
         Top             =   1455
         Width           =   1020
      End
      Begin VB.Label Label1 
         Caption         =   "External Trainer"
         Height          =   255
         Left            =   135
         TabIndex        =   1
         Top             =   705
         Width           =   1410
      End
   End
End
Attribute VB_Name = "frmExtTrainers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim datExtTrn As Recordset
Dim dattrainings As Recordset

Private Sub cmdCancel_Click()
    datExtTrn.CancelUpdate
    reset True
End Sub

Private Sub cmdEdit_Click()
    reset False
End Sub

Private Sub cmdNew_Click()
    'Calculate a new code for the leave type to be added
    Set datExtra = New Recordset
    datExtra.Open "SELECT TOP 1 * FROM Ext_Trainers ORDER BY trn_code DESC", cn, adOpenStatic, adLockOptimistic
    If datExtra.RecordCount <> 0 Then
       x = datExtra!trn_code + 1
    Else
        x = 1
    End If
    datExtTrn.AddNew
    txtTrn(6) = x
    reset False
End Sub

Private Sub cmdTrTypes_Click()
    frmTrainings.Show vbModal
    dattrainings.Requery
    cboTrn.ReFill
End Sub

Private Sub cmdUpdate_Click()
    'Validate the entries
    If Trim(txtTrn(0)) = "" Then
       MsgBox ("Must give the Trainers Name."), vbInformation, "External Trainer"
       Exit Sub
    ElseIf cboTrn.BoundText = "" Then
       MsgBox ("Must select the training offered."), vbInformation, "External Trainer"
       Exit Sub
    End If
    datExtTrn.Update
    If datExtTrn.Bookmark > 0 Then
        bk = datExtTrn.Bookmark
        datExtTrn.Bookmark = bk
    End If
    reset True
End Sub

Private Sub ffirst_Click()
On Error GoTo PrevErr
    With datExtTrn
         .MoveFirst
    End With
    Exit Sub
PrevErr:
MsgBox (Err.Description), vbInformation, "External Trainers"
Err.Clear
End Sub

Private Sub flast_Click()
On Error GoTo PrevErr
    With datExtTrn
         .MoveLast
    End With
    Exit Sub
PrevErr:
MsgBox (Err.Description), vbInformation, "External Trainers"
Err.Clear
End Sub

Private Sub fnext_Click()
On Error GoTo PrevErr
    With datExtTrn
         .MoveNext
         If .EOF Then
            .MoveLast
            MsgBox ("Last Record encountered"), vbInformation, "External Trainers"
         End If
    End With
    Exit Sub
PrevErr:
MsgBox (Err.Description), vbInformation, "External Trainers"
Err.Clear
End Sub
Private Sub fprev_Click()
On Error GoTo PrevErr
    With datExtTrn
         .MovePrevious
         If .BOF Then
            .MoveFirst
            MsgBox ("First Record encountered"), vbInformation, "External Trainers"
         End If
    End With
    Exit Sub
PrevErr:
MsgBox (Err.Description), vbInformation, "External Trainers"
Err.Clear
End Sub
Private Sub cmdClose_Click()
    Unload Me
    
End Sub

Private Sub Form_Load()
    Dim txt As TextBox
    Set datExtTrn = New Recordset
    datExtTrn.Open "SELECT * FROM ext_Trainers", cn, adOpenStatic, adLockOptimistic
    For Each txt In Me.txtTrn
        Set txt.DataSource = datExtTrn
    Next
    Set cboTrn.DataSource = datExtTrn
    
    Set dattrainings = New Recordset
    dattrainings.Open "SELECT * FROM Trainings", cn, adOpenStatic, adLockOptimistic
    Set cboTrn.RowSource = dattrainings
    
    reset True
End Sub

Private Sub grdLeave_KeyPress(KeyAscii As Integer)
    Char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(Char))
End Sub

Public Sub reset(bval As Boolean)
    Dim txt As TextBox
    cmdNew.Enabled = bval
    cmdEdit.Enabled = bval
    cmdUpdate.Enabled = Not bval
    cmdCancel.Enabled = Not bval
    ffirst.Enabled = bval
    flast.Enabled = bval
    fnext.Enabled = bval
    fprev.Enabled = bval
    For Each txt In Me.txtTrn
        txt.Locked = bval
    Next
    cboTrn.Enabled = Not bval
    
    txtTrn(6).Locked = True
    
End Sub

Private Sub txtTrn_GotFocus(Index As Integer)
    txtTrn(Index) = Trim(txtTrn(Index))
End Sub
