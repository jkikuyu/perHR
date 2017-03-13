VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmmedschemes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Medical Schemes"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4740
   Icon            =   "frmmedschemes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   4740
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
      Left            =   4305
      Picture         =   "frmmedschemes.frx":0742
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Last Record"
      Top             =   1635
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
      Left            =   2880
      Picture         =   "frmmedschemes.frx":088C
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "First Record"
      Top             =   1635
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
      Left            =   3345
      Picture         =   "frmmedschemes.frx":09D6
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Previous Record"
      Top             =   1635
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
      Left            =   3825
      Picture         =   "frmmedschemes.frx":0B20
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Next Record"
      Top             =   1635
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   4740
      TabIndex        =   5
      Top             =   2130
      Width           =   4740
      Begin VB.CommandButton cmdClose 
         Caption         =   "C&lose"
         Height          =   345
         Left            =   3675
         TabIndex        =   10
         Top             =   0
         Width           =   900
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "&New"
         Height          =   345
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   900
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   345
         Left            =   915
         TabIndex        =   8
         Top             =   0
         Width           =   900
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   345
         Left            =   1830
         TabIndex        =   7
         Top             =   0
         Width           =   900
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   345
         Left            =   2760
         TabIndex        =   6
         Top             =   0
         Width           =   900
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   1545
      Left            =   30
      TabIndex        =   0
      Top             =   15
      Width           =   4710
      _ExtentX        =   8308
      _ExtentY        =   2725
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Medical Scheme Types"
      TabPicture(0)   =   "frmmedschemes.frx":0C6A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtScheme(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtScheme(1)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      Begin VB.TextBox txtScheme 
         DataField       =   "scheme_name"
         Height          =   315
         Index           =   1
         Left            =   1305
         TabIndex        =   4
         Top             =   1080
         Width           =   3300
      End
      Begin VB.TextBox txtScheme 
         DataField       =   "scheme_id"
         Height          =   285
         Index           =   0
         Left            =   1290
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   615
         Width           =   1890
      End
      Begin VB.Label Label2 
         Caption         =   "Scheme Name"
         Height          =   270
         Left            =   105
         TabIndex        =   2
         Top             =   1110
         Width           =   1140
      End
      Begin VB.Label Label1 
         Caption         =   "Scheme Code"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   645
         Width           =   1080
      End
   End
End
Attribute VB_Name = "frmmedschemes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim datSchemes As Recordset

Private Sub ffirst_Click()
On Error GoTo PrevErr
    With datSchemes
         .MoveFirst
    End With
    Exit Sub
PrevErr:
MsgBox (Err.Description), vbInformation, "Medical Schemes"
Err.Clear
End Sub

Private Sub flast_Click()
On Error GoTo PrevErr
    With datSchemes
         .MoveLast
    End With
    Exit Sub
PrevErr:
MsgBox (Err.Description), vbInformation, "Medical Schemes"
Err.Clear
End Sub

Private Sub fnext_Click()
On Error GoTo PrevErr
    With datSchemes
         .MoveNext
         If .EOF Then
            .MoveLast
            MsgBox ("Last Record encountered"), vbInformation, "Medical Schemes"
         End If
    End With
    Exit Sub
PrevErr:
MsgBox (Err.Description), vbInformation, "Medical Schemes"
Err.Clear
End Sub
Private Sub fprev_Click()
On Error GoTo PrevErr
    With datSchemes
         .MovePrevious
         If .BOF Then
            .MoveFirst
            MsgBox ("First Record encountered"), vbInformation, "Medical Schemes"
         End If
    End With
    Exit Sub
PrevErr:
MsgBox (Err.Description), vbInformation, "Medical Schemes"
Err.Clear
End Sub
Public Sub Reset(bval As Boolean)
    cmdNew.Enabled = bval
    cmdEdit.Enabled = bval
    cmdUpdate.Enabled = Not bval
    cmdCancel.Enabled = Not bval
    txtScheme(1).Locked = bval
    
    ffirst.Enabled = bval
    flast.Enabled = bval
    fnext.Enabled = bval
    fprev.Enabled = bval
    
End Sub
Public Sub FormatGrid()
    'Format the grid and give the desired dimensions
    With grdsch
        .Columns(0).Locked = True
        .Columns(0).Width = 900
        .Columns(1).Width = 4350
        .Width = Me.Width - 60
    End With
End Sub

Private Sub cmdCancel_Click()
    datSchemes.CancelUpdate
    Reset True
End Sub

Private Sub cmdClose_Click()
    Unload Me
    
End Sub

Private Sub cmdEdit_Click()
    If datSchemes.RecordCount = 0 Then
       MsgBox ("No Record to Edit"), vbInformation, "Medical Scheme Types"
       Exit Sub
    End If
    Reset False

End Sub

Private Sub cmdNew_Click()
    'Calculate a new code for the scheme type to be added
    Set datExtra = New Recordset
    datExtra.Open "SELECT TOP 1 * FROM med_schemes ORDER BY scheme_id DESC", cn, adOpenStatic, adLockOptimistic
    If datExtra.RecordCount <> 0 Then
       x = datExtra!scheme_id + 1
    Else
        x = 1
    End If
    datSchemes.AddNew
    txtScheme(0) = x
    Reset False
End Sub

Private Sub cmdUpdate_Click()
    If txtScheme(1) = "" Then
       MsgBox ("You must give the scheme name."), vbInformation, "Medical Schemes"
       Exit Sub
    End If
    datSchemes.Update
    If datSchemes.Bookmark > 0 Then
        bk = datSchemes.Bookmark
        datSchemes.Bookmark = bk
    End If
    Reset True
End Sub

Private Sub Form_Load()
    Set datSchemes = New Recordset
    datSchemes.Open "SELECT * FROM Med_schemes", cn, adOpenStatic, adLockOptimistic
    
    
    Dim txt As TextBox
    For Each txt In Me.txtScheme
        Set txt.DataSource = datSchemes
    Next
    Reset True
End Sub

Private Sub grdsch_KeyPress(KeyAscii As Integer)
    Char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(Char))
End Sub

Private Sub grdsch_OnAddNew()
    'Calculate a new code for the leave type to be added
    Set datExtra = New Recordset
    datExtra.Open "SELECT TOP 1 * FROM med_schemes ORDER BY scheme_id DESC", cn, adOpenStatic, adLockOptimistic
    If datExtra.RecordCount <> 0 Then
       grdsch.Columns(0) = datExtra!scheme_id + 1
    Else
        grdsch.Columns(0) = 1
    End If
End Sub
