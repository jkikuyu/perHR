VERSION 5.00
Begin VB.Form frmChemists 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Chemists"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   5925
      TabIndex        =   20
      Top             =   4455
      Width           =   5925
      Begin VB.CommandButton cmdNew 
         Caption         =   "&New"
         Height          =   345
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   930
      End
      Begin VB.CommandButton cmdModify 
         Caption         =   "&Modify"
         Height          =   345
         Left            =   945
         TabIndex        =   8
         Top             =   0
         Width           =   930
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   345
         Left            =   1890
         TabIndex        =   5
         Top             =   0
         Width           =   930
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   345
         Left            =   2835
         TabIndex        =   6
         Top             =   0
         Width           =   930
      End
      Begin VB.CommandButton cmdTerminate 
         Caption         =   "&Delete"
         Height          =   345
         Left            =   3780
         TabIndex        =   9
         Top             =   0
         Width           =   930
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Quit"
         Height          =   345
         Left            =   4725
         TabIndex        =   10
         Top             =   0
         Width           =   930
      End
   End
   Begin VB.TextBox txtdoc 
      DataField       =   "c_id"
      Height          =   285
      Index           =   0
      Left            =   30
      TabIndex        =   0
      Top             =   255
      Width           =   1170
   End
   Begin VB.TextBox txtdoc 
      DataField       =   "c_name"
      Height          =   285
      Index           =   2
      Left            =   15
      MaxLength       =   50
      TabIndex        =   1
      Top             =   855
      Width           =   5895
   End
   Begin VB.TextBox txtdoc 
      DataField       =   "c_address"
      Height          =   1095
      Index           =   3
      Left            =   15
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1485
      Width           =   2895
   End
   Begin VB.TextBox txtdoc 
      DataField       =   "location"
      Height          =   1095
      Index           =   4
      Left            =   3015
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   1485
      Width           =   2895
   End
   Begin VB.TextBox txtdoc 
      DataField       =   "notes"
      Height          =   855
      Index           =   6
      Left            =   15
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   2970
      Width           =   5895
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
      Picture         =   "frmChemists.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3930
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
      Picture         =   "frmChemists.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3930
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
      Picture         =   "frmChemists.frx":0294
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3930
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
      Picture         =   "frmChemists.frx":03DE
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3930
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Chemist ID"
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
      Left            =   30
      TabIndex        =   19
      Top             =   15
      Width           =   1080
   End
   Begin VB.Label Label3 
      Caption         =   "Chemist Name"
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
      Left            =   15
      TabIndex        =   18
      Top             =   600
      Width           =   1275
   End
   Begin VB.Label Label4 
      Caption         =   "Address"
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
      Left            =   15
      TabIndex        =   17
      Top             =   1245
      Width           =   1395
   End
   Begin VB.Label Label5 
      Caption         =   "Physical Location"
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
      Left            =   3015
      TabIndex        =   16
      Top             =   1245
      Width           =   1740
   End
   Begin VB.Label Label7 
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
      Height          =   225
      Left            =   15
      TabIndex        =   15
      Top             =   2730
      Width           =   1320
   End
End
Attribute VB_Name = "frmChemists"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim datChemists As Recordset

Private Sub cmdCancel_Click()
    datChemists.CancelUpdate
    Reset True
End Sub

Private Sub cmdClose_Click()
    Unload Me
    
End Sub

Private Sub cmdModify_Click()
    If datChemists.RecordCount = 0 Then
       MsgBox ("No record to edit."), vbInformation, "Chemists"
    Else
       Reset False
    End If
End Sub

Private Sub cmdNew_Click()
    'Calculate the new doctor id
    Set datExtra = New Recordset
    datExtra.Open "SELECT TOP 1 * FROM chemists WHERE co_code ='" & ID & "' ORDER BY c_id DESC", cn, adOpenStatic, adLockOptimistic
    If datExtra.RecordCount <> 0 Then
       x = datExtra!c_id + 1
    Else
        x = 1
    End If
    
    datChemists.AddNew
    txtdoc(0) = x
    txtdoc(0).Locked = True
    txtdoc(2).SetFocus
    Reset False
End Sub

Private Sub cmdSave_Click()
    'Validate for the entries to be made
    If Trim(txtdoc(2)) = "" Then
       MsgBox "You must give the chemists' name.", vbInformation, "Chemists"
       Exit Sub
    Else
        datChemists!co_code = ID
        datChemists.Update
        Reset True
    End If
End Sub

Private Sub cmdTerminate_Click()
    If datChemists.RecordCount = 0 Then
       MsgBox "No record to delete", vbInformation, "Chemists"
       Exit Sub
    End If
    If MsgBox("Delete this chemist record?", vbYesNo + vbQuestion) = vbYes Then
        datChemists.Delete
        datChemists.Requery
    Else
       MsgBox "Deletion aborted", vbInformation, "Chemists"
    End If
End Sub

Private Sub ffirst_Click()
On Error GoTo PrevErr
    With datChemists
         .MoveFirst
    End With
    Exit Sub
PrevErr:
MsgBox (Err.Description), vbInformation
Err.Clear
End Sub

Private Sub flast_Click()
On Error GoTo PrevErr
    With datChemists
         .MoveLast
    End With
    Exit Sub
PrevErr:
MsgBox (Err.Description), vbInformation
Err.Clear
End Sub

Private Sub fnext_Click()
On Error GoTo PrevErr
    With datChemists
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

Private Sub Form_Load()
    Dim txt As TextBox
    Set datChemists = New Recordset
    datChemists.Open "SELECT * FROM chemists WHERE co_code ='" & ID & _
    "'", cn, adOpenStatic, adLockOptimistic
    
    For Each txt In Me.txtdoc
        Set txt.DataSource = datChemists
    Next
    
    Reset True
End Sub

Private Sub fprev_Click()
On Error GoTo PrevErr
    With datChemists
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

Private Sub txtdoc_GotFocus(Index As Integer)
    txtdoc(Index) = Trim(txtdoc(Index))
    
End Sub

Private Sub txtdoc_KeyPress(Index As Integer, KeyAscii As Integer)
    Char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(Char))
End Sub

Public Sub Reset(bval As Boolean)
    Dim txt As TextBox
    cmdNew.Enabled = bval
    cmdModify.Enabled = bval
    cmdSave.Enabled = Not bval
    cmdCancel.Enabled = Not bval
    cmdTerminate.Enabled = bval
    ffirst.Enabled = bval
    fprev.Enabled = bval
    fnext.Enabled = bval
    flast.Enabled = bval
    
    For Each txt In Me.txtdoc
        txt.Locked = bval
    Next
    txtdoc(0).Locked = True
    
End Sub


