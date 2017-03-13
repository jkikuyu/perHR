VERSION 5.00
Begin VB.Form frmco_doctors 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Company Doctors and Hospitals"
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6045
   Icon            =   "frmco_doctors.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   6045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame Frame1 
      Caption         =   "Data Entry Type"
      Height          =   900
      Left            =   2895
      TabIndex        =   25
      Top             =   0
      Width           =   3045
      Begin VB.OptionButton optHospital 
         Caption         =   "Hospital Entry"
         Height          =   195
         Left            =   330
         TabIndex        =   27
         Top             =   600
         Width           =   1920
      End
      Begin VB.OptionButton optDoctor 
         Caption         =   "Doctor Entry"
         Height          =   195
         Left            =   330
         TabIndex        =   26
         Top             =   270
         Value           =   -1  'True
         Width           =   1890
      End
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
      Picture         =   "frmco_doctors.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5055
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
      Picture         =   "frmco_doctors.frx":058C
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5055
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
      Picture         =   "frmco_doctors.frx":06D6
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5055
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
      Left            =   5145
      Picture         =   "frmco_doctors.frx":0820
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5055
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   6045
      TabIndex        =   24
      Top             =   5595
      Width           =   6045
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Quit"
         Height          =   345
         Left            =   4725
         TabIndex        =   16
         Top             =   0
         Width           =   930
      End
      Begin VB.CommandButton cmdTerminate 
         Caption         =   "&Delete"
         Height          =   345
         Left            =   3780
         TabIndex        =   15
         Top             =   0
         Width           =   930
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   345
         Left            =   2835
         TabIndex        =   12
         Top             =   0
         Width           =   930
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   345
         Left            =   1890
         TabIndex        =   11
         Top             =   0
         Width           =   930
      End
      Begin VB.CommandButton cmdModify 
         Caption         =   "&Modify"
         Height          =   345
         Left            =   945
         TabIndex        =   14
         Top             =   0
         Width           =   930
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "&New"
         Height          =   345
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   930
      End
   End
   Begin VB.TextBox txtdoc 
      DataField       =   "notes"
      Height          =   855
      Index           =   6
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   4095
      Width           =   5895
   End
   Begin VB.TextBox txtdoc 
      DataField       =   "speciality"
      Height          =   285
      Index           =   5
      Left            =   1080
      TabIndex        =   5
      Top             =   3495
      Width           =   3975
   End
   Begin VB.TextBox txtdoc 
      DataField       =   "location"
      Height          =   1095
      Index           =   4
      Left            =   3120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   2175
      Width           =   2895
   End
   Begin VB.TextBox txtdoc 
      DataField       =   "doctor_address"
      Height          =   1095
      Index           =   3
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   2175
      Width           =   2895
   End
   Begin VB.TextBox txtdoc 
      DataField       =   "doctor_fname"
      Height          =   285
      Index           =   2
      Left            =   120
      MaxLength       =   50
      TabIndex        =   1
      Top             =   945
      Width           =   4830
   End
   Begin VB.TextBox txtdoc 
      DataField       =   "doctor_lname"
      Height          =   285
      Index           =   1
      Left            =   105
      MaxLength       =   50
      TabIndex        =   2
      Top             =   1560
      Width           =   4845
   End
   Begin VB.TextBox txtdoc 
      DataField       =   "doctor_id"
      Height          =   285
      Index           =   0
      Left            =   135
      TabIndex        =   0
      Top             =   360
      Width           =   1500
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
      Left            =   120
      TabIndex        =   23
      Top             =   3855
      Width           =   1320
   End
   Begin VB.Label Label6 
      Caption         =   "Speciality"
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
      Left            =   120
      TabIndex        =   22
      Top             =   3495
      Width           =   900
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
      Left            =   3120
      TabIndex        =   21
      Top             =   1935
      Width           =   1740
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
      Left            =   120
      TabIndex        =   20
      Top             =   1935
      Width           =   1395
   End
   Begin VB.Label Label3 
      Caption         =   "First Name"
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
      Left            =   120
      TabIndex        =   19
      Top             =   705
      Width           =   1275
   End
   Begin VB.Label Label2 
      Caption         =   "Last Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   18
      Top             =   1305
      Width           =   1740
   End
   Begin VB.Label Label1 
      Caption         =   "Doctor/Hospital ID"
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
      Left            =   135
      TabIndex        =   17
      Top             =   120
      Width           =   1620
   End
End
Attribute VB_Name = "frmco_doctors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents datdoc As Recordset

Private Sub cmdCancel_Click()
    datdoc.CancelUpdate
    Reset True
End Sub

Private Sub cmdClose_Click()
    Unload Me
    
End Sub

Private Sub cmdModify_Click()
    If datdoc.RecordCount = 0 Then
       MsgBox ("No record to edit."), vbInformation, "Company Doctors"
    Else
       Reset False
    End If
End Sub

Private Sub cmdNew_Click()
    'Calculate the new doctor id
    Set datExtra = New Recordset
    datExtra.Open "SELECT TOP 1 * FROM co_doctors WHERE co_code ='" & ID & "' ORDER BY doctor_id DESC", cn, adOpenStatic, adLockOptimistic
    If datExtra.RecordCount <> 0 Then
       x = datExtra!doctor_id + 1
    Else
        x = 1
    End If
    
    datdoc.AddNew
    txtdoc(0) = x
    txtdoc(0).Locked = True
    txtdoc(2).SetFocus
    Reset False
End Sub

Private Sub cmdSave_Click()
    'Validate for the entries to be made
    If Trim(txtdoc(2)) = "" Then
       MsgBox "You must give the doctors first name details .", vbInformation, "Company Doctors"
       Exit Sub
    Else
        If optDoctor.Value = True Then
           datdoc!e_type = 1
        Else
           datdoc!e_type = 2
        End If
        datdoc!co_code = ID
        datdoc.Update
        Reset True
    End If
End Sub

Private Sub cmdTerminate_Click()
    If datdoc.RecordCount = 0 Then
       MsgBox "No record to delete", vbInformation, "Company Doctors"
       Exit Sub
    End If
    If MsgBox("Delete this doctor record?", vbYesNo + vbQuestion) = vbYes Then
        datdoc.Delete
        datdoc.Requery
    Else
       MsgBox "Deletion aborted", vbInformation, "Company Doctors"
    End If
End Sub

Private Sub datdoc_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error Resume Next
    'Display the entry type for records
    If datdoc!e_type = 1 Then
       optDoctor.Value = True
    ElseIf datdoc!e_type = 2 Then
       optHospital.Value = True
    End If
    
End Sub

Private Sub ffirst_Click()
On Error GoTo PrevErr
    With datdoc
         .MoveFirst
    End With
    Exit Sub
PrevErr:
MsgBox (Err.Description), vbInformation
Err.Clear
End Sub

Private Sub flast_Click()
On Error GoTo PrevErr
    With datdoc
         .MoveLast
    End With
    Exit Sub
PrevErr:
MsgBox (Err.Description), vbInformation
Err.Clear
End Sub

Private Sub fnext_Click()
On Error GoTo PrevErr
    With datdoc
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
    Set datdoc = New Recordset
    datdoc.Open "SELECT * FROM co_doctors WHERE co_code ='" & ID & _
    "'", cn, adOpenStatic, adLockOptimistic
    
    For Each txt In Me.txtdoc
        Set txt.DataSource = datdoc
    Next
    
    Reset True
End Sub

Private Sub fprev_Click()
On Error GoTo PrevErr
    With datdoc
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
