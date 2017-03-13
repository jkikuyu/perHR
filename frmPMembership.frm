VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmPMembership 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Professional Membership"
   ClientHeight    =   5385
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   6030
      TabIndex        =   21
      Top             =   5010
      Width           =   6030
      Begin VB.CommandButton cmdClose 
         Caption         =   "C&lose"
         Height          =   350
         Left            =   4905
         TabIndex        =   27
         Top             =   0
         Width           =   930
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   350
         Left            =   2025
         TabIndex        =   26
         Top             =   0
         Width           =   930
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "&New"
         Height          =   350
         Left            =   105
         TabIndex        =   25
         Top             =   0
         Width           =   930
      End
      Begin VB.CommandButton cmdModify 
         Caption         =   "Modify"
         Height          =   350
         Left            =   1065
         TabIndex        =   24
         Top             =   0
         Width           =   930
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   350
         Left            =   2985
         TabIndex        =   23
         Top             =   0
         Width           =   930
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   350
         Left            =   3945
         TabIndex        =   22
         Top             =   0
         Width           =   930
      End
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
      Left            =   5115
      Picture         =   "frmPMembership.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Next Record"
      Top             =   4500
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
      Left            =   4635
      Picture         =   "frmPMembership.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Previous Record"
      Top             =   4500
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
      Left            =   4170
      Picture         =   "frmPMembership.frx":0294
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "First Record"
      Top             =   4500
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
      Left            =   5595
      Picture         =   "frmPMembership.frx":03DE
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Last Record"
      Top             =   4500
      Width           =   375
   End
   Begin VB.TextBox txtMember 
      DataField       =   "notes"
      Height          =   1335
      Index           =   2
      Left            =   15
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   16
      Top             =   3075
      Width           =   5955
   End
   Begin VB.ComboBox cboMember 
      DataField       =   "to"
      Height          =   345
      Index           =   1
      Left            =   3585
      TabIndex        =   14
      Top             =   2310
      Width           =   1395
   End
   Begin VB.ComboBox cboMember 
      DataField       =   "from"
      Height          =   345
      Index           =   0
      Left            =   870
      TabIndex        =   13
      Top             =   2295
      Width           =   1455
   End
   Begin VB.TextBox txtMember 
      DataField       =   "level"
      Height          =   330
      Index           =   1
      Left            =   30
      MaxLength       =   100
      TabIndex        =   12
      Top             =   1740
      Width           =   5895
   End
   Begin VB.TextBox txtMember 
      DataField       =   "organztn"
      Height          =   330
      Index           =   0
      Left            =   45
      MaxLength       =   100
      TabIndex        =   8
      Top             =   1005
      Width           =   5910
   End
   Begin VB.CommandButton cmdFind 
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
      Left            =   5595
      Picture         =   "frmPMembership.frx":0528
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   270
      Width           =   375
   End
   Begin Threed.SSPanel pnlother_names 
      Height          =   375
      Left            =   2910
      TabIndex        =   1
      Top             =   270
      Width           =   2655
      _Version        =   65536
      _ExtentX        =   4683
      _ExtentY        =   661
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
   End
   Begin Threed.SSPanel pnlLast_name 
      Height          =   375
      Left            =   975
      TabIndex        =   2
      Top             =   270
      Width           =   1920
      _Version        =   65536
      _ExtentX        =   3387
      _ExtentY        =   661
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSPanel pnlStaff_code 
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   270
      Width           =   975
      _Version        =   65536
      _ExtentX        =   1720
      _ExtentY        =   661
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label5 
      Caption         =   "Notes"
      Height          =   285
      Left            =   30
      TabIndex        =   15
      Top             =   2835
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "To"
      Height          =   225
      Left            =   2925
      TabIndex        =   11
      Top             =   2370
      Width           =   645
   End
   Begin VB.Label Label3 
      Caption         =   "From"
      Height          =   240
      Left            =   45
      TabIndex        =   10
      Top             =   2340
      Width           =   660
   End
   Begin VB.Label Label2 
      Caption         =   "Membership Level"
      Height          =   195
      Left            =   30
      TabIndex        =   9
      Top             =   1515
      Width           =   3360
   End
   Begin VB.Label Label1 
      Caption         =   "Membership Organisation"
      Height          =   270
      Left            =   45
      TabIndex        =   7
      Top             =   750
      Width           =   4560
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000001&
      Caption         =   " Staff Code"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   975
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000001&
      Caption         =   " Last Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   960
      TabIndex        =   5
      Top             =   0
      Width           =   2175
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000001&
      Caption         =   " Other Names"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   3120
      TabIndex        =   4
      Top             =   0
      Width           =   2415
   End
End
Attribute VB_Name = "frmPMembership"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim datMember As Recordset
Dim e_ID As Long

Public Sub Reset(bval As Boolean)
    Dim txt As TextBox
    Dim cbo As ComboBox
    
    For Each cbo In Me.cboMember
        cbo.Enabled = bval
    Next
    cmdNew.Enabled = Not bval
    cmdSave.Enabled = bval
    cmdModify.Enabled = Not bval
    cmdCancel.Enabled = bval
    cmdDelete.Enabled = Not bval
    
    For Each txt In Me.txtMember
        txt.Locked = Not bval
    Next
    
    ffirst.Enabled = Not bval
    flast.Enabled = Not bval
    fnext.Enabled = Not bval
    fprev.Enabled = Not bval
    cmdFind.Enabled = Not bval
End Sub

Private Sub cmdModify_Click()
    If datMember.RecordCount = 0 Then
       MsgBox "No Record to Modify", vbInformation, "Editing Aborted"
       Exit Sub
    End If
    Reset True
    
End Sub

Private Sub cmdNew_Click()
    If Trim(pnlStaff_code) = "" Then
       MsgBox ("No Employee selected."), vbInformation, "Absenteeism"
       Exit Sub
    End If
    'Calculate the new e_ID
    Set datExtra = New Recordset
    datExtra.Open "SELECT TOP 1 * FROM PMember ORDER BY E_id DESC", cn, adOpenStatic, adLockOptimistic
    If datExtra.RecordCount = 0 Then
       e_ID = 1
    Else: e_ID = datExtra!e_ID + 1
    End If
    
    datMember.AddNew
    TAdd = True
    txtMember(0).SetFocus
    datMember!e_ID = e_ID
    datMember!co_code = ID
    Reset True
End Sub

Private Sub cmdSave_Click()
    datMember!staff_code = pnlStaff_code
    datMember.Update
    Reset False
    
End Sub


Private Sub cmdCancel_Click()
'On Error GoTo CancelErr
    datMember.CancelUpdate
    Reset False
    Exit Sub
CancelErr:
datMember.CancelBatch
Reset False
'datMember.Requery
End Sub

Private Sub cmdClose_Click()
    Unload Me
    
End Sub

Private Sub cmdDelete_Click()
    If datMember.RecordCount = 0 Then
       MsgBox "No Record to Delete", vbInformation, "Modification aborted"
       Exit Sub
    End If
    If MsgBox("Delete this record.", vbQuestion + vbYesNo) = vbYes Then
        datMember.Delete
        datMember.Requery
    Else
        MsgBox ("Deletion aborted."), vbInformation
    End If
End Sub


Private Sub cmdFind_Click()
    frmSelectEmp.Left = Me.Left + 40
    frmSelectEmp.Top = Me.Top + pnlStaff_code.Top + pnlStaff_code.Height
    frmSelectEmp.Show vbModal
    
    If pnlStaff_code <> "" Then
       datMember.Filter = "Staff_code ='" & Trim(pnlStaff_code) & "'"
        If datMember.RecordCount = 0 Then
           cmdModify.Enabled = False
           cmdDelete.Enabled = False
        Else
           cmdModify.Enabled = True
           cmdDelete.Enabled = True
        End If
    End If
End Sub


Private Sub ffirst_Click()
On Error GoTo PrevErr
    With datMember
         .MoveFirst
    End With
    Exit Sub
PrevErr:
MsgBox (Err.Description), vbInformation
Err.Clear
End Sub

Private Sub flast_Click()
On Error GoTo PrevErr
    With datMember
         .MoveLast
    End With
    Exit Sub
PrevErr:
MsgBox (Err.Description), vbInformation
Err.Clear
End Sub
Private Sub fnext_Click()
On Error GoTo PrevErr
    With datMember
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

Private Sub Form_Unload(Cancel As Integer)
    PMember = False
End Sub

Private Sub fprev_Click()
On Error GoTo PrevErr
    With datMember
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


Private Sub Form_Load()
    Dim txt As TextBox
    Dim cbo As ComboBox
    
    PMember = True
    'Populate the combo boxes with the years
    Dim x As Long
    For x = 1900 To 2099
        cboMember(0).AddItem x
        cboMember(1).AddItem x
    Next
    
    
    Set datMember = New Recordset
    datMember.Open "SELECT * FROM pMember WHERE co_code ='" & ID & "'", cn, adOpenStatic, adLockOptimistic
    For Each txt In Me.txtMember
        Set txt.DataSource = datMember
    Next
    For Each cbo In Me.cboMember
        Set cbo.DataSource = datMember
    Next
    
    Reset False
    Set datEmp = New Recordset
    With datEmp
        datEmp.Open "SELECT * FROM Personal_data WHERE co_code ='" & ID & "'", cn, adOpenStatic, adLockOptimistic
        If .RecordCount <> 0 Then
            pnlStaff_code = Trim(!staff_code)
            pnlLast_name = Trim(!last_name)
            If Not IsNull(!mid_name) Then
                pnlother_names = Trim(!first_name) & ", " & Trim(!mid_name)
            Else
                pnlother_names = Trim(!first_name)
            End If
            datMember.Filter = "Staff_code ='" & Trim(pnlStaff_code) & "'"
            If datMember.RecordCount = 0 Then
               cmdModify.Enabled = False
               cmdDelete.Enabled = False
            End If
        Else
            cmdModify.Enabled = False
        End If
    End With
    
End Sub

Private Sub txtMember_GotFocus(Index As Integer)
    txtMember(Index) = Trim(txtMember(Index))
End Sub
