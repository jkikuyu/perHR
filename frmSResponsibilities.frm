VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmSResponsibilities 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Social and Public Responsibilities"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   6000
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
      Left            =   5415
      Picture         =   "frmSResponsibilities.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Last Record"
      Top             =   2250
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
      Left            =   3990
      Picture         =   "frmSResponsibilities.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "First Record"
      Top             =   2250
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
      Left            =   4455
      Picture         =   "frmSResponsibilities.frx":0294
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Previous Record"
      Top             =   2250
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
      Left            =   4935
      Picture         =   "frmSResponsibilities.frx":03DE
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Next Record"
      Top             =   2250
      Width           =   375
   End
   Begin VB.ComboBox cboTo 
      DataField       =   "To"
      Height          =   315
      Left            =   3915
      TabIndex        =   19
      Top             =   1770
      Width           =   1905
   End
   Begin VB.ComboBox cboFrom 
      DataField       =   "From"
      Height          =   315
      Left            =   870
      TabIndex        =   18
      Top             =   1770
      Width           =   1905
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   6000
      TabIndex        =   11
      Top             =   2760
      Width           =   6000
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   350
         Left            =   3945
         TabIndex        =   17
         Top             =   0
         Width           =   930
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   350
         Left            =   2985
         TabIndex        =   16
         Top             =   0
         Width           =   930
      End
      Begin VB.CommandButton cmdModify 
         Caption         =   "Modify"
         Height          =   350
         Left            =   1065
         TabIndex        =   15
         Top             =   0
         Width           =   930
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "&New"
         Height          =   350
         Left            =   105
         TabIndex        =   14
         Top             =   0
         Width           =   930
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   350
         Left            =   2025
         TabIndex        =   13
         Top             =   0
         Width           =   930
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "C&lose"
         Height          =   350
         Left            =   4905
         TabIndex        =   12
         Top             =   0
         Width           =   930
      End
   End
   Begin VB.TextBox txtResponsibility 
      DataField       =   "Responsibility"
      Height          =   675
      Left            =   75
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   990
      Width           =   5730
   End
   Begin VB.CommandButton cmdFind 
      Height          =   375
      Left            =   5595
      Picture         =   "frmSResponsibilities.frx":0528
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
   Begin VB.Label Label3 
      Caption         =   "To"
      Height          =   240
      Left            =   3330
      TabIndex        =   10
      Top             =   1815
      Width           =   465
   End
   Begin VB.Label Label2 
      Caption         =   "From"
      Height          =   285
      Left            =   90
      TabIndex        =   9
      Top             =   1815
      Width           =   570
   End
   Begin VB.Label Label1 
      Caption         =   "Social/Public Responsibility"
      Height          =   300
      Left            =   75
      TabIndex        =   7
      Top             =   750
      Width           =   2535
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
      TabIndex        =   6
      Top             =   0
      Width           =   2415
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
      TabIndex        =   4
      Top             =   0
      Width           =   975
   End
End
Attribute VB_Name = "frmSResponsibilities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim datSResp As Recordset
Public Sub Reset(bval As Boolean)
    cboFrom.Enabled = bval
    cboTo.Enabled = bval
    cmdNew.Enabled = Not bval
    cmdSave.Enabled = bval
    cmdModify.Enabled = Not bval
    cmdCancel.Enabled = bval
    cmdDelete.Enabled = Not bval
    txtResponsibility.Locked = Not bval
    
    ffirst.Enabled = Not bval
    flast.Enabled = Not bval
    fnext.Enabled = Not bval
    fprev.Enabled = Not bval
    cmdFind.Enabled = Not bval
End Sub

Private Sub ffirst_Click()
On Error GoTo PrevErr
    With datSResp
         .MoveFirst
    End With
    Exit Sub
PrevErr:
MsgBox (Err.Description), vbInformation
Err.Clear
End Sub

Private Sub flast_Click()
On Error GoTo PrevErr
    With datSResp
         .MoveLast
    End With
    Exit Sub
PrevErr:
MsgBox (Err.Description), vbInformation
Err.Clear
End Sub
Private Sub fnext_Click()
On Error GoTo PrevErr
    With datSResp
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
    With datSResp
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


Private Sub cmdModify_Click()
    If datSResp.RecordCount = 0 Then
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
    
    datSResp.AddNew
    TAdd = True
    txtResponsibility.SetFocus
    
    datSResp!co_code = ID
    Reset True
End Sub

Private Sub cmdSave_Click()
    datSResp!staff_code = pnlStaff_code
    datSResp.Update
    Reset False
    
End Sub


Private Sub cmdCancel_Click()
'On Error GoTo CancelErr
    datSResp.CancelUpdate
    Reset False
    Exit Sub
CancelErr:
datSResp.CancelBatch
Reset False
'datsresp.Requery
End Sub

Private Sub cmdClose_Click()
    Unload Me
    
End Sub

Private Sub cmdDelete_Click()
    If datSResp.RecordCount = 0 Then
       MsgBox "No Record to Delete", vbInformation, "Modification aborted"
       Exit Sub
    End If
    If MsgBox("Delete this record.", vbQuestion + vbYesNo) = vbYes Then
        datSResp.Delete
        datSResp.Requery
    Else
        MsgBox ("Deletion aborted."), vbInformation
    End If
End Sub


Private Sub cmdFind_Click()
    frmSelectEmp.Left = Me.Left + 40
    frmSelectEmp.Top = Me.Top + pnlStaff_code.Top + pnlStaff_code.Height
    frmSelectEmp.Show vbModal
    
    If pnlStaff_code <> "" Then
       datSResp.Filter = "Staff_code ='" & Trim(pnlStaff_code) & "'"
        If datSResp.RecordCount = 0 Then
           cmdModify.Enabled = False
           cmdDelete.Enabled = False
        Else
           cmdModify.Enabled = True
           cmdDelete.Enabled = True
        End If
    End If
End Sub

Private Sub Form_Load()
    PSocial = True
    'Populate the combo boxes with the years
    Dim x As Long
    For x = 1900 To 2099
        cboFrom.AddItem x
        cboTo.AddItem x
    Next
    
    
    Set datSResp = New Recordset
    datSResp.Open "SELECT * FROM socialr WHERE co_code ='" & ID & "'", cn, adOpenStatic, adLockOptimistic
    Set Me.txtResponsibility.DataSource = datSResp
    Set cboFrom.DataSource = datSResp
    Set cboTo.DataSource = datSResp
    
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
            datSResp.Filter = "Staff_code ='" & Trim(pnlStaff_code) & "'"
            If datSResp.RecordCount = 0 Then
               cmdModify.Enabled = False
               cmdDelete.Enabled = False
            End If
        Else
            cmdModify.Enabled = False
        End If
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PSocial = False
End Sub
