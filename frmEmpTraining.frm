VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmEmpTraining 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Employee Training History"
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6015
   Icon            =   "frmEmpTraining.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGoTo 
      Caption         =   "&Go To..."
      Height          =   360
      Left            =   120
      TabIndex        =   27
      Top             =   4575
      Width           =   930
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   6015
      TabIndex        =   20
      Top             =   4995
      Width           =   6015
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   350
         Left            =   3945
         TabIndex        =   26
         Top             =   0
         Width           =   930
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   350
         Left            =   2985
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
      Begin VB.CommandButton cmdNew 
         Caption         =   "&New"
         Height          =   350
         Left            =   105
         TabIndex        =   23
         Top             =   0
         Width           =   930
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   350
         Left            =   2025
         TabIndex        =   22
         Top             =   0
         Width           =   930
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "C&lose"
         Height          =   350
         Left            =   4905
         TabIndex        =   21
         Top             =   0
         Width           =   930
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
      Left            =   5415
      Picture         =   "frmEmpTraining.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Last Record"
      Top             =   4020
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
      Picture         =   "frmEmpTraining.frx":058C
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "First Record"
      Top             =   4020
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
      Picture         =   "frmEmpTraining.frx":06D6
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Previous Record"
      Top             =   4020
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
      Picture         =   "frmEmpTraining.frx":0820
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Next Record"
      Top             =   4020
      Width           =   375
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   285
      Left            =   2025
      TabIndex        =   15
      Top             =   4065
      Width           =   270
      _ExtentX        =   476
      _ExtentY        =   503
      _Version        =   393216
      Format          =   24576001
      CurrentDate     =   36878
   End
   Begin VB.TextBox txtTraining 
      DataField       =   "date"
      Height          =   285
      Index           =   3
      Left            =   60
      TabIndex        =   14
      Top             =   4065
      Width           =   1965
   End
   Begin VB.TextBox txtTraining 
      DataField       =   "achievement"
      Height          =   840
      Index           =   2
      Left            =   75
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   13
      Top             =   2895
      Width           =   5670
   End
   Begin VB.TextBox txtTraining 
      DataField       =   "institution"
      Height          =   615
      Index           =   1
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Top             =   1950
      Width           =   5700
   End
   Begin VB.TextBox txtTraining 
      DataField       =   "training"
      Height          =   690
      Index           =   0
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   975
      Width           =   5730
   End
   Begin VB.CommandButton cmdFind 
      Height          =   375
      Left            =   5595
      Picture         =   "frmEmpTraining.frx":096A
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
   Begin VB.Label Label4 
      Caption         =   "Date Achieved"
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
      Left            =   60
      TabIndex        =   10
      Top             =   3825
      Width           =   2190
   End
   Begin VB.Label Label3 
      Caption         =   "Acheivement"
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
      Left            =   90
      TabIndex        =   9
      Top             =   2625
      Width           =   1905
   End
   Begin VB.Label Label2 
      Caption         =   "Institution Attended"
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
      Left            =   75
      TabIndex        =   8
      Top             =   1695
      Width           =   2010
   End
   Begin VB.Label Label1 
      Caption         =   "Training Undergone"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   45
      TabIndex        =   7
      Top             =   735
      Width           =   1770
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
Attribute VB_Name = "frmEmpTraining"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents datTraining As Recordset
Attribute datTraining.VB_VarHelpID = -1
Private Sub cmdGoTo_Click()
    If datEmp.RecordCount = 0 Then
       MsgBox ("File is empty."), vbInformation, "Find"
       Exit Sub
    End If
    Dim Response
    Response = InputBox("Go To Employee number...", "Go To")
    If Response <> "" Then
        If Len(Response) > 7 Then
           MsgBox "Invalid Staff Code", vbInformation, "Training History"
           Exit Sub
        End If
        With datEmp
            .Find "Staff_code ='" & Trim(Response) & "'", 0, adSearchForward, 1
            If datEmp.EOF Then
               MsgBox ("Employee Record not found"), vbExclamation, "Record Missing"
               datEmp.Requery
            Else
               pnlStaff_code = Trim(!staff_code)
               pnlLast_name = Trim(!last_name)
               If Not IsNull(!mid_name) Then
                  pnlother_names = Trim(!first_name) & ", " & Trim(!mid_name)
               Else
                  pnlother_names = Trim(!first_name)
               End If
            End If
        End With
    
       datTraining.Filter = "Staff_code ='" & Trim(pnlStaff_code) & "'"
    End If
End Sub
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    If datTraining.RecordCount = 0 Then
       MsgBox "No Record to Delete", vbInformation, "Deletion aborted"
       Exit Sub
    End If

    If MsgBox("Delete this training record?", vbQuestion + vbYesNo) = vbYes Then
       datTraining.Delete
       MsgBox ("Record deleted."), vbInformation
       datTraining.Requery
    Else
        MsgBox ("Deletion Aborted."), vbInformation
        
    End If
End Sub

Private Sub cmdFind_Click()
    frmSelectEmp.Left = Me.Left + 40
    frmSelectEmp.Top = Me.Top + pnlStaff_code.Top + pnlStaff_code.Height
    frmSelectEmp.Show vbModal
    
    If pnlStaff_code <> "" Then
       datTraining.Filter = "Staff_code ='" & Trim(pnlStaff_code) & "'"
        If datTraining.RecordCount = 0 Then
           cmdModify.Enabled = False
           cmdDelete.Enabled = False
        Else
           cmdModify.Enabled = True
           cmdDelete.Enabled = True
        End If
    End If
End Sub

Private Sub cmdSave_Click()
    If txtTraining(0) = "" Or txtTraining(2) = "" Or txtTraining(3) = "" Then
       MsgBox ("Must supply the following data:" & Chr(13) & _
       "              The kind of Training" & Chr(13) & _
       "              The achievement in the training" & Chr(13) & _
       "              The Date"), vbInformation
       Exit Sub
    Else
        datTraining!staff_code = pnlStaff_code
        datTraining!co_code = ID
        datTraining.Update
        bk = datTraining.Bookmark
        Reset True
        datTraining.Bookmark = bk
    End If
    
End Sub

Private Sub cmdScheme_Click()
    frmselschemes.Left = Me.Left + 150
    frmselschemes.Top = Me.Top + (Label4.Top - Label1.Top) + (pnlStaff_code.Height * 2) + 60
    frmselschemes.Show vbModal
End Sub

Private Sub datTraining_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error Resume Next
    'Find the status of this record
    If datTraining.RecordCount = 0 Then
       lblstatus.Visible = False
       Exit Sub
    Else
        lblstatus.Visible = True
        With datTraining
            If !suspend = True Then
                lblstatus = "Suspended"
            ElseIf CVDate(!End_Date) < Date Then
                lblstatus = "Expired"
            Else
                lblstatus = "Valid"
            End If
        End With
    End If
End Sub

Private Sub DTPicker1_CloseUp()
    txtTraining(3) = Format(DTPicker1.Value, "dd/mmm/yyyy")
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
Private Sub cmdCancel_Click()
    datTraining.CancelUpdate
    Reset True
End Sub

Private Sub cmdModify_Click()
    If datTraining.RecordCount = 0 Then
       MsgBox "No Record to Modify", vbInformation, "Editing Aborted"
       Exit Sub
    End If
    Reset False
    
End Sub

Private Sub cmdNew_Click()
    If Trim(pnlStaff_code) = "" Then
       MsgBox ("No Employee selected."), vbInformation, "Training History"
       Exit Sub
    End If
    datTraining.AddNew
    txtTraining(3) = Date
    Reset False
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

Private Sub Form_Load()
    Dim txt As TextBox
    PTr = True
    
    Set datTraining = New Recordset
    datTraining.Open "SELECT * FROM emp_training WHERE co_code ='" & ID & "'", cn, adOpenStatic, adLockOptimistic
    
    For Each txt In Me.txtTraining
        Set txt.DataSource = datTraining
    Next
    'Set chkSusp.DataSource = datTraining
    
    
    Reset True
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
            datTraining.Filter = "Staff_code ='" & Trim(pnlStaff_code) & "'"
            If datTraining.RecordCount = 0 Then
               cmdModify.Enabled = False
               cmdDelete.Enabled = False
            End If
        Else
            cmdModify.Enabled = False
        End If
    End With
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PTr = False
End Sub

Public Sub Reset(bval As Boolean)
    Dim txt As TextBox
    cmdNew.Enabled = bval
    cmdSave.Enabled = Not bval
    cmdModify.Enabled = bval
    cmdCancel.Enabled = Not bval
    cmdDelete.Enabled = bval
    cmdFind.Enabled = bval
    DTPicker1.Enabled = Not bval
    For Each txt In Me.txtTraining
        txt.Locked = bval
    Next
    
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

Private Sub txttraining_GotFocus(Index As Integer)
    txtTraining(Index) = Trim(txtTraining(Index))
End Sub

Private Sub txtTraining_KeyPress(Index As Integer, KeyAscii As Integer)
    'Char = Chr(KeyAscii)
    'KeyAscii = Asc(UCase(Char))
End Sub
