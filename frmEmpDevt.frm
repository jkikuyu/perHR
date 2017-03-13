VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmEmpDevt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Professional Development"
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6015
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEmpDevt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGoTo 
      Caption         =   "&Go To..."
      Height          =   360
      Left            =   15
      TabIndex        =   28
      Top             =   6045
      Width           =   960
   End
   Begin VB.TextBox txtDevt 
      DataField       =   "institution"
      Height          =   330
      Index           =   2
      Left            =   105
      MaxLength       =   50
      TabIndex        =   27
      Top             =   1650
      Width           =   5490
   End
   Begin VB.ComboBox cboYear 
      DataField       =   "d_date"
      Height          =   345
      Left            =   90
      TabIndex        =   25
      Text            =   "cboYear"
      Top             =   4170
      Width           =   1890
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
      Height          =   390
      Left            =   0
      ScaleHeight     =   390
      ScaleWidth      =   6015
      TabIndex        =   18
      Top             =   6525
      Width           =   6015
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   360
         Left            =   3915
         TabIndex        =   24
         Top             =   0
         Width           =   960
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "C&lose"
         Height          =   360
         Left            =   4995
         TabIndex        =   23
         Top             =   0
         Width           =   960
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   360
         Left            =   2925
         TabIndex        =   22
         Top             =   0
         Width           =   960
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   360
         Left            =   1950
         TabIndex        =   21
         Top             =   0
         Width           =   960
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   360
         Left            =   975
         TabIndex        =   20
         Top             =   0
         Width           =   960
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "&New"
         Height          =   360
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Width           =   960
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
      Left            =   4770
      Picture         =   "frmEmpDevt.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Next Record"
      Top             =   6015
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
      Left            =   4290
      Picture         =   "frmEmpDevt.frx":058C
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Previous Record"
      Top             =   6015
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
      Left            =   3825
      Picture         =   "frmEmpDevt.frx":06D6
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "First Record"
      Top             =   6015
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
      Left            =   5250
      Picture         =   "frmEmpDevt.frx":0820
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Last Record"
      Top             =   6015
      Width           =   375
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
      Picture         =   "frmEmpDevt.frx":096A
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   270
      Width           =   375
   End
   Begin VB.TextBox txtDevt 
      DataField       =   "notes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1110
      Index           =   3
      Left            =   75
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   4785
      Width           =   5520
   End
   Begin VB.TextBox txtDevt 
      DataField       =   "d_level"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1590
      Index           =   1
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   2265
      Width           =   5475
   End
   Begin VB.TextBox txtDevt 
      DataField       =   "d_name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   90
      MaxLength       =   50
      TabIndex        =   1
      Top             =   975
      Width           =   5505
   End
   Begin Threed.SSPanel pnlother_names 
      Height          =   375
      Left            =   2910
      TabIndex        =   8
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
      TabIndex        =   9
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
      TabIndex        =   10
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
      Caption         =   "Institution"
      Height          =   240
      Left            =   105
      TabIndex        =   26
      Top             =   1395
      Width           =   2595
   End
   Begin VB.Label Label9 
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
      Left            =   3135
      TabIndex        =   13
      Top             =   0
      Width           =   2415
   End
   Begin VB.Label Label8 
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
      TabIndex        =   12
      Top             =   0
      Width           =   2175
   End
   Begin VB.Label Label7 
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
      TabIndex        =   11
      Top             =   0
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Notes"
      Height          =   285
      Left            =   90
      TabIndex        =   4
      Top             =   4560
      Width           =   1350
   End
   Begin VB.Label Label3 
      Caption         =   "Year Attained"
      Height          =   195
      Left            =   90
      TabIndex        =   3
      Top             =   3945
      Width           =   1845
   End
   Begin VB.Label Label2 
      Caption         =   "Achievement"
      Height          =   210
      Left            =   105
      TabIndex        =   2
      Top             =   2010
      Width           =   1620
   End
   Begin VB.Label Label1 
      Caption         =   "Development Activity / Course"
      Height          =   225
      Left            =   105
      TabIndex        =   0
      Top             =   750
      Width           =   2880
   End
End
Attribute VB_Name = "frmEmpDevt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim datDevt As Recordset
Private Sub cmdGoTo_Click()
    If datEmp.RecordCount = 0 Then
       MsgBox ("File is empty."), vbInformation, "Find"
       Exit Sub
    End If
    Dim Response
    Response = InputBox("Go To Employee number...", "Go To")
    If Response <> "" Then
        If Len(Response) > 7 Then
           MsgBox "Invalid Staff Code", vbInformation, "Professional Development"
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
    
       datDevt.Filter = "Staff_code ='" & Trim(pnlStaff_code) & "'"
    End If
End Sub
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    If datDevt.RecordCount = 0 Then
       MsgBox "No Record to Delete", vbInformation, "Deletion aborted"
       Exit Sub
    End If
    If MsgBox("Delete this development record?", vbYesNo + vbQuestion, "Employee Development") = vbYes Then
       datDevt.Delete
       datDevt.Requery
    End If
End Sub

Private Sub cmdFind_Click()
    frmSelectEmp.Left = Me.Left + 40
    frmSelectEmp.Top = Me.Top + pnlStaff_code.Top + pnlStaff_code.Height
    frmSelectEmp.Show vbModal
    
    If pnlStaff_code <> "" Then
       datDevt.Filter = "Staff_code ='" & Trim(pnlStaff_code) & "'"
        If datDevt.RecordCount = 0 Then
           cmdEdit.Enabled = False
           'cmdDelete.Enabled = False
        Else
           cmdEdit.Enabled = True
           'cmdDelete.Enabled = True
        End If
    End If

End Sub

Private Sub dtDevt_CloseUp()
    txtDevt(2) = dtDevt.Value
End Sub

Private Sub fprev_Click()
On Error GoTo PrevErr
    With datDevt
         .MovePrevious
         If .BOF Then
            .MoveFirst
            MsgBox ("First Record encountered"), vbInformation, "Employee Development"
         End If
    End With
    Exit Sub
PrevErr:
MsgBox (Err.Description), vbInformation, "Employee Development"
Err.Clear
End Sub
Private Sub ffirst_Click()
On Error GoTo PrevErr
    With datDevt
         .MoveFirst
    End With
    Exit Sub
PrevErr:
MsgBox (Err.Description), vbInformation, "Employee Development"
Err.Clear
End Sub

Private Sub flast_Click()
On Error GoTo PrevErr
    With datDevt
         .MoveLast
    End With
    Exit Sub
PrevErr:
MsgBox (Err.Description), vbInformation, "Employee Development"
Err.Clear
End Sub

Private Sub fnext_Click()
On Error GoTo PrevErr
    With datDevt
         .MoveNext
         If .EOF Then
            .MoveLast
            MsgBox ("Last Record encountered"), vbInformation, "Employee Development"
         End If
    End With
    Exit Sub
PrevErr:
MsgBox (Err.Description), vbInformation, "Employee Development"

Err.Clear
End Sub
Public Sub Reset(bval As Boolean)

    Dim txt As TextBox
    cmdNew.Enabled = bval
    cmdEdit.Enabled = bval
    cmdUpdate.Enabled = Not bval
    cmdCancel.Enabled = Not bval
    cmdFind.Enabled = bval
        
    ffirst.Enabled = bval
    fnext.Enabled = bval
    fprev.Enabled = bval
    flast.Enabled = bval
    
    For Each txt In Me.txtDevt
        txt.Locked = bval
    Next
      
    cboYear.Enabled = Not bval
        
End Sub
Private Sub cmdCancel_Click()
    datDevt.CancelUpdate
    Reset True
End Sub

Private Sub cmdEdit_Click()
    If datDevt.RecordCount = 0 Then
       MsgBox "No Record to Edit", vbInformation, "Editing Aborted"
       Exit Sub
    End If
    Reset False
End Sub

Private Sub cmdNew_Click()
    If Trim(pnlStaff_code) = "" Then
       MsgBox ("No Employee selected."), vbInformation, "Employee Development"
       Exit Sub
    End If
    datDevt.AddNew
    txtDevt(2) = Date
    Reset False
End Sub

Private Sub cmdUpdate_Click()
    If Trim(txtDevt(0)) = "" Then
       MsgBox ("You must give the name of this development record."), vbInformation, "Employee Development"
       Exit Sub
    End If
    'Validate for the year give
    If cboYear <> "" Then
       If Not IsNumeric(cboYear) Then
          MsgBox "Invalid year", vbInformation, "Employee Development"
          cboYear.SetFocus
          Exit Sub
       End If
    End If
    
    
    datDevt!co_code = ID
    datDevt!staff_code = pnlStaff_code
    
    datDevt.Update
    If datDevt.Bookmark > 0 Then
        bk = datDevt.Bookmark
        datDevt.Bookmark = bk
    End If
    Reset True
End Sub

Private Sub Form_Load()
    'Load the years
    For x = 1900 To 2099
        cboYear.AddItem x
    Next
    
    PDevt = True
    Set datDevt = New Recordset
    datDevt.Open "SELECT * FROM emp_devt WHERE co_code ='" & ID & _
    "'", cn, adOpenStatic, adLockOptimistic
    For Each txt In Me.txtDevt
        Set txt.DataSource = datDevt
    Next
    Set cboYear.DataSource = datDevt
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
            datDevt.Filter = "Staff_code ='" & Trim(pnlStaff_code) & "'"
            If datDevt.RecordCount = 0 Then
               cmdEdit.Enabled = False
               'cmdDelete.Enabled = False
            End If
        Else
            cmdEdit.Enabled = False
        End If
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PDevt = False
End Sub

