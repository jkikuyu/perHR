VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmEmpHist 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Employment History"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6000
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEmpHist.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGoTo 
      Caption         =   "&Go To..."
      Height          =   360
      Left            =   15
      TabIndex        =   32
      Top             =   5535
      Width           =   960
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   390
      Left            =   0
      ScaleHeight     =   390
      ScaleWidth      =   6000
      TabIndex        =   25
      Top             =   5985
      Width           =   6000
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   360
         Left            =   3915
         TabIndex        =   31
         Top             =   0
         Width           =   960
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "C&lose"
         Height          =   360
         Left            =   5010
         TabIndex        =   30
         Top             =   0
         Width           =   960
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   360
         Left            =   2925
         TabIndex        =   29
         Top             =   0
         Width           =   960
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   360
         Left            =   1950
         TabIndex        =   28
         Top             =   0
         Width           =   960
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   360
         Left            =   975
         TabIndex        =   27
         Top             =   0
         Width           =   960
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "&New"
         Height          =   360
         Left            =   0
         TabIndex        =   26
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
      Left            =   5040
      Picture         =   "frmEmpHist.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Next Record"
      Top             =   5385
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
      Picture         =   "frmEmpHist.frx":058C
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Previous Record"
      Top             =   5385
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
      Picture         =   "frmEmpHist.frx":06D6
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "First Record"
      Top             =   5385
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
      Picture         =   "frmEmpHist.frx":0820
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Last Record"
      Top             =   5385
      Width           =   375
   End
   Begin VB.CommandButton cmdFind 
      Height          =   375
      Left            =   5595
      Picture         =   "frmEmpHist.frx":096A
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   270
      Width           =   375
   End
   Begin MSComCtl2.DTPicker dtEnd 
      Height          =   330
      Left            =   4815
      TabIndex        =   13
      Top             =   3345
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   582
      _Version        =   393216
      Format          =   24641537
      CurrentDate     =   36898
   End
   Begin MSComCtl2.DTPicker dtStart 
      Height          =   330
      Left            =   1875
      TabIndex        =   12
      Top             =   3390
      Width           =   270
      _ExtentX        =   476
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "mmm/yyyy"
      Format          =   24641537
      CurrentDate     =   36898
   End
   Begin VB.TextBox txtEHist 
      DataField       =   "notes"
      Height          =   1170
      Index           =   3
      Left            =   105
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   4140
      Width           =   5820
   End
   Begin VB.TextBox txtEHist 
      DataField       =   "ed_date"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "mmm/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   330
      Index           =   2
      Left            =   2880
      TabIndex        =   8
      Top             =   3345
      Width           =   1935
   End
   Begin VB.TextBox txtEHist 
      DataField       =   "st_date"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "mmm/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   3390
      Width           =   1755
   End
   Begin MSDataListLib.DataCombo cboEHist 
      DataField       =   "desg_code"
      Height          =   345
      Index           =   1
      Left            =   135
      TabIndex        =   5
      Top             =   2610
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   609
      _Version        =   393216
      Style           =   2
      ListField       =   "paramname"
      BoundColumn     =   "code"
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cboEHist 
      DataField       =   "dept_code"
      Height          =   345
      Index           =   0
      Left            =   105
      TabIndex        =   3
      Top             =   1875
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   609
      _Version        =   393216
      Style           =   2
      ListField       =   "paramname"
      BoundColumn     =   "code"
      Text            =   ""
   End
   Begin VB.TextBox txtEHist 
      DataField       =   "co_name"
      Height          =   315
      Index           =   0
      Left            =   105
      TabIndex        =   1
      Top             =   1050
      Width           =   5850
   End
   Begin Threed.SSPanel pnlother_names 
      Height          =   375
      Left            =   2910
      TabIndex        =   15
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
      TabIndex        =   16
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
      TabIndex        =   17
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
      Left            =   3120
      TabIndex        =   20
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
      TabIndex        =   19
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
      TabIndex        =   18
      Top             =   0
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "Notes"
      Height          =   300
      Left            =   165
      TabIndex        =   11
      Top             =   3885
      Width           =   750
   End
   Begin VB.Label Label5 
      Caption         =   "End Date"
      Height          =   300
      Left            =   2865
      TabIndex        =   10
      Top             =   3090
      Width           =   915
   End
   Begin VB.Label Label4 
      Caption         =   "Start Date"
      Height          =   225
      Left            =   120
      TabIndex        =   6
      Top             =   3090
      Width           =   945
   End
   Begin VB.Label Label3 
      Caption         =   "Designation"
      Height          =   300
      Left            =   165
      TabIndex        =   4
      Top             =   2340
      Width           =   1065
   End
   Begin VB.Label Label2 
      Caption         =   "Department"
      Height          =   300
      Left            =   150
      TabIndex        =   2
      Top             =   1605
      Width           =   1185
   End
   Begin VB.Label Label1 
      Caption         =   "Company Name"
      Height          =   300
      Left            =   135
      TabIndex        =   0
      Top             =   765
      Width           =   1530
   End
End
Attribute VB_Name = "frmEmpHist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim datHist As Recordset
Dim datDept As Recordset
Dim datDesg As Recordset

Private Sub cmdGoTo_Click()
    If datEmp.RecordCount = 0 Then
       MsgBox ("File is empty."), vbInformation, "Find"
       Exit Sub
    End If
    Dim Response
    Response = InputBox("Go To Employee number...", "Go To")
    If Response <> "" Then
        If Len(Response) > 7 Then
           MsgBox "Invalid Staff Code", vbInformation, "Employment History"
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
    
       datHist.Filter = "Staff_code ='" & Trim(pnlStaff_code) & "'"
    End If
End Sub

Public Sub Reset(bval As Boolean)
    Dim txt As TextBox
    cmdNew.Enabled = bval
    cmdEdit.Enabled = bval
    cmdUpdate.Enabled = Not bval
    cmdCancel.Enabled = Not bval
    cmdFind.Enabled = bval
    cmdDelete.Enabled = bval
        
    ffirst.Enabled = bval
    fnext.Enabled = bval
    fprev.Enabled = bval
    flast.Enabled = bval
    
    For Each txt In Me.txtEHist
        txt.Locked = bval
    Next
    cboEHist(0).Enabled = Not bval
    cboEHist(1).Enabled = Not bval
    
    dtStart.Enabled = Not bval
    dtEnd.Enabled = Not bval
    
End Sub

Private Sub cmdCancel_Click()
    datHist.CancelUpdate
    Reset True
End Sub

Private Sub cmdDelete_Click()
    If datHist.RecordCount = 0 Then
       MsgBox "No Record to Delete", vbInformation, "Deletion aborted"
       Exit Sub
    End If
    If MsgBox("Delete this employment record?", vbYesNo + vbQuestion, "Employment History") = vbYes Then
       datHist.Delete
       datHist.Requery
    End If
End Sub

Private Sub cmdEdit_Click()
    If datHist.RecordCount = 0 Then
       MsgBox "No record to edit", vbInformation, "Editing Aborted"
       Exit Sub
    End If
    Reset False
End Sub

Private Sub cmdNew_Click()
    If Trim(pnlStaff_code) = "" Then
       MsgBox ("No Employee selected."), vbInformation, "Employment History"
       Exit Sub
    End If
    datHist.AddNew
    txtEHist(1) = Format(Date, "mmm/yyyy")
    txtEHist(2) = Format(Date, "mmm/yyyy")
    Reset False
End Sub



Private Sub cmdUpdate_Click()
    If Trim(txtEHist(0)) = "" Or Trim(txtEHist(1)) = "" Or Trim(txtEHist(2)) = "" Or cboEHist(1).BoundText = "" Then
       MsgBox ("The following data are required for education history records: " & Chr(13) & _
       "           Company where employee worked" & Chr(13) & _
       "           Designation" & Chr(13) & _
       "           Date Started" & Chr(13) & _
       "           Date Completed"), vbInformation, "Employment History"
       Exit Sub
    End If

    datHist!co_code = ID
    datHist!staff_code = pnlStaff_code
    
    datHist.Update
    If datHist.Bookmark > 0 Then
        bk = datHist.Bookmark
        datHist.Bookmark = bk
    End If
    Reset True
End Sub

Private Sub dtEnd_CloseUp()
    txtEHist(2) = Format(dtEnd.Value, "mmm/yyyy")
End Sub

Private Sub dtStart_CloseUp()
    txtEHist(1) = Format(dtStart.Value, "mmm/yyyy")
End Sub

Private Sub ffirst_Click()
On Error GoTo PrevErr
    With datHist
         .MoveFirst
    End With
    Exit Sub
PrevErr:
MsgBox (Err.Description), vbInformation, "Employment History"
Err.Clear
End Sub

Private Sub flast_Click()
On Error GoTo PrevErr
    With datHist
         .MoveLast
    End With
    Exit Sub
PrevErr:
MsgBox (Err.Description), vbInformation, "Employment History"
Err.Clear
End Sub

Private Sub fnext_Click()
On Error GoTo PrevErr
    With datHist
         .MoveNext
         If .EOF Then
            .MoveLast
            MsgBox ("Last Record encountered"), vbInformation, "Employment History"
         End If
    End With
    Exit Sub
PrevErr:
MsgBox (Err.Description), vbInformation, "Employment History"

Err.Clear
End Sub
Private Sub fprev_Click()
On Error GoTo PrevErr
    With datHist
         .MovePrevious
         If .BOF Then
            .MoveFirst
            MsgBox ("First Record encountered"), vbInformation, "Employment History"
         End If
    End With
    Exit Sub
PrevErr:
MsgBox (Err.Description), vbInformation, "Employment History"
Err.Clear
End Sub
Private Sub cmdClose_Click()
    Unload Me
    
End Sub

Private Sub cmdFind_Click()
    frmSelectEmp.Left = Me.Left + 40
    frmSelectEmp.Top = Me.Top + pnlStaff_code.Top + pnlStaff_code.Height
    frmSelectEmp.Show vbModal
    
    If pnlStaff_code <> "" Then
       datHist.Filter = "Staff_code ='" & Trim(pnlStaff_code) & "'"
        If datHist.RecordCount = 0 Then
           cmdEdit.Enabled = False
           'cmdDelete.Enabled = False
        Else
           cmdEdit.Enabled = True
           'cmdDelete.Enabled = True
        End If
    End If
End Sub

Private Sub Form_Load()
    PEHist = True
    Set datHist = New Recordset
    datHist.Open "SELECT * FROM emp_history WHERE co_code ='" & ID & _
    "'", cn, adOpenStatic, adLockOptimistic
    For Each txt In Me.txtEHist
        Set txt.DataSource = datHist
    Next
    Set cboEHist(0).DataSource = datHist
    Set cboEHist(1).DataSource = datHist
    
    Set datDept = New Recordset
    datDept.Open "SELECT * FROM udfparameters WHERE co_code='" & ID & _
    "' AND paramcode='Dpt'", cn, adOpenStatic, adLockOptimistic
    Set cboEHist(0).RowSource = datDept
    
    Set datDesg = New Recordset
    datDesg.Open "SELECT * FROM udfparameters WHERE co_code='" & ID & _
    "' AND paramcode='Des'", cn, adOpenStatic, adLockOptimistic
    Set cboEHist(1).RowSource = datDesg
    
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
            datHist.Filter = "Staff_code ='" & Trim(pnlStaff_code) & "'"
            If datHist.RecordCount = 0 Then
               cmdEdit.Enabled = False
               'cmdDelete.Enabled = False
            End If
        Else
            cmdEdit.Enabled = False
        End If
    End With
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PEHist = False
End Sub
