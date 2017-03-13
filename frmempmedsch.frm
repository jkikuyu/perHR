VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmempmedsch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Employee Medical Schemes"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6060
   Icon            =   "frmempmedsch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   6060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2715
      Left            =   15
      TabIndex        =   19
      Top             =   1065
      Width           =   6015
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   350
         Left            =   4680
         TabIndex        =   34
         Top             =   1800
         Width           =   1260
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   350
         Left            =   4680
         TabIndex        =   33
         Top             =   1395
         Width           =   1260
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   350
         Left            =   4680
         TabIndex        =   32
         Top             =   990
         Width           =   1260
      End
      Begin VB.CommandButton cmdModify 
         Caption         =   "Modify"
         Height          =   350
         Left            =   4680
         TabIndex        =   31
         Top             =   585
         Width           =   1260
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "&New"
         Height          =   350
         Left            =   4680
         TabIndex        =   30
         Top             =   195
         Width           =   1260
      End
      Begin VB.TextBox txtsch 
         DataField       =   "reg_no"
         Height          =   285
         Index           =   2
         Left            =   90
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   1155
         Width           =   2175
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   285
         Left            =   3900
         TabIndex        =   20
         Top             =   1875
         Width           =   270
         _ExtentX        =   476
         _ExtentY        =   503
         _Version        =   393216
         Format          =   24641537
         CurrentDate     =   36875
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   1770
         TabIndex        =   21
         Top             =   1875
         Width           =   270
         _ExtentX        =   476
         _ExtentY        =   503
         _Version        =   393216
         Format          =   24641537
         CurrentDate     =   36875
      End
      Begin MSDataListLib.DataCombo cboSchemes 
         DataField       =   "scheme_id"
         Height          =   315
         Left            =   90
         TabIndex        =   29
         Top             =   465
         Width           =   4125
         _ExtentX        =   7276
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "scheme_name"
         BoundColumn     =   "scheme_id"
         Text            =   ""
      End
      Begin VB.TextBox txtsch 
         DataField       =   "start_date"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MMM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   90
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   1875
         Width           =   1935
      End
      Begin VB.TextBox txtsch 
         DataField       =   "end_date"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MMM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   2100
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   1875
         Width           =   2055
      End
      Begin VB.Label Label4 
         Caption         =   "Medical Scheme"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   90
         TabIndex        =   28
         Top             =   195
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Registration Number"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   90
         TabIndex        =   27
         Top             =   915
         Width           =   2295
      End
      Begin VB.Label Label6 
         Caption         =   "Start Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   90
         TabIndex        =   26
         Top             =   1635
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "Expiry Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2100
         TabIndex        =   25
         Top             =   1635
         Width           =   1455
      End
   End
   Begin VB.Frame fropts 
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   30
      TabIndex        =   17
      Top             =   3765
      Width           =   3825
      Begin VB.CheckBox chkSusp 
         Caption         =   "Suspend Scheme"
         DataField       =   "suspend"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   270
         Width           =   1935
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
      Left            =   5190
      Picture         =   "frmempmedsch.frx":0742
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3900
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
      Left            =   4710
      Picture         =   "frmempmedsch.frx":088C
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3900
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
      Left            =   4245
      Picture         =   "frmempmedsch.frx":09D6
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3900
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
      Left            =   5670
      Picture         =   "frmempmedsch.frx":0B20
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3900
      Width           =   375
   End
   Begin VB.TextBox txtsch 
      Height          =   285
      Index           =   0
      Left            =   8070
      TabIndex        =   11
      Top             =   2310
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   6060
      TabIndex        =   9
      Top             =   4515
      Width           =   6060
      Begin VB.CommandButton cmdGoTo 
         Caption         =   "Go To..."
         Height          =   345
         Left            =   60
         TabIndex        =   35
         Top             =   15
         Width           =   930
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "C&lose"
         Height          =   345
         Left            =   5070
         TabIndex        =   10
         Top             =   15
         Width           =   930
      End
   End
   Begin VB.CommandButton cmdScheme 
      Height          =   285
      Left            =   8400
      Picture         =   "frmempmedsch.frx":0C6A
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2520
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.TextBox txtsch 
      Height          =   285
      Index           =   1
      Left            =   8070
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   2520
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.CommandButton cmdFind 
      Height          =   375
      Left            =   5640
      Picture         =   "frmempmedsch.frx":0D64
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Select an Employee"
      Top             =   240
      Width           =   375
   End
   Begin Threed.SSPanel pnlother_names 
      Height          =   375
      Left            =   2910
      TabIndex        =   1
      Top             =   240
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
      Top             =   240
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
      Top             =   240
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
   Begin VB.Label lblstatus 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   705
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      X1              =   3135
      X2              =   3135
      Y1              =   0
      Y2              =   240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000009&
      X1              =   960
      X2              =   960
      Y1              =   0
      Y2              =   240
   End
   Begin VB.Label Label3 
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
   Begin VB.Label Label2 
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
   Begin VB.Label Label1 
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
Attribute VB_Name = "frmempmedsch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents datsch As Recordset
Attribute datsch.VB_VarHelpID = -1
Dim datSchemes As Recordset

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    If datsch.RecordCount = 0 Then
       MsgBox ("No Record to delete."), vbInformation, "Record Deletion"
       Exit Sub
    End If
    If MsgBox("Delete this medical scheme assignment to this employee?", vbQuestion + vbYesNo, "Record Deletion") = vbYes Then
       datsch.Delete
       MsgBox ("Record deleted."), vbInformation, "Record Deletion"
       datsch.Requery
    Else
        MsgBox ("Deletion Aborted."), vbInformation, "Record Deletion"
        
    End If
End Sub

Private Sub cmdFind_Click()
    frmSelectEmp.Left = Me.Left + 40
    frmSelectEmp.Top = Me.Top + pnlStaff_code.Top + pnlStaff_code.Height
    frmSelectEmp.Show vbModal
    
    If pnlStaff_code <> "" Then
       datsch.Filter = "Staff_code ='" & Trim(pnlStaff_code) & "'"
        If datsch.RecordCount = 0 Then
           cmdModify.Enabled = False
           cmdDelete.Enabled = False
        Else
           cmdModify.Enabled = True
           cmdDelete.Enabled = True
        End If
    End If
End Sub

Private Sub cmdGoTo_Click()
    If datEmp.RecordCount = 0 Then
       MsgBox ("File is empty."), vbInformation, "Find"
       Exit Sub
    End If
    Dim Response
    Response = InputBox("Go To Employee number...", "Go To")
    If Response <> "" Then
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
    
       datsch.Filter = "Staff_code ='" & Trim(pnlStaff_code) & "'"
    End If
End Sub

Private Sub cmdSave_Click()
    If cboSchemes.BoundText = "" Or Trim(txtsch(2)) = "" Or Trim(txtsch(3)) = "" Or Trim(txtsch(4)) = "" Then
       MsgBox ("Must supply the following data:" & Chr(13) & _
       "              A Valid Medical Scheme" & Chr(13) & _
       "              A Registration Number" & Chr(13) & _
       "              A Start Date" & Chr(13) & _
       "              An Expiry Date"), vbInformation
       Exit Sub
    ElseIf CVDate(txtsch(3)) >= CVDate(txtsch(4)) Then
        MsgBox ("Expiry date cannot be earlier or same as the start date."), vbInformation
        Exit Sub
    Else
        datsch!staff_code = pnlStaff_code
        datsch!co_code = ID
        datsch.Update
        bk = datsch.Bookmark
        Reset True
        datsch.Bookmark = bk
    End If
    
End Sub

Private Sub cmdScheme_Click()
    frmselschemes.Left = Me.Left + 150
    frmselschemes.Top = Me.Top + (Label4.Top - Label1.Top) + (pnlStaff_code.Height * 2) + 60
    frmselschemes.Show vbModal
End Sub

Private Sub datsch_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error Resume Next
    'Find the status of this record
    If datsch.RecordCount = 0 Then
       lblstatus.Visible = False
       Exit Sub
    Else
        lblstatus.Visible = True
        With datsch
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
    txtsch(3) = Format(DTPicker1.Value, "dd/mmm/yyyy")
End Sub

Private Sub DTPicker2_CloseUp()
    txtsch(4) = Format(DTPicker2.Value, "dd/mmm/yyyy")
End Sub

Private Sub ffirst_Click()
On Error GoTo PrevErr
    With datsch
         .MoveFirst
    End With
    Exit Sub
PrevErr:
MsgBox (Err.Description), vbInformation
Err.Clear
End Sub

Private Sub flast_Click()
On Error GoTo PrevErr
    With datsch
         .MoveLast
    End With
    Exit Sub
PrevErr:
MsgBox (Err.Description), vbInformation
Err.Clear
End Sub
Private Sub cmdCancel_Click()
    datsch.CancelUpdate
    Reset True
End Sub

Private Sub cmdModify_Click()
    If datsch.RecordCount = 0 Then
       MsgBox ("No Record to modify."), vbInformation, "Record Modification"
       Exit Sub
    End If
    Reset False
    
End Sub

Private Sub cmdNew_Click()
    If Trim(pnlStaff_code) = "" Then
       MsgBox ("No Employee selected."), vbInformation, "Medical Schemes"
       Exit Sub
    End If
    datsch.AddNew
    chkSusp.Value = 0
    txtsch(3) = Date
    txtsch(4) = Date
    Reset False
End Sub

Private Sub fnext_Click()
On Error GoTo PrevErr
    With datsch
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
    PSch = True
    
    Set datsch = New Recordset
    datsch.Open "SELECT * FROM emp_medsch WHERE co_code ='" & ID & "'", cn, adOpenStatic, adLockOptimistic
    
    For Each txt In Me.txtsch
        Set txt.DataSource = datsch
    Next
    Set chkSusp.DataSource = datsch
    Set cboSchemes.DataSource = datsch
    
    Set datSchemes = New Recordset
    datSchemes.Open "SELECT * FROM med_schemes", cn, adOpenStatic, adLockOptimistic
    Set cboSchemes.RowSource = datSchemes
    
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
            datsch.Filter = "Staff_code ='" & Trim(pnlStaff_code) & "'"
            If datsch.RecordCount = 0 Then
               cmdModify.Enabled = False
               cmdDelete.Enabled = False
            End If
        Else
            cmdModify.Enabled = False
        End If
    End With
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PSch = False
End Sub

Public Sub Reset(bval As Boolean)
    cmdNew.Enabled = bval
    cmdSave.Enabled = Not bval
    cmdModify.Enabled = bval
    cmdCancel.Enabled = Not bval
    cmdDelete.Enabled = bval
    cmdScheme.Enabled = Not bval
    cmdFind.Enabled = bval
    fropts.Enabled = Not bval
    cboSchemes.Enabled = Not bval
    
    DTPicker1.Enabled = Not bval
    DTPicker2.Enabled = Not bval
    txtsch(2).Locked = bval
    
End Sub

Private Sub fprev_Click()
On Error GoTo PrevErr
    With datsch
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

Private Sub txtsch_Change(Index As Integer)
    If Index = 0 Then
       If Trim(txtsch(Index)) <> "" Then
            'Find this medical scheme
            Set datExtra = New Recordset
            datExtra.Open "SELECT * FROM med_schemes WHERE scheme_id =" & _
            txtsch(0), cn, adOpenStatic, adLockOptimistic
            If datExtra.RecordCount <> 0 Then
               txtsch(1) = datExtra!scheme_name
            Else
                txtsch(1) = ""
            End If
        Else
            txtsch(1) = ""
        End If
    End If
End Sub

Private Sub txtsch_GotFocus(Index As Integer)
    txtsch(Index) = Trim(txtsch(Index))
End Sub

Private Sub txtsch_KeyPress(Index As Integer, KeyAscii As Integer)
    Char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(Char))
End Sub
