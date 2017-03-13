VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmEmpApp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Employee Appraisals"
   ClientHeight    =   5190
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
   Icon            =   "frmEmpApp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   390
      Left            =   0
      ScaleHeight     =   390
      ScaleWidth      =   6030
      TabIndex        =   20
      Top             =   4800
      Width           =   6030
      Begin VB.CommandButton cmdCancel 
         Caption         =   "C&ancel"
         Height          =   375
         Left            =   2040
         TabIndex        =   24
         Top             =   0
         Width           =   945
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   375
         Left            =   5025
         TabIndex        =   23
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   15
         TabIndex        =   22
         Top             =   0
         Width           =   945
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   375
         Left            =   1035
         TabIndex        =   21
         Top             =   0
         Width           =   945
      End
   End
   Begin VB.CommandButton cmdLast 
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
      Left            =   5610
      Picture         =   "frmEmpApp.frx":044A
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Move to the Last Record"
      Top             =   4305
      Width           =   375
   End
   Begin VB.CommandButton cmdFirst 
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
      Picture         =   "frmEmpApp.frx":0594
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Move to the First Record"
      Top             =   4305
      Width           =   375
   End
   Begin VB.CommandButton cmdPrev 
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
      Left            =   4650
      Picture         =   "frmEmpApp.frx":06DE
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Show Previous Record"
      Top             =   4305
      Width           =   375
   End
   Begin VB.CommandButton cmdNext 
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
      Left            =   5130
      Picture         =   "frmEmpApp.frx":0828
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Show Next Record"
      Top             =   4305
      Width           =   375
   End
   Begin VB.TextBox txtApp 
      DataField       =   "param_code"
      Height          =   330
      Index           =   0
      Left            =   3315
      TabIndex        =   15
      Top             =   1440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtApp 
      DataField       =   "notes"
      Height          =   1035
      Index           =   3
      Left            =   15
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   14
      Top             =   3165
      Width           =   6015
   End
   Begin VB.TextBox txtApp 
      DataField       =   "points_awarded"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   330
      Index           =   2
      Left            =   45
      TabIndex        =   13
      Top             =   2490
      Width           =   1290
   End
   Begin VB.TextBox txtApp 
      Height          =   330
      Index           =   1
      Left            =   45
      TabIndex        =   12
      Top             =   1755
      Width           =   5445
   End
   Begin VB.CommandButton cmdFind 
      Height          =   375
      Left            =   5655
      Picture         =   "frmEmpApp.frx":0972
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Select an Employee"
      Top             =   990
      Width           =   375
   End
   Begin MSDataListLib.DataCombo dbDefn 
      Height          =   345
      Left            =   45
      TabIndex        =   1
      Top             =   330
      Width           =   4260
      _ExtentX        =   7514
      _ExtentY        =   609
      _Version        =   393216
      Style           =   2
      ListField       =   "app_name"
      BoundColumn     =   "defn_code"
      Text            =   ""
   End
   Begin Threed.SSPanel pnlother_names 
      Height          =   375
      Left            =   2925
      TabIndex        =   3
      Top             =   990
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
      Left            =   990
      TabIndex        =   4
      Top             =   990
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
      Left            =   15
      TabIndex        =   5
      Top             =   990
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
   Begin VB.Label Label7 
      Caption         =   "Notes"
      Height          =   270
      Left            =   30
      TabIndex        =   11
      Top             =   2940
      Width           =   2985
   End
   Begin VB.Label Label6 
      Caption         =   "Points Awarded"
      Height          =   285
      Left            =   45
      TabIndex        =   10
      Top             =   2250
      Width           =   2760
   End
   Begin VB.Label Label5 
      Caption         =   "Appraisal Parameter"
      Height          =   300
      Left            =   60
      TabIndex        =   9
      Top             =   1515
      Width           =   3030
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      X1              =   3150
      X2              =   3150
      Y1              =   750
      Y2              =   990
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000009&
      X1              =   975
      X2              =   975
      Y1              =   750
      Y2              =   990
   End
   Begin VB.Label Label4 
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
      TabIndex        =   8
      Top             =   750
      Width           =   2415
   End
   Begin VB.Label Label3 
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
      Left            =   975
      TabIndex        =   7
      Top             =   750
      Width           =   2175
   End
   Begin VB.Label Label2 
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
      Left            =   15
      TabIndex        =   6
      Top             =   750
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Select an Appraisal Definition"
      Height          =   300
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   3000
   End
End
Attribute VB_Name = "frmEmpApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim datDefn As Recordset
Dim datApp As Recordset

Public Sub reset(bval As Boolean)
    'Enable and disbale as appropriate
    Dim txt As TextBox
    cmdEdit.Enabled = bval
    cmdUpdate.Enabled = Not bval
    cmdCancel.Enabled = Not bval
    cmdNext.Enabled = bval
    cmdPrev.Enabled = bval
    cmdFirst.Enabled = bval
    cmdLast.Enabled = bval
    
    
    For Each txt In Me.txtApp
        txt.Locked = bval
    Next
    txtApp(1).Locked = True
End Sub

Private Sub cmdCancel_Click()
    datApp.CancelUpdate
    reset True
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdEdit_Click()
    reset False
End Sub

Private Sub cmdfirst_Click()
On Error GoTo PrevErr
    With datApp
         .MoveFirst
    End With
    Exit Sub
PrevErr:
MsgBox (Err.Description), vbInformation
Err.Clear
End Sub

Private Sub cmdlast_Click()
On Error GoTo PrevErr
    With datApp
         .MoveLast
    End With
    Exit Sub
PrevErr:
MsgBox (Err.Description), vbInformation
Err.Clear
End Sub


Private Sub cmdNext_Click()
On Error GoTo NextErr
    With datApp
        .MoveNext
         If .EOF Then
            .MoveLast
            MsgBox ("End of file"), vbInformation
         End If
    End With
    Exit Sub
NextErr:
MsgBox (Err.Description), vbInformation
    
End Sub

Private Sub cmdPrev_Click()
On Error GoTo PrevErr
    With datApp
        .MovePrevious
         If .BOF Then
            .MoveFirst
            MsgBox ("End of file"), vbInformation
         End If
    End With
    Exit Sub
PrevErr:
MsgBox (Err.Description), vbInformation
End Sub

Private Sub cmdFind_Click()
    Dim txt As TextBox
    frmSelectEmp.Left = Me.Left + 40
    frmSelectEmp.Top = Me.Top + pnlStaff_code.Top + pnlStaff_code.Height - 50
    frmSelectEmp.Show vbModal
    
    'Assign data sources
    For Each txt In Me.txtApp
        Set txt.DataSource = datApp
    Next
    'Get the appraisal entries for this employee
    If dbDefn.Text <> "" Then
        If pnlStaff_code <> "" Then
           datApp.Filter = "Staff_code ='" & Trim(pnlStaff_code) & "' AND defn_code =" & dbDefn.BoundText
           reset True
        End If
    End If
End Sub

Private Sub cmdUpdate_Click()
    datApp.Update
    reset True
    
End Sub

Private Sub dbDefn_Change()
    'Assign data sources
    For Each txt In Me.txtApp
        Set txt.DataSource = datApp
    Next
    'Get the appraisal entries for this employee
    If dbDefn.Text <> "" Then
        If pnlStaff_code <> "" Then
           datApp.Filter = "Staff_code ='" & Trim(pnlStaff_code) & "' AND defn_code =" & dbDefn.BoundText
           reset True
        End If
    End If
End Sub

Private Sub Form_Load()
    pApp = True
    Set datDefn = New Recordset
    datDefn.Open "SELECT * FROM ApprDefn WHERE co_code ='" & _
    ID & "'", cn, adOpenStatic, adLockOptimistic
    
    Set dbDefn.DataSource = datDefn
    Set dbDefn.RowSource = datDefn
    
    Set datEmp = New Recordset
    With datEmp
        .Open "SELECT * FROM Personal_data WHERE co_code ='" & ID & _
        "' AND Terminated <> 1", cn, adOpenStatic, adLockOptimistic
        If .RecordCount <> 0 Then
            If frmEmp.Visible = True Then
               datEmp.Find "Staff_code ='" & Trim(frmEmp.txtField(7)) & "'", 0, adSearchForward, 1
            End If
            pnlStaff_code = Trim(!staff_code)
            pnlLast_name = Trim(!last_name)
            If Not IsNull(!mid_name) Then
                pnlother_names = Trim(!first_name) & ", " & Trim(!mid_name)
            Else
                pnlother_names = Trim(!first_name)
            End If
        Else
            'cmdModify.Enabled = False
        End If
    End With
    
    Set datApp = New Recordset
    datApp.Open "SELECT * FROM empAppraisals WHERE co_code ='" & ID & _
    "'", cn, adOpenStatic, adLockOptimistic
    reset True
    cmdEdit.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    pApp = False
End Sub

Private Sub txtApp_Change(Index As Integer)
    If Index = 0 Then
        If txtApp(0) <> "" Then
            Set datExtra = New Recordset
            datExtra.Open "SELECT * FROM AppParams WHERE Param_Code =" & _
            txtApp(0) & " AND co_code ='" & ID & "'", cn, adOpenStatic, adLockOptimistic
            txtApp(1) = datExtra!param_name
        Else
            txtApp(1) = ""
        End If
    End If
End Sub
