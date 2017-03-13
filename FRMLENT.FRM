VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmlent 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Leave Entitlement"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5970
   Icon            =   "frmlent.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   5970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3585
      Left            =   0
      TabIndex        =   16
      Top             =   705
      Width           =   5955
      Begin VB.TextBox txtleave 
         DataField       =   "leave_dur"
         Height          =   285
         Index           =   2
         Left            =   1620
         MaxLength       =   2
         TabIndex        =   24
         Top             =   1395
         Width           =   720
      End
      Begin VB.TextBox txtleave 
         DataField       =   "notes"
         Height          =   990
         Index           =   3
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   23
         Top             =   2010
         Width           =   5805
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
         Picture         =   "frmlent.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   3120
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
         Picture         =   "frmlent.frx":058C
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   3120
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
         Picture         =   "frmlent.frx":06D6
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   3120
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
         Picture         =   "frmlent.frx":0820
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   3120
         Width           =   375
      End
      Begin VB.ComboBox cboYear 
         DataField       =   "year"
         Enabled         =   0   'False
         Height          =   315
         Left            =   2130
         TabIndex        =   18
         Top             =   855
         Width           =   1680
      End
      Begin MSDataListLib.DataCombo cboleave 
         DataField       =   "leave_code"
         Height          =   315
         Left            =   90
         TabIndex        =   17
         Top             =   465
         Width           =   5820
         _ExtentX        =   10266
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "leave_name"
         BoundColumn     =   "leave_code"
         Text            =   ""
      End
      Begin VB.Label Label1 
         Caption         =   "Leave Employee is Entitled to:"
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
         Left            =   105
         TabIndex        =   29
         Top             =   225
         Width           =   3495
      End
      Begin VB.Label Label9 
         Caption         =   "working days"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2445
         TabIndex        =   28
         Top             =   1410
         Width           =   1200
      End
      Begin VB.Label Label5 
         Caption         =   "Entitled Duration"
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
         Left            =   105
         TabIndex        =   27
         Top             =   1395
         Width           =   1470
      End
      Begin VB.Label Label4 
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
         Height          =   255
         Left            =   105
         TabIndex        =   26
         Top             =   1770
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Entitled Starting Year"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   105
         TabIndex        =   25
         Top             =   870
         Width           =   1845
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   5970
      TabIndex        =   10
      Top             =   4380
      Width           =   5970
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   350
         Left            =   3975
         TabIndex        =   30
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   350
         Left            =   1005
         TabIndex        =   15
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   350
         Left            =   1995
         TabIndex        =   14
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   350
         Left            =   2985
         TabIndex        =   13
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "C&lose"
         Height          =   350
         Left            =   4980
         TabIndex        =   12
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   350
         Left            =   15
         TabIndex        =   11
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.TextBox txtleave 
      Height          =   285
      Index           =   1
      Left            =   7440
      TabIndex        =   3
      Top             =   840
      Width           =   345
   End
   Begin VB.CommandButton cmdFind 
      Height          =   375
      Left            =   5595
      Picture         =   "frmlent.frx":096A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   270
      Width           =   375
   End
   Begin VB.CommandButton cmdLeave 
      Height          =   285
      Left            =   7830
      Picture         =   "frmlent.frx":0AEC
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   840
      Width           =   285
   End
   Begin VB.TextBox txtleave 
      DataField       =   "leave_code"
      Height          =   285
      Index           =   0
      Left            =   7500
      TabIndex        =   0
      Top             =   570
      Visible         =   0   'False
      Width           =   195
   End
   Begin Threed.SSPanel pnlother_names 
      Height          =   375
      Left            =   2910
      TabIndex        =   4
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
      TabIndex        =   5
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
      TabIndex        =   6
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
      TabIndex        =   9
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
      TabIndex        =   8
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
      TabIndex        =   7
      Top             =   0
      Width           =   2415
   End
End
Attribute VB_Name = "frmlent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim datLeave As Recordset
Dim datLeaveTypes As Recordset
Dim LAdd As Boolean
Dim ecode As Long


Private Sub cmdAdd_Click()
    'Calculate the new entry code
    Set datExtra = New Recordset
    datExtra.Open "SELECT TOP 1 * FROM Leave_entitlement ORDER BY e_code DESC", cn, adOpenStatic, adLockOptimistic
    If datExtra.RecordCount = 0 Then
       ecode = 1
    Else
       ecode = datExtra!e_code + 1
    End If
    If Trim(pnlStaff_code) = "" Then
       MsgBox ("No employee selected."), vbInformation, "Leave Entitlements"
       Exit Sub
    End If
    datLeave.AddNew
    cboYear = Operation_Year
    Reset False
    LAdd = True
End Sub

Private Sub cmdCancel_Click()
    datLeave.CancelUpdate
    Reset True
    LAdd = False
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    If datLeave.RecordCount = 0 Then
       MsgBox "No record to delete", vbInformation, "Leave Entitlements"
       Exit Sub
    End If
    If MsgBox("Delete this leave entitlement." & Chr(13) & _
        "Please note that deleting a leave entitlement may affect the approval of any leave requests already made." & Chr(13) & _
        "Do you still want to delete this entitlement?", vbYesNo + vbQuestion, "Leave Entitlements") = vbYes Then
       datLeave.Delete
       datLeave.Requery
    Else
       MsgBox ("Deletion aborted."), vbInformation, "Leave Entitlements"
    End If
End Sub

Private Sub cmdEdit_Click()
    If datLeave.RecordCount = 0 Then
       MsgBox ("No Record to Edit"), vbInformation, "Leave Entitlements"
       Exit Sub
    End If
    Reset False
    
End Sub

Private Sub cmdFind_Click()
    frmSelectEmp.Left = Me.Left + 40
    frmSelectEmp.Top = Me.Top + pnlStaff_code.Top + pnlStaff_code.Height
    frmSelectEmp.Show vbModal
    
    If pnlStaff_code <> "" Then
       datLeave.Filter = "Staff_code ='" & Trim(pnlStaff_code) & "'"
    End If
End Sub

Private Sub cmdLeave_Click()
    frmSelLeave.Top = Me.Top + (Label4.Top - Label6.Top) + 100
    frmSelLeave.Left = Me.Left + 60
    frmSelLeave.Show vbModal
End Sub

Private Sub cmdSave_Click()
    'Check if leave is selected
    If cboleave.BoundText = "" Then
       MsgBox ("No valid leave type selected."), vbInformation, "Leave Entitlements"
       Exit Sub
    ElseIf Trim(txtleave(2)) = "" Then
       MsgBox ("Give the number of days employee is entitled to."), vbInformation, "Leave Entitlements"
       Exit Sub
    End If
    
    If LAdd = True Then
        Set datExtra = New Recordset
        datExtra.Open "SELECT * FROM leave_entitlement WHERE staff_Code ='" & pnlStaff_code & _
        "' AND co_code ='" & ID & "' AND leave_code =" & CLng(cboleave.BoundText), cn, adOpenStatic, adLockOptimistic
        If datExtra.RecordCount <> 0 Then
            MsgBox ("This employee has already been assigned this leave." & Chr(13) & _
            "Edit the leave record if you want to change the leave record."), vbInformation
            Exit Sub
        End If
    End If
    
    
    
    'Get the leave balances file and add this leave for this employee for the year
    Set datExtra = New Recordset
    datExtra.Open "SELECT * FROM leave_bal WHERE staff_code ='" & _
    pnlStaff_code & "' AND leave_code =" & CLng(cboleave.BoundText) & " AND " & _
    " co_code ='" & ID & "' AND Year =" & CLng(Operation_Year), cn, adOpenStatic, adLockOptimistic
    With datExtra
        If .RecordCount = 0 Then
           'Leave has never been put for this employee for this period
           .AddNew
           !co_code = ID
           
           !staff_code = Trim(pnlStaff_code)
           !Leave_code = CLng(cboleave.BoundText)
           !Year = CLng(cboYear)
           !bal_bf = 0
           !days_ent = CDbl(txtleave(2))
           !days_taken = 0
           !bal = CDbl(txtleave(2))
            .Update
        Else
            'Leave has been found, adjust the entitled days and the balance
            !days_ent = CDbl(txtleave(2))
            !bal = CDbl(!days_ent) + CDbl(!bal_bf)
            .Update
        End If
    End With
    datLeave!e_code = ecode
    datLeave!co_code = ID
    datLeave!staff_code = pnlStaff_code
    datLeave.Update
    Reset True
End Sub

Private Sub ffirst_Click()
On Error GoTo PrevErr
    With datLeave
         .MoveFirst
    End With
    Exit Sub
PrevErr:
MsgBox (Err.Description), vbInformation
Err.Clear
End Sub

Private Sub flast_Click()
On Error GoTo PrevErr
    With datLeave
         .MoveLast
    End With
    Exit Sub
PrevErr:
MsgBox (Err.Description), vbInformation
Err.Clear
End Sub

Private Sub fnext_Click()
On Error GoTo PrevErr
    With datLeave
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
    PLeave2 = True
    'Load the year combo box
    For y = 2000 To 2099
        cboYear.AddItem y
    Next

    
    Dim txt As TextBox
    Set datLeave = New Recordset
    datLeave.Open "SELECT * FROM leave_entitlement WHERE co_code ='" & ID & _
    "'", cn, adOpenStatic, adLockOptimistic
    
    For Each txt In Me.txtleave
        Set txt.DataSource = datLeave
    Next
    Set cboYear.DataSource = datLeave
    Set cboleave.DataSource = datLeave
    
    Set datLeaveTypes = New Recordset
    datLeaveTypes.Open "SELECT * FROM leave_mast WHERE co_code ='" & _
    ID & "'", cn, adOpenStatic, adLockOptimistic
    Set cboleave.RowSource = datLeaveTypes
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
            datLeave.Filter = "Staff_code ='" & Trim(pnlStaff_code) & "'"
        Else
            'cmdModify.Enabled = False
        End If
    End With
    Reset True
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PLeave2 = False
End Sub

Private Sub fprev_Click()
On Error GoTo PrevErr
    With datLeave
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

Private Sub txtleave_Change(Index As Integer)
    If Index = 0 Then
        If Trim(txtleave(0)) <> "" Then
            'Find this leave from the leaves master file
            Set datExtra = New Recordset
            datExtra.Open "SELECT * FROM leave_mast WHERE leave_code =" & _
            Trim(txtleave(0)), cn, adOpenStatic, adLockOptimistic
            If datExtra.RecordCount <> 0 Then
               txtleave(1) = datExtra!leave_name
            Else
                txtleave(1) = ""
            End If
        Else
            txtleave(1) = ""
        End If
    End If
End Sub

Public Sub Reset(bval As Boolean)
    cmdAdd.Enabled = bval
    cmdEdit.Enabled = bval
    cmdSave.Enabled = Not bval
    cmdCancel.Enabled = Not bval
    txtleave(1).Locked = bval
    cboleave.Enabled = Not bval
    cmdDelete.Enabled = bval
    cboYear.Enabled = Not bval
    cmdFind.Enabled = bval
    
    ffirst.Enabled = bval
    flast.Enabled = bval
    fnext.Enabled = bval
    fprev.Enabled = bval
End Sub
