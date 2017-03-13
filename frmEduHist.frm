VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmEduHist 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Education History"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5970
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEduHist.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   5970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGoTo 
      Caption         =   "Go To..."
      Height          =   360
      Left            =   45
      TabIndex        =   30
      Top             =   4860
      Width           =   960
   End
   Begin VB.ComboBox cboEnd 
      DataField       =   "ed_date"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   345
      Left            =   2850
      TabIndex        =   29
      Text            =   "cboEnd"
      Top             =   2925
      Width           =   2145
   End
   Begin VB.ComboBox cboStart 
      DataField       =   "st_date"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   345
      Left            =   90
      TabIndex        =   28
      Text            =   "cboStart"
      Top             =   2925
      Width           =   2175
   End
   Begin VB.TextBox txtEHist 
      DataField       =   "level"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   5
      Left            =   60
      TabIndex        =   26
      Top             =   2280
      Width           =   4935
   End
   Begin VB.TextBox txtEHist 
      DataField       =   "course"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   4
      Left            =   45
      TabIndex        =   25
      Top             =   1665
      Width           =   4935
   End
   Begin VB.TextBox txtEHist 
      DataField       =   "institution"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   30
      TabIndex        =   12
      Top             =   975
      Width           =   5850
   End
   Begin VB.TextBox txtEHist 
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
      Height          =   1170
      Index           =   3
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   3585
      Width           =   5820
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
      Picture         =   "frmEduHist.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   270
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
      Left            =   5505
      Picture         =   "frmEduHist.frx":05C4
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Last Record"
      Top             =   4875
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
      Left            =   4080
      Picture         =   "frmEduHist.frx":070E
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "First Record"
      Top             =   4875
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
      Left            =   4545
      Picture         =   "frmEduHist.frx":0858
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Previous Record"
      Top             =   4875
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
      Left            =   5025
      Picture         =   "frmEduHist.frx":09A2
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Next Record"
      Top             =   4875
      Width           =   375
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
      ScaleWidth      =   5970
      TabIndex        =   0
      Top             =   5355
      Width           =   5970
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   360
         Left            =   3915
         TabIndex        =   27
         Top             =   0
         Width           =   960
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "&New"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   960
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   975
         TabIndex        =   4
         Top             =   0
         Width           =   960
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1950
         TabIndex        =   3
         Top             =   0
         Width           =   960
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2925
         TabIndex        =   2
         Top             =   0
         Width           =   960
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "C&lose"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4950
         TabIndex        =   1
         Top             =   0
         Width           =   960
      End
   End
   Begin Threed.SSPanel pnlother_names 
      Height          =   375
      Left            =   2910
      TabIndex        =   13
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
      TabIndex        =   14
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
      TabIndex        =   15
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
   Begin VB.Label Label1 
      Caption         =   "Institution"
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
      Left            =   45
      TabIndex        =   24
      Top             =   765
      Width           =   1530
   End
   Begin VB.Label Label2 
      Caption         =   "Course"
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
      Left            =   60
      TabIndex        =   23
      Top             =   1440
      Width           =   1185
   End
   Begin VB.Label Label3 
      Caption         =   "Level"
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
      Left            =   75
      TabIndex        =   22
      Top             =   2070
      Width           =   1065
   End
   Begin VB.Label Label4 
      Caption         =   "From"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   90
      TabIndex        =   21
      Top             =   2655
      Width           =   945
   End
   Begin VB.Label Label5 
      Caption         =   "To"
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
      Left            =   2850
      TabIndex        =   20
      Top             =   2640
      Width           =   915
   End
   Begin VB.Label Label6 
      Caption         =   "Notes"
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
      Left            =   120
      TabIndex        =   19
      Top             =   3330
      Width           =   750
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
      TabIndex        =   17
      Top             =   0
      Width           =   2175
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
      TabIndex        =   16
      Top             =   0
      Width           =   2415
   End
End
Attribute VB_Name = "frmEduHist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim datHist As Recordset

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
      
    cboStart.Enabled = Not bval
    cboEnd.Enabled = Not bval
    
End Sub

Private Sub cmdCancel_Click()
    datHist.CancelUpdate
    cn.RollbackTrans
    Reset True
End Sub

Private Sub cmdDelete_Click()
    If datHist.RecordCount = 0 Then
       MsgBox "No Record to Delete", vbInformation, "Deletion aborted"
       Exit Sub
    End If
    If MsgBox("Delete this Education Record?", vbYesNo + vbQuestion, "Education History") = vbYes Then
       datHist.Delete
       datHist.Requery
    End If
End Sub

Private Sub cmdEdit_Click()
    If datHist.RecordCount = 0 Then
       MsgBox "No Record to Edit", vbInformation, "Editing Aborted"
       Exit Sub
    End If
    Reset False
    cn.BeginTrans
End Sub

Private Sub cmdGoTo_Click()
    If datEmp.RecordCount = 0 Then
       MsgBox ("File is empty."), vbInformation, "Find"
       Exit Sub
    End If
    Dim Response
    Response = InputBox("Go To Employee number...", "Go To")
    If Response <> "" Then
        If Len(Response) > 7 Then
           MsgBox "Invalid Staff Code", vbInformation, "Education History"
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

Private Sub cmdNew_Click()
    If Trim(pnlStaff_code) = "" Then
       MsgBox ("No Employee selected."), vbInformation, "Education History"
       Exit Sub
    End If
    cn.BeginTrans
    datHist.AddNew
    Reset False
End Sub

Private Sub cmdUpdate_Click()
    If Trim(txtEHist(0)) = "" Or Trim(txtEHist(4)) = "" Or Trim(txtEHist(5)) = "" Then
       MsgBox ("The following data are required for education history records: " & Chr(13) & _
       "           Educational Instituition Name" & Chr(13) & _
       "           Course Taken" & Chr(13) & _
       "           Level Attained"), vbInformation, "Education History"
       Exit Sub
    End If
    'Validate for the year
    If cboStart <> "" Then
       If Not IsNumeric(cboStart) Then
          MsgBox ("Invalid start year"), vbInformation
          cboStart.SetFocus
          Exit Sub
       End If
    End If
    If cboEnd <> "" Then
       If Not IsNumeric(cboEnd) Then
          MsgBox ("Invalid end year"), vbInformation
          cboEnd.SetFocus
          Exit Sub
       End If
    End If
    
    datHist!co_code = ID
    datHist!staff_code = pnlStaff_code
    
    datHist.Update
    cn.CommitTrans
    If datHist.Bookmark > 0 Then
        bk = datHist.Bookmark
        datHist.Bookmark = bk
    End If
    Reset True
End Sub

Private Sub dtEnd_CloseUp()
    txtEHist(2) = dtEnd.Value
End Sub

Private Sub dtStart_CloseUp()
    txtEHist(1) = dtStart.Value
End Sub

Private Sub ffirst_Click()
On Error GoTo PrevErr
    With datHist
         .MoveFirst
    End With
    Exit Sub
PrevErr:
MsgBox (Err.Description), vbInformation, "Education History"
Err.Clear
End Sub

Private Sub flast_Click()
On Error GoTo PrevErr
    With datHist
         .MoveLast
    End With
    Exit Sub
PrevErr:
MsgBox (Err.Description), vbInformation, "Education History"
Err.Clear
End Sub

Private Sub fnext_Click()
On Error GoTo PrevErr
    With datHist
         .MoveNext
         If .EOF Then
            .MoveLast
            MsgBox ("Last Record encountered"), vbInformation, "Education History"
         End If
    End With
    Exit Sub
PrevErr:
MsgBox (Err.Description), vbInformation, "Education History"

Err.Clear
End Sub

Private Sub Form_Load()
    'Add years to the from and to combo boxes
    For x = 1900 To 2099
        cboStart.AddItem x
    Next
    For y = 1900 To 2099
        cboEnd.AddItem y
    Next
    
    PDHist = True
    Set datHist = New Recordset
    datHist.Open "SELECT * FROM edu_history WHERE co_code ='" & ID & _
    "'", cn, adOpenStatic, adLockOptimistic
    For Each txt In Me.txtEHist
        Set txt.DataSource = datHist
    Next
    Set cboStart.DataSource = datHist
    Set cboEnd.DataSource = datHist
    
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
    PDHist = False
End Sub

Private Sub fprev_Click()
On Error GoTo PrevErr
    With datHist
         .MovePrevious
         If .BOF Then
            .MoveFirst
            MsgBox ("First Record encountered"), vbInformation, "Education History"
         End If
    End With
    Exit Sub
PrevErr:
MsgBox (Err.Description), vbInformation, "Education History"
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

Private Sub txtEHist_GotFocus(Index As Integer)
    txtEHist(Index) = Trim(txtEHist(Index))
    
End Sub
