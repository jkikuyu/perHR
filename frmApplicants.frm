VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmApplicants 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Applicants"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6795
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmApplicants.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   6795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdApp 
      Caption         =   "Applicants Qualifications"
      Height          =   375
      Left            =   30
      TabIndex        =   48
      Top             =   4860
      Width           =   6720
   End
   Begin VB.Frame Frame1 
      Caption         =   "Applicant's Interests"
      Height          =   975
      Left            =   30
      TabIndex        =   39
      Top             =   2820
      Width           =   6735
      Begin VB.TextBox txtapp 
         DataField       =   "desg"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   14
         Left            =   4845
         TabIndex        =   47
         Text            =   "Text4"
         Top             =   330
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.TextBox txtapp 
         DataField       =   "dept"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   13
         Left            =   1785
         TabIndex        =   46
         Text            =   "Text3"
         Top             =   360
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.CommandButton cmdDesg 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   5745
         Picture         =   "frmApplicants.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   600
         Width           =   270
      End
      Begin VB.CommandButton cmdDept 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2685
         Picture         =   "frmApplicants.frx":053C
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   615
         Width           =   270
      End
      Begin VB.TextBox txtapp 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   16
         Left            =   3180
         TabIndex        =   43
         Text            =   "Text2"
         Top             =   600
         Width           =   2535
      End
      Begin VB.TextBox txtapp 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   15
         Left            =   135
         TabIndex        =   42
         Text            =   "Text1"
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label Label13 
         Caption         =   "Job Title"
         Height          =   255
         Left            =   3180
         TabIndex        =   41
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label12 
         Caption         =   "Department"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   360
         Width           =   1335
      End
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
      Left            =   5925
      Picture         =   "frmApplicants.frx":0636
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   60
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
      Left            =   5445
      Picture         =   "frmApplicants.frx":0780
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   60
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
      Left            =   4965
      Picture         =   "frmApplicants.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   60
      Width           =   375
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
      Left            =   6405
      Picture         =   "frmApplicants.frx":0A14
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   60
      Width           =   375
   End
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   3375
      Top             =   150
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdMarital 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4050
      Picture         =   "frmApplicants.frx":0B5E
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   2460
      Width           =   285
   End
   Begin VB.CommandButton cmdSex 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4050
      Picture         =   "frmApplicants.frx":0C58
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   1740
      Width           =   285
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
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   6795
      TabIndex        =   26
      Top             =   5370
      Width           =   6795
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
         Height          =   350
         Left            =   5400
         TabIndex        =   32
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   4320
         TabIndex        =   31
         Top             =   0
         Width           =   975
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
         Height          =   350
         Left            =   3240
         TabIndex        =   30
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Sa&ve"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   2160
         TabIndex        =   29
         Top             =   0
         Width           =   975
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
         Height          =   350
         Left            =   1080
         TabIndex        =   28
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   0
         TabIndex        =   27
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdPath 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6450
      TabIndex        =   25
      Top             =   4485
      Width           =   285
   End
   Begin VB.TextBox txtapp 
      DataField       =   "app_cvpath"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   10
      Left            =   45
      TabIndex        =   24
      Text            =   "Text13"
      Top             =   4485
      Width           =   6360
   End
   Begin VB.TextBox txtapp 
      DataField       =   "app_exp"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   9
      Left            =   1695
      MaxLength       =   2
      TabIndex        =   23
      Text            =   "Text12"
      Top             =   3855
      Width           =   975
   End
   Begin VB.TextBox txtapp 
      DataField       =   "app_address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Index           =   8
      Left            =   4425
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   22
      Text            =   "frmApplicants.frx":0D52
      Top             =   1710
      Width           =   2325
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   300
      Left            =   1560
      TabIndex        =   21
      Top             =   2445
      Width           =   270
      _ExtentX        =   476
      _ExtentY        =   529
      _Version        =   393216
      Format          =   24576001
      CurrentDate     =   36873
   End
   Begin VB.TextBox txtapp 
      DataField       =   "app_marital"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   7
      Left            =   3630
      TabIndex        =   20
      Text            =   "Text10"
      Top             =   2220
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtapp 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   12
      Left            =   2070
      TabIndex        =   19
      Text            =   "Text9"
      Top             =   2460
      Width           =   1935
   End
   Begin VB.TextBox txtapp 
      DataField       =   "app_b_date"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MMM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   6
      Left            =   30
      TabIndex        =   18
      Text            =   "Text8"
      Top             =   2460
      Width           =   1530
   End
   Begin VB.TextBox txtapp 
      DataField       =   "app_id"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   30
      TabIndex        =   17
      Text            =   "Text7"
      Top             =   1740
      Width           =   1935
   End
   Begin VB.TextBox txtapp 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   11
      Left            =   2070
      TabIndex        =   16
      Text            =   "Text6"
      Top             =   1740
      Width           =   1935
   End
   Begin VB.TextBox txtapp 
      DataField       =   "app_sex"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   3150
      TabIndex        =   15
      Text            =   "Text5"
      Top             =   1500
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtapp 
      DataField       =   "app_f_name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   3750
      MaxLength       =   12
      TabIndex        =   14
      Text            =   "Text4"
      Top             =   1020
      Width           =   2535
   End
   Begin VB.TextBox txtapp 
      DataField       =   "app_m_name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   2070
      MaxLength       =   7
      TabIndex        =   13
      Text            =   "Text3"
      Top             =   1020
      Width           =   1575
   End
   Begin VB.TextBox txtapp 
      DataField       =   "app_s_name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   30
      MaxLength       =   12
      TabIndex        =   12
      Text            =   "Text2"
      Top             =   1020
      Width           =   1935
   End
   Begin VB.TextBox txtapp 
      DataField       =   "app_no"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   30
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   300
      Width           =   1575
   End
   Begin VB.Label Label11 
      Caption         =   "CV Path"
      Height          =   225
      Left            =   45
      TabIndex        =   10
      Top             =   4245
      Width           =   1680
   End
   Begin VB.Label Label10 
      Caption         =   "Experience(Years)"
      Height          =   225
      Left            =   60
      TabIndex        =   9
      Top             =   3870
      Width           =   1650
   End
   Begin VB.Label Label9 
      Caption         =   "Marital Status"
      Height          =   240
      Left            =   2070
      TabIndex        =   8
      Top             =   2220
      Width           =   1710
   End
   Begin VB.Label Label8 
      Caption         =   "Sex"
      Height          =   315
      Left            =   2070
      TabIndex        =   7
      Top             =   1500
      Width           =   705
   End
   Begin VB.Label Label7 
      Caption         =   "Address"
      Height          =   225
      Left            =   4425
      TabIndex        =   6
      Top             =   1470
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "Birth Date"
      Height          =   255
      Left            =   30
      TabIndex        =   5
      Top             =   2220
      Width           =   1395
   End
   Begin VB.Label Label5 
      Caption         =   "ID Number"
      Height          =   315
      Left            =   30
      TabIndex        =   4
      Top             =   1500
      Width           =   1875
   End
   Begin VB.Label Label4 
      Caption         =   "First Name"
      Height          =   240
      Left            =   3750
      TabIndex        =   3
      Top             =   780
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Middle Name"
      Height          =   285
      Left            =   2070
      TabIndex        =   2
      Top             =   780
      Width           =   1545
   End
   Begin VB.Label Label2 
      Caption         =   "Last Name"
      Height          =   270
      Left            =   30
      TabIndex        =   1
      Top             =   780
      Width           =   1320
   End
   Begin VB.Label Label1 
      Caption         =   "Applicant Number"
      Height          =   270
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Width           =   1605
   End
End
Attribute VB_Name = "frmApplicants"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents datApp As Recordset
Attribute datApp.VB_VarHelpID = -1
Dim datQlf As Recordset
Dim datqf As Recordset

Private Sub cmdAdd_Click()
    'Calculate the new app_no
    Set datExtra = New Recordset
    datExtra.Open "SELECT * FROM Applicants WHERE co_code ='" & _
    ID & "' ORDER BY app_no", cn, adOpenStatic, adLockOptimistic
    If datExtra.RecordCount <> 0 Then
       datExtra.MoveLast
       x = datExtra!app_no + 1
    Else
        x = 1
    End If
    datApp.AddNew
    txtapp(0) = x
    txtapp(6) = Date
    Reset False
    
End Sub

Private Sub cmdApp_Click()
    If datApp.RecordCount = 0 Then
       MsgBox "No Applicant Selected", vbInformation, "Qualifications"
       Exit Sub
    End If
    frmAppqf.Left = Screen.Width - frmAppqf.Width
    frmAppqf.Top = 100
    frmAppqf.Show vbModal
    
End Sub

Private Sub cmdCancel_Click()
    datApp.CancelUpdate
    Reset True
    
End Sub

Private Sub cmdClose_Click()
    Unload Me
    
End Sub

Private Sub cmdDel_Click()
    datQlf.Delete
End Sub

Private Sub cmdDept_Click()
    DataSelection = 1
    frmSelection.Top = 4650
    frmSelection.Left = 2500
    frmSelection.Show vbModal
End Sub

Private Sub cmdDesg_Click()
    DataSelection = 3
    frmSelection.Top = 2850
    frmSelection.Left = 2500
    frmSelection.Show vbModal
End Sub

Private Sub cmdEdit_Click()
    If datApp.RecordCount = 0 Then
       MsgBox ("No record to Edit"), vbInformation, "Editing aborted"
       Exit Sub
    End If
    Reset False
End Sub

Private Sub cmdfirst_Click()
On Error GoTo FirstErr
    datApp.MoveFirst
    Exit Sub
FirstErr:
MsgBox (Err.Description), vbInformation
Err.Clear
End Sub

Private Sub cmdlast_Click()
On Error GoTo FirstErr
    datApp.MoveLast
    Exit Sub
FirstErr:
MsgBox (Err.Description), vbInformation
Err.Clear
End Sub

Private Sub cmdMarital_Click()
    SexMar = "Mar"
    frmSexMar.Left = Me.Left + txtapp(4).Width + txtapp(7).Width - 50
    frmSexMar.Top = Me.Top + ((Label7.Top - Label1.Top)) * 2
    frmSexMar.Show vbModal
End Sub

Private Sub cmdNew_Click()
    datQlf.AddNew
End Sub

Private Sub cmdNext_Click()
On Error GoTo NextErr
    With datApp
        DbMove = True
        .MoveNext
         If .EOF Then
            .MoveLast
            MsgBox ("End of file"), vbInformation
         End If
    End With
    Exit Sub
NextErr:
MsgBox (Err.Description), vbInformation
Err.Clear
End Sub

Private Sub cmdNot_Click()
    datQlf.CancelUpdate
End Sub

Private Sub cmdPath_Click()
    cdlg.DialogTitle = "Select the Applicants CV"
    cdlg.InitDir = App.Path
    cdlg.Flags = &H1000
    cdlg.Filter = "Word Documents(*.doc)|*.doc|Rich Text Files(*.rtf)|*.rtf|"
    cdlg.ShowOpen
    If cdlg.FileName <> "" Then
       txtapp(10) = cdlg.FileName
    End If
    
End Sub

Private Sub cmdPrev_Click()
On Error GoTo PrevErr
    With datApp
         .MovePrevious
         If .BOF Then
            .MoveFirst
            MsgBox ("Beginning of file"), vbInformation
         End If
    End With
    Exit Sub
PrevErr:
MsgBox (Err.Description), vbInformation
Err.Clear
End Sub

Private Sub cmdqf_Click()

End Sub

Private Sub cmdSave_Click()
    'validate the entries made
    If txtapp(1) = "" Or txtapp(3) = "" Or txtapp(4) = "" Or txtapp(5) = "" Or txtapp(6) = "" Or txtapp(7) = "" Or txtapp(8) = "" Then
       MsgBox ("The following information must be supplied:" & Chr(13) & _
       "Last Name" & Chr(13) & _
       "First Name" & Chr(13) & _
       "ID Number" & Chr(13) & _
       "Sex" & Chr(13) & _
       "Date of Birth" & Chr(13) & _
       "Marital Status" & Chr(13) & _
       "Address"), vbInformation
       Exit Sub
    End If
    datApp!co_code = ID
    datApp.Update
    Reset True
    
End Sub

Private Sub cmdSex_Click()
    SexMar = "Sex"
    frmSexMar.Left = Me.Left + txtapp(4).Width + txtapp(7).Width - 50
    frmSexMar.Top = Me.Top + (((Label7.Top - Label1.Top) - 350) * 2) - 50
    frmSexMar.Show vbModal
End Sub

Private Sub cmdUpdate_Click()
    If IsNull(dtqf.Value) Then
       MsgBox ("Give a date when the qualification was attained."), vbInformation
       Exit Sub
    End If
    datQlf!co_code = ID
    datQlf!app_code = datApp!app_no
    datQlf.Update
End Sub

Private Sub Command1_Click()

End Sub

Private Sub datApp_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error Resume Next
    datQlf.Filter = "app_code =" & datApp!app_no
End Sub

Private Sub DTPicker2_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)

End Sub

Private Sub dtqf_Change()
    'txtqf(2) = Format(dtqf.Value, "dd/mmm/yyyy")
    
End Sub

Private Sub DTPicker1_CloseUp()
    txtapp(6) = Format(DTPicker1.Value, "dd/mmm/yyyy")
End Sub

Private Sub Form_Load()
    Appt = True
    Dim txt As TextBox
    Set datApp = New Recordset
    datApp.Open "SELECT * FROM applicants WHERE co_code ='" & ID & "'", cn, adOpenStatic, adLockOptimistic
    
    For Each txt In Me.txtapp
        Set txt.DataSource = datApp
    Next
    
    
    'If datApp.RecordCount <> 0 Then
    '   datqlf.Filter = "app_code =" & datApp!app_no
    'End If
    Reset True
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Appt = False
End Sub


Private Sub txtApp_Change(Index As Integer)
    If Index = 5 Then
        If Trim(txtapp(5)) <> "" And Not IsNull(Trim(txtapp(5))) Then
            Set datExtra = New Recordset
             datExtra.Open "SELECT * FROM UDFParameters WHERE ParamCode = 'Sex' " & _
             " AND code = " & txtapp(5) & _
             " AND co_code ='" & ID & "'", cn, adOpenStatic, adLockOptimistic
             If datExtra.RecordCount <> 0 Then
                txtapp(11) = datExtra!ParamName
             Else
                txtapp(11) = ""
             End If
        Else
            txtapp(11) = ""
        End If
    ElseIf Index = 7 Then
        If Trim(txtapp(7)) <> "" And Not IsNull(Trim(txtapp(7))) Then
            Set datExtra = New Recordset
             datExtra.Open "SELECT * FROM UDFParameters WHERE ParamCode = 'Mar' " & _
             " AND code = " & txtapp(7) & _
             " AND co_code ='" & ID & "'", cn, adOpenStatic, adLockOptimistic
             If datExtra.RecordCount <> 0 Then
                txtapp(12) = datExtra!ParamName
             Else
                txtapp(12) = ""
             End If
        Else
            txtapp(12) = ""
        End If
    ElseIf Index = 13 Then
            If Trim(txtapp(13)) <> "" Then
                Set datExtra = New Recordset
                datExtra.Open "SELECT * FROM UDFParameters WHERE Co_Code ='" & ID & "'" & _
                " AND ParamCode ='Dpt' AND Code =" & _
                CLng(txtapp(Index)), cn, adOpenStatic, adLockOptimistic
                
                If datExtra.RecordCount <> 0 Then
                    txtapp(15) = datExtra!ParamName
                Else
                    txtapp(15) = ""
                End If
            Else
                txtapp(15) = ""
            End If
    ElseIf Index = 14 Then
            If Trim(txtapp(14)) <> "" Then
                Set datExtra = New Recordset
                datExtra.Open "SELECT * FROM UDFParameters WHERE Co_Code ='" & ID & "'" & _
                " AND ParamCode ='Des' AND Code =" & _
                CLng(txtapp(Index)), cn, adOpenStatic, adLockOptimistic
                'txtapp(2) = datExtra!ParamName
                
                If datExtra.RecordCount <> 0 Then
                    txtapp(16) = datExtra!ParamName
                Else
                    txtapp(16) = ""
                End If
            Else
            txtapp(16) = ""
            End If
    End If
            
End Sub

Private Sub txtApp_KeyPress(Index As Integer, KeyAscii As Integer)
    Char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(Char))
End Sub

Public Sub Reset(bval As Boolean)
    'Procedure to enable and disable buttons and text boxes as required
    Dim txt As TextBox
    cmdAdd.Enabled = bval
    cmdEdit.Enabled = bval
    cmdSearch.Enabled = bval
    cmdSave.Enabled = Not bval
    cmdCancel.Enabled = Not bval
    cmdSex.Enabled = Not bval
    cmdMarital.Enabled = Not bval
    cmdPath.Enabled = Not bval
    DTPicker1.Enabled = Not bval
    cmdDept.Enabled = Not bval
    cmdDesg.Enabled = Not bval
    cmdApp.Enabled = bval
    cmdFirst.Enabled = bval
    cmdLast.Enabled = bval
    cmdNext.Enabled = bval
    cmdPrev.Enabled = bval
    
    For Each txt In Me.txtapp
        txt.Locked = bval
    Next
    
    'Override the above configurations
    txtapp(0).Locked = True
    txtapp(10).Locked = True
    txtapp(11).Locked = True
    txtapp(12).Locked = True
    txtapp(6).Locked = True
End Sub
