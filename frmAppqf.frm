VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmAppqf 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Applicants Qualifications"
   ClientHeight    =   2805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5895
   Icon            =   "frmAppqf.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Applicant's Qualifications"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2790
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5895
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4815
         TabIndex        =   17
         Top             =   1935
         Width           =   1005
      End
      Begin VB.TextBox txtqf 
         Height          =   300
         Index           =   1
         Left            =   7620
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   2115
         Width           =   2595
      End
      Begin VB.TextBox txtqf 
         DataField       =   "level"
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
         Index           =   3
         Left            =   150
         TabIndex        =   11
         Text            =   "Text2"
         Top             =   1905
         Width           =   4080
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "&New"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4800
         TabIndex        =   10
         Top             =   210
         Width           =   1005
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4800
         TabIndex        =   9
         Top             =   645
         Width           =   1005
      End
      Begin VB.CommandButton cmdNot 
         Caption         =   "Ca&ncel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4815
         TabIndex        =   8
         Top             =   1095
         Width           =   1005
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4815
         TabIndex        =   7
         Top             =   1515
         Width           =   1005
      End
      Begin VB.CommandButton qfNext 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3480
         Picture         =   "frmAppqf.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2340
         Width           =   360
      End
      Begin VB.CommandButton qfPrev 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3090
         Picture         =   "frmAppqf.frx":058C
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2340
         Width           =   360
      End
      Begin VB.CommandButton qfFirst 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2700
         Picture         =   "frmAppqf.frx":06D6
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2340
         Width           =   360
      End
      Begin VB.CommandButton qfLast 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3870
         Picture         =   "frmAppqf.frx":0820
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2340
         Width           =   360
      End
      Begin VB.TextBox txtqf 
         DataField       =   "qf_code"
         Height          =   285
         Index           =   0
         Left            =   7605
         TabIndex        =   2
         Text            =   "Text4"
         Top             =   1770
         Visible         =   0   'False
         Width           =   405
      End
      Begin MSComCtl2.DTPicker dtqf 
         DataField       =   "qf_date"
         Height          =   300
         Left            =   120
         TabIndex        =   1
         Top             =   1200
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CheckBox        =   -1  'True
         Format          =   24444929
         CurrentDate     =   36526
         MinDate         =   2
      End
      Begin MSDataListLib.DataCombo dbf 
         DataField       =   "qf_code"
         Height          =   345
         Left            =   90
         TabIndex        =   13
         Top             =   555
         Width           =   4140
         _ExtentX        =   7303
         _ExtentY        =   609
         _Version        =   393216
         Style           =   2
         ListField       =   "qualification"
         BoundColumn     =   "qf_code"
         Text            =   ""
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
      Begin VB.Label Label14 
         Caption         =   "Qualification"
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
         Left            =   135
         TabIndex        =   16
         Top             =   315
         Width           =   1230
      End
      Begin VB.Label Label15 
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
         Height          =   180
         Left            =   135
         TabIndex        =   15
         Top             =   990
         Width           =   1320
      End
      Begin VB.Label Label16 
         Caption         =   "Level Attained"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   150
         TabIndex        =   14
         Top             =   1680
         Width           =   1500
      End
   End
End
Attribute VB_Name = "frmAppqf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim datqlf As Recordset
Dim datqf As Recordset

Private Sub cmdClose_Click()
    Unload Me
    
End Sub

Private Sub cmdDel_Click()
    datqlf.Delete
    datqlf.Requery
End Sub

Private Sub cmdNew_Click()
    datqlf.AddNew
    dtqf.Value = Date
    Reset False
End Sub

Private Sub cmdNot_Click()
    datqlf.CancelUpdate
    Reset True
End Sub

Private Sub cmdUpdate_Click()
    datqlf!co_code = ID
    datqlf!app_code = frmApplicants.txtApp(0)
    datqlf.Update
    Reset True
End Sub

Private Sub qfFirst_Click()
On Error GoTo FirstErr
    datqlf.MoveFirst
    Exit Sub
FirstErr:
MsgBox (Err.Description), vbInformation, "Applicants Qualifications"
Err.Clear
End Sub

Private Sub qfLast_Click()
On Error GoTo FirstErr
    datqlf.MoveLast
    Exit Sub
FirstErr:
MsgBox (Err.Description), vbInformation, "Applicants Qualifications"
Err.Clear
End Sub

Private Sub qfNext_Click()
On Error GoTo NextErr
    With datqlf
        DbMove = True
        .MoveNext
         If .EOF Then
            .MoveLast
            MsgBox ("End of file"), vbInformation, "Applicants Qualifications"
         End If
    End With
    Exit Sub
NextErr:
MsgBox (Err.Description), vbInformation, "Applicants Qualifications"
Err.Clear
End Sub

Private Sub qfPrev_Click()
On Error GoTo PrevErr
    With datqlf
         .MovePrevious
         If .BOF Then
            .MoveFirst
            MsgBox ("Beginning of file"), vbInformation, "Applicants Qualifications"
         End If
    End With
    Exit Sub
PrevErr:
MsgBox (Err.Description), vbInformation, "Applicants Qualifications"
Err.Clear
End Sub

Private Sub Form_Load()
    Set datqf = New Recordset
    datqf.Open "SELECT * FROM Qualifications", cn, adOpenStatic, adLockOptimistic
    
    Set dbf.RowSource = datqf
    
    Set datqlf = New Recordset
    datqlf.Open "SELECT * FROM appQualifications WHERE co_code ='" & ID & _
    "' AND app_Code ='" & frmApplicants.txtApp(0) & "'", cn, adOpenStatic, adLockOptimistic
    Set dbf.DataSource = datqlf
    For Each txt In Me.txtqf
        Set txt.DataSource = datqlf
    Next
    Set dtqf.DataSource = datqlf
    Reset True
End Sub

Public Sub Reset(bval As Boolean)
    cmdNew.Enabled = bval
    cmdUpdate.Enabled = Not bval
    cmdNot.Enabled = Not bval
    cmdDel.Enabled = bval
    
    dbf.Enabled = Not bval
    dtqf.Enabled = Not bval
    txtqf(3).Enabled = Not bval
    
End Sub
