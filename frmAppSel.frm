VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmAppSel 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Applicants' Selection Criteria"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5730
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   5730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   5040
      Left            =   0
      TabIndex        =   0
      Top             =   -90
      Width           =   5715
      Begin VB.ComboBox cboExp 
         Height          =   345
         ItemData        =   "frmAppSel.frx":0000
         Left            =   2025
         List            =   "frmAppSel.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1110
         Width           =   630
      End
      Begin VB.ComboBox cboAge 
         Height          =   345
         ItemData        =   "frmAppSel.frx":001C
         Left            =   165
         List            =   "frmAppSel.frx":0029
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1110
         Width           =   630
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   345
         Left            =   4260
         TabIndex        =   13
         Top             =   750
         Width           =   1305
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Height          =   345
         Left            =   4260
         TabIndex        =   12
         Top             =   330
         Width           =   1305
      End
      Begin MSComctlLib.ImageList img 
         Left            =   4680
         Top             =   1905
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAppSel.frx":0038
               Key             =   "qlf"
            EndProperty
         EndProperty
      End
      Begin MSDataListLib.DataCombo cboDesg 
         Height          =   345
         Left            =   165
         TabIndex        =   10
         Top             =   2580
         Width           =   3600
         _ExtentX        =   6350
         _ExtentY        =   609
         _Version        =   393216
         Style           =   2
         ListField       =   "paramname"
         BoundColumn     =   "code"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cboDept 
         Height          =   345
         Left            =   165
         TabIndex        =   9
         Top             =   1830
         Width           =   3600
         _ExtentX        =   6350
         _ExtentY        =   609
         _Version        =   393216
         Style           =   2
         ListField       =   "paramname"
         BoundColumn     =   "code"
         Text            =   ""
      End
      Begin VB.TextBox txtExp 
         Height          =   330
         Left            =   2640
         TabIndex        =   8
         Top             =   1110
         Width           =   975
      End
      Begin VB.TextBox txtAge 
         Height          =   330
         Left            =   780
         TabIndex        =   7
         Top             =   1110
         Width           =   945
      End
      Begin VB.Frame Frame2 
         Caption         =   "Qualification Selection"
         Height          =   1995
         Left            =   135
         TabIndex        =   6
         Top             =   2940
         Width           =   5490
         Begin MSComctlLib.ListView lstApp 
            Height          =   1680
            Left            =   60
            TabIndex        =   11
            Top             =   225
            Width           =   5370
            _ExtentX        =   9472
            _ExtentY        =   2963
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            Icons           =   "img"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Qlf Code"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Qualifications"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.Label Label5 
         Caption         =   "Job of Interest"
         Height          =   345
         Left            =   165
         TabIndex        =   5
         Top             =   2310
         Width           =   2565
      End
      Begin VB.Label Label4 
         Caption         =   "Department of Interest"
         Height          =   270
         Left            =   165
         TabIndex        =   4
         Top             =   1605
         Width           =   2295
      End
      Begin VB.Label Label3 
         Caption         =   "Experience(Years)"
         Height          =   285
         Left            =   2070
         TabIndex        =   3
         Top             =   870
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Applicants Age(Yrs)"
         Height          =   300
         Left            =   165
         TabIndex        =   2
         Top             =   855
         Width           =   1725
      End
      Begin VB.Label Label1 
         Caption         =   "Make your selection criteria and print a report showing applicants that match the given criteria"
         Height          =   675
         Left            =   150
         TabIndex        =   1
         Top             =   225
         Width           =   4185
      End
   End
End
Attribute VB_Name = "frmAppSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim datQlf As Recordset
Dim SelF As String  'Selection formula
Dim datDept As Recordset
Dim datDesg As Recordset


Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Dim fr As New frmPayslips
    Dim myfile As String
    myfile = Dir(App.Path + "\Reports\applicant selections.rpt")
    If myfile = "" Then
       MsgBox ("Cannot find the Report File required."), vbInformation
       Exit Sub
    End If
    fr.rpt1.ReportFileName = App.Path + "\Reports\applicant selections.rpt"
    
    'Check for any age specification
    If txtAge <> "" Then
       SelF = "{@Age}" & cboAge & txtAge
    End If
    'Check for any experience selection
    If txtExp <> "" Then
       If SelF = "" Then
          SelF = "{applicants.app_exp}" & cboExp & txtExp
       Else
          SelF = SelF & " AND {applicants.app_exp}" & cboExp & txtExp
       End If
    End If
    'Check for Dept selection
    If cboDept.BoundText <> "" Then
       If SelF = "" Then
          SelF = "{applicants.dept}=" & cboDept.BoundText
       Else: SelF = SelF & " AND {applicants.dept}=" & CLng(cboDept.BoundText)
       End If
    End If
    
    'Check for designation selection
    If cboDesg.BoundText <> "" Then
       If SelF = "" Then
          SelF = "{applicants.desg}=" & cboDesg.BoundText
       Else: SelF = SelF & " AND {applicants.desg}=" & CLng(cboDesg.BoundText)
       End If
    End If
    'Now get the Qualifications selection
    Dim ItemCount As Long
    Dim AppSel As String
    Dim CommaFound
    For ItemCount = 1 To lstApp.ListItems.count
        Set lstApp.SelectedItem = lstApp.ListItems(ItemCount)
        lstApp.SetFocus
        If lstApp.SelectedItem.CHECKED = True Then
           If AppSel = "" Then
              AppSel = CLng(lstApp.SelectedItem.Text)
           Else
              AppSel = AppSel & "," & CLng(lstApp.SelectedItem.Text)
           End If
        End If
    Next ItemCount
    If AppSel <> "" Then
       CommaFound = InStr(1, AppSel, ",")
       If CommaFound Then
            If SelF = "" Then
               SelF = "{appqualifications.qf_code} in [" & AppSel & "]"
            Else
                SelF = SelF & " AND {appqualifications.qf_code} in [" & AppSel & "]"
            End If
       Else
            If SelF = "" Then
               SelF = "{appqualifications.qf_code} in [" & AppSel & "]"
            Else
                SelF = SelF & " AND {appqualifications.qf_code} = [" & AppSel & "]"
            End If
       End If
    End If
    
    
    If SelF <> "" Then
       fr.rpt1.SelectionFormula = SelF
    End If
    fr.rpt1.WindowParentHandle = fr.hwnd
    fr.rpt1.Action = 1
    SelF = ""
    fr.Caption = myfile
    
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
    'Get the qualifications
    Dim nCount As Long
    Set datQlf = New Recordset
    With datQlf
        .Open "SELECT * FROM qualifications", cn, adOpenStatic, adLockOptimistic
        nCount = 1
        While Not .EOF
            lstApp.ListItems.Add , , !qf_code, qlf, qlf
            lstApp.ListItems(nCount).ListSubItems.Add , , Trim(!Qualification), qlf
            .MoveNext
            nCount = nCount + 1
        Wend
    End With
    
    'Get the departments
    'Get the row sources of the combo boxes
    Set datDept = New Recordset
    datDept.Open "SELECT * FROM UDFParameters WHERE Co_Code ='" & ID & "'" & _
    " AND ParamCode ='Dpt'", cn, adOpenStatic, adLockOptimistic
    Set cboDept.RowSource = datDept
    
    Set datDesg = New Recordset
    datDesg.Open "SELECT * FROM UDFParameters WHERE Co_Code ='" & ID & "'" & _
    " AND ParamCode ='Des'", cn, adOpenStatic, adLockOptimistic
    Set cboDesg.RowSource = datDesg
    
    'Load operands
    cboAge.Text = cboAge.List(0)
    cboExp.Text = cboExp.List(0)
End Sub
