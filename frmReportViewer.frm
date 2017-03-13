VERSION 5.00
Begin VB.Form frmReportViewer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Report Viewer"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5925
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmReportViewer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Select a Report to Print"
      Height          =   2670
      Left            =   45
      TabIndex        =   2
      Top             =   30
      Width           =   5865
      Begin VB.TextBox txtRpt 
         DataField       =   "report_desc"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Index           =   1
         Left            =   90
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   1260
         Width           =   5700
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Preview"
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
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   195
         Width           =   1110
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
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
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   630
         Width           =   1110
      End
      Begin VB.TextBox txtRpt 
         DataField       =   "report_name"
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
         Left            =   90
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   540
         Width           =   3855
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
         Height          =   300
         Left            =   4050
         Picture         =   "frmReportViewer.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   540
         Width           =   300
      End
      Begin VB.Label Label1 
         Caption         =   "Report Selection"
         Height          =   255
         Left            =   90
         TabIndex        =   8
         Top             =   300
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Report Description"
         Height          =   255
         Left            =   90
         TabIndex        =   7
         Top             =   1020
         Width           =   2175
      End
   End
   Begin VB.TextBox txtRpt 
      DataField       =   "report_path"
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
      Left            =   1380
      TabIndex        =   1
      Top             =   1575
      Width           =   3450
   End
End
Attribute VB_Name = "frmReportViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim datRpt As Recordset

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    frmShowRptHr.Left = Me.Left + 100
    frmShowRptHr.Top = Me.Top + (Label2.Top - Label1.Top) + txtRpt(0).Height
    frmShowRptHr.Show vbModal
    If SelRpt <> "" Then
       datRpt.Requery
       datRpt.Find "Report_id =" & CLng(SelRpt), 0, adSearchForward, 1
    End If
    SelRpt = ""
End Sub

Private Sub cmdPrint_Click()
    'Print the report selected
    'Dim fr As New frmPayslips
    Dim SelFormula
    Dim myfile As String
    If Trim(txtRpt(2)) = "" Then
       MsgBox ("No report for printing"), vbInformation, "Report Viewer"
       Exit Sub
    End If
    
    myfile = Dir(Trim(txtRpt(2)))
    If myfile = "" Then
       MsgBox ("Cannot find the Report File specified."), vbInformation
       Exit Sub
    End If
    
    
    If datRpt!emp_filter = True Then
        frmSelectEmp.Left = Me.Left + 100
        frmSelectEmp.Top = Me.Top + (Label2.Top - Label1.Top) + txtRpt(0).Height
        frmSelectEmp.Show vbModal
        SelFormula = "{Payroll_Results.staff_code}='" & Trim(EmpSelect) & "'"
    End If
    If datRpt!wage_filter = True Then
        frmSelectWages.Left = Me.Left + 100
        frmSelectWages.Top = Me.Top + (Label2.Top - Label1.Top) + txtRpt(0).Height
        frmSelectWages.Show vbModal
        'SelFormula = ""
        If SelFormula <> "" Then
           SelFormula = SelFormula & " AND {wage_types.wage_code}='" & WageSelect & "'"
        Else
           SelFormula = "{wage_types.wage_code}='" & WageSelect & "'"
        End If
        WageSelect = ""
    End If
    
    If datRpt!date_range = True Then
        rptMonthsFilter.Left = Me.Left + 100
        rptMonthsFilter.Top = Me.Top + (Label2.Top - Label1.Top) + txtRpt(0).Height
        rptMonthsFilter.Show vbModal
        If SelFormula <> "" Then
            If SelF <> "" Then
               SelFormula = SelFormula & " AND " & SelF
            End If
        Else
           SelFormula = SelF
        End If
    End If
    'fr.rpt1.PrintFileType = crptExcel50
    'fr.rpt1.Destination = crptToFile
    fr.rpt1.ReportFileName = Trim(txtRpt(2))
    fr.rpt1.WindowParentHandle = fr.hwnd
    If SelFormula <> "" Then
       fr.rpt1.SelectionFormula = SelFormula & " AND {Personal_data.co_code}='" & ID & "'"
    Else
        fr.rpt1.SelectionFormula = "{company_data.co_code}='" & ID & "'"
    End If
    fr.rpt1.Action = 1
    fr.Caption = txtRpt(0)
End Sub

Private Sub Form_Load()
    PRpt = True
    Dim txt As TextBox
    Set datRpt = New Recordset
    datRpt.Open "SELECT * FROM hr_Reports", cn, adOpenStatic, adLockOptimistic
    For Each txt In Me.txtRpt
        Set txt.DataSource = datRpt
    Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PRpt = False
End Sub
