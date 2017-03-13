VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmShowRpts 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4590
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   350
      Left            =   3600
      TabIndex        =   3
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      Height          =   350
      Left            =   2520
      TabIndex        =   2
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Select"
      Height          =   350
      Left            =   1440
      TabIndex        =   1
      Top             =   1920
      Width           =   975
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdRpts 
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   3201
      _Version        =   393216
      FixedRows       =   0
      FixedCols       =   0
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   2
      GridLinesFixed  =   1
      ScrollBars      =   2
      SelectionMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   0
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
End
Attribute VB_Name = "frmShowRpts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim datRpt As Recordset
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdNew_Click()
    NewRpt = True
    Unload Me
End Sub

Private Sub cmdSelect_Click()
    If grdRpts.Text <> "" Then
       SelRpt = grdRpts.Text
       Unload Me
    Else
        MsgBox ("No report Selected."), vbInformation, "Report Viewer"
    End If
End Sub

Private Sub Form_Load()
    If App.FileDescription <> "Personnel Manager Administrator" Then
        If frmReportViewer.Visible = True Then
           cmdNew.Enabled = False
        End If
    End If
    Set datRpt = New Recordset
    datRpt.Open "SELECT report_id, report_name FROM hr_reports ORDER BY report_id", cn, adOpenStatic, adLockOptimistic
    Set grdRpts.DataSource = datRpt
    
    'Format the grid
With grdRpts

        .Redraw = False
        ' set grid's column widths
        .ColWidth(0) = 1250
        .ColWidth(1) = 3800
        
        ' set grid's style
        .AllowBigSelection = True
        .FillStyle = flexFillRepeat

        ' grey every other row
        For i = .FixedRows + 1 To .Rows - 1 Step 2
            .Row = i
            .RowHeight(i) = 300
            .Col = .FixedCols
            .ColSel = .Cols() - .FixedCols - 1
            .CellBackColor = &H80000018       '&HC0C0C0   ' light grey
        Next i

        .AllowBigSelection = False
        .FillStyle = flexFillSingle
        .Redraw = True

    End With
End Sub

Private Sub grdRpts_DblClick()
    cmdSelect_Click
End Sub
