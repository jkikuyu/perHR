VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmSelLeave 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   4680
      TabIndex        =   1
      Top             =   1440
      Width           =   4680
      Begin VB.CommandButton cmdSelect 
         Caption         =   "&Select"
         Height          =   330
         Left            =   1635
         TabIndex        =   4
         Top             =   30
         Width           =   975
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   330
         Left            =   3645
         TabIndex        =   3
         Top             =   30
         Width           =   975
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "&New"
         Height          =   330
         Left            =   2640
         TabIndex        =   2
         Top             =   30
         Width           =   975
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdLeave 
      Height          =   1350
      Left            =   -15
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   2381
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
Attribute VB_Name = "frmSelLeave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim datLeave As Recordset
Public Sub FormatGrid()
    With grdLeave
        .Redraw = False
        ' set grid's column widths
        .ColWidth(0) = 1250
        .ColWidth(1) = 3900
        
        ' set grid's style
        .AllowBigSelection = True
        .FillStyle = flexFillRepeat

        ' grey every other row
        For i = .FixedRows + 1 To .Rows - 1 Step 2
            .Row = i
            .Col = .FixedCols
            .ColSel = .Cols() - .FixedCols - 1
            .CellBackColor = &H80000018       '&HC0C0C0   ' light grey
        Next i
        
        'Set the row height for all the rows
        i = 0
        For i = .FixedRows + 1 To .Rows - 1
            .Row = i
            .RowHeight(i) = 300
        Next i
       
        .AllowBigSelection = False
        .FillStyle = flexFillSingle
        .Redraw = True

    End With
End Sub
Private Sub cmdClose_Click()
    Unload Me
    
End Sub

Private Sub cmdNew_Click()
    frmLeaves.Show vbModal
    datLeave.Requery
End Sub

Private Sub cmdSelect_Click()
    If frmlent.Visible = True Then
       frmlent.txtleave(0) = grdLeave.Text
       Unload Me
    Else
        frmEmpLeave.txtleave(0) = grdLeave.Text
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Set datLeave = New Recordset
    datLeave.Open "SELECT * FROM Leave_mast", cn, adOpenStatic, adLockOptimistic
    
    Set grdLeave.DataSource = datLeave
    FormatGrid
    
End Sub
