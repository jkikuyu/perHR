VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmselschemes 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4110
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   4110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   4110
      TabIndex        =   1
      Top             =   1200
      Width           =   4110
      Begin VB.CommandButton cmdNew 
         Caption         =   "&New"
         Height          =   350
         Left            =   2040
         TabIndex        =   4
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   350
         Left            =   3120
         TabIndex        =   3
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "&Select"
         Height          =   350
         Left            =   960
         TabIndex        =   2
         Top             =   0
         Width           =   975
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdsch 
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   1931
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
Attribute VB_Name = "frmselschemes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim datsch As Recordset
Public Sub FormatGrid()
    With grdsch
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
    frmmedschemes.Show vbModal
    
    datsch.Requery
    Set grdsch.DataSource = datsch
    FormatGrid
End Sub

Private Sub cmdSelect_Click()
    If grdsch.Text <> "" Then
       frmempmedsch.txtsch(0) = grdsch.Text
       Unload Me
    Else
        MsgBox ("No scheme to select."), vbInformation
    End If
End Sub

Private Sub Form_Load()
    Set datsch = New Recordset
    datsch.Open "SELECT * FROM Med_schemes", cn, adOpenStatic, adLockOptimistic
    
    Set grdsch.DataSource = datsch
    
    'Format the appearance of this grid
    FormatGrid
    
End Sub
