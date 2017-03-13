VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmTheTrainings 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4560
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   4560
      TabIndex        =   1
      Top             =   1935
      Width           =   4560
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   345
         Left            =   3450
         TabIndex        =   4
         ToolTipText     =   "Close the screen"
         Top             =   0
         Width           =   1005
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "&New"
         Height          =   345
         Left            =   2370
         TabIndex        =   3
         ToolTipText     =   "Add a Training"
         Top             =   0
         Width           =   1005
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "&Select"
         Height          =   345
         Left            =   1290
         TabIndex        =   2
         ToolTipText     =   "Select a Training"
         Top             =   0
         Width           =   1005
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdTrainings 
      Height          =   1890
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4560
      _ExtentX        =   8043
      _ExtentY        =   3334
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
Attribute VB_Name = "frmTheTrainings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dattrainings As Recordset

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdNew_Click()
    Trsb = True
    Unload Me
    
End Sub

Private Sub cmdSelect_Click()
    If grdTrainings.Text <> "" Then
        Trs = grdTrainings.Text
        Unload Me
    Else
        MsgBox ("No Training to select.")
    End If
    
End Sub

Private Sub Form_Load()
    Set dattrainings = New Recordset
    dattrainings.Open "SELECT training_code, training_name FROM co_training WHERE co_code ='" & ID & _
    "'", cn, adOpenStatic, adLockOptimistic
    
    Set Me.grdTrainings.DataSource = dattrainings
    With grdTrainings

        .Redraw = False
        ' set grid's column widths
        .ColWidth(0) = 600
        .ColWidth(1) = 4000
        
        
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

Private Sub grdTrainings_DblClick()
    cmdSelect_Click
End Sub
