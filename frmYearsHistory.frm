VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmYearsHistory 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Periods History"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3330
   Icon            =   "frmYearsHistory.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   3330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel SSPanel4 
      Height          =   375
      Left            =   2205
      TabIndex        =   4
      Top             =   0
      Width           =   1095
      _Version        =   65536
      _ExtentX        =   1931
      _ExtentY        =   661
      _StockProps     =   15
      Caption         =   "Close Date"
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
   Begin Threed.SSPanel SSPanel3 
      Height          =   375
      Left            =   1110
      TabIndex        =   3
      Top             =   0
      Width           =   1095
      _Version        =   65536
      _ExtentX        =   1931
      _ExtentY        =   661
      _StockProps     =   15
      Caption         =   "Open Date"
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
   Begin Threed.SSPanel SSPanel2 
      Height          =   375
      Left            =   15
      TabIndex        =   2
      Top             =   0
      Width           =   1095
      _Version        =   65536
      _ExtentX        =   1931
      _ExtentY        =   661
      _StockProps     =   15
      Caption         =   "Year"
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
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   2100
      TabIndex        =   1
      Top             =   3930
      Width           =   1215
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdPeriods 
      Height          =   3375
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   5953
      _Version        =   393216
      FixedRows       =   0
      FixedCols       =   0
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   2
      GridLinesFixed  =   1
      ScrollBars      =   2
      SelectionMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   0
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
End
Attribute VB_Name = "frmYearsHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim datPeriods As Recordset
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Set datPeriods = New Recordset
    datPeriods.Open "SELECT Year,start_Date,end_date FROM Year_History WHERE co_code ='" & ID & "'", cn, adOpenStatic, adLockOptimistic
    Set grdPeriods.DataSource = datPeriods
    With grdPeriods

        .Redraw = False
        ' set grid's column widths
        .ColWidth(0) = 1100
        .ColWidth(1) = 1100
        .ColWidth(2) = 1150
        
        
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

        .AllowBigSelection = False
        .FillStyle = flexFillSingle
        .Redraw = True

    End With
End Sub
