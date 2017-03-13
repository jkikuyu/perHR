VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmLeaveBal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Leave Balances Analysis"
   ClientHeight    =   5130
   ClientLeft      =   3030
   ClientTop       =   5790
   ClientWidth     =   10050
   Icon            =   "frmLeaveBal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   10050
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSetup 
      Caption         =   "&Setup"
      Height          =   345
      Left            =   30
      TabIndex        =   10
      Top             =   4770
      Width           =   1050
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   345
      Left            =   8940
      TabIndex        =   9
      Top             =   4785
      Width           =   1080
   End
   Begin VB.CommandButton cmdFind 
      Height          =   345
      Left            =   4290
      Picture         =   "frmLeaveBal.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Find an Employee"
      Top             =   0
      Width           =   345
   End
   Begin Threed.SSPanel SSPanel7 
      Height          =   345
      Left            =   9240
      TabIndex        =   7
      Top             =   0
      Width           =   795
      _Version        =   65536
      _ExtentX        =   1402
      _ExtentY        =   609
      _StockProps     =   15
      Caption         =   "Bal"
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSPanel SSPanel6 
      Height          =   345
      Left            =   8565
      TabIndex        =   6
      Top             =   0
      Width           =   660
      _Version        =   65536
      _ExtentX        =   1164
      _ExtentY        =   609
      _StockProps     =   15
      Caption         =   "Taken"
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSPanel SSPanel5 
      Height          =   345
      Left            =   7785
      TabIndex        =   5
      Top             =   0
      Width           =   765
      _Version        =   65536
      _ExtentX        =   1349
      _ExtentY        =   609
      _StockProps     =   15
      Caption         =   "Entitled"
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSPanel SSPanel4 
      Height          =   345
      Left            =   7020
      TabIndex        =   4
      Top             =   0
      Width           =   750
      _Version        =   65536
      _ExtentX        =   1323
      _ExtentY        =   609
      _StockProps     =   15
      Caption         =   "Bal b/f"
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSPanel SSPanel3 
      Height          =   345
      Left            =   4650
      TabIndex        =   3
      Top             =   0
      Width           =   2355
      _Version        =   65536
      _ExtentX        =   4154
      _ExtentY        =   609
      _StockProps     =   15
      Caption         =   "Leave Type"
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   345
      Left            =   750
      TabIndex        =   2
      Top             =   0
      Width           =   3540
      _Version        =   65536
      _ExtentX        =   6244
      _ExtentY        =   609
      _StockProps     =   15
      Caption         =   "Employee"
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   345
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   735
      _Version        =   65536
      _ExtentX        =   1296
      _ExtentY        =   609
      _StockProps     =   15
      Caption         =   "Year"
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flex 
      Height          =   4365
      Left            =   0
      TabIndex        =   0
      Top             =   375
      Width           =   10050
      _ExtentX        =   17727
      _ExtentY        =   7699
      _Version        =   393216
      FixedRows       =   0
      FixedCols       =   0
      HighLight       =   2
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
Attribute VB_Name = "frmLeaveBal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim datBal As Recordset

Private Sub Command1_Click()

End Sub

Private Sub cmdClose_Click()
    Unload Me
    
End Sub

Private Sub cmdFind_Click()
    frmSelectEmp.Left = Me.Left + SSPanel1.Width
    frmSelectEmp.Top = Me.Top + (SSPanel1.Height * 2)
    frmSelectEmp.Show vbModal
    
    If PLeaveEmp <> "" Then
        Set datBal = New Recordset
        DoInitialSettings
        DoSql
        DoSort
    End If
End Sub

Private Sub cmdSetup_Click()
    frmLSetup.Show vbModal
    datBal.Requery
    Set datBal = New Recordset
    DoInitialSettings
    DoSql
    DoSort
End Sub

Private Sub flex_DragDrop(Source As VB.Control, x As Single, y As Single)
    'If flex.Tag = "" Then Exit Sub
    'flex.Redraw = False
    'flex.ColPosition(Val(flex.Tag)) = flex.MouseCol
    'DoSort
    'flex.Redraw = True
End Sub

Private Sub flex_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'flex.Tag = ""
    'If flex.MouseRow <> 0 Then Exit Sub
    'flex.Tag = Str(flex.MouseCol)
    'flex.Drag 1
End Sub

Private Sub Form_Load()
  PLeave3 = True
  'Get the first available employee
  Set datEmp = New Recordset
  datEmp.Open "SELECT * FROM Personal_data WHERE co_code ='" & _
  ID & "'", cn, adOpenStatic, adLockOptimistic
  If datEmp.RecordCount <> 0 Then
     PLeaveEmp = Trim(datEmp!staff_code)
  End If
  
  Set datBal = New Recordset
    
    
  flex.Redraw = False
  DoInitialSettings
  DoSql
  DoSort
  flex.Redraw = True
End Sub

Sub DoInitialSettings()
On Error Resume Next
    Dim i%
    flex.Row = 0
    flex.ColAlignment(0) = 7
    
    For i = 0 To flex.Cols - 1
    
        flex.Col = i
        flex.CellFontSize = 14
        flex.CellAlignment = 4
        
        flex.MergeCol(i) = True     ' Allow merge on Columns 0 thru 3
    Next i
    
    flex.ColWidth(0) = 600
    flex.ColWidth(1) = 800
    flex.ColWidth(2) = 1600
    flex.ColWidth(3) = 1600
    flex.ColWidth(4) = 2350
    flex.ColWidth(5) = 800
    flex.ColWidth(6) = 800
    flex.ColWidth(7) = 800
    flex.ColWidth(8) = 800
    
    flex.MergeCells = flexMergeRestrictColumns
End Sub
Sub DoSql()
    
    Dim mysql$
    
    mysql = "SELECT Leave_bal.Year, leave_bal.staff_code, personal_data.last_name, personal_data.first_name, " & _
    " leave_mast.Leave_name, leave_bal.bal_bf, leave_bal.days_ent, " & _
    " leave_bal.days_taken, leave_bal.bal FROM (leave_bal INNER JOIN Personal_data ON leave_bal.staff_code = personal_data.staff_code) " & _
    " INNER JOIN leave_mast ON " & _
    " leave_bal.leave_code = leave_mast.leave_code WHERE leave_bal.co_code ='" & ID & "'"
    
    datBal.Open mysql, cn, adOpenStatic, adLockOptimistic
    
    If PLeaveEmp <> "" Then
       datBal.Filter = "staff_code ='" & Trim(PLeaveEmp) & "'"
    End If
    Set flex.DataSource = datBal
    
    With flex
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
    flex.MergeCol(5) = False
    flex.MergeCol(6) = False
    flex.MergeCol(7) = False
    flex.MergeCol(8) = False
End Sub
Sub DoSort()
On Error Resume Next
    flex.Col = 0
    flex.ColSel = flex.Cols - 1
    flex.Sort = 1 ' Generic Ascending
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PLeave3 = False
End Sub
