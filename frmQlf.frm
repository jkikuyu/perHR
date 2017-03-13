VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmQlf 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Qualifications"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5520
   Icon            =   "frmQlf.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "C&lose"
      Height          =   345
      Left            =   4500
      TabIndex        =   2
      Top             =   4455
      Width           =   990
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5475
      _Version        =   65536
      _ExtentX        =   9657
      _ExtentY        =   635
      _StockProps     =   15
      Caption         =   "Qualifications"
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
   Begin MSDataGridLib.DataGrid grdqlf 
      Height          =   4005
      Left            =   0
      TabIndex        =   0
      Top             =   405
      Width           =   5520
      _ExtentX        =   9737
      _ExtentY        =   7064
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   -2147483648
      ColumnHeaders   =   0   'False
      ForeColor       =   128
      HeadLines       =   1
      RowHeight       =   20
      RowDividerStyle =   4
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         ScrollBars      =   2
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmQlf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim datQlf As Recordset
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Set datQlf = New Recordset
    datQlf.Open "SELECT * FROM Qualifications", cn, adOpenStatic, adLockOptimistic
    
    Set grdqlf.DataSource = datQlf
    FormatGrid
End Sub

Public Sub FormatGrid()
    'Format the grid and give the desired dimensions
    With grdqlf
        .Columns(0).Locked = True
        .Columns(0).Width = 900
        .Columns(1).Width = 4020
        .Width = Me.Width - 60
    End With
End Sub

Private Sub grdqlf_KeyPress(KeyAscii As Integer)
    Char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(Char))
End Sub

Private Sub grdqlf_OnAddNew()
    'Calculate a new code for the qualification to be added
    Set datExtra = New Recordset
    datExtra.Open "SELECT TOP 1 * FROM Qualifications ORDER BY qf_code DESC", cn, adOpenStatic, adLockOptimistic
    If datExtra.RecordCount <> 0 Then
       grdqlf.Columns(0) = datExtra!qf_code + 1
    Else
        grdqlf.Columns(0) = 1
    End If
End Sub
