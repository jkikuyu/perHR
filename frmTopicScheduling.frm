VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmTopicScheduling 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Schedule Your Training Topics and Timings"
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7095
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTopicScheduling.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   7095
      TabIndex        =   1
      Top             =   5205
      Width           =   7095
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   345
         Left            =   6015
         TabIndex        =   2
         Top             =   30
         Width           =   1035
      End
   End
   Begin MSDataGridLib.DataGrid grdtp 
      Height          =   4770
      Left            =   15
      TabIndex        =   0
      Top             =   435
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   8414
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   -2147483648
      HeadLines       =   1
      RowHeight       =   23
      RowDividerStyle =   4
      FormatLocked    =   -1  'True
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "co_code"
         Caption         =   "co_code"
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
         DataField       =   "training_code"
         Caption         =   "training_code"
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
      BeginProperty Column02 
         DataField       =   "date"
         Caption         =   "date"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "dd/MMM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "start_time"
         Caption         =   "start_time"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "HH.mm"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "end_time"
         Caption         =   "end_time"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "HH:mm"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "topic"
         Caption         =   "topic"
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
            Object.Visible         =   0   'False
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column01 
            Object.Visible         =   0   'False
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   945.071
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   945.071
         EndProperty
         BeginProperty Column05 
            WrapText        =   -1  'True
            ColumnWidth     =   3404.977
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmTopicScheduling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim datTp As Recordset

Private Sub cmdClose_Click()
    Unload Me
    
End Sub

Private Sub Form_Load()
    Set datTp = New Recordset
    datTp.Open "SELECT * FROM tr_tp_sch WHERE co_code ='" & ID & _
    "' AND training_code =" & Tp, cn, adOpenStatic, adLockOptimistic
    
    Set grdtp.DataSource = datTp
    FormatGrid
End Sub

Public Sub FormatGrid()
    'Format the grid to give the desired look
    With grdtp
        .Columns(0).Visible = False
        .Columns(1).Visible = False
        .Columns(2).Width = 1200
        .Columns(3).Width = 950
        .Columns(4).Width = 950
        .Columns(5).Width = 3400
        .Columns(5).WrapText = True
    End With
End Sub

Private Sub grdtp_OnAddNew()
    'Compute a new value for the tp_id
    Set datExtra = New Recordset
    datExtra.Open "SELECT TOP 1 * FROM tr_tp_sch ORDER BY tp_id DESC", cn, adOpenStatic, adLockOptimistic
    If datExtra.RecordCount = 0 Then
       datTp!tp_id = 1
    Else
        datTp!tp_id = datExtra!tp_id + 1
    End If
    
    
    grdtp.Columns(0) = ID
    grdtp.Columns(1) = Tp
End Sub
