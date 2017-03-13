VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4020
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   5760
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   5760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   30
      Top             =   450
   End
   Begin Threed.SSPanel SSPanel1 
      Align           =   2  'Align Bottom
      Height          =   510
      Left            =   0
      TabIndex        =   3
      Top             =   3510
      Width           =   5760
      _Version        =   65536
      _ExtentX        =   10160
      _ExtentY        =   900
      _StockProps     =   15
      Caption         =   "An ICS Project developed by Allan Obiero"
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
      BevelOuter      =   1
      Alignment       =   1
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "HUMAN RESOURCES SYSTEM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   -1  'True
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   2895
      Left            =   75
      TabIndex        =   4
      Top             =   120
      Width           =   5610
   End
   Begin VB.Label lblloadst 
      BackStyle       =   0  'Transparent
      Caption         =   "Loading......."
      Height          =   255
      Left            =   30
      TabIndex        =   2
      Top             =   4530
      Width           =   2775
   End
   Begin VB.Label lblPlatform 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Windows 95/98"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3930
      TabIndex        =   1
      Top             =   2715
      Width           =   1725
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4770
      TabIndex        =   0
      Top             =   3075
      Width           =   885
   End
   Begin VB.Image Image1 
      Height          =   3495
      Left            =   30
      Picture         =   "frmSplash.frx":000C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5925
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SourceSt As Boolean
Dim FileSt As Boolean
Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    'lblProductName.Caption = App.Title
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub

Private Sub Timer1_Timer()
On Error GoTo StartErr
    'Use this timer to check for file integrity and load the main program
    'If SourceSt = False Then
        'If GetSetting("PrMngr", "Registry Setting", "SerialNo") = "" Then
        '   frmActKey.Show vbModal
        '   If valid_key = True Then
        '
        '   Else
        '        End
        '   End If
        'ElseIf GetSetting("PrMngr", "Registry Setting", "SerialNo") <> "PJC50042" Then
        '    MsgBox ("Program not well registered."), vbInformation
        '    End
        'ElseIf GetSetting("PrMngr", "Registry Setting", "ActiveKey") = "TV-206273" Then
        '   If DateDiff("d", GetSetting("PrMngr", "Registry Setting", "StartDate"), Now) > 30 Then
        '      MsgBox ("You Evaluation period of this software(30 days) has expired." & Chr(13) & _
        '      "Please get an upgraded working copy of this software from your dealers to continue using it."), vbInformation
        '      End
        '   Else
        '      MsgBox ("This is an evaluation copy of the Personnel Manager System." & Chr(13) & _
        '      "This copy has an Evaluation Period of 30 days and is bound to expire on " & _
        '      Format(DateAdd("d", 30, GetSetting("PrMngr", "Registry Setting", "StartDate")), "dd/mmmm/yyyy")), vbExclamation, "System Notice"
        '   End If
        'ElseIf GetSetting("PrMngr", "Registry Setting", "ActiveKey") <> "ZN-10094V" Then
        '    MsgBox ("Program not well registered."), vbInformation
        '    End
        'End If
        
    
        'lblloadst = "Checking for data source"
        'lblloadst.Refresh
        'Me.Refresh
        'Check for the datSource
        Set cn = New Connection
        cn.CursorLocation = adUseClient
        cn.Open "PROVIDER=MSDASQL;dsn=PMngr;uid=Admin;pwd=soni;"
        SourceSt = True
        
        'lblloadst = "Checking file integrity"
        'Set datExtra = New Recordset
        'datExtra.Open "SELECT * FROM Personal_data", cn, adOpenStatic, adLockOptimistic
        'datExtra.Close
        
        
        'Call PackTables
        
        frmEntry.Show
        Unload Me
        
    'End If
    Exit Sub
StartErr:
'An Error has been found when starting
MsgBox (Err.Description & Chr(13) & "Program Execution Terminated."), vbCritical
End
End Sub

Private Sub Timer2_Timer()
    frmEntry.Show
    Unload Me
End Sub

Private Sub Timer3_Timer()
On Error GoTo StartErr
    'Timer to check for file integrity
    lblloadst = "Checking file integrity"
    Set datExtra = New Recordset
    datExtra.Open "SELECT * FROM Personal_data", cn, adOpenStatic, adLockOptimistic
    datExtra.Close
    
    
    Call PackTables
    Exit Sub
StartErr:
MsgBox (Err.Description & Chr(13) & "Program Execution halted."), vbCritical
End
End Sub
