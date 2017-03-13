VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmLSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Leave Setup"
   ClientHeight    =   1275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5280
   Icon            =   "frmLSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1275
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Post the Leave Set-up According to the following"
      Height          =   1155
      Left            =   15
      TabIndex        =   0
      Top             =   105
      Width           =   5250
      Begin VB.TextBox txtbal 
         Height          =   285
         Left            =   825
         TabIndex        =   7
         Top             =   765
         Width           =   1470
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   330
         Left            =   4140
         TabIndex        =   6
         Top             =   585
         Width           =   1050
      End
      Begin MSDataListLib.DataCombo cboleave 
         Height          =   345
         Left            =   855
         TabIndex        =   4
         Top             =   330
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   609
         _Version        =   393216
         Style           =   2
         ListField       =   "leave_name"
         BoundColumn     =   "leave_code"
         Text            =   ""
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
      Begin VB.CommandButton cmdPost 
         Caption         =   "&Post"
         Height          =   330
         Left            =   4155
         TabIndex        =   3
         Top             =   195
         Width           =   1050
      End
      Begin VB.Label Label3 
         Caption         =   "Leave"
         Height          =   255
         Left            =   135
         TabIndex        =   5
         Top             =   360
         Width           =   720
      End
      Begin VB.Label Label2 
         Caption         =   "days"
         Height          =   270
         Left            =   2385
         TabIndex        =   2
         Top             =   780
         Width           =   705
      End
      Begin VB.Label Label1 
         Caption         =   "Bal b/f:"
         Height          =   225
         Left            =   105
         TabIndex        =   1
         Top             =   780
         Width           =   585
      End
   End
End
Attribute VB_Name = "frmLSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim datLeave As Recordset

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPost_Click()
On Error GoTo PostErr
    'Validate entries
    If Trim(txtbal) = "" Then
       MsgBox ("You must give the number of days to post"), vbInformation, "Leave Setup"
       Exit Sub
    ElseIf cboleave.BoundText = "" Then
       MsgBox ("You must give the leave to post"), vbInformation, "Leave Setup"
       Exit Sub
    End If
    Set datExtra = New Recordset
    If MsgBox("Post this value as leave balance b/f ?", vbYesNo + vbQuestion) = vbYes Then
       'If frmLeaveBal.flex.Row = 0 Then
       '   MsgBox ("No leave entitlement for this employee"), vbInformation, "Leave Setup"
       '   Exit Sub
       'End If
       frmLeaveBal.flex.Col = 1
       datExtra.Open "SELECT * FROM leave_bal WHERE staff_code='" & _
       frmLeaveBal.flex.Text & "' AND leave_code =" & _
       cboleave.BoundText & " AND year =" & Operation_Year, cn, adOpenStatic, adLockOptimistic
       If datExtra.RecordCount = 0 Then
          MsgBox ("Employee has not been assigned this leave type under " & Chr(13) & _
          "leave entitlements for this operation year"), vbInformation, "Leave Setup"
          Exit Sub
       Else
          'Post this leave
          With datExtra
               !bal_bf = !bal_bf + CDbl(txtbal)
               !bal = !bal + CDbl(txtbal)
               .Update
          End With
          MsgBox ("Leave Balance Posted"), vbInformation, "Leave Setup"
          Unload Me
       End If
    End If
    Exit Sub
PostErr:
MsgBox ("Cannot post Leave Setup" & Chr(13) & Err.Description), vbInformation, "Leave Setup"
End Sub

Private Sub Form_Load()
    Set datLeave = New Recordset
    datLeave.Open "SELECT * FROM Leave_Mast WHERE co_code='" & _
    ID & "'", cn, adOpenStatic, adLockOptimistic
    
    Set cboleave.RowSource = datLeave
End Sub

Private Sub txtbal_LostFocus()
    'Convert this value to a whole value
    If Trim(txtbal) <> "" Then
       If IsNumeric(txtbal) Then
          txtbal = Round(txtbal, 3)
       Else
          MsgBox ("Invalid Number of Days"), vbInformation, "Leave Setup"
          txtbal.SetFocus
          SendKeys "{HOME}"
          SendKeys "+{END}"
       End If
    End If
End Sub
