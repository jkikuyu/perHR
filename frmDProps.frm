VERSION 5.00
Begin VB.Form frmDProps 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Day's Properties"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4515
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   4515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdQuit 
      Caption         =   "&Quit"
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   1515
      Width           =   1020
   End
   Begin VB.CommandButton cmdPost 
      Caption         =   "&Post Properties"
      Height          =   375
      Left            =   30
      TabIndex        =   3
      Top             =   1515
      Width           =   1830
   End
   Begin VB.Frame Frame1 
      Caption         =   "Properties for the "
      Height          =   1425
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   4485
      Begin VB.CheckBox chkNoWork 
         Caption         =   "Set this day as a non-working day"
         Height          =   240
         Left            =   195
         TabIndex        =   2
         Top             =   1005
         Width           =   3525
      End
      Begin VB.CheckBox chkHoliday 
         Caption         =   "Set this days as a holiday"
         Height          =   360
         Left            =   180
         TabIndex        =   1
         Top             =   405
         Width           =   3420
      End
   End
End
Attribute VB_Name = "frmDProps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim datHoliday As Recordset
Dim datNoWork As Recordset

Private Sub cmdPost_Click()
    Dim hID As Long
    If MsgBox("Assign the specified properties to " & frmCalendar.clnd.Value, vbYesNo + vbQuestion, "Day Properties") = vbYes Then
       If chkHoliday.Value = 1 Then
          frmHoliday.Show vbModal
       Else
          If chkNoWork.Value = 1 Then
            'Put a solitary no work day without a holiday
            'Add this no work day to the list
            'Get the last noworkday_id
            Set datExtra = New Recordset
            datExtra.Open "SELECT TOP 1 * FROM non_working ORDER BY eID DESC", cn, adOpenStatic, adLockOptimistic
            If datExtra.RecordCount = 0 Then
               hID = 1
            Else
               hID = datExtra!eid + 1
            End If
            
            With datNoWork
                .AddNew
                !eid = hID
                !co_code = ID
                !Year = Year(frmCalendar.clnd.Value)
                !Month = Month(frmCalendar.clnd.Value)
                !Date = Day(frmCalendar.clnd.Value)
                .Update
            End With
          End If
       End If
       MsgBox ("Properties set"), vbInformation, "Day Properties"
    End If
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    'Get the holidays
    Set datHolidays = New Recordset
    datHolidays.Open "SELECT * FROM Holidays WHERE co_code ='" & _
    ID & "'", cn, adOpenStatic, adLockOptimistic
    
    'Get the non-working days
    Set datNoWork = New Recordset
    datNoWork.Open "SELECT * FROM Non_Working WHERE co_code ='" & _
    ID & "'", cn, adOpenStatic, adLockOptimistic
End Sub
