VERSION 5.00
Begin VB.Form frmCloseYear 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Close Payroll Period"
   ClientHeight    =   1665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmCloseYear.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   6
      ToolTipText     =   "Exit this Screen"
      Top             =   1245
      Width           =   975
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      ToolTipText     =   "Close the Payroll Period"
      Top             =   1245
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Current Payroll Period"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1080
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.TextBox txtPeriod 
         Height          =   285
         Index           =   3
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox txtPeriod 
         Height          =   285
         Index           =   0
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label4 
         Caption         =   "Close Date:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Current Year"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmCloseYear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim datYear As Recordset
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdStart_Click()
    'This procedure is added to close a payroll period.
    Set datExtra = New Recordset
    datExtra.Open "SELECT * FROM current_year WHERE Co_code='" & ID & "'", cn, adOpenStatic, adLockOptimistic
    With datExtra
         If .RecordCount <> 0 Then
            If MsgBox("Close the current operational year?", vbQuestion + vbYesNo, "Close Year?") = vbYes Then
               'Do validations for the period
               MsgBar "Validating operational year", True
               Set datYear = New Recordset
                    .Delete
                    'operation_Year = 0
                    
                    Set datExtra = New Recordset
                    datExtra.Open "SELECT * FROM company_data WHERE co_code ='" & ID & "'", cn, adOpenStatic, adLockOptimistic
                    frmMain.Caption = "Human Resource System     " & Trim(datExtra!co_name)
                      
            Else
                 MsgBox ("Closure aborted."), vbInformation, "Year Closure"
            End If
               Set datExtra = New Recordset
               datExtra.Open "SELECT TOP 1 * FROM year_History WHERE co_code ='" & ID & _
               "' AND year =" & Operation_Year & " ORDER BY co_code DESC", cn, adOpenStatic, adLockOptimistic
               datExtra!end_date = Format(Now, "dd/mmm/yyyy")
               datExtra.Update
               Operation_Year = 0
               MsgBox ("Operation Year Closed. Ensure you open a new period before proceeding."), vbInformation, "Year Closed"
               Unload Me
         Else
            MsgBox ("There's no current operation selected for this company."), vbInformation
         End If
    End With
End Sub

Private Sub Form_Load()
    Set datExtra = New Recordset
    datExtra.Open "SELECT * FROM current_year WHERE co_code ='" & ID & "'", cn, adOpenStatic, adLockOptimistic
    
    With datExtra
        If .RecordCount <> 0 Then
            txtPeriod(0) = !Year
            txtPeriod(3) = Format(Now, "dd/mmm/yyyy")
        Else
            MsgBox ("No previous period to close"), vbInformation
            Me.cmdStart.Enabled = False
        End If
    End With
End Sub
