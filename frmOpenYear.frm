VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmOpenYear 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Open Payroll Periods"
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4035
   Icon            =   "frmOpenYear.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   4035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar prg 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   1980
      Width           =   4035
      _ExtentX        =   7117
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdCancel 
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
      Left            =   3045
      TabIndex        =   3
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "&Open"
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
      Left            =   1845
      TabIndex        =   2
      Top             =   1560
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   "Open Payroll Processing Period For:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3900
      Begin MSComCtl2.DTPicker dtOpen 
         Height          =   300
         Left            =   2205
         TabIndex        =   5
         Top             =   600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   24576001
         CurrentDate     =   36838
      End
      Begin VB.ComboBox cboYear 
         Height          =   315
         Left            =   135
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   585
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Start Date:"
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
         Left            =   2205
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmOpenYear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lastperiod
Dim datPeriods As Recordset
Dim datTransactions As Recordset
Dim datLoan As Recordset

Private Sub cmdCancel_Click()
     Unload Me
End Sub

Private Sub cmdOpen_Click()
'On Error GoTo OpErr

Dim prgvalue
Dim prevmonth As String
Dim prevyear As Long
Dim p_code As Long
'Validate

'Check if this period had already been opened before
Set datExtra = New Recordset
datExtra.Open "SELECT * FROM year_History WHERE co_code ='" & ID & _
"' AND year =" & CLng(cboYear), cn, adOpenStatic, adLockOptimistic
If datExtra.RecordCount <> 0 Then
   MsgBox ("Violation of Operation Year Integrity!" & Chr(13) & _
          "This operation year had been opened before and cannot be opened again."), vbCritical, "Violation"
   Exit Sub
End If

If MsgBox("Open a new operation year?", vbQuestion + vbYesNo, "Year Opening") = vbYes Then
    'First get the previous period
    cn.BeginTrans
    Set datExtra = New Recordset
    datExtra.Open "SELECT TOP 1 * FROM year_History WHERE co_code ='" & ID & "'" & _
    " ORDER BY p_code DESC", cn, adOpenStatic, adLockOptimistic
    
    If datExtra.RecordCount <> 0 Then
        prevyear = Trim(datExtra!Year)
        p_code = datExtra!p_code
        
        'Check if the period to be opened is earlier than the previous period
        'Stop opening if true is returned
        'If GetMonthNum(cboMonth) < GetMonthNum(Trim(prevmonth)) Then
            If cboYear <= prevyear Then
               MsgBox ("Period Integrity Violation!" & Chr(13) & _
               "You are trying to open an operation year earlier than your previous period."), vbCritical
               Exit Sub
            End If
            If cboYear < prevyear Then
               MsgBox ("Period Integrity Violation!" & Chr(13) & _
               "You are trying to open an operation year earlier than your previous period."), vbCritical
               Exit Sub
            End If
    End If
    
    
    With datExtra
        If IsNull(!End_Date) Then  'Previous period not yet closed
           MsgBox ("You must close the current operation year before opening a new period."), vbCritical
           Exit Sub
        Else
            Screen.MousePointer = vbHourglass
            .AddNew
            !p_code = p_code + 1
            !co_code = ID
            !Year = CLng(cboYear)
            !Start_Date = Format(Now, "dd/mmm/yyyy")
            .Update
            Operation_Year = cboYear
            ThisYear = Operation_Year
            
            Set datPeriods = New Recordset
            
            datPeriods.Open "SELECT TOP 1 * FROM Current_year WHERE co_code ='" & ID & "'" & _
            " ORDER BY co_code", cn, adOpenStatic, adLockOptimistic
            
            If datPeriods.RecordCount <> 0 Then
               datPeriods.Delete
            End If
               datPeriods.AddNew
               datPeriods!co_code = ID
               datPeriods!Year = CLng(cboYear)
               datPeriods.Update
               
               Set datExtra = New Recordset
               datExtra.Open "SELECT * FROM company_data WHERE co_code ='" & ID & "'", cn, adOpenStatic, adLockOptimistic
               frmMain.Caption = "Human Resources System     " & Trim(datExtra!co_name) & " - " & Operation_Year
        End If
    End With
    
    'Transfer the leave balances and the medical limits and amount to be
    'recovered
    
    'Get the leave that needs to be carried forward
    Dim datLeave As Recordset
    Dim datEnt As Recordset
    Set datLeave = New Recordset
    datLeave.Open "SELECT * FROM leave_mast WHERE forward = 1", cn, adOpenStatic, adLockOptimistic
    If datLeave.RecordCount <> 0 Then
       While Not datLeave.EOF
            'Get the leave data for employees with this iterative leave type in the previous year
            Set datExtra = New Recordset
            datExtra.Open "SELECT * FROM leave_bal WHERE co_code ='" & ID & "' AND leave_code=" & _
            datLeave!Leave_code & " AND Year =" & prevyear, cn, adOpenStatic, adLockOptimistic
            
            'Loop through each employee data, get their entitlements, carry forward balances
            'and put in a new entry for this year's leave
            If datExtra.RecordCount <> 0 Then
               While Not datExtra.EOF
                     'Get the leave entitlement for this particular leave type for this employee
                     Set datEnt = New Recordset
                     datEnt.Open "SELECT * FROM leave_entitlement WHERE co_code ='" & _
                     ID & "' AND leave_code =" & datExtra!Leave_code & " AND staff_Code ='" & _
                     datExtra!staff_code & "'", cn, adOpenStatic, adLockOptimistic
                     
                     'Put a new leave entry for this employee bringing forward any balance
                     Set datExtra2 = New Recordset
                     With datExtra2
                         .Open "SELECT * FROM leave_bal", cn, adOpenStatic, adLockOptimistic
                         .AddNew
                         !co_code = ID
                         !staff_code = Trim(datExtra!staff_code)
                         !Leave_code = CDbl(datExtra!Leave_code)
                         !Year = Operation_Year
                                                
       
                         !bal_bf = CDbl(datExtra!bal)
                         !bal = CDbl(datEnt!leave_dur) + CDbl(datExtra!bal)
                         !days_taken = 0
                         .Update
                     End With
                     datExtra.MoveNext
               Wend
            End If
            datLeave.MoveNext
       Wend
    End If
                         
    'Transfer the medical limits to the new year
    Dim datMed As Recordset
    Dim datTrans As Recordset
    Dim inAmt As Currency
    Dim outAmt As Currency
    Dim outbfbal As Currency
    Dim inbfbal As Currency
    Dim outcf As Currency
    Dim incf As Currency
    Dim outrec As Currency
    Dim inrec As Currency
    Set datMed = New Recordset
    Set datExtra = New Recordset
     
    datMed.Open "SELECT * FROM med_limits WHERE co_code ='" & ID & _
    "' AND Year =" & prevyear, cn, adOpenStatic, adLockOptimistic
    
    datExtra.Open "SELECT * FROM med_limits WHERE co_code ='" & ID & _
    "' AND Year =" & prevyear, cn, adOpenStatic, adLockOptimistic
    
    If datMed.RecordCount <> 0 Then
       While Not datMed.EOF
            outcf = 0
            incf = 0
            inAmt = 0
            outAmt = 0
            outbfbal = 0
            inbfbal = 0
            outrec = 0
            inrec = 0
            
            'Get the amount recovered from this guy this year and compare it to the
            'amount of recovery b/f
            outbfbal = datMed!outrec - datMed!outrecbf
            'Evaluate the recovery of the bf
            If outbfbal < 0 Then   'We recovered less than we brought forward
               outcf = Abs(outbfbal) 'Get the absolute balance of bf for recovery
            ElseIf outbfbal > 0 Then 'We recovered more than we brought forward so there are more recoveries
               outAmt = outbfbal
            Else
            End If
                        
            'Get the amount expended by this employee for the outpatient
            Set datTrans = New Recordset
            datTrans.Open "SELECT sum(amt) AS 'Amt' FROM emp_mtrans WHERE " & _
            " co_code ='" & ID & "' AND Year =" & prevyear & _
            " AND staff_code ='" & Trim(datMed!staff_code) & "' AND opt = 1", cn, adOpenStatic, adLockOptimistic
            
            If datTrans.RecordCount <> 0 Then
               If CCur(datTrans!amt) > (CCur(datMed!outlimit)) Then
                  outrec = CCur(datTrans!amt) - (CCur(datMed!outlimit))
                  outrec = outrec - outAmt
               End If
            Else
            End If
            outcf = outcf + outrec
                        
            'Now do the same for the In-Patient Limits
            
            'Get the amount recovered from this guy this year and compare it to the
            'amount of recovery b/f
            inbfbal = datMed!inrec - datMed!inrecbf
            'Evaluate the recovery of the bf
            If inbfbal < 0 Then   'We recovered less than we brought forward
               incf = Abs(inbfbal) 'Get the absolute balance of bf for recovery
            ElseIf inbfbal > 0 Then 'We recovered more than we brought forward so there are more recoveries
               inAmt = inbfbal
            Else
            End If
                        
            'Get the amount expended by this employee for the inpatient
            Set datTrans = New Recordset
            datTrans.Open "SELECT sum(amt) AS 'Amt' FROM emp_mtrans WHERE " & _
            " co_code ='" & ID & "' AND Year =" & prevyear & _
            " AND staff_code ='" & Trim(datMed!staff_code) & "' AND opt = 2", cn, adOpenStatic, adLockOptimistic
            
            If datTrans.RecordCount <> 0 Then
               If CCur(datTrans!amt) > (CCur(datMed!inlimit)) Then
                  inrec = CCur(datTrans!amt) - (CCur(datMed!inlimit))
                  inrec = inrec - inAmt
               End If
            Else
            End If
            incf = incf + inrec
                        
                        
                        
            'Transfer the limits for the new year
            With datExtra
                .AddNew
                !co_code = ID
                !staff_code = Trim(datMed!staff_code)
                !Year = Operation_Year
                !inlimit = CCur(datMed!inlimit)
                !outlimit = CCur(datMed!outlimit)
                !outrecbf = outcf
                !inrecbf = incf
                !Date = Format(Date, "dd/mmm/yyyy")
                .Update
            End With
            
            'Put in a zero entry for the sake of the report printing every time a year is started
            'NB: Without any transaction entry, the report will not show this employee's limits thus
            'the need to put in a zero entry
            Set datExtra = New Recordset
            datExtra.Open "SELECT * FROM emp_mtrans", cn, adOpenStatic, adLockOptimistic
            With datExtra
                .AddNew
                !co_code = ID
                !staff_code = Trim(datMed!staff_code)
                !Date = Format(Date, "dd/mm/yyyy")
                !Year = Operation_Year
                !From = "O/Bal"
                !amt = 0
                !Opt = 1
                .Update
            End With
            
            datMed.MoveNext
        Wend
    End If
            
    cn.CommitTrans
    MsgBox ("Operation Year opening process complete "), vbInformation, "Year Opening"
    Unload Me
End If
    Screen.MousePointer = vbDefault
    Exit Sub
OpErr:
MsgBox "An error was encountered while opening the year." & Chr(13) & _
Err.Description & Chr(13) & _
"Operation has been aborted", vbInformation
cn.RollbackTrans
End Sub

Private Sub Form_Load()
    'Set datPeriods = New Recordset
    'datPeriods.Open "SELECT * FROM Periods_History", cn, adOpenStatic, adLockOptimistic
    Static x(12)
    x(1) = "January"
    x(2) = "February"
    x(3) = "March"
    x(4) = "April"
    x(5) = "May"
    x(6) = "June"
    x(7) = "July"
    x(8) = "August"
    x(9) = "September"
    x(10) = "October"
    x(11) = "November"
    x(12) = "December"
    'For y = 1 To 12
    '    cboMonth.AddItem x(y)
    'Next y
    'cboMonth = Format(Date, "mmmm")
    For yr = 1995 To 2099
        cboYear.AddItem yr
    Next
    cboYear = Year(Date)
    
    Set datExtra = New Recordset
    datExtra.Open "SELECT * FROM Options", cn, adOpenStatic, adLockOptimistic
    If datExtra!Consecutive = True Then
           'cboMonth.Enabled = False
           'cboYear.Enabled = False
           Set datPeriods = New Recordset
           datPeriods.Open "SELECT * FROM Current_Periods WHERE Co_code ='" & ID & "'", cn, adOpenStatic
            With datPeriods
                 If .RecordCount <> 0 Then
                    .MoveLast
                    lastperiod = !Period_Month
                    Select Case !Period_Month
                           Case "January"
                                cboMonth = "February"
                                cboYear = !Period_Year
                           Case "February"
                                cboMonth = "March"
                                cboYear = !Period_Year
                           Case "March"
                                cboMonth = "April"
                                cboYear = !Period_Year
                           Case "April"
                                cboMonth = "May"
                                cboYear = !Period_Year
                           Case "May"
                                cboMonth = "June"
                                cboYear = !Period_Year
                           Case "June"
                                cboMonth = "July"
                                cboYear = !Period_Year
                           Case "July"
                                cboMonth = "August"
                                cboYear = !Period_Year
                           Case "August"
                                cboMonth = "September"
                                cboYear = !Period_Year
                           Case "September"
                                cboMonth = "October"
                                cboYear = !Period_Year
                           Case "October"
                                cboMonth = "November"
                                cboYear = !Period_Year
                           Case "November"
                                cboMonth = "December"
                                cboYear = !Period_Year
                           Case "December"
                                cboMonth = "January"
                                cboYear = !Period_Year + 1
                    End Select
                End If
            End With
    End If
End Sub
