VERSION 5.00
Begin VB.Form frmARec 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Adjust Medical Recovery"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3930
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
   ScaleHeight     =   2505
   ScaleWidth      =   3930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   345
      Left            =   2790
      TabIndex        =   7
      ToolTipText     =   "Close Screen"
      Top             =   2145
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdjust 
      Caption         =   "&Adjust"
      Height          =   345
      Left            =   1650
      TabIndex        =   6
      ToolTipText     =   "Post Adjustment"
      Top             =   2145
      Width           =   1095
   End
   Begin VB.TextBox txtAdj 
      Height          =   330
      Left            =   2550
      TabIndex        =   5
      Top             =   1725
      Width           =   1350
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select a Recovery to Adjust"
      Height          =   750
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   3810
      Begin VB.ComboBox cboRec 
         Height          =   345
         ItemData        =   "frmARec.frx":0000
         Left            =   75
         List            =   "frmARec.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   270
         Width           =   3630
      End
   End
   Begin VB.Label lblAmt 
      Caption         =   "Label4"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   2550
      TabIndex        =   9
      Top             =   915
      Width           =   1170
   End
   Begin VB.Label Label3 
      Caption         =   "Maximum Possible Recovery"
      Height          =   255
      Left            =   45
      TabIndex        =   8
      Top             =   900
      Width           =   2415
   End
   Begin VB.Label lblAdjAmt 
      Caption         =   "Label3"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2550
      TabIndex        =   4
      Top             =   1290
      Width           =   1245
   End
   Begin VB.Label Label2 
      Caption         =   "Adjust amount to"
      Height          =   300
      Left            =   30
      TabIndex        =   3
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Amount Posted for Recovery"
      Height          =   285
      Left            =   45
      TabIndex        =   2
      Top             =   1290
      Width           =   2355
   End
End
Attribute VB_Name = "frmARec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboRec_Click()
    'Get the amount already posted for recovery
    Set datExtra2 = New Recordset
    datExtra2.Open "SELECT * FROM med_limits WHERE staff_code ='" & _
    Trim(frmempmed.pnlStaff_code) & "'", cn, adOpenStatic, adLockOptimistic
    If datExtra2.RecordCount <> 0 Then
       If cboRec.ListIndex = 0 Then
            frmARec.lblAmt = datExtra2!outpost + frmempmed.txtOutRecbal
            frmARec.lblAdjAmt = datExtra2!outpost
       ElseIf cboRec.ListIndex = 1 Then
            frmARec.lblAmt = datExtra2!inpost + frmempmed.txtInRecBal
            frmARec.lblAdjAmt = datExtra2!inpost
       End If
    End If
End Sub

Private Sub cmdAdjust_Click()
    'Make the desired adjustment
    Dim lkWage As String
    Dim lkname As String
    Dim exAmt As Currency
    If Trim(txtAdj) = "" Then
       MsgBox ("Give the adjustment amount."), vbInformation, "Medical Rec. Adjustment"
       txtAdj.SetFocus
       Exit Sub
    ElseIf CCur(txtAdj) > CCur(lblAmt) Then
       MsgBox ("You are trying to recover an amount greater than what the employee has expended."), vbInformation, "Medical Rec. Adjustment"
       txtAdj.SetFocus
       Exit Sub
    ElseIf CCur(txtAdj) < 0 Then
       MsgBox ("Cannot make adjustment to an amount less than 0"), vbInformation, "Medical Rec. Adjustment"
       txtAdj.SetFocus
       Exit Sub
    Else
       ' check if a wage linkage is made
       Set datExtra = New Recordset
       datExtra.Open "SELECT * FROM med_opts WHERE co_code ='" & ID & _
       "'", cn, adOpenStatic, adLockOptimistic
       If datExtra.RecordCount <> 0 Then
          If Trim(datExtra!link_wage) = "" Then
             MsgBox ("Please assign a link wage to this medical scheme under 'Utilites-Options'"), vbInformation, "Medical Rec. Adjustment"
             Exit Sub
          Else
             lkWage = Trim(datExtra!link_wage)
             
             Set datExtra = New Recordset
             datExtra.Open "SELECT * FROM Wage_types WHERE wage_code ='" & _
             lkWage & "'", cn, adOpenStatic, adLockOptimistic
             lkname = Trim(datExtra!wage_name)
          End If
       Else
           MsgBox ("Please assign a link wage to this medical scheme under 'Utilites-Options'"), vbInformation, "Medical Rec. Adjustment"
           Exit Sub
       End If
       'A wage linkage has been found
       Set datExtra = New Recordset
       With datExtra
            .Open "SELECT TOP 1 * FROM Transactions WHERE staff_code ='" & _
            frmempmed.pnlStaff_code & "' AND wage_code ='" & _
            lkWage & "' ORDER BY wage_code", cn, adOpenStatic, adLockOptimistic
            If .RecordCount = 0 Then 'No record so add a new wage
               'We cannot make an adjustment on a non-existing record
               MsgBox ("No medical recovery record found for adjustment"), vbInformation, "Medical Rec. Adjustment"
               Exit Sub
            ElseIf !out_source = False Then
               'The entry is not from the HR system
               MsgBox ("No medical recovery record found for adjustment" & _
               "The record found did not originate from this system and cannot be adjusted."), vbInformation, "Medical Rec. Adjustment"
               Exit Sub
            Else
            
                If MsgBox("Post this amount for adjustment to " & lkname, vbYesNo + vbQuestion, "Medical Recovery") = vbNo Then
                   MsgBox ("Adjustment Aborted"), vbInformation, "Medical Recovery"
                   Exit Sub
                End If
                'Find the amount already posted for recovery this month
                Set datExtra2 = New Recordset
                datExtra2.Open "SELECT * FROM med_limits WHERE staff_code ='" & _
                frmempmed.pnlStaff_code & "'", cn, adOpenStatic, adLockOptimistic
                If cboRec.ListIndex = 0 Then
                   'Make an outpost
                   If CCur(txtAdj) < datExtra2!outpost Then
                      'Decrease the amount set for recovery
                      !amount = !amount - CCur(lblAdjAmt)
                      !amount = !amount + CCur(txtAdj)
                      .Update
                      
                      'Decrease the outpost in the med_limits file
                      With datExtra2
                           !outpost = !outpost - CCur(lblAdjAmt)
                           !outpost = !outpost + CCur(txtAdj)
                           
                           !outrec = !outrec - CCur(lblAdjAmt)
                           !outrec = !outrec + CCur(txtAdj)
                           
                           .Update
                      End With
                   Else
                      'Adjust to put an amount greater than original amt
                      exAmt = CCur(txtAdj) - lblAdjAmt   'Find the excess
                      !amount = !amount + exAmt
                      .Update
                      
                      'Increase outpost
                      datExtra2!outpost = datExtra2!outpost + exAmt
                      datExtra2!outrec = datExtra2!outrec + exAmt
                      datExtra2.Update
                   End If
                Else
                   'Make an outpost
                   If CCur(txtAdj) < datExtra2!inpost Then
                      'Decrease the amount set for recovery
                      !amount = !amount - CCur(lblAdjAmt)
                      !amount = !amount + CCur(txtAdj)
                      .Update
                      
                      'Decrease the outpost in the med_limits file
                      With datExtra2
                           !inpost = !inpost - CCur(lblAdjAmt)
                           !inpost = !inpost + CCur(txtAdj)
                           
                           !inrec = !inrec - CCur(lblAdjAmt)
                           !inrec = !inrec + CCur(txtAdj)
                           
                           .Update
                      End With
                   Else
                      'Adjust to put an amount greater than original amt
                      exAmt = CCur(txtAdj) - lblAdjAmt   'Find the excess
                      !amount = !amount + exAmt
                      .Update
                      
                      'Increase outpost
                      datExtra2!inpost = datExtra2!inpost + exAmt
                      datExtra2!inrec = datExtra2!inrec + exAmt
                      datExtra2.Update
                   End If
                End If
                'If inpost = 0 and outpost = 0 then check if Transactions file still shows
                'a value, if 0 delete else change out source to false
                If datExtra2!outpost = 0 And datExtra2!inpost = 0 Then
                   If !amount = 0 Then
                      .Delete
                      .Requery
                   Else
                      !out_source = False
                      .Update
                   End If
                End If
            End If
       End With
       
        'Amount posted, show this as posted in the medical limits file
       
       
       
       MsgBox ("Amount Posted for Recovery Adjusted"), vbInformation, "Medical Rec. Adjustment"
       Unload Me
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    cboRec.Text = cboRec.List(0)
    
End Sub

Private Sub txtAdj_LostFocus()
    If Trim(txtAdj) <> "" Then
        If Not IsNumeric(txtAdj) Then
           MsgBox ("Invalid Numeric Value"), vbInformation, "Medical Rec. Adjustment"
           txtAdj.SetFocus
           SendKeys "{HOME}"
           SendKeys "+{END}"
        End If
    End If
End Sub
