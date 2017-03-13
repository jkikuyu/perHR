VERSION 5.00
Begin VB.Form frmMRec 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Make Medical Recovery"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4260
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMRec.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   4260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Select a Recovery to make"
      Height          =   765
      Left            =   30
      TabIndex        =   4
      Top             =   0
      Width           =   4230
      Begin VB.ComboBox cboRec 
         Height          =   345
         ItemData        =   "frmMRec.frx":0442
         Left            =   135
         List            =   "frmMRec.frx":044C
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   300
         Width           =   3960
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   330
      Left            =   3210
      TabIndex        =   3
      Top             =   1815
      Width           =   1035
   End
   Begin VB.CommandButton cmdRec 
      Caption         =   "&Recover"
      Height          =   330
      Left            =   2100
      TabIndex        =   2
      Top             =   1815
      Width           =   1035
   End
   Begin VB.TextBox txtRec 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   330
      Left            =   1965
      TabIndex        =   1
      Top             =   1335
      Width           =   1425
   End
   Begin VB.Label lblAmt 
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   2475
      TabIndex        =   7
      Top             =   855
      Width           =   1500
   End
   Begin VB.Label Label2 
      Caption         =   "Amount Pending Recovery"
      Height          =   255
      Left            =   75
      TabIndex        =   6
      Top             =   900
      Width           =   2280
   End
   Begin VB.Label Label1 
      Caption         =   "Amount to Recover"
      Height          =   300
      Left            =   60
      TabIndex        =   0
      Top             =   1320
      Width           =   1830
   End
End
Attribute VB_Name = "frmMRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cboRec_Click()
    If cboRec.ListIndex = 0 Then
       lblAmt = frmempmed.txtOutRecbal
    ElseIf cboRec.ListIndex = 1 Then
       lblAmt = frmempmed.txtInRecBal
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdRec_Click()
    Dim lkWage As String
    Dim lkname As String
    If CCur(txtRec) > CCur(lblAmt) Then
       MsgBox ("You are trying to recover an amount greater than what the employee has expended."), vbInformation, "Medical Recovery"
       Exit Sub
    Else
       Set datExtra = New Recordset
       datExtra.Open "SELECT * FROM med_opts WHERE co_code ='" & ID & _
       "'", cn, adOpenStatic, adLockOptimistic
       If datExtra.RecordCount <> 0 Then
          If Trim(datExtra!link_wage) = "" Then
             MsgBox ("Please assign a link wage to this medical scheme under 'Utilites-Options'"), vbInformation, "Medical Recovery"
             Exit Sub
          Else
             lkWage = Trim(datExtra!link_wage)
             
             Set datExtra = New Recordset
             datExtra.Open "SELECT * FROM Wage_types WHERE wage_code ='" & _
             lkWage & "'", cn, adOpenStatic, adLockOptimistic
             lkname = Trim(datExtra!wage_name)
          End If
       Else
           MsgBox ("Please assign a link wage to this medical scheme under 'Utilites-Options'"), vbInformation, "Medical Recovery"
           Exit Sub
       End If
       Set datExtra = New Recordset
       With datExtra
            .Open "SELECT TOP 1 * FROM Transactions WHERE staff_code ='" & _
            frmempmed.pnlStaff_code & "' AND wage_code ='" & _
            lkWage & "' ORDER BY wage_code", cn, adOpenStatic, adLockOptimistic
            If MsgBox("Post this amount for recovery to " & lkname, vbYesNo + vbQuestion, "Medical Recovery") = vbNo Then
               MsgBox ("Recovery Aborted"), vbInformation, "Medical Recovery"
               Exit Sub
            End If
            cn.BeginTrans
            If datExtra.RecordCount = 0 Then 'No record so add a new wage
               .AddNew
               !co_code = ID
               !staff_code = frmempmed.pnlStaff_code
               !wage_code = lkWage
               !amount = CCur(txtRec)
               !date_input = Format(Now, "dd/mmm/yyyy")
               !out_source = True
               .Update
            Else
               'Record found, increment the value
               !amount = !amount + CCur(txtRec)
               !date_input = Format(Now, "dd/mmm/yyyy")
               !out_source = True
               .Update
            End If
       End With
        'Amount posted, show this as posted in the medical limits file
        Set datExtra = New Recordset
        datExtra.Open "SELECT * FROM med_limits WHERE staff_code ='" & _
        frmempmed.pnlStaff_code & "'", cn, adOpenStatic, adLockOptimistic
        If cboRec.ListIndex = 0 Then
           'Make an outpost
           datExtra!outpost = datExtra!outpost + CCur(txtRec)
           datExtra!outrec = datExtra!outrec + CCur(txtRec)
        Else
           'Make an inpost
           datExtra!inpost = datExtra!inpost + CCur(txtRec)
           datExtra!inrec = datExtra!inrec + CCur(txtRec)
        End If
        datExtra.Update
        cn.CommitTrans
               
       
       
       MsgBox ("Amount Posted for recovery"), vbInformation, "Medical Recovery"
       Unload Me
    End If
       
       
End Sub

Private Sub Form_Load()
    cboRec.Text = cboRec.List(0)
End Sub
