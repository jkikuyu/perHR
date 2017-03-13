VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRAppr 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reconcile Appraisals"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4995
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRAppr.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   4995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "C&ancel"
      Height          =   375
      Left            =   75
      TabIndex        =   15
      Top             =   1560
      Width           =   945
   End
   Begin VB.CommandButton cmdNext 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4095
      Picture         =   "frmRAppr.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Show Next Record"
      Top             =   1635
      Width           =   375
   End
   Begin VB.CommandButton cmdPrev 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3615
      Picture         =   "frmRAppr.frx":058C
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Show Previous Record"
      Top             =   1635
      Width           =   375
   End
   Begin VB.CommandButton cmdFirst 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3135
      Picture         =   "frmRAppr.frx":06D6
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Move to the First Record"
      Top             =   1635
      Width           =   375
   End
   Begin VB.CommandButton cmdLast 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4575
      Picture         =   "frmRAppr.frx":0820
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Move to the Last Record"
      Top             =   1635
      Width           =   375
   End
   Begin MSComCtl2.DTPicker dtapp 
      DataField       =   "Date"
      Height          =   330
      Left            =   45
      TabIndex        =   7
      Top             =   1065
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   582
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   24576001
      CurrentDate     =   36892
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   390
      Left            =   0
      ScaleHeight     =   390
      ScaleWidth      =   4995
      TabIndex        =   3
      Top             =   2070
      Width           =   4995
      Begin VB.CommandButton cmdNew 
         Caption         =   "&New"
         Height          =   375
         Left            =   60
         TabIndex        =   10
         Top             =   0
         Width           =   945
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   375
         Left            =   1995
         TabIndex        =   9
         Top             =   0
         Width           =   945
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   1035
         TabIndex        =   8
         Top             =   0
         Width           =   945
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   375
         Left            =   4005
         TabIndex        =   5
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdReconcile 
         Caption         =   "&Reconcile"
         Height          =   375
         Left            =   2985
         TabIndex        =   4
         ToolTipText     =   "Starts a new appraisal session and adds the parameters to all employees"
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.TextBox txtapp 
      DataField       =   "app_name"
      Height          =   330
      Index           =   0
      Left            =   60
      TabIndex        =   2
      Top             =   405
      Width           =   4875
   End
   Begin VB.TextBox txtapp 
      DataField       =   "defn_code"
      Height          =   330
      Index           =   1
      Left            =   4140
      TabIndex        =   1
      Top             =   60
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label2 
      Caption         =   "Date of Appraisal"
      Height          =   225
      Left            =   45
      TabIndex        =   6
      Top             =   825
      Width           =   1890
   End
   Begin VB.Label Label1 
      Caption         =   "Give a Definition for this Appraisal"
      Height          =   270
      Left            =   60
      TabIndex        =   0
      Top             =   105
      Width           =   3915
   End
End
Attribute VB_Name = "frmRAppr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim datrApp As Recordset

Private Sub cmdCancel_Click()
    datrApp.CancelUpdate
    reset True
    
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdEdit_Click()
    reset False
    
End Sub

Private Sub cmdfirst_Click()
On Error GoTo PrevErr
    With datrApp
         .MoveFirst
    End With
    Exit Sub
PrevErr:
MsgBox (Err.Description), vbInformation
Err.Clear
End Sub

Private Sub cmdlast_Click()
On Error GoTo PrevErr
    With datrApp
         .MoveLast
    End With
    Exit Sub
PrevErr:
MsgBox (Err.Description), vbInformation
Err.Clear
End Sub

Private Sub cmdNew_Click()
    'Calculate the new defn_code
    Set datExtra = New Recordset
    datExtra.Open "SELECT * FROM apprdefn ORDER BY defn_code ASC", cn, adOpenStatic, adLockOptimistic
    If datExtra.RecordCount = 0 Then
       y = 1
    Else
       datExtra.MoveLast
       y = datExtra!defn_code + 1
    End If

    datrApp.AddNew
    txtApp(1) = y
    reset False
    
End Sub

Private Sub cmdNext_Click()
On Error GoTo NextErr
    With datrApp
        .MoveNext
         If .EOF Then
            .MoveLast
            MsgBox ("End of file"), vbInformation
         End If
    End With
    Exit Sub
NextErr:
MsgBox (Err.Description), vbInformation
    
End Sub

Private Sub cmdPrev_Click()
On Error GoTo PrevErr
    With datrApp
        .MovePrevious
         If .BOF Then
            .MoveFirst
            MsgBox ("End of file"), vbInformation
         End If
    End With
    Exit Sub
PrevErr:
MsgBox (Err.Description), vbInformation
End Sub

Private Sub cmdReconcile_Click()
    Dim datReconcile As Recordset
    Dim datParams As Recordset
    If MsgBox("Add this appraisal definitions to all registered employees?", vbQuestion + vbYesNo, "Appraisals") = vbYes Then
       Set datEmp = New Recordset
       Set datReconcile = New Recordset
       Set datParams = New Recordset
       
       'Check if this appraisal defn has already been assigned
       Set datExtra = New Recordset
       datExtra.Open "SELECT TOP 1 * FROM empappraisals WHERE co_code ='" & _
       ID & "' AND defn_code =" & CLng(datrApp!defn_code) & " ORDER BY co_code", cn, adOpenStatic, adLockOptimistic
       If datExtra.RecordCount <> 0 Then
          If MsgBox("Reconciliation for this appraisal has already been done." & _
          Chr(13) & "The previous reconciliation will be overwritten.", vbOKCancel + vbExclamation, "Appraisal Reconciliation") = vbCancel Then
             Exit Sub
          Else
              cn.Execute "DELETE FROM EmpAppraisals WHERE co_code ='" & _
              ID & "' AND defn_code =" & CLng(datrApp!defn_code)
          End If
       End If
       
       
       datEmp.Open "SELECT * FROM Personal_data WHERE co_code ='" & _
       ID & "'", cn, adOpenStatic, adLockOptimistic
       
       datReconcile.Open "SELECT * FROM empAppraisals", cn, adOpenStatic, adLockOptimistic
       datParams.Open "SELECT * FROM AppParams WHERE co_code ='" & ID & "' AND Valid =1", cn, adOpenStatic, adLockOptimistic
       
       'Check if there are any appraisal parameters for this company
       If datParams.RecordCount <> 0 Then
          'Parameters found, check for the employees
          If datEmp.RecordCount <> 0 Then
             'For each employee, assign the parameters
             
             While Not datEmp.EOF
                  datParams.Requery
                  While Not datParams.EOF
                       With datReconcile
                            .AddNew
                            !co_code = ID
                            !staff_code = Trim(datEmp!staff_code)
                            !defn_code = CLng(datrApp!defn_code)
                            !param_code = CLng(datParams!param_code)
                            !points_awarded = 0
                            .Update
                       End With
                       datParams.MoveNext
                  Wend
                  datEmp.MoveNext
             Wend
          Else
              MsgBox "There are no employees to appraise in this company.", vbInformation, "Appraisals"
              Exit Sub
          End If
       Else
            MsgBox ("There are no valid appraisal parameters defined for this company."), vbInformation, "Appraisals"
            Exit Sub
       End If
       MsgBox "Appraisal Reconciliation completed.", vbInformation, "Appraisals"
    End If
                  
                
       
End Sub

Private Sub cmdUpdate_Click()
    datrApp!co_code = ID
    datrApp.Update
    reset True
    
End Sub

Private Sub Form_Load()
    Dim txt As TextBox
    Set datrApp = New Recordset
    datrApp.Open "SELECT * FROM apprdefn WHERE co_code ='" & ID & "'", cn, adOpenStatic, adLockOptimistic
    
    For Each txt In Me.txtApp
        Set txt.DataSource = datrApp
    Next
    Set dtapp.DataSource = datrApp
    reset True
    
End Sub

Public Sub reset(bval As Boolean)
    'Enable and disbale as appropriate
    Dim txt As TextBox
    cmdNew.Enabled = bval
    cmdEdit.Enabled = bval
    cmdUpdate.Enabled = Not bval
    cmdCancel.Enabled = Not bval
    cmdReconcile.Enabled = bval
    
    For Each txt In Me.txtApp
        txt.Locked = bval
    Next
    dtapp.Enabled = Not bval
    
End Sub

Private Sub txtapp_GotFocus(Index As Integer)
    txtApp(Index) = Trim(txtApp(Index))
End Sub
