VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmLAna 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Leave Analysis Results"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6915
   Icon            =   "frmLAna.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   6915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox rtf 
      Height          =   4140
      Left            =   0
      TabIndex        =   0
      Top             =   -15
      Width           =   6930
      _ExtentX        =   12224
      _ExtentY        =   7303
      _Version        =   393217
      TextRTF         =   $"frmLAna.frx":0442
   End
End
Attribute VB_Name = "frmLAna"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    'Load the text file
    rtf.LoadFile App.Path + "\Leave.txt", rtfText
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Delete the file
    Kill App.Path + "\Leave.txt"
End Sub
