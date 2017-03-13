VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmempmed 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Employee Medical Entitlements"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9210
   Icon            =   "frmempmed.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   9210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtnotes 
      DataField       =   "notes"
      Height          =   900
      Left            =   15
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   17
      Top             =   5670
      Width           =   9150
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   9210
      TabIndex        =   12
      Top             =   6615
      Width           =   9210
      Begin VB.CommandButton cmdMRec 
         Caption         =   "&Make Recovery"
         Height          =   330
         Left            =   30
         TabIndex        =   63
         Top             =   30
         Width           =   1470
      End
      Begin VB.CommandButton cmdARec 
         Caption         =   "Adjust Recovery"
         Height          =   330
         Left            =   1545
         TabIndex        =   62
         Top             =   30
         Width           =   1470
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   350
         Left            =   8070
         TabIndex        =   16
         Top             =   30
         Width           =   1095
      End
      Begin VB.CommandButton cmdModify 
         Caption         =   "&Modify"
         Height          =   350
         Left            =   4695
         TabIndex        =   15
         Top             =   30
         Width           =   1095
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   350
         Left            =   5820
         TabIndex        =   14
         Top             =   30
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "C&ancel"
         Height          =   350
         Left            =   6945
         TabIndex        =   13
         Top             =   30
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dependants Covered"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1800
      Left            =   0
      TabIndex        =   7
      Top             =   660
      Width           =   9210
      Begin MSDataListLib.DataList lstc 
         Height          =   1230
         Left            =   4935
         TabIndex        =   35
         Top             =   495
         Width           =   4185
         _ExtentX        =   7382
         _ExtentY        =   2170
         _Version        =   393216
      End
      Begin MSDataListLib.DataList lstnc 
         Height          =   1230
         Left            =   105
         TabIndex        =   34
         Top             =   510
         Width           =   4200
         _ExtentX        =   7408
         _ExtentY        =   2170
         _Version        =   393216
      End
      Begin VB.CommandButton cmdfRemove 
         Caption         =   "<<"
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
         Left            =   4395
         TabIndex        =   11
         ToolTipText     =   "Remove Dependant from Covered Group"
         Top             =   1125
         Width           =   435
      End
      Begin VB.CommandButton cmdfAdd 
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4395
         TabIndex        =   10
         ToolTipText     =   "Add this Dependant as Covered"
         Top             =   645
         Width           =   435
      End
      Begin VB.Label Label3 
         Caption         =   "Covered"
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
         Left            =   5325
         TabIndex        =   9
         Top             =   225
         Width           =   2505
      End
      Begin VB.Label Label2 
         Caption         =   "Not Covered"
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
         Left            =   105
         TabIndex        =   8
         Top             =   255
         Width           =   2220
      End
   End
   Begin VB.CommandButton cmdFind 
      Height          =   375
      Left            =   5595
      Picture         =   "frmempmed.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   270
      Width           =   375
   End
   Begin Threed.SSPanel pnlother_names 
      Height          =   375
      Left            =   2910
      TabIndex        =   1
      Top             =   270
      Width           =   2655
      _Version        =   65536
      _ExtentX        =   4683
      _ExtentY        =   661
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
   End
   Begin Threed.SSPanel pnlLast_name 
      Height          =   375
      Left            =   975
      TabIndex        =   2
      Top             =   270
      Width           =   1920
      _Version        =   65536
      _ExtentX        =   3387
      _ExtentY        =   661
      _StockProps     =   15
      BackColor       =   12632256
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
   Begin Threed.SSPanel pnlStaff_code 
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   270
      Width           =   975
      _Version        =   65536
      _ExtentX        =   1720
      _ExtentY        =   661
      _StockProps     =   15
      BackColor       =   12632256
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   2880
      Left            =   0
      TabIndex        =   19
      Top             =   2520
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   5080
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "Medical Limits"
      TabPicture(0)   =   "frmempmed.frx":05C4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Out-Patient Medical Transactions"
      TabPicture(1)   =   "frmempmed.frx":05E0
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grdMTrans"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "SSPanel1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "SSPanel2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "SSPanel3"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "SSPanel4"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "SSPanel5"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "SSPanel7"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "In-Patient Medical Transactions"
      TabPicture(2)   =   "frmempmed.frx":05FC
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "SSPanel9"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "SSPanel8"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "SSPanel6"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "grdInTrans"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "SSPanel14"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "SSPanel13"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "SSPanel12"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).ControlCount=   7
      Begin VB.Frame Frame3 
         Caption         =   "In-Patient"
         Height          =   2445
         Left            =   4605
         TabIndex        =   37
         Top             =   375
         Width           =   4500
         Begin VB.TextBox txtInRecBal 
            Alignment       =   1  'Right Justify
            DataField       =   "amtrec"
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
            Height          =   300
            Left            =   2745
            Locked          =   -1  'True
            TabIndex        =   60
            Top             =   1905
            Width           =   1710
         End
         Begin VB.TextBox txtInbf 
            Alignment       =   1  'Right Justify
            DataField       =   "inrecbf"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   300
            Left            =   2745
            Locked          =   -1  'True
            TabIndex        =   57
            Top             =   1275
            Width           =   1710
         End
         Begin VB.TextBox txtinRec 
            Alignment       =   1  'Right Justify
            DataField       =   "inrec"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   300
            Left            =   2745
            Locked          =   -1  'True
            TabIndex        =   56
            Top             =   1590
            Width           =   1710
         End
         Begin VB.TextBox txtInbal 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   300
            Left            =   2730
            Locked          =   -1  'True
            TabIndex        =   54
            Top             =   885
            Width           =   1710
         End
         Begin VB.TextBox txtInTotal 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   300
            Left            =   2745
            Locked          =   -1  'True
            TabIndex        =   50
            Top             =   555
            Width           =   1680
         End
         Begin VB.TextBox txtInLimit 
            Alignment       =   1  'Right Justify
            DataField       =   "inlimit"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   300
            Left            =   2745
            TabIndex        =   48
            Top             =   210
            Width           =   1695
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            Caption         =   "Balance to be Recovered"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004040&
            Height          =   300
            Left            =   75
            TabIndex        =   61
            Top             =   1905
            Width           =   2610
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            Caption         =   "Amount to be Recovered b/f"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004040&
            Height          =   300
            Left            =   90
            TabIndex        =   59
            Top             =   1260
            Width           =   2580
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            Caption         =   "Amount already Recovered"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004040&
            Height          =   300
            Left            =   75
            TabIndex        =   58
            Top             =   1590
            Width           =   2610
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            Caption         =   "Balance"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004040&
            Height          =   255
            Left            =   75
            TabIndex        =   55
            Top             =   900
            Width           =   2610
         End
         Begin VB.Label Label11 
            Caption         =   "In-Patient Amount Expended"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004040&
            Height          =   270
            Left            =   285
            TabIndex        =   51
            Top             =   570
            Width           =   2550
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            Caption         =   "In-Patient Medical Allowance"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004040&
            Height          =   300
            Left            =   75
            TabIndex        =   49
            Top             =   240
            Width           =   2625
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Out Patient"
         Height          =   2445
         Left            =   60
         TabIndex        =   36
         Top             =   375
         Width           =   4500
         Begin VB.TextBox txtOutRecbal 
            Alignment       =   1  'Right Justify
            DataField       =   "amtrec"
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
            Height          =   300
            Left            =   2730
            Locked          =   -1  'True
            TabIndex        =   52
            Top             =   1920
            Width           =   1710
         End
         Begin VB.TextBox txtoutbal 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   300
            Left            =   2730
            Locked          =   -1  'True
            TabIndex        =   46
            Top             =   885
            Width           =   1710
         End
         Begin VB.TextBox txtoutRec 
            Alignment       =   1  'Right Justify
            DataField       =   "outrec"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   300
            Left            =   2730
            Locked          =   -1  'True
            TabIndex        =   43
            Top             =   1605
            Width           =   1710
         End
         Begin VB.TextBox txtoutbf 
            Alignment       =   1  'Right Justify
            DataField       =   "outrecbf"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   300
            Left            =   2730
            Locked          =   -1  'True
            TabIndex        =   42
            Top             =   1290
            Width           =   1710
         End
         Begin VB.TextBox txtOutTotal 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   285
            Left            =   2745
            Locked          =   -1  'True
            TabIndex        =   40
            Top             =   585
            Width           =   1695
         End
         Begin VB.TextBox txtOutlimit 
            Alignment       =   1  'Right Justify
            DataField       =   "outlimit"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   300
            Left            =   2745
            Locked          =   -1  'True
            TabIndex        =   38
            Top             =   270
            Width           =   1695
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            Caption         =   "Balance to be Recovered"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004040&
            Height          =   300
            Left            =   60
            TabIndex        =   53
            Top             =   1920
            Width           =   2610
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "Balance"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004040&
            Height          =   255
            Left            =   75
            TabIndex        =   47
            Top             =   900
            Width           =   2610
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            Caption         =   "Amount already Recovered"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004040&
            Height          =   300
            Left            =   60
            TabIndex        =   45
            Top             =   1605
            Width           =   2610
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            Caption         =   "Amount to be Recovered b/f"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004040&
            Height          =   300
            Left            =   75
            TabIndex        =   44
            Top             =   1275
            Width           =   2580
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            Caption         =   "Out-Patient Amount Expended"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004040&
            Height          =   270
            Left            =   75
            TabIndex        =   41
            Top             =   585
            Width           =   2580
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Out-Patient Medical Allowance"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004040&
            Height          =   300
            Left            =   60
            TabIndex        =   39
            Top             =   285
            Width           =   2640
         End
      End
      Begin Threed.SSPanel SSPanel7 
         Height          =   345
         Left            =   -67650
         TabIndex        =   20
         Top             =   405
         Width           =   1245
         _Version        =   65536
         _ExtentX        =   2196
         _ExtentY        =   609
         _StockProps     =   15
         Caption         =   "Amount"
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.01
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   345
         Left            =   -68745
         TabIndex        =   21
         Top             =   405
         Width           =   1080
         _Version        =   65536
         _ExtentX        =   1905
         _ExtentY        =   609
         _StockProps     =   15
         Caption         =   "Cheque"
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.01
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   345
         Left            =   -69810
         TabIndex        =   22
         Top             =   405
         Width           =   1050
         _Version        =   65536
         _ExtentX        =   1852
         _ExtentY        =   609
         _StockProps     =   15
         Caption         =   "Receipt"
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.01
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   345
         Left            =   -70890
         TabIndex        =   23
         Top             =   405
         Width           =   1065
         _Version        =   65536
         _ExtentX        =   1879
         _ExtentY        =   609
         _StockProps     =   15
         Caption         =   "Invoice"
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.01
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   345
         Left            =   -73635
         TabIndex        =   24
         Top             =   405
         Width           =   2730
         _Version        =   65536
         _ExtentX        =   4815
         _ExtentY        =   609
         _StockProps     =   15
         Caption         =   "Source"
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.01
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   345
         Left            =   -74685
         TabIndex        =   25
         Top             =   405
         Width           =   1035
         _Version        =   65536
         _ExtentX        =   1826
         _ExtentY        =   609
         _StockProps     =   15
         Caption         =   "Date"
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.01
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataGridLib.DataGrid grdMTrans 
         Height          =   2040
         Left            =   -74955
         TabIndex        =   26
         Top             =   765
         Width           =   9090
         _ExtentX        =   16034
         _ExtentY        =   3598
         _Version        =   393216
         AllowUpdate     =   -1  'True
         BackColor       =   -2147483624
         ColumnHeaders   =   0   'False
         HeadLines       =   1
         RowHeight       =   20
         FormatLocked    =   -1  'True
         AllowAddNew     =   -1  'True
         AllowDelete     =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   9
         BeginProperty Column00 
            DataField       =   "staff_code"
            Caption         =   "staff_code"
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
            DataField       =   "date"
            Caption         =   "date"
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
            DataField       =   "from"
            Caption         =   "from"
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
         BeginProperty Column03 
            DataField       =   "inv"
            Caption         =   "inv"
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
         BeginProperty Column04 
            DataField       =   "rec"
            Caption         =   "rec"
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
         BeginProperty Column05 
            DataField       =   "chq"
            Caption         =   "chq"
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
         BeginProperty Column06 
            DataField       =   "amt"
            Caption         =   "amt"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column07 
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
         BeginProperty Column08 
            DataField       =   "year"
            Caption         =   "year"
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
               ColumnWidth     =   915.024
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1289.764
            EndProperty
            BeginProperty Column06 
               Alignment       =   1
               ColumnWidth     =   1289.764
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   689.953
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   840.189
            EndProperty
         EndProperty
      End
      Begin Threed.SSPanel SSPanel12 
         Height          =   345
         Left            =   -70890
         TabIndex        =   27
         Top             =   405
         Width           =   1065
         _Version        =   65536
         _ExtentX        =   1879
         _ExtentY        =   609
         _StockProps     =   15
         Caption         =   "Invoice"
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.01
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSPanel SSPanel13 
         Height          =   345
         Left            =   -73635
         TabIndex        =   28
         Top             =   405
         Width           =   2730
         _Version        =   65536
         _ExtentX        =   4815
         _ExtentY        =   609
         _StockProps     =   15
         Caption         =   "Source"
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.01
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSPanel SSPanel14 
         Height          =   345
         Left            =   -74685
         TabIndex        =   29
         Top             =   405
         Width           =   1035
         _Version        =   65536
         _ExtentX        =   1826
         _ExtentY        =   609
         _StockProps     =   15
         Caption         =   "Date"
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.01
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataGridLib.DataGrid grdInTrans 
         Height          =   2040
         Left            =   -74955
         TabIndex        =   30
         Top             =   765
         Width           =   9090
         _ExtentX        =   16034
         _ExtentY        =   3598
         _Version        =   393216
         AllowUpdate     =   -1  'True
         BackColor       =   -2147483624
         ColumnHeaders   =   0   'False
         HeadLines       =   1
         RowHeight       =   20
         FormatLocked    =   -1  'True
         AllowAddNew     =   -1  'True
         AllowDelete     =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   9
         BeginProperty Column00 
            DataField       =   "staff_code"
            Caption         =   "staff_code"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "date"
            Caption         =   "date"
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
            DataField       =   "from"
            Caption         =   "from"
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
         BeginProperty Column03 
            DataField       =   "inv"
            Caption         =   "inv"
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
         BeginProperty Column04 
            DataField       =   "rec"
            Caption         =   "rec"
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
         BeginProperty Column05 
            DataField       =   "chq"
            Caption         =   "chq"
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
         BeginProperty Column06 
            DataField       =   "amt"
            Caption         =   "amt"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column07 
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
         BeginProperty Column08 
            DataField       =   "year"
            Caption         =   "year"
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
               ColumnWidth     =   915.024
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1289.764
            EndProperty
            BeginProperty Column06 
               Alignment       =   1
               ColumnWidth     =   1289.764
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   689.953
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   840.189
            EndProperty
         EndProperty
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   345
         Left            =   -67650
         TabIndex        =   31
         Top             =   405
         Width           =   1245
         _Version        =   65536
         _ExtentX        =   2196
         _ExtentY        =   609
         _StockProps     =   15
         Caption         =   "Amount"
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.01
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSPanel SSPanel8 
         Height          =   345
         Left            =   -68745
         TabIndex        =   32
         Top             =   405
         Width           =   1080
         _Version        =   65536
         _ExtentX        =   1905
         _ExtentY        =   609
         _StockProps     =   15
         Caption         =   "Cheque"
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.01
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSPanel SSPanel9 
         Height          =   345
         Left            =   -69810
         TabIndex        =   33
         Top             =   405
         Width           =   1050
         _Version        =   65536
         _ExtentX        =   1852
         _ExtentY        =   609
         _StockProps     =   15
         Caption         =   "Receipt"
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.01
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Label Label9 
      Caption         =   "Additional Notes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   60
      TabIndex        =   18
      Top             =   5445
      Width           =   2745
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000001&
      Caption         =   " Staff Code"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   975
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000001&
      Caption         =   " Last Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   960
      TabIndex        =   5
      Top             =   0
      Width           =   2175
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000001&
      Caption         =   " Other Names"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   3120
      TabIndex        =   4
      Top             =   0
      Width           =   2415
   End
End
Attribute VB_Name = "frmempmed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim datOutTrans As Recordset
Dim datInTrans As Recordset
Dim datKin As Recordset
Dim datMed As Recordset
Dim datlstnc As Recordset
Dim datlstc As Recordset

Private Sub cmdARec_Click()
    If datMed.RecordCount = 0 Then
       MsgBox ("Medical Entitlements not defined."), vbInformation, "Medical Entitlements"
       Exit Sub
    End If
    'Get the amount already posted for out patient recovery
    Set datExtra = New Recordset
    datExtra.Open "SELECT * FROM med_limits WHERE staff_code ='" & _
    Trim(pnlStaff_code) & "'", cn, adOpenStatic, adLockOptimistic
    If datExtra.RecordCount <> 0 Then
       frmARec.lblAmt = datExtra!outpost + txtOutRecbal
       frmARec.lblAdjAmt = datExtra!outpost
       frmARec.Show vbModal
       datMed.Requery
       CalcTotal
    End If
    
End Sub

Private Sub cmdCancel_Click()
    Reset True
End Sub

Private Sub cmdClose_Click()
    Unload Me
    
End Sub

Private Sub cmdFind_Click()
    frmSelectEmp.Left = Me.Left + 20
    frmSelectEmp.Top = Me.Top + (Frame1.Top - Label6.Top) + pnlStaff_code.Height - 50
    frmSelectEmp.Show vbModal
    If pnlStaff_code = "" Then
       Exit Sub
    End If
    
    datMed.Filter = "Staff_code ='" & pnlStaff_code & "'"
    datOutTrans.Filter = "Staff_code ='" & pnlStaff_code & "'"
    datInTrans.Filter = "Staff_code ='" & pnlStaff_code & "'"
    CalcTotal
    'Get this employee dependants
    'Set datExtra = New Recordset
    'datExtra.Open "SELECT * FROM family_data WHERE co_code ='" & ID & _
    '"' AND Staff_code ='" & pnlStaff_code & "'", cn, adOpenStatic, adLockOptimistic
    'If datExtra.RecordCount <> 0 Then
    '   While Not datExtra.EOF
    '     lstnc.AddItem Trim(datExtra!f_lName) & " " & Trim(datExtra!f_fname)
    '     datExtra.MoveNext
    '   Wend
    'End If
    
End Sub

Private Sub cmdModify_Click()
    If Trim(pnlStaff_code) = "" Then
       MsgBox ("No Employee selected."), vbInformation, "Medical Entitlements"
       Exit Sub
    End If
    If datMed.RecordCount = 0 Then
       datMed.AddNew
    Else
        'Just Edit the Current Record
    End If
    txtnotes.Locked = False
    txtOutlimit.Locked = False
    txtInLimit.Locked = False
    Reset False
    
End Sub

Private Sub cmdMRec_Click()
    If datMed.RecordCount = 0 Then
       MsgBox ("Medical Entitlements not defined."), vbInformation, "Medical Entitlements"
       Exit Sub
    End If
    frmMRec.lblAmt = txtOutRecbal
    frmMRec.Show vbModal
    datMed.Requery
    CalcTotal
End Sub

Private Sub cmdRecover_Click()
    
End Sub

Private Sub cmdSave_Click()
    'Store the new limit and notes for this employee
    'Find his/her medical limit
    Set datExtra = New Recordset
    datExtra.Open "SELECT * FROM med_limits WHERE co_code ='" & ID & _
    "' AND staff_code ='" & pnlStaff_code & "' AND Year=" & _
    Operation_Year, cn, adOpenStatic, adLockOptimistic
    
    If datExtra.RecordCount <> 0 Then
       datMed!co_code = ID
       datMed.Update
       bk = datMed.Bookmark
       datMed.Bookmark = bk
       
       'datExtra.MoveLast
       'Check if the out patient limit amount shown is the same as the previous amount
       'If CCur(datExtra!outlimit) = CCur(txtOutlimit) And CCur(datExtra!inlimit) = CCur(txtInLimit) Then
       '   'Store in the edit of this record
       '   datExtra!notes = Trim(txtnotes)
       '   datExtra.Update
       '   bk = datExtra.Bookmark
       '   datExtra.Bookmark = bk
       '   Exit Sub
       'Else
       '   x = CLng(datExtra!entry_code) + 1
          'Add a new record of the medical limit
       '   datExtra.AddNew
       '   datExtra!entry_code = CLng(x)
       '   datExtra!co_code = ID
       '   datExtra!staff_code = pnlStaff_code
       '   datExtra!outlimit = CCur(txtOutlimit)
       '   datExtra!inlimit = CCur(txtInLimit)
       '   datExtra!notes = Trim(txtnotes)
       '   datExtra!Date = Format(Now, "dd/mmm/yyyy")
       '   datExtra!Year = Operation_Year
       '   datExtra.Update
       'End If
    Else
        'A record has never been made for this employee so make it
          'datMed.AddNew
          datMed!co_code = ID
          datMed!staff_code = pnlStaff_code
          'datMed!outlimit = CCur(txtOutlimit)
          'datMed!inlimit = CCur(txtInLimit)
          'datMed!notes = Trim(txtnotes)
          datMed!Date = Format(Now, "dd/mmm/yyyy")
          datMed!Year = Operation_Year
          datMed.Update
    End If
    txtOutlimit.Locked = True
    txtInLimit.Locked = True
    txtnotes.Locked = True
    CalcTotal
    Reset True
End Sub

Private Sub cmdfAdd_Click()
    'This procedure add the dependant as covered by the medical scheme
    
End Sub

Private Sub Form_Load()
    PMed = True
    Set datEmp = New Recordset
    datEmp.Open "SELECT * FROM Personal_data WHERE co_code ='" & _
    ID & "'", cn, adOpenStatic, adLockOptimistic
    
    Set datOutTrans = New Recordset
    datOutTrans.Open "SELECT staff_code, Date, from, inv, rec, chq, amt, co_code, year, opt FROM emp_mtrans WHERE " & _
    " co_code ='" & ID & "' AND Year =" & Operation_Year & " AND opt = 1 AND From <> 'O/Bal'", cn, adOpenStatic, adLockOptimistic
    
    Set datInTrans = New Recordset
    datInTrans.Open "SELECT staff_code, Date, from, inv, rec, chq, amt, co_code, year, opt FROM emp_mtrans WHERE " & _
    " co_code ='" & ID & "' AND Year =" & Operation_Year & " AND opt = 2 AND From <> 'O/Bal'", cn, adOpenStatic, adLockOptimistic
    
    Set datMed = New Recordset
    datMed.Open "SELECT * FROM med_limits WHERE co_code ='" & ID & _
    "' AND Year =" & Operation_Year, cn, adOpenStatic, adLockOptimistic
    
    'Assign the datasources
    Set txtInLimit.DataSource = datMed
    Set txtOutlimit.DataSource = datMed
    Set txtoutbf.DataSource = datMed
    Set txtoutRec.DataSource = datMed
    Set txtInbf.DataSource = datMed
    Set txtinRec.DataSource = datMed
    Set txtnotes.DataSource = datMed
    
    If datEmp.RecordCount <> 0 Then
       pnlStaff_code = datEmp!staff_code
       pnlLast_name = datEmp!last_name
       If Not IsNull(datEmp!mid_name) Then
           pnlother_names = Trim(datEmp!first_name) & ", " & Trim(datEmp!mid_name)
       Else
           pnlother_names = Trim(datEmp!first_name)
       End If
       'Get this employee dependants
       Set datExtra = New Recordset
       datExtra.Open "SELECT * FROM family_data WHERE co_code ='" & ID & _
       "' AND Staff_code ='" & pnlStaff_code & "'", cn, adOpenStatic, adLockOptimistic
       'If datExtra.RecordCount <> 0 Then
       '   While Not datExtra.EOF
       '     lstnc.AddItem datExtra!last_name & " " & datExtra!first_name
       '     datExtra.MoveNext
       '   Wend
       'End If
       
       datMed.Filter = "Staff_code ='" & datEmp!staff_code & "'"
       datOutTrans.Filter = "Staff_code ='" & datEmp!staff_code & "'"
       datInTrans.Filter = "Staff_code ='" & datEmp!staff_code & "'"
       
       Set grdInTrans.DataSource = datInTrans
       Set grdMTrans.DataSource = datOutTrans
       CalcTotal
    End If
    
    FormatGrid
    Reset True
End Sub

Public Sub FormatGrid()
    With grdMTrans
        .Columns(0).Visible = False
        
        .Columns(1).Width = 1000
        .Columns(2).Width = 2700
        .Columns(3).Width = 1100
        .Columns(4).Width = 1050
        .Columns(5).Width = 1100
        
        .Columns(6).Width = 1200
        .Columns(7).Width = 805
        
        
        .Columns(8).Visible = False
        .Columns(7).Visible = False
    End With
    
    With grdInTrans
        .Columns(0).Visible = False
        
        .Columns(1).Width = 1000
        .Columns(2).Width = 2700
        .Columns(3).Width = 1100
        .Columns(4).Width = 1050
        .Columns(5).Width = 1100
        
        .Columns(6).Width = 1200
        .Columns(7).Width = 805
        
        
        .Columns(8).Visible = False
        .Columns(7).Visible = False
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PMed = False
End Sub

Private Sub grdInTrans_AfterUpdate()
    CalcTotal
End Sub

Private Sub grdInTrans_BeforeUpdate(Cancel As Integer)
On Error Resume Next
    With datInTrans
        !co_code = ID
        !staff_code = Trim(pnlStaff_code)
        !Year = Operation_Year
        !Opt = 2
    End With
End Sub

Private Sub grdMTrans_AfterUpdate()
    CalcTotal
End Sub

Private Sub grdMTrans_BeforeUpdate(Cancel As Integer)
On Error Resume Next
    With datOutTrans
        !co_code = ID
        !staff_code = Trim(pnlStaff_code)
        !Year = Operation_Year
        !Opt = 1
    End With
    
End Sub


Public Sub CalcTotal()
    
    'Get the out patient amount expended
    Set datExtra = New Recordset
    datExtra.Open "SELECT sum(amt) AS 'Amt' FROM emp_mtrans WHERE " & _
    " co_code ='" & ID & "' AND Year =" & Operation_Year & _
    " AND staff_code ='" & pnlStaff_code & "' AND opt =1", cn, adOpenStatic, adLockOptimistic
       
    If datExtra.RecordCount <> 0 Then
       txtOutTotal = datExtra!Amt
    Else
        txtOutTotal = 0
    End If
    'get the balance of the out patient
    If Trim(txtOutlimit) = "" Then
       txtoutbal = 0
    Else
       txtoutbal = CCur(txtOutlimit) - CCur(txtOutTotal)
    End If
    If CCur(txtoutbal) < 0 Then
       txtOutRecbal = (CCur(txtoutbf) - CCur(txtoutbal)) - CCur(txtoutRec)
    Else
       txtOutRecbal = 0
    End If
    
    'If there has been no excessive expenditure, show the bal bf if any
    If txtoutbf <> "" Then
        If CCur(txtOutRecbal) = 0 And CCur(txtoutbf) <> 0 Then
           txtOutRecbal = txtoutbf
        End If
    End If
    
    'Get the in patient amount expended
    Set datExtra = New Recordset
    datExtra.Open "SELECT sum(amt) AS 'Amt' FROM emp_mtrans WHERE " & _
    " co_code ='" & ID & "' AND Year =" & Operation_Year & _
    " AND staff_code ='" & pnlStaff_code & "' AND opt =2", cn, adOpenStatic, adLockOptimistic
       
    If datExtra.RecordCount <> 0 Then
       txtInTotal = datExtra!Amt
    Else
        txtInTotal = 0
    End If
    'get the balance of the in patient
    If Trim(txtInLimit) = "" Then
       txtInbal = 0
    Else
       txtInbal = CCur(txtInLimit) - CCur(txtInTotal)
    End If
    If CCur(txtInbal) < 0 Then
       txtInRecBal = (CCur(txtInbf) - CCur(txtInbal)) - CCur(txtinRec)
    Else
       txtInRecBal = 0
    End If
    
        'If there has been no excessive expenditure, show the bal bf if any
    If txtInbf <> "" Then
        If CCur(txtInRecBal) = 0 And CCur(txtInbf) <> 0 Then
           txtInRecBal = txtInbf
        End If
    End If
End Sub

Public Sub Reset(bval As Boolean)
    cmdModify.Enabled = bval
    cmdCancel.Enabled = Not bval
    cmdSave.Enabled = Not bval
    
    txtOutlimit.Locked = bval
    txtInLimit.Locked = bval
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    If SSTab1.TabIndex = 1 Or SSTab1.TabIndex = 2 Then
       If datMed.RecordCount = 0 Then
          Me.grdInTrans.Enabled = False
          Me.grdMTrans.EditActive = False
       End If
    End If
End Sub

