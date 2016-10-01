VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmVoucher 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Voucher"
   ClientHeight    =   5550
   ClientLeft      =   2325
   ClientTop       =   1770
   ClientWidth     =   7170
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5550
   ScaleWidth      =   7170
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   5430
      Left            =   90
      TabIndex        =   0
      Top             =   45
      Width           =   6915
      Begin VB.Frame Frame2 
         BackColor       =   &H8000000B&
         BorderStyle     =   0  'None
         Height          =   3075
         Left            =   1905
         TabIndex        =   17
         Top             =   1410
         Width           =   4935
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2535
            TabIndex        =   20
            Top             =   2640
            Width           =   1065
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3615
            TabIndex        =   19
            Top             =   2640
            Width           =   1080
         End
         Begin MSFlexGridLib.MSFlexGrid FGBill 
            Height          =   2505
            Left            =   90
            TabIndex        =   18
            Top             =   75
            Width           =   4740
            _ExtentX        =   8361
            _ExtentY        =   4419
            _Version        =   393216
            Cols            =   5
            FixedCols       =   0
            FocusRect       =   0
            HighLight       =   2
            GridLinesFixed  =   1
            ScrollBars      =   2
            Appearance      =   0
            FormatString    =   "Bill No  |  Bill Date|      Amount  |     Balance    |             Pay"
         End
      End
      Begin VB.TextBox txtCashBank 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1425
         MaxLength       =   50
         TabIndex        =   15
         Top             =   1695
         Width           =   5370
      End
      Begin VB.ComboBox cboVrtye 
         Height          =   315
         ItemData        =   "frmVoucher.frx":0000
         Left            =   1425
         List            =   "frmVoucher.frx":0010
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   315
         Width           =   1185
      End
      Begin VB.TextBox txtVrdate 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1455
         MaxLength       =   5
         TabIndex        =   6
         Top             =   915
         Width           =   1125
      End
      Begin VB.TextBox TxtName 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   5
         Top             =   2700
         Width           =   5370
      End
      Begin VB.TextBox txtAmount 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1425
         MaxLength       =   5
         TabIndex        =   4
         Top             =   3270
         Width           =   1125
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Save"
         Height          =   420
         Left            =   4230
         TabIndex        =   3
         Top             =   4755
         Width           =   1260
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cancel"
         Height          =   420
         Left            =   5550
         TabIndex        =   2
         Top             =   4740
         Width           =   1260
      End
      Begin VB.TextBox txtNarr 
         Appearance      =   0  'Flat
         Height          =   615
         Left            =   1425
         MaxLength       =   5
         TabIndex        =   1
         Top             =   3780
         Width           =   5355
      End
      Begin MSMask.MaskEdBox mskVrDt 
         Height          =   285
         Left            =   5310
         TabIndex        =   11
         Top             =   855
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Format          =   "dd-mmm-yy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cash\Bank"
         Height          =   195
         Left            =   195
         TabIndex        =   16
         Top             =   1755
         Width           =   810
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Voucher Type"
         Height          =   195
         Left            =   285
         TabIndex        =   14
         Top             =   390
         Width           =   1005
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Voucher Date"
         Height          =   195
         Left            =   4155
         TabIndex        =   12
         Top             =   900
         Width           =   990
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Voucher No."
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   960
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account Name"
         Height          =   195
         Left            =   210
         TabIndex        =   9
         Top             =   2760
         Width           =   1065
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         Height          =   195
         Left            =   210
         TabIndex        =   8
         Top             =   3315
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Narration"
         Height          =   195
         Left            =   210
         TabIndex        =   7
         Top             =   3825
         Width           =   645
      End
   End
End
Attribute VB_Name = "frmVoucher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Change()

End Sub


Private Sub MaskEdBox1_Change()

End Sub


Private Sub Form_Load()
cboVrtye.ListIndex = 0
End Sub


