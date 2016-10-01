VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmqrtchng 
   Caption         =   "Product Rate Change"
   ClientHeight    =   7995
   ClientLeft      =   1230
   ClientTop       =   1845
   ClientWidth     =   11580
   LinkTopic       =   "Form1"
   ScaleHeight     =   7995
   ScaleWidth      =   11580
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   4950
      Left            =   4515
      TabIndex        =   1
      Top             =   120
      Width           =   7035
      Begin VB.TextBox TxtName 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1350
         MaxLength       =   50
         TabIndex        =   12
         Top             =   300
         Width           =   5370
      End
      Begin VB.TextBox txtGrp 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1350
         TabIndex        =   11
         Top             =   795
         Width           =   5370
      End
      Begin VB.ComboBox cboUof 
         Height          =   315
         Left            =   1680
         TabIndex        =   10
         Text            =   "Combo1"
         Top             =   1890
         Width           =   1455
      End
      Begin VB.TextBox txtPack 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   5280
         TabIndex        =   9
         Top             =   1935
         Width           =   1335
      End
      Begin VB.TextBox txtRL 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1260
         TabIndex        =   8
         Top             =   2640
         Width           =   1335
      End
      Begin VB.TextBox txtTax 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   5280
         TabIndex        =   7
         Top             =   2730
         Width           =   1335
      End
      Begin VB.TextBox txtSr 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1260
         TabIndex        =   6
         Top             =   3525
         Width           =   1335
      End
      Begin VB.TextBox txtPrt 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   5265
         TabIndex        =   5
         Top             =   3495
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cancel"
         Height          =   420
         Left            =   5430
         TabIndex        =   4
         Top             =   4395
         Width           =   1260
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Apply"
         Height          =   420
         Left            =   4080
         TabIndex        =   3
         Top             =   4410
         Width           =   1260
      End
      Begin VB.TextBox TxtBar 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1365
         MaxLength       =   20
         TabIndex        =   2
         Top             =   1305
         Width           =   3495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Item Name"
         Height          =   195
         Left            =   285
         TabIndex        =   21
         Top             =   345
         Width           =   765
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Item Group"
         Height          =   195
         Left            =   270
         TabIndex        =   20
         Top             =   840
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unit of Measurement"
         Height          =   195
         Left            =   105
         TabIndex        =   19
         Top             =   1980
         Width           =   1470
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Package"
         Height          =   195
         Left            =   4155
         TabIndex        =   18
         Top             =   1980
         Width           =   645
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reorder Level"
         Height          =   195
         Left            =   105
         TabIndex        =   17
         Top             =   2670
         Width           =   1005
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tax %"
         Height          =   195
         Left            =   4170
         TabIndex        =   16
         Top             =   2745
         Width           =   435
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Selling Rate"
         Height          =   195
         Left            =   105
         TabIndex        =   15
         Top             =   3600
         Width           =   855
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Purchase Rate"
         Height          =   195
         Left            =   4125
         TabIndex        =   14
         Top             =   3570
         Width           =   1065
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Barcode"
         Height          =   195
         Left            =   285
         TabIndex        =   13
         Top             =   1380
         Width           =   600
      End
   End
   Begin MSFlexGridLib.MSFlexGrid ItemLst 
      Height          =   7425
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4290
      _ExtentX        =   7567
      _ExtentY        =   13097
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      GridLinesFixed  =   1
      Appearance      =   0
      FormatString    =   "Product  Name                                  | Code |  Barcode"
   End
End
Attribute VB_Name = "frmqrtchng"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
