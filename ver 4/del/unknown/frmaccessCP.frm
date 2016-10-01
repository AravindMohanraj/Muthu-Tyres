VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmAccCp 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Access Control Panel"
   ClientHeight    =   8655
   ClientLeft      =   2280
   ClientTop       =   1650
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8655
   ScaleWidth      =   11880
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   7770
      Left            =   30
      TabIndex        =   3
      Top             =   0
      Width           =   11790
      Begin VB.CheckBox Check8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Report"
         Height          =   315
         Index           =   2
         Left            =   6075
         TabIndex        =   22
         Top             =   5055
         Width           =   1770
      End
      Begin VB.CheckBox Check14 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Issue"
         Height          =   315
         Left            =   6345
         TabIndex        =   20
         Top             =   3660
         Width           =   1770
      End
      Begin VB.CheckBox Check13 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Receipt"
         Height          =   315
         Left            =   6345
         TabIndex        =   19
         Top             =   4110
         Width           =   1770
      End
      Begin VB.CheckBox Check12 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Purchase"
         Height          =   315
         Left            =   6345
         TabIndex        =   18
         Top             =   4575
         Width           =   1950
      End
      Begin VB.CheckBox Check8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Accounts"
         Height          =   315
         Index           =   1
         Left            =   6045
         TabIndex        =   16
         Top             =   585
         Width           =   1770
      End
      Begin VB.CheckBox Check9 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sales Return Entry"
         Height          =   315
         Left            =   765
         TabIndex        =   15
         Top             =   6840
         Width           =   2475
      End
      Begin VB.CheckBox Check8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Master"
         Height          =   315
         Index           =   0
         Left            =   465
         TabIndex        =   13
         Top             =   630
         Width           =   1770
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Edit Option"
         Height          =   315
         Left            =   765
         TabIndex        =   10
         Top             =   3660
         Width           =   1770
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Bill Delete Option"
         Height          =   315
         Left            =   765
         TabIndex        =   9
         Top             =   4110
         Width           =   1770
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "New product Addition"
         Height          =   315
         Left            =   765
         TabIndex        =   8
         Top             =   4575
         Width           =   1950
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Change Product Rate Option"
         Height          =   315
         Left            =   765
         TabIndex        =   7
         Top             =   5025
         Width           =   2730
      End
      Begin VB.CheckBox Check5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Goods Receipt Entry"
         Height          =   315
         Left            =   765
         TabIndex        =   6
         Top             =   5475
         Width           =   1770
      End
      Begin VB.CheckBox Check6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Accounts Receipt Entry"
         Height          =   315
         Left            =   765
         TabIndex        =   5
         Top             =   5925
         Width           =   2355
      End
      Begin VB.CheckBox Check7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Accounts Payment Entry"
         Height          =   315
         Left            =   765
         TabIndex        =   4
         Top             =   6390
         Width           =   2475
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   1950
         Left            =   450
         TabIndex        =   14
         Top             =   1080
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   3440
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         BackColorBkg    =   16777215
         AllowBigSelection=   0   'False
         Appearance      =   0
         FormatString    =   "Access|        Menu Name        | Add    | Edit  |   Delete"
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   1950
         Left            =   6030
         TabIndex        =   17
         Top             =   1035
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   3440
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         BackColorBkg    =   16777215
         AllowBigSelection=   0   'False
         Appearance      =   0
         FormatString    =   "Access|        Menu Name        | Add    | Edit  |   Delete"
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid3 
         Height          =   1950
         Left            =   6060
         TabIndex        =   23
         Top             =   5505
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   3440
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         BackColorBkg    =   16777215
         AllowBigSelection=   0   'False
         Appearance      =   0
         FormatString    =   "Access|                             Menu Name        | Print"
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Inventory"
         Height          =   195
         Left            =   6075
         TabIndex        =   21
         Top             =   3315
         Width           =   660
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000080FF&
         X1              =   5595
         X2              =   5595
         Y1              =   270
         Y2              =   6780
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         Caption         =   "Menu Access Control"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   435
         TabIndex        =   12
         Top             =   240
         Width           =   4440
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bill Setup"
         Height          =   195
         Left            =   495
         TabIndex        =   11
         Top             =   3315
         Width           =   660
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Ok"
      Height          =   375
      Left            =   9360
      TabIndex        =   2
      Top             =   7935
      Width           =   1080
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   10530
      TabIndex        =   1
      Top             =   7935
      Width           =   1080
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Apply"
      Height          =   375
      Left            =   8160
      TabIndex        =   0
      Top             =   7935
      Width           =   1080
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000A&
      BorderWidth     =   2
      Height          =   750
      Left            =   45
      Top             =   7830
      Width           =   11775
   End
End
Attribute VB_Name = "frmAccCp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Width = 11985
Me.Height = 9165

End Sub


