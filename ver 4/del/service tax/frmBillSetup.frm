VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmBillSetup 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   5475
   ClientLeft      =   3225
   ClientTop       =   1950
   ClientWidth     =   6600
   LinkTopic       =   "Form1"
   ScaleHeight     =   5475
   ScaleWidth      =   6600
   Begin VB.CommandButton Command2 
      Caption         =   "Command1"
      Height          =   390
      Left            =   5370
      TabIndex        =   5
      Top             =   4950
      Width           =   945
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   390
      Left            =   4320
      TabIndex        =   4
      Top             =   4965
      Width           =   945
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   1950
      Left            =   225
      TabIndex        =   1
      Top             =   465
      Width           =   6120
      _ExtentX        =   10795
      _ExtentY        =   3440
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      BackColorBkg    =   16777215
      AllowBigSelection=   0   'False
      Appearance      =   0
      FormatString    =   "Row |                                          Company Name /Address               | S  | N |  B"
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Height          =   1950
      Left            =   210
      TabIndex        =   3
      Top             =   2760
      Width           =   6120
      _ExtentX        =   10795
      _ExtentY        =   3440
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      BackColorBkg    =   16777215
      AllowBigSelection=   0   'False
      Appearance      =   0
      FormatString    =   "Row |                                          Company Name /Address               | S  | N |  B"
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Character Size"
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
      Left            =   5220
      TabIndex        =   6
      Top             =   30
      Width           =   1170
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BOTTOM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   210
      TabIndex        =   2
      Top             =   2475
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TOP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   165
      TabIndex        =   0
      Top             =   180
      Width           =   390
   End
End
Attribute VB_Name = "frmBillSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
