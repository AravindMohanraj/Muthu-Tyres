VERSION 5.00
Begin VB.Form frmChngPass 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Change Password"
   ClientHeight    =   2760
   ClientLeft      =   3825
   ClientTop       =   3930
   ClientWidth     =   4290
   LinkTopic       =   "Form1"
   ScaleHeight     =   2760
   ScaleWidth      =   4290
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   1500
      MaxLength       =   15
      PasswordChar    =   "#"
      TabIndex        =   5
      Top             =   1380
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Apply"
      Height          =   375
      Left            =   2685
      TabIndex        =   4
      Top             =   2190
      Width           =   1380
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   1515
      MaxLength       =   15
      PasswordChar    =   "#"
      TabIndex        =   2
      Top             =   795
      Width           =   2535
   End
   Begin VB.TextBox txtBno 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   1530
      MaxLength       =   15
      TabIndex        =   0
      Top             =   210
      Width           =   2535
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      Height          =   195
      Left            =   540
      TabIndex        =   6
      Top             =   1440
      Width           =   690
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      Height          =   195
      Left            =   555
      TabIndex        =   3
      Top             =   855
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Login"
      Height          =   195
      Left            =   570
      TabIndex        =   1
      Top             =   270
      Width           =   765
   End
End
Attribute VB_Name = "frmChngPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
