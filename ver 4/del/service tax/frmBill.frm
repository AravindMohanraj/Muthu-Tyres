VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBill 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Bill"
   ClientHeight    =   7740
   ClientLeft      =   270
   ClientTop       =   885
   ClientWidth     =   10785
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   NegotiateMenus  =   0   'False
   ScaleHeight     =   7740
   ScaleWidth      =   10785
   WindowState     =   2  'Maximized
   Begin VB.Frame fSp 
      Appearance      =   0  'Flat
      BackColor       =   &H006F472B&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2010
      Left            =   7875
      TabIndex        =   56
      Top             =   4890
      Visible         =   0   'False
      Width           =   2490
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Spelling Mistake"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   255
         TabIndex        =   57
         Top             =   825
         Width           =   1980
      End
   End
   Begin VB.Frame fremCust 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6315
      Left            =   10785
      TabIndex        =   43
      Top             =   3240
      Visible         =   0   'False
      Width           =   4560
      Begin VB.TextBox txtCustlist 
         Appearance      =   0  'Flat
         Height          =   435
         Left            =   45
         TabIndex        =   45
         Top             =   75
         Width           =   4455
      End
      Begin MSFlexGridLib.MSFlexGrid CustLst 
         Height          =   5535
         Left            =   60
         TabIndex        =   47
         Top             =   615
         Width           =   4440
         _ExtentX        =   7832
         _ExtentY        =   9763
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         BackColorFixed  =   2701557
         ForeColorFixed  =   0
         BackColorSel    =   -2147483636
         BackColorBkg    =   14737632
         GridColor       =   0
         FocusRect       =   0
         GridLinesFixed  =   1
         ScrollBars      =   2
         SelectionMode   =   1
         Appearance      =   0
         FormatString    =   "               Customer  Name                              "
      End
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5670
      Left            =   840
      TabIndex        =   16
      Top             =   675
      Visible         =   0   'False
      Width           =   7035
      Begin MSFlexGridLib.MSFlexGrid GrpGrid 
         Height          =   2745
         Left            =   375
         TabIndex        =   52
         Top             =   1680
         Visible         =   0   'False
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   4842
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         FocusRect       =   0
         GridLinesFixed  =   1
         ScrollBars      =   2
         SelectionMode   =   1
         Appearance      =   0
         FormatString    =   "                                                          Account Name                                | Code"
      End
      Begin VB.TextBox txtNote 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   4485
         TabIndex        =   32
         Top             =   135
         Width           =   2145
      End
      Begin MSFlexGridLib.MSFlexGrid tAxGrid 
         Height          =   2550
         Left            =   4365
         TabIndex        =   50
         Top             =   2790
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   4498
         _Version        =   393216
         FixedCols       =   0
         FocusRect       =   0
         GridLinesFixed  =   1
         ScrollBars      =   2
         AllowUserResizing=   2
         Appearance      =   0
         FormatString    =   "Tax %     |        Tax Value"
      End
      Begin VB.TextBox txtPcP 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1695
         TabIndex        =   51
         Top             =   4305
         Width           =   705
      End
      Begin VB.TextBox txtPcR 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2475
         TabIndex        =   38
         Top             =   4305
         Width           =   1665
      End
      Begin VB.TextBox txtCardNo 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1695
         TabIndex        =   33
         Top             =   525
         Visible         =   0   'False
         Width           =   4935
      End
      Begin VB.ComboBox cboPaytype 
         Height          =   315
         ItemData        =   "frmBill.frx":0000
         Left            =   1695
         List            =   "frmBill.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   165
         Width           =   1320
      End
      Begin VB.TextBox txtBcP 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1695
         TabIndex        =   49
         Top             =   3900
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.TextBox txtBcR 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2475
         TabIndex        =   29
         Top             =   3915
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.TextBox txtPartyName 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1695
         TabIndex        =   35
         Top             =   900
         Width           =   4935
      End
      Begin VB.TextBox txtAdd1 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1695
         TabIndex        =   37
         Top             =   1275
         Width           =   4935
      End
      Begin VB.TextBox txtAdd2 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1695
         TabIndex        =   39
         Top             =   1635
         Width           =   4935
      End
      Begin VB.TextBox txtAdd3 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1695
         TabIndex        =   40
         Top             =   1995
         Width           =   4935
      End
      Begin VB.TextBox txtadd4 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1695
         TabIndex        =   42
         Top             =   2355
         Width           =   4935
      End
      Begin VB.TextBox txtTrnsChg 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1695
         TabIndex        =   48
         Top             =   3510
         Width           =   705
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2475
         TabIndex        =   21
         Top             =   3525
         Width           =   1665
      End
      Begin VB.TextBox txtdedP 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1695
         TabIndex        =   44
         Top             =   2730
         Width           =   705
      End
      Begin VB.TextBox txtDedR 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2475
         TabIndex        =   20
         Top             =   2745
         Width           =   1665
      End
      Begin VB.TextBox txtTaxP 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1695
         TabIndex        =   46
         Top             =   3120
         Width           =   705
      End
      Begin VB.TextBox txtTaxR 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   2475
         TabIndex        =   19
         Top             =   3135
         Width           =   1665
      End
      Begin VB.TextBox txtAmtR 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2475
         TabIndex        =   18
         Top             =   4710
         Width           =   1665
      End
      Begin VB.TextBox txtAmtB 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2475
         TabIndex        =   17
         Top             =   5100
         Width           =   1665
      End
      Begin VB.Label lblNtDes 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vehical Number"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3120
         TabIndex        =   69
         Top             =   195
         Width           =   1125
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Packing Chgs"
         Height          =   195
         Left            =   555
         TabIndex        =   41
         Top             =   4335
         Width           =   990
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Card Number"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   615
         TabIndex        =   36
         Top             =   540
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pay Type"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   630
         TabIndex        =   34
         Top             =   180
         Width           =   675
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Charges"
         Height          =   195
         Left            =   555
         TabIndex        =   30
         Top             =   3945
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Party Name"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   600
         TabIndex        =   28
         Top             =   900
         Width           =   825
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   615
         TabIndex        =   27
         Top             =   1305
         Width           =   570
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transport Chgs"
         Height          =   195
         Left            =   555
         TabIndex        =   26
         Top             =   3570
         Width           =   1080
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Deduction"
         Height          =   195
         Left            =   570
         TabIndex        =   25
         Top             =   2760
         Width           =   735
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tax"
         Height          =   195
         Left            =   570
         TabIndex        =   24
         Top             =   3135
         Width           =   270
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount Received"
         Height          =   195
         Left            =   555
         TabIndex        =   23
         Top             =   4830
         Width           =   1275
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cash Balance"
         Height          =   195
         Left            =   570
         TabIndex        =   22
         Top             =   5220
         Width           =   990
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   405
      Left            =   12015
      Top             =   5700
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   714
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame fremIte 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6315
      Left            =   3585
      TabIndex        =   13
      Top             =   375
      Visible         =   0   'False
      Width           =   6015
      Begin VB.TextBox txtItemLst 
         Appearance      =   0  'Flat
         Height          =   435
         Left            =   30
         TabIndex        =   15
         Top             =   60
         Width           =   5940
      End
      Begin MSFlexGridLib.MSFlexGrid ItemLst 
         Height          =   5805
         Left            =   60
         TabIndex        =   14
         Top             =   660
         Width           =   5910
         _ExtentX        =   10425
         _ExtentY        =   10239
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         BackColorFixed  =   2701557
         ForeColorFixed  =   0
         BackColorSel    =   8388608
         BackColorBkg    =   14737632
         GridColor       =   0
         FocusRect       =   0
         GridLinesFixed  =   1
         ScrollBars      =   2
         SelectionMode   =   1
         Appearance      =   0
         FormatString    =   "Product  Name                                                    | Stock       | MRP"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   11910
      Top             =   2895
   End
   Begin MSComctlLib.StatusBar sb1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   11
      Top             =   7410
      Width           =   10785
      _ExtentX        =   19024
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3175
            MinWidth        =   3175
            Text            =   "User name"
            TextSave        =   "User name"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Time"
            TextSave        =   "Time"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   6675
      Left            =   75
      TabIndex        =   2
      Top             =   150
      Width           =   10335
      Begin VB.TextBox txtMdt 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   120
         MaxLength       =   144
         TabIndex        =   73
         Top             =   5955
         Width           =   10095
      End
      Begin VB.ComboBox cboBranch 
         Height          =   315
         ItemData        =   "frmBill.frx":0038
         Left            =   120
         List            =   "frmBill.frx":003A
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   645
         Width           =   1440
      End
      Begin VB.TextBox txtSTot 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   4680
         TabIndex        =   68
         Top             =   5400
         Width           =   1170
      End
      Begin VB.Frame frmOno 
         BackColor       =   &H00FFFFFF&
         Height          =   1440
         Left            =   3270
         TabIndex        =   63
         Top             =   105
         Visible         =   0   'False
         Width           =   1665
         Begin VB.TextBox txtOno 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   75
            TabIndex        =   64
            Top             =   420
            Width           =   1170
         End
         Begin MSMask.MaskEdBox txtODt 
            Height          =   285
            Left            =   90
            TabIndex        =   65
            Top             =   1065
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   8
            Format          =   "dd-mmm-yy"
            Mask            =   "##/##/##"
            PromptChar      =   "_"
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Order No"
            Height          =   195
            Left            =   60
            TabIndex        =   67
            Top             =   150
            Width           =   645
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Order Date"
            Height          =   195
            Left            =   75
            TabIndex        =   66
            Top             =   780
            Width           =   780
         End
      End
      Begin VB.Frame FrmDc 
         BackColor       =   &H00FFFFFF&
         Height          =   1440
         Left            =   9360
         TabIndex        =   58
         Top             =   6600
         Visible         =   0   'False
         Width           =   1770
         Begin VB.TextBox txtdNo 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   75
            TabIndex        =   59
            Top             =   420
            Width           =   1170
         End
         Begin MSMask.MaskEdBox txtDcDt 
            Height          =   285
            Left            =   90
            TabIndex        =   60
            Top             =   1065
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   8
            Format          =   "dd-mmm-yy"
            Mask            =   "##/##/##"
            PromptChar      =   "_"
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Del. Date"
            Height          =   195
            Left            =   75
            TabIndex        =   62
            Top             =   780
            Width           =   675
         End
         Begin VB.Label dcLblno 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "D.C. Number"
            Height          =   195
            Left            =   60
            TabIndex        =   61
            Top             =   150
            Width           =   915
         End
      End
      Begin VB.Frame fremGdwn 
         Caption         =   "Godown"
         Height          =   2730
         Left            =   6135
         TabIndex        =   53
         Top             =   2670
         Visible         =   0   'False
         Width           =   2115
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   735
            MaxLength       =   5
            TabIndex        =   55
            Top             =   2850
            Width           =   960
         End
         Begin MSFlexGridLib.MSFlexGrid GdwnGrid 
            Height          =   2310
            Left            =   90
            TabIndex        =   54
            Top             =   300
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   4075
            _Version        =   393216
            Cols            =   3
            FixedCols       =   0
            BackColorFixed  =   -2147483648
            BackColorBkg    =   16777215
            FocusRect       =   0
            GridLinesFixed  =   1
            ScrollBars      =   2
            SelectionMode   =   1
            Appearance      =   0
            FormatString    =   "Godown               | Qty  |GC"
         End
      End
      Begin MSFlexGridLib.MSFlexGrid flxgrd 
         Height          =   3645
         Left            =   45
         TabIndex        =   1
         Top             =   1755
         Width           =   10230
         _ExtentX        =   18045
         _ExtentY        =   6429
         _Version        =   393216
         Cols            =   11
         FixedCols       =   0
         BackColorFixed  =   -2147483648
         BackColorBkg    =   16777215
         FocusRect       =   0
         HighLight       =   2
         GridLinesFixed  =   1
         ScrollBars      =   2
         Appearance      =   0
         FormatString    =   $"frmBill.frx":003C
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Height          =   1440
         Left            =   1770
         TabIndex        =   6
         Top             =   120
         Width           =   1380
         Begin VB.TextBox txtBno 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   75
            TabIndex        =   7
            Top             =   420
            Width           =   1170
         End
         Begin MSMask.MaskEdBox dBilldt 
            Height          =   285
            Left            =   90
            TabIndex        =   8
            Top             =   1065
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   8
            Format          =   "dd-mmm-yy"
            Mask            =   "##/##/##"
            PromptChar      =   "_"
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bill Number"
            Height          =   195
            Left            =   60
            TabIndex        =   10
            Top             =   150
            Width           =   795
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bill Date"
            Height          =   195
            Left            =   75
            TabIndex        =   9
            Top             =   780
            Width           =   585
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Height          =   1440
         Left            =   4950
         TabIndex        =   3
         Top             =   105
         Width           =   5280
         Begin VB.TextBox txtBstk 
            Appearance      =   0  'Flat
            BackColor       =   &H00008000&
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   60
            Locked          =   -1  'True
            TabIndex        =   71
            Top             =   990
            Width           =   1380
         End
         Begin VB.TextBox txtTotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   48
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   1275
            Left            =   1545
            TabIndex        =   12
            Text            =   "0.00"
            Top             =   135
            Width           =   3735
         End
         Begin VB.TextBox txtStock 
            Appearance      =   0  'Flat
            BackColor       =   &H000080FF&
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   60
            Locked          =   -1  'True
            TabIndex        =   4
            Top             =   390
            Width           =   1380
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Branch Stock"
            Height          =   195
            Left            =   135
            TabIndex        =   72
            Top             =   735
            Width           =   975
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Stock In hand"
            Height          =   195
            Left            =   120
            TabIndex        =   5
            Top             =   135
            Width           =   1005
         End
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Note"
         Height          =   195
         Left            =   135
         TabIndex        =   74
         Top             =   5580
         Width           =   345
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Branch"
         Height          =   195
         Left            =   135
         TabIndex        =   70
         Top             =   330
         Width           =   510
      End
   End
   Begin VB.Menu mnu1 
      Caption         =   "Files"
      Begin VB.Menu mnu2 
         Caption         =   "New Bill"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnu3 
         Caption         =   "Edit Bill"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnu4 
         Caption         =   "Cancel Bill"
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnu5 
         Caption         =   "Delete Bill"
      End
      Begin VB.Menu mnu6 
         Caption         =   "Print Bill"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnue6 
         Caption         =   "-"
      End
      Begin VB.Menu mnu7 
         Caption         =   "Park a Bill"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnu8 
         Caption         =   "Save"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnu9 
         Caption         =   "Save As"
         Shortcut        =   {F3}
         Visible         =   0   'False
      End
      Begin VB.Menu mnu10 
         Caption         =   "-"
      End
      Begin VB.Menu mnu11 
         Caption         =   "Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnu12 
      Caption         =   "Edit"
      Begin VB.Menu mnu13 
         Caption         =   "Delete Current Line"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnu14 
         Caption         =   "Clear This Line"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnu15 
         Caption         =   "Search Party"
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu mnu16 
      Caption         =   "Master"
      Begin VB.Menu mnu17 
         Caption         =   "Add New Customer"
      End
      Begin VB.Menu mnu18 
         Caption         =   "Add New Product"
      End
      Begin VB.Menu mnu19 
         Caption         =   "Change Rate for Product"
      End
   End
   Begin VB.Menu mnu20 
      Caption         =   "Accounts"
      Begin VB.Menu mnu21 
         Caption         =   "Receipt"
      End
      Begin VB.Menu mnu22 
         Caption         =   "Payment"
      End
   End
   Begin VB.Menu mnu23 
      Caption         =   "Inventory"
      Begin VB.Menu mnu24 
         Caption         =   "Goods Receipt"
      End
   End
   Begin VB.Menu mnu25 
      Caption         =   "Windows"
   End
End
Attribute VB_Name = "frmBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lFnd As Boolean
Dim nOpt As Single, nTax(10) As Double
Dim nBcP As Double, nTxP As Double
Dim cCustCode As String, lANew As Boolean
Dim cTaxDisp As String

Private Sub CallTot()
Dim nSubTot As Double, nTax As Double, nTotal As Double
Dim nDed As Double, nBc As Double, nPc As Double
For I = 1 To FlxGrd.Rows - 1
    nSubTot = nSubTot + Val(FlxGrd.TextMatrix(I, 4))
Next
txtSTot = nSubTot

If lSuDed Then  ' if deduction is allowed
If Val(txtdedP) > 0 Then
nDed = nSubTot * Val(txtdedP) / 100
ElseIf Val(txtDedR) > 0 Then
nDed = Val(txtDedR)

End If
End If

If lSuTax Then   ' if tax available then
    If lSuSTax Then
       nTax = (nSubTot - nDed) * Val(txtTaxP) / 100
       
    End If
End If


If lSuBc And cboPaytype.ListIndex = 2 Then ' if bank chgs awailable then
  nBc = ((nSubTot - nDed) + nTax) * Val(txtBcP) / 100
End If


nTotal = (nSubTot - nDed) + nTax + nBc
txtTotal = nearRnd(nTotal)
txtTaxR = nTax
txtBcR = nBc
txtDedR = nDed
txtAmtR = txtTotal

End Sub


Private Sub ClearData()
FlxGrd.Rows = 2
FlxGrd.Clear
FlxGrd.FormatString = "SlNo. |                  Product Name              | Quantity   |  Rate     |   Amount   | Code  |                    Language  |UOM |Gdwn  |Weight|Tax"
If nOpt = 1 Then txtBno = ""
txtdNo = ""
txtOno = ""
txtTotal = ""
txtCardNo = ""
txtSTot = ""
txtPartyName = ""
txtAdd1 = ""
txtAdd2 = ""
txtAdd3 = ""
txtadd4 = ""
txtdedP = ""
txtDedR = ""
txtTaxP = ""
txtTaxR = ""
txtTrnsChg = ""
txtBcP = ""
txtBcR = ""
txtPcP = ""
txtPcR = ""
txtAmtR = ""
txtAmtB = ""
txtTaxP = nTxP
txtMdt = ""
txtNote = ""
cboPaytype.ListIndex = 0
txtCardNo = ""
End Sub

Private Function DelData()
    Set rsBill = New ADODB.Recordset
    Set rsIled = New ADODB.Recordset
    Set rsGdTrn = New ADODB.Recordset
    Set rsBLst = New ADODB.Recordset
    Set rsBrStk = New ADODB.Recordset
    rsBLst.Open "delete from billlist where fbillno='" & txtBno & "' and fbranch='" & CboBranch.Text & "'", Con, adOpenDynamic, adLockPessimistic

    rsBill.Open "Select * from bil0203d where fbillno='" & txtBno & "'and fbranch='" & CboBranch.Text & "'", Con, adOpenStatic
    If Not rsBill.EOF Then
    If Not rsBill.BOF Then rsBill.MoveFirst
    Do While Not rsBill.EOF
    
        AddStock rsBill!faccode, rsBill!fqty
        rsBrStk.Open "update brstk set fbal=fbal+ '" & rsBill!fqty & "' where faccode='" & rsBill!faccode & "' and fbranch='" & rsBill!fbranch & "' ", Con, adOpenDynamic, adLockPessimistic

    
    rsBill.MoveNext
    Loop
    End If
    rsBill.Close
    If lSuGd Then
        rsGdTrn.Open "select * from gdtrans where fvrno='" & txtBno & "' and fvrtype=1", Con, adOpenStatic
        If Not rsGdTrn.EOF Then
        If Not rsGdTrn.BOF Then rsGdTrn.MoveFirst
        Do While Not rsGdTrn.EOF
            AddGdStk rsGdTrn!faccode, rsGdTrn!fqty, rsGdTrn!fgodown
        rsGdTrn.MoveNext
        Loop
        
        End If
        rsGdTrn.Close
            rsGdTrn.Open "delete from gdtrans where fvrno='" & txtBno & "' and fvrtype=1", Con, adOpenDynamic, adLockPessimistic

    End If
    
    rsBill.Open "delete from bil0203d where fbillno='" & txtBno & "'and fbranch='" & CboBranch.Text & "'", Con, adOpenDynamic, adLockPessimistic
    
    rsIled.Open "delete from ile0203d where fvrno='" & txtBno & "' and fvrtype=1 and fbranch='" & CboBranch.Text & "'", Con, adOpenDynamic, adLockPessimistic
    

    

End Function

Private Sub DelPrint()
Dim cPaid As String, cBbal As String, cWord As String, cWord1 As String, cWord2 As String
Dim nWgt As String, nQty As String, cTotwgt As String, cLess As String
Dim nLM As Integer, cVl As String, nBG As Integer, cHl As String
Dim nTotAmt As Double, nTotQty, nTRAmt As Double
Dim nTlen, nLoop, nInc As Integer, cString As String
Dim nCnt As Integer
nLM = 10
cVl = " "
cHl = "-"
Open "c:\files\testfile.TXT" For Output As #1  ' Open file for output.
For l = 1 To nNoLR
Print #1, Chr(27) + "j" + "n"
Next
'############### Head Printing #################
  '  Print #1,
    'Print #1,
    Print #1, Space(0) + Space(24) + RPad(Left(Trim(txtBno.Text), 6), 6) + " (H.O) " + Space(4) + CPad(Mid(cboPaytype.Text, 4, Len(cboPaytype.Text)), 11) + Space(13) + dBilldt.FormattedText
        Print #1,

    Print #1, Space(0) + Space(53) + ""
    Print #1, Space(0) + Space(53) + ""
 '   Print #1,
    Print #1, Space(0) + Space(53) + ""
    For I = 1 To 7
    Print #1,
    Next
    If txtPartyName.Text <> "" Then
    Print #1, Space(nLM + 7) + Chr$(27) + Chr$(71) + "M/s " + txtPartyName.Text + Chr(27) + Chr$(72)
   ' ElseIf TxtName.Text <> "" Then
    'Print #1, Space(nLM + 7) + Chr$(27) + Chr$(71) + "M/s " + TxtName.Text + Chr(27) + Chr$(72)
    Else
    Print #1,
    End If
    
    If txtAdd1.Text <> "" Then
    Print #1, Space(nLM + 7) + txtAdd1.Text
    Else
    Print #1,
    End If
    If txtAdd2.Text <> "" Then
    Print #1, Space(nLM + 7) + txtAdd2.Text
    Else
    Print #1,
    End If
    
    If txtAdd3.Text <> "" Then
    Print #1, Space(nLM + 7) + txtAdd3.Text; Tab(60); txtNote
    Else
    Print #1, Tab(60); txtNote
    End If
    Print #1,
    Print #1,
    Print #1,
    nCnt = 1
     For I = 1 To FlxGrd.Rows - 1
     If FlxGrd.TextMatrix(I, 5) <> "" Then
       If Val(FlxGrd.TextMatrix(I, 2)) <> 0 Then
       nCnt = nCnt + 1
'         If Len(flxgrd.TextMatrix(i, 1)) < 25 Then
           Print #1, Space(nLM) + Space(6) + RPad(Left(FlxGrd.TextMatrix(I, 1), 35), 35) + Space(4) + RPad(FlxGrd.TextMatrix(I, 2), 3) + Space(1) + LPad(Format(FlxGrd.TextMatrix(I, 3), "####0"), 5) + Space(2) + LPad(Format(FlxGrd.TextMatrix(I, 4), "#####0.00"), 9)
 '        Else
  '         Print #1, Space(nLM) + Space(3) + RPad(flxgrd.TextMatrix(i, 0), 2) + RPad(flxgrd.TextMatrix(i, 1), 6) + RPad(flxgrd.TextMatrix(i, 2), 5) + RPad(flxgrd.TextMatrix(i, 3), 6) + RPad(flxgrd.TextMatrix(i, 4), 6)
   '        Print #1, Space(nLM) + Space(3) + RPad(Mid(flxgrd.TextMatrix(i, 1), 15, Len(flxgrd.TextMatrix(i, 1))), 5)
    '     End If
    End If
    nTotAmt = nTotAmt + Val(FlxGrd.TextMatrix(I, 4))
    nTotQty = nTotQty + Val(FlxGrd.TextMatrix(I, 2))
    End If
    Next
        
        nCnt = nCnt + 1
        Print #1, Space(nLM) + Space(11) + "Product Value" + Space(32) + LPad(Format(txtSTot.Text, "#####0.00"), 9)
    
    If Val(txtDedR) <> 0 Then
        nCnt = nCnt + 1
        
        Print #1, Space(nLM) + Space(11) + "Discount" + Space(37) + LPad(Format(txtDedR, "#####0.00"), 9)
        nCnt = nCnt + 1
        Print #1, Space(nLM) + Space(11) + "Sub Total" + Space(36) + LPad(Format(Val(txtSTot.Text) - Val(txtDedR), "#####0.00"), 9)
    Else
       nCnt = nCnt + 2
       Print #1,
       Print #1,
       
       
    End If
    
        nCnt = nCnt + 1
        'Print #1, Space(nLM) + Space(1) + "VAT @ " + LPad(txtTaxP, 4) + " %" + Space(33) + LPad(Format(txtTaxR, "#####0.00"), 9)
                Print #1, Space(nLM) + Space(11) + RPad(cTaxDisp, 20) + "  " + Space(23) + LPad(Format(txtTaxR, "#####0.00"), 9)

    If Val(txtBcP) <> 0 Then
        nCnt = nCnt + 1
        Print #1, Space(nLM) + Space(11) + "Bank Chg" + Space(37) + LPad(Format(txtBcR, "#####0.00"), 9)
    Else
        nCnt = nCnt + 1
       Print #1,
    End If
        
        nCnt = nCnt + 1
        Print #1,
    If nCnt <= nNoP Then
       nCnt = nNoP - nCnt
       For I = 1 To nCnt
       Print #1,
       Next
    End If
'Print #1,
       Print #1,

Print #1, Space(nLM + 56) + LPad(Format(txtTotal, "#####0.00"), 9)
cWord = "Rupees " + NumToWord(txtTotal) + " Only"

'Print #1,
If Len(cWord) > 50 Then
cWord1 = Left(cWord, 50)
cWord2 = Mid(cWord, 51, Len(cWord))
Print #1, Space(nLM) + Space(6) + cWord1
Print #1, Space(nLM) + Space(6) + cWord2
Else
Print #1, Space(nLM) + Space(6) + "Rupees " + NumToWord(txtTotal) + " Only"
End If

If txtMdt <> "" Then
Print #1,
Print #1,
Print #1,

If Len(txtMdt) <= 48 Then
Print #1, Space(32) + Chr(15) + txtMdt + Chr(18)
Else
a = Left(txtMdt, 48)
B = Mid(txtMdt, 49, 48)
c = Mid(txtMdt, 96, 48)
Print #1, Space(32) + Chr(15) + Left(txtMdt, 48) + Chr(18)
Print #1, Space(32) + Chr(15) + Mid(txtMdt, 49, 48) + Chr(18)
Print #1, Space(32) + Chr(15) + Mid(txtMdt, 97, 48) + Chr(18)

End If

End If


        
For K = 1 To nNolE
Print #1,
Next
Close #1

Do While True
RetVal = Shell("c:\files\dosprint.bat", 0)
If MsgBox("Print Again", vbOKCancel) = vbCancel Then
Exit Do
End If
Loop



End Sub

Private Sub FillAdd(cCode As String)
Set rsAdd = New ADODB.Recordset
txtAdd1 = ""
txtAdd2 = ""
txtAdd3 = ""
txtadd4 = ""
rsAdd.Open "select * from add0203d where faccode='" & cCode & " '", Con, adOpenStatic
If Not rsAdd.EOF Then
If Not rsAdd.BOF Then rsAdd.MoveFirst
If Not IsNull(rsAdd!add1) Then txtAdd1 = rsAdd!add1
If Not IsNull(rsAdd!add2) Then txtAdd2 = rsAdd!add2
If Not IsNull(rsAdd!add3) Then txtAdd3 = rsAdd!add3
If Not IsNull(rsAdd!add4) Then txtadd4 = rsAdd!add4


End If
rsAdd.Close

Set rsAdd = Nothing

End Sub

Private Sub FillGdStk(cCode As String)
Dim nR As Integer
GdwnGrid.Rows = 2
GdwnGrid.Clear
GdwnGrid.FormatString = "Godown               | Qty  |GC"

Set rsGdSTk = New ADODB.Recordset
rsGdSTk.Open "select * from gdstk where faccode='" & cCode & "' and fbal>0", Con, adOpenStatic
If Not rsGdSTk.EOF Then
If Not rsGdSTk.BOF Then rsGdSTk.MoveFirst
nR = 1

Do While Not rsGdSTk.EOF
GdwnGrid.TextMatrix(nR, 0) = rsGdSTk!fgd
GdwnGrid.TextMatrix(nR, 1) = rsGdSTk!fbal
nR = nR + 1
GdwnGrid.AddItem ""
rsGdSTk.MoveNext
Loop

End If
rsGdSTk.Close

End Sub

Private Sub fillGrdflst()
Set rsItem = New ADODB.Recordset
Set rsBrStk = New ADODB.Recordset
rsItem.Open "select * from ite0203d where faccode='" & ItemLst.TextMatrix(ItemLst.Row, 3) & "'", Con, adOpenStatic
If Not rsItem.EOF Then

    FlxGrd.TextMatrix(FlxGrd.Row, 0) = FlxGrd.Row
    FlxGrd.TextMatrix(FlxGrd.Row, 1) = rsItem!facname
    FlxGrd.TextMatrix(FlxGrd.Row, 3) = rsItem!fSp
    FlxGrd.TextMatrix(FlxGrd.Row, 5) = ItemLst.TextMatrix(ItemLst.Row, 3)
    txtStock = rsItem!fclbal
    FlxGrd.TextMatrix(FlxGrd.Row, 7) = IIf(IsNull(rsItem!funit), "", rsItem!funit)
    FlxGrd.TextMatrix(FlxGrd.Row, 9) = rsItem!fweight
    
    If lSuBra Then
     rsBrStk.Open "select * from brstk where faccode='" & ItemLst.TextMatrix(ItemLst.Row, 3) & "' and fbranch='" & CboBranch.Text & "'", Con, adOpenStatic
     If Not rsBrStk.EOF Then
     If Not rsBrStk.BOF Then rsBrStk.MoveFirst
      txtBstk.Text = rsBrStk!fbal
     End If
     rsBrStk.Close
    End If
    If lSuTax Then
        If lSuItax Then
        FlxGrd.TextMatrix(FlxGrd.Row, 10) = IIf(IsNull(rsItem!ftax), "", rsItem!ftax)
        End If
    End If
    fremIte.Visible = False
    FlxGrd.Col = 2
    FlxGrd.SetFocus
End If
rsItem.Close
Set rsItem = Nothing
End Sub

Private Sub LoadProduct()
Dim nR As Integer
Set rsItem = New ADODB.Recordset
ItemLst.Rows = 2
ItemLst.Clear
ItemLst.FormatString = "Product  Name                                                    | Stock       | MRP"
nR = 1
rsItem.Open "select * from ite0203d where faclevel<0 order by facname", Con, adOpenStatic
If Not rsItem.EOF Then
If Not rsItem.BOF Then rsItem.MoveFirst
Do While Not rsItem.EOF
ItemLst.TextMatrix(nR, 0) = rsItem!facname
ItemLst.TextMatrix(nR, 1) = rsItem!fclbal
ItemLst.TextMatrix(nR, 2) = rsItem!fmrp
ItemLst.TextMatrix(nR, 3) = rsItem!faccode
ItemLst.AddItem ""
nR = nR + 1
rsItem.MoveNext
Loop
End If
rsItem.Close
Set rsItem = Nothing
End Sub

Private Function SaveData() As Boolean
Dim nStg As Single, nVrtype As Single, cCode As String
nVrtype = VrType("SAL")


Set rsBill = New ADODB.Recordset
Set rsIled = New ADODB.Recordset
Set rsNum = New ADODB.Recordset
Set rsBLst = New ADODB.Recordset
Set rsItem = New ADODB.Recordset
Set rsGdTrn = New ADODB.Recordset
Set rsAcc = New ADODB.Recordset
Set rsArea = New ADODB.Recordset
Set rsBrStk = New ADODB.Recordset

    
    nStg = 0
    If nOpt = 1 Then
      If lSuBra Then
             Set rsBBnum = New ADODB.Recordset
            rsBBnum.Open "select * from bbnum where fbranch='" & CboBranch.Text & "'", Con, adOpenDynamic, adLockPessimistic
            If Not rsBBnum Then
                If Not rsBBnum.BOF Then rsBBnum.MoveFirst
                   txtBno.Text = rsBBnum!Fbillno + 1
                   rsBBnum!Fbillno = txtBno.Text
                   rsBBnum.Update
                End If
                rsBBnum.Close
            Set rsBBnum = Nothing
      Else
            rsNum.Open "select * from num0203d ", Con, adOpenDynamic, adLockPessimistic
            If nBM = 0 Then  '0 bill starts 1 to financial year
             txtBno = rsNum!frsales + 1
             rsNum!frsales = Val(txtBno)
             nStg = 1
            ElseIf nBM = 1 Then ' combination bill number and date
             nStg = 2
            ElseIf nBM = 2 Then ' every month reset to 1 with add chr
             nStg = 3
            End If
            rsNum.Update
            rsNum.Close
      End If
    End If
    
    Set rsBLst = New ADODB.Recordset
    rsBLst.Open "select * from billlist", Con, adOpenDynamic, adLockPessimistic
            
            rsBLst.AddNew
                   rsBLst!Fbillno = txtBno
                   rsBLst!fbilldt = dBilldt.FormattedText
                    rsBLst!fpaytype = cboPaytype.ListIndex
                    rsBLst!fdtime = Time
'                    rsBLst!fddate = ""
                    rsBLst!fcucode = ""
                    rsBLst!facname = txtPartyName.Text
                    rsBLst!fadd1 = txtAdd1
                    rsBLst!fadd2 = txtAdd2
                    rsBLst!fadd3 = txtAdd3
                    rsBLst!fadd4 = txtadd4
                    rsBLst!ftotal = Val(txtTotal)
                    rsBLst!ftax = Val(txtTaxR)
                    rsBLst!ftaxp = Val(txtTaxP)
                    rsBLst!fnote = txtNote
                    rsBLst!ftrns = Val(txtTrnsChg)
                    rsBLst!fded = Val(txtDedR)
                    rsBLst!fdedp = Val(txtdedP)
                    rsBLst!fbranch = CboBranch.Text
                    
                    rsBLst!fbcp = Val(txtBcP)
                    rsBLst!fbc = Val(txtBcR)
                    rsBLst!fpc = Val(txtPcR)
                    rsBLst!fpcp = Val(txtPcP)
                    rsBLst!fslno = txtMdt.Text
                    rsBLst!FCARDNO = Left(txtCardNo, 16)

                    If nOpt = 1 Then
                        If Val(txtAmtR) < Val(txtTotal) Then
                            rsBLst!fbalance = Val(txtAmtB)
                            rsBLst!fpaid = Val(txtAmtR)
                        Else
                            rsBLst!fbalance = 0
                            rsBLst!fpaid = Val(txtTotal)
                        End If
                    Else
                        
                    End If
            rsBLst.Update
        nStg = 1
    rsBLst.Close
    
    Set rsBill = New ADODB.Recordset
    Set rsIled = New ADODB.Recordset
    Set rsGdTrn = New ADODB.Recordset
    rsBill.Open "Select * from bil0203d", Con, adOpenDynamic, adLockPessimistic
    rsIled.Open "select * from ile0203d", Con, adOpenDynamic, adLockPessimistic
    rsGdTrn.Open "select * from gdtrans", Con, adOpenDynamic, adLockPessimistic
    
    For I = 1 To FlxGrd.Rows - 1
        If FlxGrd.TextMatrix(I, 5) <> "" Then
            '/update bill/
            rsBill.AddNew
                rsBill!Fbillno = txtBno
                rsBill!fbilldt = dBilldt.FormattedText
                rsBill!faccode = FlxGrd.TextMatrix(I, 5)
                rsBill!fqty = Val(FlxGrd.TextMatrix(I, 2))
                rsBill!frate = Val(FlxGrd.TextMatrix(I, 3))
                rsBill!facname = FlxGrd.TextMatrix(I, 1)
                rsBill!fpaytype = cboPaytype.ListIndex
                rsBill!fweight = Val(FlxGrd.TextMatrix(I, 9))
                rsBill!funit = Val(FlxGrd.TextMatrix(I, 7))
                rsBill!fgd = FlxGrd.TextMatrix(I, 8)
                rsBill!ftax = Val(FlxGrd.TextMatrix(I, 10))
                rsBill!fbranch = CboBranch.Text
                
                rsBrStk.Open "update brstk set fbal=fbal- '" & Val(FlxGrd.TextMatrix(I, 2)) & "' where faccode='" & FlxGrd.TextMatrix(I, 5) & "' and fbranch='" & CboBranch.Text & "' ", Con, adOpenDynamic, adLockPessimistic
                
            rsBill.Update
            
            '/Update Iledger/
            
            rsIled.AddNew
            
                rsIled!fvrtype = VrType("SAL")
                rsIled!fvrno = txtBno
                rsIled!fvrdate = dBilldt.FormattedText
                rsIled!faccode = FlxGrd.TextMatrix(I, 5)
                rsIled!fqty = Val(FlxGrd.TextMatrix(I, 2))
                rsIled!frate = Val(FlxGrd.TextMatrix(I, 3))
                rsIled!fval = Val(FlxGrd.TextMatrix(I, 4))
                rsIled!fgodown = FlxGrd.TextMatrix(I, 8)
                rsIled!fbranch = CboBranch.Text
                rsIled!ftag = 10
            rsIled.Update
            '/Update Godown transaction/
            
            rsGdTrn.AddNew
               rsGdTrn!fvrno = txtBno
               rsGdTrn!fgodown = FlxGrd.TextMatrix(I, 8)
               rsGdTrn!fqty = Val(FlxGrd.TextMatrix(I, 2))
               rsGdTrn!faccode = FlxGrd.TextMatrix(I, 5)
               rsGdTrn!fvrtype = VrType("SAL")
            rsGdTrn.Update
                MinusStock FlxGrd.TextMatrix(I, 5), Val(FlxGrd.TextMatrix(I, 2))
                MinusGdStk FlxGrd.TextMatrix(I, 5), Val(FlxGrd.TextMatrix(I, 2)), FlxGrd.TextMatrix(I, 8)
        End If
    Next
    nStg = 2
    Set rsLed = New ADODB.Recordset
    rsLed.Open "select * from led0203d", Con, adOpenDynamic, adLockPessimistic
                rsLed.AddNew
                rsLed!fvrtype = nVrtype
                rsLed!fvrno = Val(txtBno)
                rsLed!fvrdate = dBilldt.FormattedText
                If cboPaytype.ListIndex = 0 Or cboPaytype.ListIndex = 2 Then
                   rsLed!faccode = cCash
                Else
                   rsLed!faccode = cCustCode
                End If
                rsLed!fcrdb = "CR"
                rsLed!famount = Val(txtTotal)
                rsLed!faccode2 = cSales
                
                rsLed.Update
                
                
                If lSuTax Then
                    rsLed.AddNew
                    rsLed!fvrtype = nVrtype
                    rsLed!fvrno = Val(txtBno)
                    rsLed!fvrdate = dBilldt.FormattedText
                    rsLed!faccode = cVat
                    rsLed!fcrdb = "CR"
                    rsLed!famount = Val(txtTotal)
                    If cboPaytype.ListIndex = 0 Or cboPaytype.ListIndex = 2 Then
                       rsLed!faccode2 = cCash
                    Else
                       rsLed!faccode2 = cCustCode
                    End If
                    rsLed.Update
        
                End If

    
    
                If lSuDed Then
                    rsLed.AddNew
                    rsLed!fvrtype = nVrtype
                    rsLed!fvrno = Val(txtBno)
                    rsLed!fvrdate = dBilldt.FormattedText
                    rsLed!faccode = cDiscount
                    rsLed!fcrdb = "CR"
                    rsLed!famount = Val(txtTotal)
                    If cboPaytype.ListIndex = 0 Or cboPaytype.ListIndex = 2 Then
                       rsLed!faccode2 = cCash
                    Else
                       rsLed!faccode2 = cCustCode
                    End If
                    rsLed.Update
                End If
    

                If lSuBc Then
                    rsLed.AddNew
                    rsLed!fvrtype = nVrtype
                    rsLed!fvrno = Val(txtBno)
                    rsLed!fvrdate = dBilldt.FormattedText
                    rsLed!faccode = cBankchg
                    rsLed!fcrdb = "CR"
                    rsLed!famount = Val(txtTotal)
                    If cboPaytype.ListIndex = 0 Or cboPaytype.ListIndex = 2 Then
                       rsLed!faccode2 = cCash
                    Else
                       rsLed!faccode2 = cCustCode
                    End If
                    rsLed.Update
        
                End If
  
  
               If lSuPc Then
                    rsLed.AddNew
                    rsLed!fvrtype = nVrtype
                    rsLed!fvrno = Val(txtBno)
                    rsLed!fvrdate = dBilldt.FormattedText
                    rsLed!faccode = cTrans
                    rsLed!fcrdb = "CR"
                    rsLed!famount = Val(txtTotal)
                    If cboPaytype.ListIndex = 0 Then
                       rsLed!faccode2 = cCash
                    Else
                       rsLed!faccode2 = cCustCode
                    End If
                    rsLed.Update
               End If
  
  
  
rsLed.Close

Set rsBrStk = Nothing
End Function



Private Sub StartAccessLmt()

End Sub


Private Sub stuffData()
Dim nR As Integer
Set rsBill = New ADODB.Recordset
Set rsIled = New ADODB.Recordset
nR = 1
ClearData
rsBill.Open "select * from lstbills where fbillno='" & txtBno & "' AND fbranch='" & CboBranch.Text & "'", Con, adOpenDynamic, adLockPessimistic


If Not rsBill.EOF Then
    If Not rsBill.BOF Then rsBill.MoveFirst
    cboPaytype.ListIndex = rsBill!fpaytype
    If rsBill!fpaytype = 0 Then
       If Not IsNull(rsBill!acname) Then txtPartyName = rsBill!acname
    ElseIf rsBill!fpaytype = 1 Then
       If Not IsNull(rsBill!acname) Then txtPartyName = rsBill!acname
    ElseIf rsBill!fpaytype = 2 Then
       If Not IsNull(rsBill!acname) Then txtPartyName = rsBill!acname
    End If
    dBilldt.Text = datecon(rsBill!fbilldt)
    If Not IsNull(rsBill!fadd1) Then txtAdd1 = rsBill!fadd1
    If Not IsNull(rsBill!fadd2) Then txtAdd2 = rsBill!fadd2
    If Not IsNull(rsBill!fadd3) Then txtAdd3 = rsBill!fadd3
    If Not IsNull(rsBill!fadd4) Then txtadd4 = rsBill!fadd4
    If Not IsNull(rsBill!fnote) Then txtNote = rsBill!fnote
    If Not IsNull(rsBill!fslno) Then txtMdt = rsBill!mdt
    If Not IsNull(rsBill!FCARDNO) Then txtCardNo = rsBill!FCARDNO
    If lSuDed Then
       If Not IsNull(rsBill!fded) Then txtDedR = rsBill!fded
       If Not IsNull(rsBill!fdedp) Then txtdedP = rsBill!fdedp
       
    End If
    
    If lSuTax Then
       If Not IsNull(rsBill!ftax) Then txtTaxR = rsBill!ftax
       If Not IsNull(rsBill!ftaxp) Then txtTaxP = rsBill!ftaxp
    End If
    
    If lSuBc Then
       If Not IsNull(rsBill!fbc) Then txtBcR = rsBill!fbc
       If Not IsNull(rsBill!fbcp) Then txtBcP = rsBill!fbcp
    End If
    
    If lSuPc Then
       If Not IsNull(rsBill!fpc) Then txtPcR = rrsbill!fpc
       If Not IsNull(rsBill!fpcp) Then txtPcP = rrsbill!fpcp
    End If
        Do While Not rsBill.EOF
           FlxGrd.TextMatrix(nR, 0) = nR
           FlxGrd.TextMatrix(nR, 1) = rsBill!product
           FlxGrd.TextMatrix(nR, 2) = Format(rsBill!fqty, cDP)
           FlxGrd.TextMatrix(nR, 3) = Format(rsBill!frate, "#####0.00")
           FlxGrd.TextMatrix(nR, 4) = Format(rsBill!fqty * rsBill!frate, "######0.00")
           FlxGrd.TextMatrix(nR, 5) = rsBill!faccode
           FlxGrd.TextMatrix(nR, 7) = rsBill!funit
           FlxGrd.TextMatrix(nR, 8) = rsBill!fgd
           FlxGrd.TextMatrix(nR, 9) = rsBill!fweight
           FlxGrd.TextMatrix(nR, 10) = rsBill!ftax
           FlxGrd.AddItem ""
           nR = nR + 1
        rsBill.MoveNext
        Loop
End If
rsBill.Close
CallTot
Set rsBill = Nothing
Set rsIled = Nothing
End Sub

Private Sub cboBranch_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
'   Set rsBBnum = New ADODB.Recordset
'   rsBBnum.Open "select * from bbnum where fbranch='" & cboBranch.Text & "'", Con, adOpenStatic
'   If Not rsBBnum Then
'   If Not rsBBnum.BOF Then rsBBnum.MoveFirst
'      MsgBox rsBBnum!fbillno
'      txtBno.Text = rsBBnum!fbillno
If nOpt = 1 Then
      FlxGrd.SetFocus
Else
      txtBno.SetFocus
End If
'   End If
'   rsBBnum.Close
'   Set rsBBnum = Nothing
End If
End Sub


Private Sub cboPaytype_Click()
If cboPaytype.ListIndex = 2 Then
Label14.Visible = True
txtCardNo.Visible = True


Else
Label12.Visible = False
txtBcP.Visible = False
txtBcR.Visible = False

Label14.Visible = False
txtCardNo.Visible = False
txtBcP = ""
End If
End Sub

Private Sub cboPaytype_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
 Frame4.Visible = False
End If

If KeyAscii = 13 Then
   txtNote.SetFocus
End If
   
End Sub


Private Sub dBilldt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If lSuDc Then
        txtdNo.SetFocus
    Else
        If lSuOn Then
          txtOno.SetFocus
        Else
           FlxGrd.SetFocus
        End If
    End If
End If
End Sub


Private Sub FlxGrd_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
            If FlxGrd.Col = 1 Or FlxGrd.Col = 0 Then
                If KeyAscii = 8 Then      ' for backspace
                    If (IsEmpty(FlxGrd.Text)) Then
                        No = FlxGrd.Row
                        If (No > 1) Then
                            FlxGrd.RemoveItem (FlxGrd.Row)
                            FlxGrd.Row = No - 1
                            FlxGrd.Col = 2
                        End If
                    Else
                        If FlxGrd.Text <> "" Then
                            Stg = FlxGrd.Text
                            Strg = Mid(Stg, 1, (Len(Stg) - 1))
                            FlxGrd.Text = Strg
                        Else
                            If vbYes = MsgBox("Do you want to remove this row", vbYesNo, "Alert") Then
                                If FlxGrd.Row = 1 Then
                                    FlxGrd.AddItem ""
                                    FlxGrd.RemoveItem (1)
                                    FlxGrd.Col = 1
                                Else
                                    n = FlxGrd.Row
                                    FlxGrd.RemoveItem (FlxGrd.Row)
                                    FlxGrd.Row = n - 1
                                    FlxGrd.Col = 6
                                End If
                            End If
                        End If
                    End If
                ElseIf KeyAscii = 32 Then
                
                Else                             'for normal
                    
                    FlxGrd.TextMatrix(FlxGrd.Row, 1) = FlxGrd.TextMatrix(FlxGrd.Row, 1) + Chr(KeyAscii)
                    If nLst = 2 Then
                        txtItemLst.Text = ""
                        txtItemLst.Text = FlxGrd.TextMatrix(FlxGrd.Row, 1)
                                            fremIte.Top = 600
                    fremIte.Left = 5000

                        fremIte.Visible = True
                        txtItemLst.SetFocus
                        SendKeys "{END}"
                        Find ItemLst, UCase(txtItemLst.Text), 0, frmBill.fSp
                        
                    End If
                End If
           ElseIf FlxGrd.Col = 2 Then
                If KeyAscii = 8 Then      ' for backspace
                    If (IsEmpty(FlxGrd.Text)) Then
                        No = FlxGrd.Row
                        If (No > 1) Then
                            FlxGrd.RemoveItem (FlxGrd.Row)
                            FlxGrd.Row = No - 1
                            FlxGrd.Col = 2
                        End If
                    Else
                        If FlxGrd.Text <> "" Then
                            Stg = FlxGrd.Text
                            Strg = Mid(Stg, 1, (Len(Stg) - 1))
                            FlxGrd.Text = Strg
                        End If
                    End If
                 Else
                  If IsNumeric(Chr(KeyAscii)) Then
                     FlxGrd.TextMatrix(FlxGrd.Row, 2) = FlxGrd.TextMatrix(FlxGrd.Row, 2) + Chr(KeyAscii)
                  End If
                 End If
           ElseIf FlxGrd.Col = 3 Then
                 If KeyAscii = 8 Then      ' for backspace
                    If (IsEmpty(FlxGrd.Text)) Then
                        No = FlxGrd.Row
                        If (No > 1) Then
                            FlxGrd.RemoveItem (FlxGrd.Row)
                            FlxGrd.Row = No - 1
                            FlxGrd.Col = 3
                        End If
                    Else
                        If FlxGrd.Text <> "" Then
                            Stg = FlxGrd.Text
                            Strg = Mid(Stg, 1, (Len(Stg) - 1))
                            FlxGrd.Text = Strg
                        End If
                    End If
                Else
         If (Val(FlxGrd.TextMatrix(FlxGrd.Row, 3)) <= 0 And FlxGrd.TextMatrix(FlxGrd.Row, 3) <> ".") And Chr(KeyAscii) <> "." Then FlxGrd.TextMatrix(FlxGrd.Row, 3) = ""

                 FlxGrd.TextMatrix(FlxGrd.Row, 3) = FlxGrd.TextMatrix(FlxGrd.Row, 3) + Chr(KeyAscii)
                End If
           End If
ElseIf KeyAscii = 13 Then
        If FlxGrd.Col = 1 Or FlxGrd.Col = 0 Then
           If FlxGrd.TextMatrix(FlxGrd.Row, 1) = "" And FlxGrd.Row >= 2 Then
              txtMdt.SetFocus
           End If
        End If
        
     If FlxGrd.Col = 1 Then
            If nLst = 0 Then
            Set rsItem = New ADODB.Recordset
               rsItem.Open "select * from ite0203d where faccode='" & FlxGrd.TextMatrix(FlxGrd.Row, 1) & "'", Con, adOpenStatic
               If Not rsItem.EOF Then
                If Not rsItem.BOF Then rsItem.MoveFirst
                  FlxGrd.TextMatrix(FlxGrd.Row, 1) = rsItem!facname
                  FlxGrd.TextMatrix(FlxGrd.Row, 5) = rsItem!faccode
                  FlxGrd.Col = 6
                  FlxGrd.CellFontName = "Mylaiplain"
                  FlxGrd.TextMatrix(FlxGrd.Row, 6) = rsItem!facname
               Else
                   MsgBox "Code Not Found"
               End If
               rsItem.Close
            Set rsItem = Nothing
            End If
              txtStock = GetStock(FlxGrd.TextMatrix(FlxGrd.Row, 5))
              FlxGrd.Col = 2
     ElseIf FlxGrd.Col = 2 Then
        '    If Val(flxgrd.TextMatrix(flxgrd.Row, 2)) <= Val(txtBstk) Then
                FlxGrd.TextMatrix(FlxGrd.Row, 4) = Val(FlxGrd.TextMatrix(FlxGrd.Row, 2)) * Val(FlxGrd.TextMatrix(FlxGrd.Row, 3))
                If lSuGd Then
                FillGdStk FlxGrd.TextMatrix(FlxGrd.Row, 5)
                 fremGdwn.Visible = True
                 GdwnGrid.SetFocus
                End If
                FlxGrd.Col = 3
         '   Else
           '   MsgBox "Check for Stock"
          '  End If
     ElseIf FlxGrd.Col = 3 Then
     
            FlxGrd.TextMatrix(FlxGrd.Row, 4) = Val(FlxGrd.TextMatrix(FlxGrd.Row, 2)) * Val(FlxGrd.TextMatrix(FlxGrd.Row, 3))
            FlxGrd.AddItem ""
            FlxGrd.Row = FlxGrd.Row + 1
            CallTot
            FlxGrd.Col = 0
     End If
End If


End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
fSp.Visible = False
End Sub

Private Sub Form_Load()
nOpt = 1
lANew = False  'to add new address from bill
LoadProduct
StartAccessLmt
BranchLoad CboBranch
    CboBranch.ListIndex = 0
    CboBranch.Locked = False

StartSet
FillAcc Grpgrid, -1, "                                                          Account Name                                | Code"

lblNtDes.Caption = cNtDes
dBilldt.Text = datecon(Date)
cboPaytype.ListIndex = 0
FrmDc.Visible = lSuDc
frmOno.Visible = lSuOn

  Label8.Visible = lSuDed
  txtdedP.Visible = lSuDed
  txtDedR.Visible = lSuDed
  
  Label9.Visible = lSuTax
  txtTaxP.Visible = lSuTax
  txtTaxR.Visible = lSuTax
    tAxGrid.Visible = lSuItax
  
  Label12.Visible = lSuBc
  txtBcP.Visible = lSuBc
  txtBcR.Visible = lSuBc
  
  Label15.Visible = lSuPc
  txtPcP.Visible = lSuPc
  txtPcR.Visible = lSuPc
  
If lSuTax Then
    Set rsTxM = New ADODB.Recordset
    rsTxM.Open "select * from tax", Con, adOpenStatic
        If lSuSTax Then
          If Not rsTxM.EOF Then
          If Not rsTxM.BOF Then rsTxM.MoveFirst
             nTxP = rsTxM!taxp
             cTaxDisp = rsTxM!display
          End If
        ElseIf lSuItax Then
          
        End If
    rsTxM.Close
    Set rsTxM = Nothing
End If
txtTaxP = nTxP
If lSuBc Then
Set rsBC = New ADODB.Recordset
rsBC.Open "select * from bnkchgs", Con, adOpenStatic
If Not rsBC.EOF Then
If Not rsBC.BOF Then rsBC.MoveFirst
nBcP = rsBC!fbc
End If
rsBC.Close
Set rsBC = Nothing
End If


End Sub



Private Sub GdwnGrid_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    Text1 = ""
    FlxGrd.TextMatrix(FlxGrd.Row, 8) = GdwnGrid.TextMatrix(GdwnGrid.Row, 0)
    fremGdwn.Visible = False
    FlxGrd.SetFocus
ElseIf KeyAscii = 27 Then
    fremGdwn.Visible = False
End If
End Sub


Private Sub GrpGrid_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   lFnd = True
   txtPartyName.SetFocus
  txtPartyName_KeyPress 13
End If
End Sub


Private Sub ItemLst_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
fillGrdflst
End If
End Sub



Private Sub mnu11_Click()
Unload Me
End Sub

Private Sub mnu13_Click()
If MsgBox("Remove The Line", vbYesNo + vbDefaultButton2) = vbYes Then
   FlxGrd.RemoveItem (FlxGrd.Row)
End If
End Sub

Private Sub mnu2_Click()
Frame4.Visible = False
ClearData
nOpt = 1

'
'      If lSuBra Then
'             Set rsBBnum = New ADODB.Recordset
'            rsBBnum.Open "select * from bbnum where fbranch='" & cboBranch.Text & "'", Con, adOpenStatic
'            If Not rsBBnum Then
'                If Not rsBBnum.BOF Then rsBBnum.MoveFirst
'                   txtBno.Text = rsBBnum!fbillno + 1
'                End If
'                rsBBnum.Close
'            Set rsBBnum = Nothing
'      Else
'            rsNum.Open "select * from num0203d ", Con, adOpenStatic
'            If nBM = 0 Then  '0 bill starts 1 to financial year
'             txtBno = rsNum!frsales + 1
'             nStg = 1
'            ElseIf nBM = 1 Then ' combination bill number and date
'             nStg = 2
'            ElseIf nBM = 2 Then ' every month reset to 1 with add chr
'             nStg = 3
'            End If
'            rsNum.Update
'            rsNum.Close
'      End If
'    End If









    If lSuDc Then
        txtdNo.SetFocus
    Else
        If lSuOn Then
          txtOno.SetFocus
        Else
           FlxGrd.SetFocus
        End If
    End If

End Sub

Private Sub mnu3_Click()
ClearData
nOpt = 2
txtBno.SetFocus
End Sub

Private Sub mnu4_Click()
ClearData
nOpt = 3
txtBno.SetFocus
End Sub

Private Sub mnu5_Click()
ClearData
nOpt = 4
txtBno.SetFocus

End Sub


Private Sub mnu6_Click()
ClearData
nOpt = 5
CboBranch.SetFocus

End Sub


Private Sub Timer1_Timer()
sb1.Panels(2).Text = Time
End Sub


Private Sub txtAdd1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And txtAdd1 <> "" Then
txtAdd2.SetFocus
ElseIf KeyAscii = 13 And txtadd = "" Then
    If lSuDed Then
    txtDedR.SetFocus
    ElseIf lSuTax Then
    txtTaxP.SetFocus
    ElseIf lSuBc Then
    txtBcP.SetFocus
    ElseIf lSuPc Then
    txtPcP.SetFocus
    End If
End If
End Sub


Private Sub txtAdd2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And txtAdd2 <> "" Then
txtAdd3.SetFocus
ElseIf KeyAscii = 13 And txtAdd2 = "" Then
    If lSuDed Then
    txtDedR.SetFocus
    ElseIf lSuTax Then
    txtTaxP.SetFocus
    ElseIf lSuBc Then
    txtBcP.SetFocus
    ElseIf lSuPc Then
    txtPcP.SetFocus
    End If
End If
End Sub


Private Sub txtAdd3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And txtAdd3 <> "" Then
txtadd4.SetFocus
ElseIf KeyAscii = 13 And txtAdd3 = "" Then
    If lSuDed Then
    txtDedR.SetFocus
    ElseIf lSuTax Then
    txtTaxP.SetFocus
    ElseIf lSuBc Then
    txtBcP.SetFocus
    ElseIf lSuPc Then
    txtPcP.SetFocus
    End If
End If

End Sub


Private Sub txtadd4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If lANew Then
  If MsgBox("Add to Master", vbYesNo) = vbYes Then
     StoreAccData
  End If
End If
    If lSuDed Then
    txtDedR.SetFocus
    ElseIf lSuTax Then
    txtTaxP.SetFocus
    ElseIf lSuBc Then
    txtBcP.SetFocus
    ElseIf lSuPc Then
    txtPcP.SetFocus
    End If
End If

End Sub


Private Sub StoreAccData()
Dim cNum As String
Set rsAcc = New ADODB.Recordset
rsAcc.Open "select * from acc0203d", Con, adOpenDynamic, adLockPessimistic
rsAcc.AddNew
Set rsNum = New ADODB.Recordset
    rsNum.Open "select * from num0203d", Con, adOpenDynamic, adLockPessimistic
    cNum = Right(String(5, "0") + Trim(Str(Val(rsNum!facnum) + 1)), 5)
    rsNum!facnum = cNum
    rsNum.Update
rsAcc!faccode = cNum
cCustCode = cNum
rsAcc!facname = txtPartyName
rsAcc!facparent = "000080001000015" + cNum
rsAcc!faclevel = (Len(cGroup + cNum) / 5) * -1
rsAcc!fopbal = 0
rsAcc.Update
rsAcc.Close
Set rsAdd = New ADODB.Recordset
rsAdd.Open "select * from add0203d", Con, adOpenDynamic, adLockPessimistic
rsAdd.AddNew
rsAdd!faccode = cNum
rsAdd!add1 = txtAdd1
rsAdd!add2 = txtAdd2
rsAdd!add3 = txtAdd3
rsAdd!add4 = txtadd4
rsAdd.Update
rsAdd.Close
End Sub

Private Sub txtAmtR_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Val(txtAmtR) <> 0 Then
   txtAmtB = Val(txtAmtR) - Val(txtTotal)
   
   If nOpt = 1 Then
      SaveData
       If MsgBox("Print the bill", vbYesNo) = vbYes Then DelPrint
   ElseIf nOpt = 2 Then
      DelData
      SaveData
      If MsgBox("Print the bill", vbYesNo) = vbYes Then
        DelPrint
      End If
   ElseIf nOpt = 3 Then
   ElseIf nOpt = 5 Then
   DelPrint
   End If
   ClearData
   nOpt = 1
       Frame4.Visible = False
       FlxGrd.SetFocus

ElseIf KeyAscii = 27 Then
    Frame4.Visible = False
       FlxGrd.SetFocus

End If
End Sub


Private Sub txtBcP_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And txtBcP <> "" Then

CallTot
If lSuPc Then
txtPcP.SetFocus
Else
txtAmtR.SetFocus
End If
ElseIf KeyAscii = 27 Then
    Frame4.Visible = False
       FlxGrd.SetFocus

End If

End Sub


Private Sub txtBno_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And txtBno <> "" Then
    If nOpt > 1 Then
         stuffData
         If nOpt = 2 Then
            dBilldt.SetFocus
         ElseIf nOpt = 3 Then
            If MsgBox("cancel this bill", vbYesNo) = vbYes Then
                
            End If
         ElseIf nOpt = 4 Then
            If MsgBox("Delete this bill", vbYesNo) = vbYes Then
                
            End If
         ElseIf nOpt = 5 Then
            If MsgBox("Print this bill", vbYesNo) = vbYes Then
                DelPrint
            End If
         End If
         
    End If
End If
End Sub


Private Sub txtCardNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And txtCardNo <> "" Then
txtBcP = nBcP
Label12.Visible = True
txtBcP.Visible = True
txtBcR.Visible = True
CallTot
txtPartyName.SetFocus
End If
End Sub


Private Sub txtDcDt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        If lSuOn Then
          txtOno.SetFocus
        Else
           FlxGrd.SetFocus
        End If
  

End If
End Sub


Private Sub txtdedP_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
CallTot
txtTaxP.SetFocus
ElseIf KeyAscii = 27 Then
    Frame4.Visible = False

End If
End Sub


Private Sub txtDedR_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
CallTot
txtTaxP.SetFocus
ElseIf KeyAscii = 27 Then
    Frame4.Visible = False

End If

End Sub


Private Sub txtdNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtDcDt.SetFocus
End If
End Sub


Private Sub txtItemLst_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
ItemLst.SetFocus
ElseIf KeyCode = vbKeyUp Then
ItemLst.SetFocus
End If
End Sub

Private Sub txtItemLst_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And lFnd Then

fillGrdflst


ElseIf KeyAscii = 27 Then
    
    fremIte.Visible = False


End If
End Sub


Private Sub txtItemLst_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode <> 13 Then
 lFnd = Find(ItemLst, UCase(txtItemLst.Text), 0, frmBill.fSp)
End If

End Sub

Private Sub txtMdt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
              Frame4.Top = 1800
              Frame4.Left = 3285
              Frame4.Visible = True
              cboPaytype.SetFocus
End If
End Sub


Private Sub txtNote_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtNote = UCase(txtNote)
   If cboPaytype.ListIndex = 2 Then
   txtCardNo.SetFocus
   Else
   txtPartyName.SetFocus
   End If
End If
End Sub


Private Sub txtODt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
FlxGrd.SetFocus
End If
End Sub


Private Sub txtOno_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtODt.SetFocus
End If
End Sub


Private Sub txtPartyName_GotFocus()
    SendKeys "{HOME}"
    SendKeys "+{END}"

End Sub

Private Sub txtPartyName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
Grpgrid.SetFocus
ElseIf KeyCode = vbKeyUp Then
Grpgrid.SetFocus
End If
End Sub


Private Sub txtPartyName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And txtPartyName <> "" And lFnd Then
lANew = False
   'Find GrpGrid, Format(txtPartyName.Text, ">"), 0  for change option we have check this
    txtPartyName.Text = Grpgrid.TextMatrix(Grpgrid.Row, 0)
    cCustCode = Grpgrid.TextMatrix(Grpgrid.Row, 1)
    cGroup = Grpgrid.TextMatrix(Grpgrid.Row, 2)
    If lFnd Then FillAdd cCustCode
    Grpgrid.Visible = False
    If lSuDed Then
    txtDedR.SetFocus
    ElseIf lSuTax Then
    txtTaxP.SetFocus
    ElseIf lSuBc Then
    txtBcP.SetFocus
    ElseIf lSuPc Then
    txtPcP.SetFocus
    End If

ElseIf KeyAscii = 13 And txtPartyName <> "" And Not lFnd Then
    Grpgrid.Visible = False
If nOpt = 1 Then
    txtAdd1 = ""
    txtAdd2 = ""
    txtAdd3 = ""
    txtadd4 = ""
    txtAdd5 = ""
End If
    txtAdd1.SetFocus
    lANew = True
ElseIf KeyAscii = 13 And txtPartyName = "" Then
    If lSuDed Then
    txtDedR.SetFocus
    ElseIf lSuTax Then
    txtTaxP.SetFocus
    ElseIf lSuBc Then
    txtBcP.SetFocus
    ElseIf lSuPc Then
    txtPcP.SetFocus
    End If

ElseIf KeyAscii = 27 Then
    Grpgrid.Visible = False

End If

If KeyAscii = 13 And Not lFnd Then
Grpgrid.Visible = False
End If
End Sub


Private Sub txtPartyName_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode <> vbKeyEscape And KeyCode <> vbKeyRight And KeyCode <> vbKeyLeft _
            And txtPartyName.Text <> "" And KeyCode <> vbKeyReturn Then
   lFnd = Find(Grpgrid, Format(txtPartyName.Text, ">"), 0, frmBill.fSp)
If lFnd Then FillAdd Grpgrid.TextMatrix(Grpgrid.Row, 1)
End If

End Sub


Private Sub txtPcP_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtAmtR.SetFocus
ElseIf KeyAscii = 27 Then
    Frame4.Visible = False
       FlxGrd.SetFocus

End If

End Sub


Private Sub txtTaxP_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
CallTot
txtTrnsChg.SetFocus
ElseIf KeyAscii = 27 Then
    Frame4.Visible = False
    FlxGrd.SetFocus

End If

End Sub


Private Sub txtTrnsChg_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If lSuBc And cboPaytype.ListIndex = 2 Then
   
     
    txtBcP.SetFocus
   ElseIf lSuPc Then
     txtPcP.SetFocus
   Else
     txtAmtR.SetFocus
   End If
ElseIf KeyAscii = 27 Then
    Frame4.Visible = False
       FlxGrd.SetFocus

End If

End Sub


