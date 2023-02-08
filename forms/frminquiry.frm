VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frminquiry 
   Caption         =   "RF-Member Statement Inquiry"
   ClientHeight    =   10170
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15120
   Icon            =   "frminquiry.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10170
   ScaleWidth      =   15120
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtnotes 
      Height          =   615
      Left            =   5760
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   8880
      Width           =   5895
   End
   Begin VB.CheckBox chkissued 
      Caption         =   "Card Issued"
      Height          =   615
      Left            =   14160
      TabIndex        =   77
      Top             =   1080
      Width           =   855
   End
   Begin VB.ComboBox yyear 
      Height          =   315
      ItemData        =   "frminquiry.frx":0442
      Left            =   6720
      List            =   "frminquiry.frx":046D
      TabIndex        =   76
      Text            =   "2011"
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton cmdarchived 
      Caption         =   "Archived"
      Height          =   255
      Left            =   14160
      TabIndex        =   75
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton cmdPhoto 
      Caption         =   "Photo"
      Height          =   285
      Left            =   14295
      TabIndex        =   74
      Top             =   105
      Width           =   900
   End
   Begin VB.CommandButton cmdprintTransaction 
      Caption         =   "&Old Statement "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8940
      TabIndex        =   61
      Top             =   1800
      Width           =   1170
   End
   Begin VB.Frame FraSig 
      Caption         =   "Signatories"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   5520
      TabIndex        =   47
      Top             =   2385
      Visible         =   0   'False
      Width           =   6135
      Begin VB.CommandButton cmdoksig 
         Caption         =   "Ok"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3840
         TabIndex        =   60
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox txtsigid3 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4440
         TabIndex        =   55
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtsig3 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         TabIndex        =   54
         Top             =   960
         Width           =   3015
      End
      Begin VB.TextBox txtsigid2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4440
         TabIndex        =   53
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtsig2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         TabIndex        =   52
         Top             =   600
         Width           =   3015
      End
      Begin VB.TextBox txtsigid1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4440
         TabIndex        =   51
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtsig1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         TabIndex        =   50
         Top             =   240
         Width           =   3015
      End
      Begin VB.TextBox txtsigid4 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4440
         TabIndex        =   49
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox txtsig4 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         TabIndex        =   48
         Top             =   1320
         Width           =   3015
      End
      Begin VB.Label Label22 
         Caption         =   "3 rd Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   59
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label23 
         Caption         =   "2 nd Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   58
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label24 
         Caption         =   "1 st Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   57
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label25 
         Caption         =   "4 th Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   56
         Top             =   1320
         Width           =   855
      End
   End
   Begin VB.Frame FraNominees 
      Caption         =   "Nominees"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   6840
      TabIndex        =   36
      Top             =   2400
      Visible         =   0   'False
      Width           =   4815
      Begin VB.CommandButton cmdok 
         Caption         =   "Ok"
         Height          =   375
         Left            =   2640
         TabIndex        =   46
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox txtnomi1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   720
         TabIndex        =   42
         Top             =   240
         Width           =   2655
      End
      Begin VB.TextBox txtnomiid1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3480
         TabIndex        =   41
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtnomi2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   720
         TabIndex        =   40
         Top             =   600
         Width           =   2655
      End
      Begin VB.TextBox txtnomiid2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3480
         TabIndex        =   39
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtnomi3 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   720
         TabIndex        =   38
         Top             =   960
         Width           =   2655
      End
      Begin VB.TextBox txtnomiid3 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3480
         TabIndex        =   37
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label19 
         Caption         =   "Name1"
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label20 
         Caption         =   "Name2"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label21 
         Caption         =   "Name3"
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   960
         Width           =   1095
      End
   End
   Begin VB.TextBox txtothercharges 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   12780
      TabIndex        =   31
      Top             =   1815
      Width           =   1200
   End
   Begin VB.TextBox txtwithrawalcharges 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   12780
      TabIndex        =   30
      Top             =   1470
      Width           =   1200
   End
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11250
      Begin VB.TextBox txtpayno 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         TabIndex        =   81
         Top             =   240
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.TextBox txtidno 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1800
         TabIndex        =   80
         Top             =   240
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.TextBox Txtaccno 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1800
         MaxLength       =   15
         TabIndex        =   79
         Top             =   240
         Width           =   2070
      End
      Begin VB.CommandButton cmdBen 
         Caption         =   "Beneficiaries"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   10110
         TabIndex        =   73
         Top             =   1800
         Width           =   1125
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "..."
         Height          =   300
         Left            =   7410
         TabIndex        =   68
         Top             =   240
         Width           =   315
      End
      Begin VB.CommandButton cmdsignatories 
         Caption         =   "Signatories"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7800
         TabIndex        =   35
         Top             =   1800
         Width           =   945
      End
      Begin VB.CommandButton cmdnominee 
         Caption         =   "Nominees"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6840
         TabIndex        =   34
         Top             =   1800
         Width           =   900
      End
      Begin VB.TextBox txtwithrawableamount 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   3810
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   1725
         Width           =   3015
      End
      Begin MSComCtl2.DTPicker dtpfromtransdate 
         Height          =   315
         Left            =   5160
         TabIndex        =   27
         Top             =   1410
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   " dd-MM-yyyy"
         Format          =   100466691
         CurrentDate     =   38950
      End
      Begin VB.CommandButton cmdrefresh 
         Caption         =   "Refresh"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5025
         TabIndex        =   25
         Top             =   960
         Width           =   1095
      End
      Begin VB.ComboBox Cbodetail 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frminquiry.frx":04BF
         Left            =   30
         List            =   "frminquiry.frx":04CF
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   240
         Width           =   1740
      End
      Begin VB.TextBox TXTTGNO 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1800
         TabIndex        =   6
         Top             =   240
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.CommandButton cmdrebuild 
         Caption         =   "Update"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3960
         MaskColor       =   &H0080FF80&
         TabIndex        =   5
         Top             =   960
         Width           =   1065
      End
      Begin VB.CommandButton cmdprintstatement 
         Caption         =   "Print Statement"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6120
         MaskColor       =   &H0080FF80&
         TabIndex        =   4
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtdate 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9015
         TabIndex        =   3
         Top             =   1485
         Width           =   2040
      End
      Begin VB.TextBox txtoverdraft 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   1755
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   1575
         Width           =   2055
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Frozen Balance"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   465
         TabIndex        =   72
         Top             =   1275
         Width           =   1245
      End
      Begin VB.Label lblFrozen 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   300
         Left            =   1755
         TabIndex        =   71
         Top             =   1245
         Width           =   2055
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H000000FF&
         Caption         =   "WITHDRAWALABLE AMOUNT "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   225
         Left            =   1305
         TabIndex        =   32
         Top             =   1920
         Width           =   2505
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "From Trans Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3840
         TabIndex        =   26
         Top             =   1455
         Width           =   1200
      End
      Begin VB.Label lblbookbalance 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   300
         Left            =   1755
         TabIndex        =   23
         Top             =   930
         Width           =   2055
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Avail Balance"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   645
         TabIndex        =   22
         Top             =   960
         Width           =   1065
      End
      Begin VB.Label lblavail 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1755
         TabIndex        =   21
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Current Balance"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   405
         TabIndex        =   20
         Top             =   615
         Width           =   1320
      End
      Begin VB.Label lblaccname 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   3840
         TabIndex        =   19
         Top             =   600
         Width           =   3735
      End
      Begin VB.Label lblname 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   3840
         TabIndex        =   18
         Top             =   240
         Width           =   3540
      End
      Begin VB.Label Lblidno 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   9015
         TabIndex        =   17
         Top             =   795
         Width           =   2040
      End
      Begin VB.Label lblaccno 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   9015
         TabIndex        =   16
         Top             =   120
         Width           =   2040
      End
      Begin VB.Label lblmemno 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   9015
         TabIndex        =   15
         Top             =   450
         Width           =   2040
      End
      Begin VB.Label Label1 
         Caption         =   "Account Number"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   7740
         TabIndex        =   14
         Top             =   135
         Width           =   1470
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Member No"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   8100
         TabIndex        =   13
         Top             =   465
         Width           =   810
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "ID N0."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   8505
         TabIndex        =   12
         Top             =   825
         Width           =   420
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Last Withdr Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   7770
         TabIndex        =   11
         Top             =   1515
         Width           =   1200
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "OverDraft"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   930
         TabIndex        =   10
         Top             =   1575
         Width           =   780
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Phone No"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   8295
         TabIndex        =   9
         Top             =   1140
         Width           =   690
      End
      Begin VB.Label LBLTSCNO 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   9015
         TabIndex        =   8
         Top             =   1125
         Width           =   2040
      End
   End
   Begin MSComctlLib.ListView lvememtrans 
      Height          =   5895
      Left            =   0
      TabIndex        =   78
      ToolTipText     =   "Shows actual /available balances for the period specified"
      Top             =   2400
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   10398
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483624
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Image Image2 
      Height          =   975
      Left            =   2400
      Stretch         =   -1  'True
      Top             =   8880
      Width           =   3255
   End
   Begin VB.Image Image1 
      Height          =   1455
      Left            =   360
      Stretch         =   -1  'True
      Top             =   8400
      Width           =   1695
   End
   Begin VB.Label lblLoabBalance 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   12765
      TabIndex        =   70
      Top             =   405
      Width           =   1200
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Loan Balance"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   11760
      TabIndex        =   69
      Top             =   435
      Width           =   990
   End
   Begin VB.Label lblFrequency 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   12765
      TabIndex        =   67
      Top             =   1125
      Width           =   1200
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Frequency Charge"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   11400
      TabIndex        =   66
      Top             =   1170
      Width           =   1350
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Advance Balance"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   11460
      TabIndex        =   65
      Top             =   60
      Width           =   1290
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Lumpsum Charge"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   11490
      TabIndex        =   64
      Top             =   810
      Width           =   1260
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Withdawal Charges"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   11325
      TabIndex        =   63
      Top             =   1530
      Width           =   1425
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Excise Duty"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   11895
      TabIndex        =   62
      Top             =   1905
      Width           =   855
   End
   Begin VB.Label lblstandingorderbalance 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   12765
      TabIndex        =   29
      Top             =   765
      Width           =   1200
   End
   Begin VB.Label lbladvancebalance 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   12765
      TabIndex        =   28
      Top             =   45
      Width           =   1200
   End
   Begin VB.Label Label9 
      BackColor       =   &H80000015&
      Caption         =   "Label9"
      Height          =   1815
      Left            =   -15
      TabIndex        =   24
      Top             =   0
      Width           =   11235
   End
End
Attribute VB_Name = "frminquiry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim myclass As cdbase
Dim cn As Object
Dim Provider As String
Dim rsAdv As New Recordset
Dim rs As Object
Dim withcharges As Currency
Dim lblacname As String
Dim minBal As Currency
Dim wnoticecharges As Currency
Dim rscode As Recordset
Dim sql As String
Dim dateValid As Boolean
Dim lastWithDate As Date
Dim accno As String
Dim PAYNO As String
Dim isgl As Integer
Public Sub checkglacc(accno As String)
On Error GoTo errormesg
sql = ""
sql = "select * from cub where accno='" & accno & "' and ismain=1"
Set rs = oEasySacco.GetRecordset(sql)
If Not rs.EOF Then
isgl = 1
Else
isgl = 0
End If
Exit Sub
errormesg:
MsgBox Err.Description
End Sub
Private Sub cbodetail_Change()
gl_lastmove = Now()
    Cbodetail_Click
    gl_lastmove = Now()
End Sub

Private Sub Cbodetail_Click()
    On Error Resume Next
    gl_lastmove = Now()
    Select Case Cbodetail
        Case "Account Number"
        txtAccno.Enabled = True
        txtAccno.Visible = True
        txtAccno.SetFocus
        txtAccno = lblaccno
        txtidno.Enabled = False
        txtidno.Visible = False
        TXTTGNO.Visible = False
        TXTTGNO.Enabled = False
        txtpayno.Enabled = False
        txtpayno.Visible = False
        Case "IDNo"
        txtidno.Enabled = True
        txtidno.Visible = True
        txtidno.SetFocus
        TXTTGNO.Enabled = False
        TXTTGNO.Visible = False
        txtAccno.Enabled = False
        txtAccno.Visible = False
        txtpayno.Enabled = False
        txtpayno.Visible = False
        txtpayno = ""
        txtAccno = ""
        TXTTGNO = ""
        Case "MemberNo"
        TXTTGNO.Enabled = True
        TXTTGNO.Visible = True
        TXTTGNO.SetFocus
        txtAccno.Enabled = False
        txtAccno.Visible = False
        txtidno.Enabled = False
        txtidno.Visible = False
        txtpayno.Enabled = False
        txtpayno.Visible = False
        txtidno = ""
        txtpayno = ""
        txtAccno = ""
        Case "TeaGrowerNo"
        txtpayno.Enabled = True
        txtpayno.Visible = True
        txtpayno.SetFocus
        txtAccno.Enabled = False
        txtAccno.Visible = False
        txtidno.Enabled = False
        txtidno.Visible = False
        TXTTGNO.Visible = False
        TXTTGNO.Visible = False
        txtidno = ""
        txtpayno = ""
        txtAccno = ""
    End Select
    gl_lastmove = Now()
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdarchived_Click()
frmarchivedinquiry.Show vbModal
End Sub

Private Sub cmdBen_Click()
    On Error GoTo SysError
    If lblaccno = "" Then
        MsgBox "Please enter the Account No", vbInformation, "Beneficiaries"
        Exit Sub
    End If
    MyRecord = lblaccno
    With frmAccBen
        .cmdEdit.Visible = False
        .cmdNew.Visible = False
        .cmdsave.Visible = False
    End With
    frmAccBen.Show , Me
    Exit Sub
SysError:
    MsgBox Err.Description, vbInformation
End Sub

Private Sub cmdnominee_Click()
    FraNominees.Visible = True
    FraSig.Visible = False
End Sub

Private Sub cmdOk_Click()
    FraNominees.Visible = False
End Sub

Private Sub cmdoksig_Click()
    FraSig.Visible = False
End Sub

Private Sub cmdPhoto_Click()
    If txtAccno <> "" Then
        MyRecord = txtAccno
    End If
    frmPhoto.Show
End Sub

Private Sub cmdprintstatement_Click()
    On Error GoTo errorhandler
'    strformula = "{CUB.AccNo}='" & Txtaccno & "'"
'    Show_Sales_Crystal_Report strformula, "Statement1.rpt", ""
'    strformula = ""
    Dim FileName As String, xlApp As Excel.APPLICATION, Myfos As New FileSystemObject, _
    XLFile As TextStream, XLSheet As Excel.Worksheet, xlBook As New Excel.Workbook, VoucherNo _
    As String, Credit As Double, Debit As Double, balance As Double, mTransDate As Date, _
    TransDescription As String
'    Set xlApp = New Excel.APPLICATION
'    Set xlBook = xlApp.Workbooks.Add
'    Cells.Clear
'    xlApp.Visible = True

      'get charges for stmt
'If Current_User.Has_Authority = True Then GoTo t
If Current_User.Can_PrintStmt <> "YES" Then
 MsgBox "Access Denied,You have not been Authorized to print Account Statement!!!!", vbInformation
 Exit Sub
End If


If frminquiry.lblbookbalance < 200 Then
    MsgBox "Available Balance is less than required amount(200 for Statement)!!!", vbExclamation
    Exit Sub
 End If
    
    
   'deduct charges
    Get_transactionNo
    TransactionID = TransactionID
   If Not Save_To_Customer_Balance(frminquiry.txtAccno, frminquiry.txtidno, frminquiry.txtAccno, frminquiry.lblaccname, 200, 0, frminquiry.txtAccno, "Account Statement Charges", Format(Get_Server_Date, "dd/mm/yyyy") _
   , 0, trxno, month(Date), 0, 0, "DR", 0, trxno, Current_User.UserName, "Member Inquiry", "673-099", Date, 0, 0, 0, 0, 0, TransactionID, ErrorMessage) Then
        If ErrorMessage <> "" Then
            MsgBox ErrorMessage, vbInformation, Me.Caption
            ErrorMessage = ""
        End If
   End If
   If Not Save_To_Customer_Balance("673-099", "673-099", "673-099", "SUNDRY INCOME", 200, 0, "673-099", "Account Statement Charges", Format(Get_Server_Date, "dd/mm/yyyy") _
   , 0, trxno, month(Date), 0, 0, "CR", 0, trxno, Current_User.UserName, "Member Inquiry", frminquiry.txtAccno, Date, 0, 0, 0, 0, 0, TransactionID, ErrorMessage) Then
        If ErrorMessage <> "" Then
            MsgBox ErrorMessage, vbInformation, Me.Caption
            ErrorMessage = ""
        End If
   End If
  
t:
    frmDateRange.Show vbModal
    If Not Continue Then
        Exit Sub
    End If
    'anadaiwa nini
    Dim rstded As New Recordset
    'advance
    sql = ""
    sql = "select isnull(sum(stobalance),0)as advance from deduction where accno='" & frminquiry.txtAccno & "' and stobalance>0"
    Set rstded = oEasySacco.GetRecordset(sql)
    If Not rstded.EOF Then
    sql = ""
    sql = "update cub set adbal=" & rstded!advance & " where accno='" & frminquiry.txtAccno & "'"
    oEasySacco.ExecuteThis (sql)
    Else
    sql = ""
    sql = "update cub set adbal=0 where accno='" & frminquiry.txtAccno & "'"
    oEasySacco.ExecuteThis (sql)
    End If
    
    'loan
    
    sql = ""
    sql = "select isnull(sum(balance),0)as loan from loansto where accno='" & frminquiry.txtAccno & "' and balance>0"
    Set rstded = oEasySacco.GetRecordset(sql)
    If Not rstded.EOF Then
    sql = ""
    sql = "update cub set loans=" & rstded!loan & " where accno='" & frminquiry.txtAccno & "'"
    oEasySacco.ExecuteThis (sql)
    Else
    sql = ""
    sql = "update cub set loans=0 where accno='" & frminquiry.txtAccno & "'"
    oEasySacco.ExecuteThis (sql)
    End If
    
    
    txtAccNo_Change
    MousePointer = vbHourglass
    If Not Execute_Command("Exec Delete_Statement", ErrorMessage) Then
        If ErrorMessage <> "" Then
            MsgBox ErrorMessage, vbInformation, Me.Caption
            ErrorMessage = ""
        End If
    End If
    If Not Execute_Command("Exec Update_FrozenAmnt '" & lblaccno & "'," & _
    CDbl(lblFrozen), ErrorMessage) Then
        If ErrorMessage <> "" Then
            MsgBox ErrorMessage, vbInformation, Me.Caption
            ErrorMessage = ""
        End If
    End If
    accno = lblaccno
    For I = 1 To lvememtrans.ListItems.Count
        Set li = lvememtrans.ListItems(I)
        mTransDate = li
        If startdate <= mTransDate Then
            If FinishDate >= mTransDate Then
                VoucherNo = li.SubItems(6)
                Credit = li.SubItems(3)
                Debit = li.SubItems(2)
                balance = IIf(li.SubItems(4) = "", 0, li.SubItems(4))
                TransDescription = li.SubItems(1)
                If Not SAVE_STATEMENT(accno, VoucherNo, Credit, Debit, balance, _
                TransDescription, startdate, FinishDate, mTransDate, lblname, ErrorMessage) Then
                    If ErrorMessage <> "" Then
                        MsgBox ErrorMessage, vbInformation, Me.Caption
                        ErrorMessage = ""
                    End If
                End If
            End If
        End If
    Next I
  
    accno = Left(accno, 3)
    Report_Title = "Account Statement As At" & "  " & FinishDate
    If accno = "200" Then
    Show_Sales_Crystal_Report "", "Member Statement.rpt", Report_Title, True
    ElseIf accno = "300" Then
    Show_Sales_Crystal_Report "", "Member Statement.rpt", Report_Title, True
    Else
    Show_Sales_Crystal_Report "", "Member Statement1.rpt", Report_Title, True
    End If
    MousePointer = vbDefault
    Exit Sub
errorhandler:
    MsgBox Err.Description
    MousePointer = vbDefault
End Sub

Private Sub cmdprintTransaction_Click()
    On Error GoTo SysError
    If Trim$(lblaccno) = "" Then
        MsgBox "Please select the Account to Inquire", vbInformation, "Check Account"
        Exit Sub
    End If
    gl_lastmove = Now()
    MyRecord = lblaccno
    frmOldInquiry.cmdAdjust.Visible = False
    frmOldInquiry.cmdTransfer.Visible = False
    frmOldInquiry.Show vbModal
    gl_lastmove = Now()
    Exit Sub
SysError:
    MsgBox Err.Description, vbInformation, Me.Caption
End Sub

Private Sub cmdrebuild_Click()
    On Error Resume Next
    
    frmrebuilder.Show vbModeless
End Sub

Private Sub cmdrefresh_Click()
    'rebuild_accno txtaccno
    Dim rsProc As Recordset
    Set rsProc = New ADODB.Recordset
    Dim sssql As String
    rebuild_accno txtAccno
txtAccNo_Change
End Sub

Private Sub cmdSearch_Click()
    On Error GoTo SysError
    frmNewAcctsSearch.Show vbModal, Me
    If Continue Then
        If SearchValue <> "" Then
            txtAccno = SearchValue
            SearchValue = ""
        End If
    End If
    Exit Sub
SysError:
    MsgBox Err.Description, vbInformation, Me.Caption
End Sub

Private Sub cmdsignatories_Click()
gl_lastmove = Now()
    FraSig.Visible = True
    FraNominees.Visible = False
gl_lastmove = Now()
End Sub

Private Sub Command1_Click()

End Sub

Private Sub dtpfromtransdate_Change()
On Error Resume Next
    Select Case Cbodetail
        Case "Account Number"
        txtAccNo_Change
        Case "IDNo"
        txtIDNo_Change
        Case "MemberNo"
       
        Case "TeaGrowerNo"
        txttgno_Change
    End Select
End Sub

Private Sub dtpfromtransdate_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Select Case Cbodetail
        Case "Account Number"
        txtAccNo_Change
        Case "IDNo"
        txtIDNo_Change
        Case "MemberNo"
       
        Case "TeaGrowerNo"
        txttgno_Change
    End Select
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Dim NA As String
    chkissued = vbUnchecked
    NA = TRANSACTIONS.Company_Name("name")
    Set myclass = New load
    serverdate = Format(Get_Server_Date, "dd/MM/yyyy")
    frminquiry.Caption = NA & "  RF-Member Statement Inquiry" & "---------" & serverdate & "-----Time Login" & TIME
    With lvememtrans
        .ColumnHeaders.Clear
        .ListItems.Clear
    End With
    dtpfromtransdate = Format(Get_Server_Date, " dd-MM-yyyy")
    dtpfromtransdate = DateSerial(Year(dtpfromtransdate) - 1, month(dtpfromtransdate), Day(dtpfromtransdate))
    With lvememtrans
        .ColumnHeaders.Clear
        .ColumnHeaders.Add 1, , "Trans_Date"
        .ColumnHeaders.Add 2, , "TransDescription"
        .ColumnHeaders.Add 3, , "Dr", , lvwColumnRight
        .ColumnHeaders.Add 4, , "Cr", , lvwColumnRight
        .ColumnHeaders.Add 5, , "Avai_Bal", , lvwColumnRight
        .ColumnHeaders.Add 6, , "Actual_Bal", , lvwColumnRight
        .ColumnHeaders.Add 7, , "Voucher No"
        .ColumnHeaders.Add 8, , "Posted By"
        .ColumnHeaders.Add 9, , "Value_Date"
        .ColumnHeaders.Add 10, , "Branch", 1800
    End With
    lvememtrans.View = lvwReport
    Cbodetail.ListIndex = 0
    lblname = ""
    lblaccname = ""
    lblavail = "0.00"
    lblname = ""
    TXTTGNO = ""
    lblbookbalance = "0.00"
    lbladvancebalance = "0.00"
    lblLoabBalance = "0.00"
    lblstandingorderbalance = "0.00"
    lblFrequency = "0.00"
    txtwithrawalcharges = "0.00"
    txtothercharges = "0.00"
    txtAccno.SetFocus
    If MyRecord <> "" Then
        txtAccno = MyRecord
        MyRecord = ""
    End If
    If Not save_loginhistory(Current_User.UserName, Format(Get_Server_Date, "dd/mm/yyyy"), "Open Member Inquiry module", Format(Get_Server_Date, "HH:MM:SS"), terminal, ErrorMessage) Then
    If ErrorMessage <> "" Then
        MsgBox ErrorMessage, vbInformation, Me.Caption
        ErrorMessage = ""
      End If
    End If
End Sub

Private Sub Form_Resize()
 lvememtrans.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not save_loginhistory(Current_User.UserName, Format(Get_Server_Date, "dd/mm/yyyy"), "Exit Member Inquiry module", Format(Get_Server_Date, "HH:MM:SS"), terminal, ErrorMessage) Then
    If ErrorMessage <> "" Then
        MsgBox ErrorMessage, vbInformation, Me.Caption
        ErrorMessage = ""
        
    End If
End If
End Sub

Private Sub lvememtrans_KeyPress(KeyAscii As Integer)
If KeyAscii = 32 Then
txtidno = ""
txtAccno = ""

TXTTGNO = ""
lvememtrans.Visible = False
End If
On Error Resume Next
txtAccno.SetFocus
End Sub

Private Sub getLastTransDate()
    Dim r As Object
    Dim ans As VbMsgBoxResult
    Dim mynull As Long
    sql = "select transdate as date from customerbalance where  accno='" & txtAccno & "'  and transdescription like '%Withdrawal%' order by transdate desc"
    Set r = CreateObject("adodb.recordset")
    r.Open sql, cn
    If Not r.EOF And Not IsNull(r!Date) Then
        mynull = CLng(Date) - CLng(r!Date)
        lastWithDate = CDate(r!Date)
        If mynull >= 8 Then
            dateValid = True
            'ans = MsgBox("The fall within Withdrawal Interval of Seven days?", vbCritical Or vbYesNo, "Member Transactions Inquiry")
        Else
            dateValid = False
        End If
    Else ' there is no transaction it the custbal table eg. a new account
        dateValid = True
        lastWithDate = Date
    End If
    r.Close
End Sub

Private Sub get_notice_givencharges(charges As Currency)
    Set myclass = New cdbase
    Dim Glaccount As String
    Set cn = CreateObject("adodb.connection")
    Provider = myclass.OpenCon
    cn.Open Provider
    sql = "SELECT OverAmount, UptoAmount, RateO, Period, Glaccount, userange, chfee, overchf FROM WNCharges"
    Dim rsnotice As Recordset
    Set rsnotice = New ADODB.Recordset
    rsnotice.Open sql, cn
    If Not rsnotice.EOF Then
        If Not IsNull(rsnotice.Fields(4)) Then Glaccount = rsnotice.Fields(4)
        If CCur(lblbookbalance) >= rsnotice.Fields(7) Then 'cash handling fee for those who are have given notice.
            wnoticecharges = lblbookbalance * rsnotice.Fields(6) / 100
        Else
            If lblbookbalance > 0 Then
            While Not rsnotice.EOF
                If (rsnotice.Fields(1) >= CCur(lblbookbalance)) Then
                    wnoticecharges = rsnotice.Fields(2)
                    GoTo piussigei
                End If
                rsnotice.MoveNext
            Wend
        End If
    End If
piussigei:
    txtothercharges = wnoticecharges
    End If
End Sub

Private Sub WITHDRAWN_CHARGES()
    Set myclass = New cdbase
    Provider = myclass.OpenCon
    Set cn = CreateObject("adodb.connection")
    cn.Open Provider
    Dim rscharge As Recordset
    Set rscharge = CreateObject("ADODB.Recordset")
    rscharge.Open "SELECT Intonauthorisedod,intonunauthorisedod,intonclearedchqs,intonloans,countercharges,Bankerscheque,CashWithcharges from  SavingsAccountsParameters", cn
    If rscharge.EOF Then Exit Sub
    withcharges = rscharge.Fields(6)
    If lblbookbalance > 0 Then
        Set cn = CreateObject("adodb.connection")
        If Provider = "" Then Provider = myclass.OpenCon
        Set myclass = New cdbase
        cn.Open Provider
        sql = ""
        Dim Comm As Recordset
        sql = "SELECT overamount,uptoamount,rateO FROM WCharges order by pkid"
        Set Comm = New ADODB.Recordset
        Comm.Open sql, cn
        While Not Comm.EOF
            '/ check the first amount and move next
            If (Comm.Fields(1) >= CCur(lblbookbalance)) Then
                withcharges = Comm.Fields(2)
                GoTo piussigei
            End If
            Comm.MoveNext
        Wend
   End If
piussigei:
   txtwithrawalcharges = withcharges
        'End If
End Sub

Private Sub Load_GLAccount(accno As String, ErrMsg As String)
    On Error GoTo SysError
    
    Exit Sub
SysError:
    MsgBox Err.Description, vbInformation, Me.Caption
End Sub

Private Sub Get_Member_Transactions(SearchField As MySearchField, SearchValue As String, Optional _
startdate As Date)
    On Error GoTo ErrHandler
    gl_lastmove = Now()
    Dim adac As String
    Dim Customer As Member_Details, rsMinBal As New Recordset, minBal As Double, _
    WithdrawalCharge As Double, WithdrawalCharge1 As Double, WithdrawableAmt As Double, _
    FrequencyCharge As Double, AvalBalance As Double, ComAmount As Double, FrozenBalance _
    As Double, Exciseduty As Double
    Dim cutbal As Boolean
    cutbal = False
    lblname = ""
    lblaccname = ""
    Lblmemno = ""
    lblaccno = ""
    Lblidno = ""
    lbladvancebalance = "0.00"
    txtsig1 = ""
    txtsig2 = ""
    txtsig3 = ""
    txtsig4 = ""
    txtsigid1 = ""
    lblFrozen = "0.00"
    txtsigid2 = ""
    txtsigid3 = ""
    txtsigid4 = ""
    FrozenBalance = 0
    txtnomi1 = ""
    txtnomi2 = ""
    txtnomi3 = ""
    txtnomiid1 = ""
    txtnomiid2 = ""
    txtnomiid3 = ""
    txtwithrawableamount = 0
    lblavail = 0
    lblbookbalance = 0
    lblFrozen = 0
    LBLTSCNO = ""
    Dim balance As Double
    Dim actualB As Double
    Select Case SearchField
        Case 1 'MemberNo
        Customer = Get_Member_Details(memberno, SearchValue, ErrorMessage)
        Case 2 'AccountNo
        Customer = Get_Member_Details(AccountNo, SearchValue, ErrorMessage)
        Case 3 'IDNo
        Customer = Get_Member_Details(IDNO, SearchValue, ErrorMessage)
        Case 4 'growerno
        Customer = Get_Member_Details(growerno, SearchValue, ErrorMessage)
    End Select
    If ErrorMessage <> "" Then
        MsgBox ErrorMessage, vbInformation, Me.Caption
        ErrorMessage = ""
        Exit Sub
    End If
    
    gl_lastmove = Now()
    lvememtrans.ColumnHeaders.Clear
    If Customer.AccountNo <> "" Then
            
        If Trim(Customer.AccTypeName) = "GENERAL LEDGER" Then
            Current_User = Get_User_Details(Current_User.UserID, ErrorMessage)
            If Not Current_User.Has_Authority Then
                MsgBox "You do not have Access to this Account", vbExclamation, "ACCOUNT INQUIRY"
                Editing = False
                Exit Sub
            End If
            
        ElseIf Customer.IsStaff Then
            Current_User = Get_User_Details(Current_User.UserID, ErrorMessage)
            If Not Current_User.Can_viewStaff Then
                MsgBox "You do not have Access to this Account", vbExclamation, "ACCOUNT INQUIRY"
               Editing = False
                Exit Sub
            End If
       
       
        Else
            DateCut = dtpfromtransdate
            Update_Acc_Balances Customer.AccountNo
            'Update_Acc_BalancesCut Customer.AccountNo, DateCut
        End If
        
        Editing = True
        lblname = Customer.MemberName
        lblaccname = Trim(Customer.AccTypeName)
        Lblmemno = Customer.memberno
        lblaccno = Customer.AccountNo
        Lblidno = Customer.IDNO
        LBLTSCNO = Customer.PhoneNo
        If Customer.Authorized = False Then
            MsgBox "The Account has not been authorized,forward to be authorized", vbInformation
            Editing = False
            Exit Sub
        End If
        If Trim(Customer.AccTypeName) = "SAVINGS ACCOUNT" Then
            If Customer.IsDormant = True Then
                'MsgBox "This Account Is Dormant. Please Activate the Account.", vbExclamation, "DORMANT ACCOUNT"
                 MsgBox "This Account Is Dormant. Please Activate the Account.", vbInformation
               Editing = False
                Exit Sub
                
            End If
        End If
        lblavail = Format(Customer.BookBalance, CfMt)
        txtdate = Customer.LastWithdrawalDate
        txtnomi1 = Customer.Nominee1
        txtnomi2 = Customer.Nominee2
        txtnomi3 = Customer.Nominee3
        txtnomiid1 = Customer.NomID1
        txtnomiid2 = Customer.NomID2
        txtnomiid3 = Customer.NomID3
        txtsig1 = Customer.Signatory1
        txtsig2 = Customer.Signatory2
        txtsig3 = Customer.Signatory3
        txtsig4 = Customer.Signatory4
        txtsigid1 = Customer.SigID1
        txtsigid2 = Customer.SigID2
        txtsigid3 = Customer.SigID3
        txtsigid4 = Customer.SigID4
        txtnotes = Customer.notes
        If Trim(Customer.AccTypeName) <> "GENERAL LEDGER" Then
                If Trim(Customer.AccTypeName) = "SAVINGS ACCOUNT" Then
                   If Customer.IsDormant And Customer.BookBalance < 1000 Then
                      MsgBox "This Account Is Dormant. You are required to" & vbCrLf _
                       & "  Activate your Account", vbExclamation, Me.Caption
                       Editing = False
                       Exit Sub
                   
                   End If
                End If
            End If
        Dim RSPIC As New ADODB.Recordset
        
            'rs.Open "Select * from CUB where ACCNO = '" & Trim(txtAccNo) & "'", cnn, 1, 2
            sql = "Select * from CUB where ACCNO = '" & Trim(txtAccno) & "'"
            Set RSPIC = oEasySacco.GetRecordset(sql)
            On Error GoTo dd
            If Not RSPIC.EOF Then
                On Error GoTo dd
                Set Image1.Datasource = RSPIC
                Image1.DataField = "" & "PICTURE"

            Else
            
            End If
                 sql = "Select * from CUB where ACCNO = '" & Trim(txtAccno) & "'"
            Set RSPIC = oEasySacco.GetRecordset(sql)
            If Not RSPIC.EOF Then
            'On Error GoTo CC
            Set Image2.Datasource = RSPIC
                Image2.DataField = "signature"
'CC:
            Else
               
            End If
dd:
        If Customer.IDNO = "" Then
            If Trim(Customer.AccTypeName) <> "GENERAL LEDGER" Then
                MsgBox "The IDNo for this Account is not updated. Please update the IDNo", _
                vbInformation, "Accounts Update"
                Editing = False
                MyRecord = lblaccno
                frmcustomeropenningbalances.cmdNew.Visible = False
                frmcustomeropenningbalances.Show vbModal
                If MyRecord <> "" Then
                    Cbodetail = "Account Number"
                    txtAccno = ""
                    txtAccno = MyRecord
                    MyRecord = ""
                End If
                Exit Sub
            End If
        End If
        lbladvancebalance = Format(Customer.AdvBalance, CfMt)
        lblLoabBalance = Format(Customer.LoanBalance, CfMt)
        FrozenBalance = Get_Frozen_Balance(Customer.AccountNo, ErrorMessage)
        If FrozenBalance = 0 Then
            If ErrorMessage <> "" Then
                MsgBox ErrorMessage, vbInformation, "Frozen Balance"
                ErrorMessage = ""
            End If
        End If
        lblFrozen = Format(FrozenBalance, CfMt)
        If Customer.Frozen Then
            MsgBox "The Account Is Frozen !!!!!!!! " & vbCrLf _
            & "Reason: " & Customer.mReason, vbExclamation, "FROZEN ACCOUNT"
            Editing = False
            'Unload Me
            Exit Sub
        End If
        'XXXXXXXXXXX Check The Account Type Details XXXXXXXXXXXXXXXXX
        If Trim(Customer.AccTypeName) = "SAVINGS ACCOUNT" Then
            If DateDiff("d", Customer.LastWithdrawalDate, Get_Server_Date) < 7 Then
                FrequencyCharge = Get_Charge_Amount("006", ErrorMessage)
                lblFrequency = Format(FrequencyCharge, CfMt)
            Else
                lblFrequency = "0.00"
            End If
            WithdrawalCharge = Get_Charge_Amount("001", ErrorMessage, Customer.WithdrawableAmount)
            Exciseduty = Get_Charge_Amount("048", ErrorMessage, WithdrawalCharge)
            WithdrawableAmt = Format(Get_Withdrawable_Amount(Customer.WithdrawableAmount - _
            FrequencyCharge - Exciseduty, "001"), CfMt)
        End If
        If Trim(Customer.AccTypeName) = "SACCO B SAVINGS ACCOUNT" Then
            If DateDiff("d", Customer.LastWithdrawalDate, Get_Server_Date) < 7 Then
                FrequencyCharge = Get_Charge_Amount("006", ErrorMessage)
                lblFrequency = Format(FrequencyCharge, CfMt)
            Else
                lblFrequency = "0.00"
            End If
            WithdrawalCharge = Get_Charge_Amount("047", ErrorMessage, Customer.WithdrawableAmount)
            Exciseduty = Get_Charge_Amount("048", ErrorMessage, WithdrawalCharge)
            WithdrawableAmt = Format(Get_Withdrawable_Amount(Customer.WithdrawableAmount - _
            FrequencyCharge - Exciseduty, "047"), CfMt)
        End If
        If Trim(Customer.AccTypeName) = "CURRENT ACCOUNT" Then
            If Not Customer.IsStaff Then
                WithdrawalCharge = Get_Charge_Amount("001", ErrorMessage, Customer.WithdrawableAmount)
                WithdrawableAmt = Get_Withdrawable_Amount(Customer.WithdrawableAmount - Exciseduty, "001")
            End If
        End If
        Dim cn As New connection, rsc As New Recordset, CAS As String, SAC As String, rsac As New Recordset
        
        If Trim(Customer.AccTypeName) = "ADVANCE ACCOUNT" Then
        CAS = "200" & Right(txtAccno, Len(txtAccno) - 3)
        If cn.State = adStateClosed Then
        cn.Open "KONOIN"
        End If
        Set rsAdv = cn.Execute("Select Sum(StoBalance) as Balance From DEDUCTION where" _
    & " AccNo='" & CAS & "'")
    lbladvancebalance = IIf(IsNull(rsAdv!balance), 0, rsAdv!balance)
        WithdrawableAmt = 0
        End If
        If Trim(Customer.AccTypeName) = "SACCO B ADVANCE ACCOUNT" Then
        SAC = "300" & Right(txtAccno, Len(txtAccno) - 3)
        If cn.State = adStateClosed Then
        cn.Open "KONOIN"
        End If
        Set rsac = cn.Execute("Select Sum(StoBalance) as Balance From DEDUCTION where" _
    & " AccNo='" & SAC & "'")
    lbladvancebalance = IIf(IsNull(rsac!balance), 0, rsac!balance)
        WithdrawableAmt = 0
        End If
        
        If Trim(Customer.AccTypeName) = "CASH MEMBER" Then
            WithdrawalCharge = Get_Charge_Amount("001", ErrorMessage, Customer.WithdrawableAmount)
            txtwithrawalcharges = Format(WithdrawalCharge, CfMt)
            WithdrawableAmt = Get_Withdrawable_Amount(Customer.WithdrawableAmount - 2, "001")
        End If
        If WithdrawableAmt = 0 Then
            WithdrawableAmt = Format(Customer.WithdrawableAmount - Exciseduty, CfMt)
        End If
        lblbookbalance = Format(Customer.WithdrawableAmount, CfMt)
        txtwithrawableamount = Format(WithdrawableAmt, CfMt)
        If Trim(Customer.AccTypeName) = "SAVINGS ACCOUNT" Then
            txtwithrawalcharges = Format(Get_Charge_Amount("001", ErrorMessage, _
            CDbl(txtwithrawableamount)) + FrequencyCharge, CfMt)
        ElseIf Trim(Customer.AccTypeName) = "CURRENT ACCOUNT" Then
            If Not Customer.IsStaff Then
                txtwithrawalcharges = Format(Get_Charge_Amount("001", ErrorMessage, _
                CDbl(txtwithrawableamount)), CfMt)
            Else
                txtwithrawalcharges = "0.00"
            End If
        ElseIf Trim(Customer.AccTypeName) = "SACCO B SAVINGS ACCOUNT" Then
            If Not Customer.IsStaff Then
                txtwithrawalcharges = Format(Get_Charge_Amount("047", ErrorMessage, _
                CDbl(txtwithrawableamount)), CfMt)
            Else
                txtwithrawalcharges = "0.00"
            End If
        End If
'        If Trim(Customer.AccTypeName) = "SAVINGS ACCOUNT" Then
'        WithdrawalCharge1 = Get_Charge_Amount("001", ErrorMessage, CDbl(txtwithrawableamount))
        
        Dim rsadvance As New Recordset, strAccNo As String
        If Customer.AccountNo <> "" Then
            strAccNo = Left(Customer.AccountNo, 4) & "1" & Right(Customer.AccountNo, _
            Len(Customer.AccountNo) - 5)
        End If
        txtothercharges = Exciseduty
        With lvememtrans
            .ColumnHeaders.Clear
            .ListItems.Clear
            .ColumnHeaders.Add 1, , "Trans_Date"
            .ColumnHeaders.Add 2, , "TransDescription", 2500
            .ColumnHeaders.Add 3, , "Dr", , lvwColumnRight
            .ColumnHeaders.Add 4, , "Cr", , lvwColumnRight
            .ColumnHeaders.Add 5, , "Avai_Bal", , lvwColumnRight
            .ColumnHeaders.Add 6, , "Actual_Bal", , lvwColumnRight
            .ColumnHeaders.Add 7, , "Voucher No"
            .ColumnHeaders.Add 8, , "Posted By"
            .ColumnHeaders.Add 9, , "Value_Date"
            .ColumnHeaders.Add 10, , "Branch", 1800
            .ColumnHeaders.Add 11, , "TransactionID", 1800
            .Visible = True
        End With
        If Trim(Customer.AccTypeName) = "ADVANCE ACCOUNT" Then
            txtwithrawableamount = "0.00"
            'lbladvancebalance = !Balance
        End If
        Dim I As Long, Glaccount As GLAccount_Details
        I = 0
    
        If Trim(Customer.AccTypeName) = "GENERAL LEDGER" Then
            Set rsLoanGuar = oEasySacco.GetRecordset("Exec Get_Teller_Balances '" & Customer.AccountNo & "'")
            'Set rsLoanGuar = oEasySacco.GetRecordset("set dateformat dmy Exec Get_Balances_GL '" & Customer.AccountNo & "'," & yyear & "")
        ElseIf Trim(Customer.AccTypeName) = "ADVANCE ACCOUNT" Then
            Set rsLoanGuar = oEasySacco.GetRecordset("set dateformat dmy Exec Get_Adv_Balances '" & Customer.AccountNo & "'")
        Else
            Set rsLoanGuar = oEasySacco.GetRecordset("set dateformat dmy Exec Get_Balances '" & Customer.AccountNo & "'")
            'Set rsLoanGuar = oEasySacco.GetRecordset("set dateformat dmy Exec Get_BalancesCut '" & Customer.AccountNo & "','" & dtpfromtransdate & "'")
        End If
        If Trim(Customer.AccTypeName) = "GENERAL LED" Then
           Trim(Customer.AccTypeName) = "GENERAL LEDGER"
        End If
        With rsLoanGuar
            If .State = adStateOpen Then
            Select Case Trim(Customer.AccTypeName)
                Case "GENERAL LEDGER"
                balance = Customer.OpeningBalance
                Glaccount = Get_GLAccount_Details(Customer.AccountNo, ErrorMessage)
                If Glaccount.accno = "" Then
                    If ErrorMessage <> "" Then
                        MsgBox ErrorMessage, vbInformation, Me.Caption
                        ErrorMessage = ""
                        Exit Sub
                    End If
                Else
                    If Glaccount.OpeningBalance > 0 Then
                        Set li = lvememtrans.ListItems.Add(, , IIf(IsNull(!transdate), "", !transdate))
                        li.SubItems(1) = "Opening Balance"
                        Select Case Glaccount.NormalBalance
                            Case "DR"
                            Select Case Glaccount.OpeningBalance
                                Case Is > 0
                                li.SubItems(2) = Format(Glaccount.OpeningBalance, CfMt)
                                li.SubItems(3) = "0.00"
                                li.SubItems(4) = li.SubItems(2)
                                li.SubItems(5) = li.SubItems(2)
                                Case Is < 0
                                li.SubItems(3) = Format(Glaccount.OpeningBalance, CfMt)
                                li.SubItems(2) = "0.00"
                                li.SubItems(4) = Format(CDbl(li.SubItems(2)) * (-1), CfMt)
                                li.SubItems(5) = li.SubItems(4)
                            End Select
                            Case "CR"
                            Select Case Glaccount.OpeningBalance
                                Case Is > 0
                                li.SubItems(3) = Format(Glaccount.OpeningBalance, CfMt)
                                li.SubItems(2) = "0.00"
                                li.SubItems(4) = li.SubItems(2)
                                li.SubItems(5) = li.SubItems(2)
                                Case Is < 0
                                li.SubItems(2) = Format(Glaccount.OpeningBalance, CfMt)
                                li.SubItems(3) = "0.00"
                                li.SubItems(4) = Format(CDbl(li.SubItems(2)) * (-1), CfMt)
                                li.SubItems(5) = li.SubItems(4)
                            End Select
                        End Select
                    End If
                End If
            End Select
            If (Customer.AccTypeName) <> "GENERAL LEDGER" Then
            
            Dim rsw As New Recordset
                    Dim availopdate As Date
                    availopdate = "23/10/2008"
                    sql = ""
                    sql = " SET dateformat dmy Select top 1 * From CUSTOMERBALANCE where AccNo='" & accno & "' and transdate <'" & availopdate & "'order by TransDate desc,CustomerBalanceID desc"
                    Set rsw = oEasySacco.GetRecordset(sql)
                    If Not rsw.EOF Then
                    AvalBalance = rsw!AvailableBalance
                    Else
                    AvalBalance = 0
                    End If
            
            End If
            'check previous Avai;able balance
'                    Dim rsw As New Recordset
'                    sql = ""
'                    sql = " SET dateformat dmy Select top 1 * From CUSTOMERBALANCE where AccNo='" & accno & "' and transdate <'" & dtpfromtransdate & "'order by TransDate desc,CustomerBalanceID desc"
'                    Set rsw = oEasySacco.GetRecordset(sql)
'                    If Not rsw.EOF Then
'                        If cutbal = False Then
'                            actualB = Format(IIf(IsNull(!actualbalance), 0, !actualbalance), CfMt)
'                            cutbal = True
''                        Else
''                            actualB = actualB + Format(IIf(IsNull(!Amount), 0, !Amount), CfMt)
'                        End If
'                    End If
            While Not .EOF
            gl_lastmove = Now()
                
                I = I + 1
                Set li = lvememtrans.ListItems.Add(, , IIf(IsNull(!transdate), "", !transdate))
                If Not IsNull(!vno) Then
                    If Not IsNumeric(!vno) Then
                        li.SubItems(1) = IIf(IsNull(!TransDescription), "", !TransDescription)
                    Else
                        li.SubItems(1) = IIf(IsNull(!TransDescription), "", !TransDescription) & ", " & IIf(IsNull(!vno), "", !vno)
                    End If
                Else
                    li.SubItems(1) = IIf(IsNull(!TransDescription), "", !TransDescription)
                End If
                Select Case Trim(Customer.AccTypeName)
                    Case "GENERAL LEDGER"
                    Select Case Glaccount.NormalBalance
                        Case "DR"
                        If UCase(!TransType) = "DR" Then
                            balance = balance + IIf(IsNull(!Amount), 0, !Amount)
                            li.SubItems(2) = Format(IIf(IsNull(!Amount), 0, !Amount), CfMt)
                            li.SubItems(3) = "0.00"
'                                    If Left(UCase(!TransDescription), 10) = "CHEQUE DEP(UNCLEARED)" Then
'                                        MsgBox "NN"
'                                    Else
'                                    Balance = Balance + IIf(IsNull(!Amount), 0, !Amount)
'                                    End If
'                                    li.SubItems(3) = Format(IIf(IsNull(!Amount), 0, !Amount), CfMt)
'                                    li.SubItems(2) = "0.00"
                        Else
                            balance = balance - IIf(IsNull(!Amount), 0, !Amount)
                            li.SubItems(3) = Format(IIf(IsNull(!Amount), 0, !Amount), CfMt)
                            li.SubItems(2) = "0.00"
                        End If
                        Case "CR"
                        If UCase(!TransType) = "DR" Then
                            balance = balance - IIf(IsNull(!Amount), 0, !Amount)
                            li.SubItems(2) = Format(IIf(IsNull(!Amount), 0, !Amount), CfMt)
                            li.SubItems(3) = "0.00"
                        Else
                            balance = balance + IIf(IsNull(!Amount), 0, !Amount)
                            li.SubItems(3) = Format(IIf(IsNull(!Amount), 0, !Amount), CfMt)
                            li.SubItems(2) = "0.00"
                        End If
                    End Select
                    Case "SACCO B ADVANCE ACCOUNT"
                    If UCase(!TransType) = "DR" Then
                        balance = balance + IIf(IsNull(!Amount), 0, !Amount)
                        
                        'AvalBalance = AvalBalance + IIf(IsNull(!Amount), 0, !Amount)
                        If !transdate >= availopdate Then
                            AvalBalance = AvalBalance + IIf(IsNull(!Amount), 0, !Amount)
                            End If
                        li.SubItems(2) = Format(IIf(IsNull(!Amount), 0, !Amount), CfMt)
                        li.SubItems(3) = "0.00"
                        If cutbal = False Then
                            actualB = actualB + Format(IIf(IsNull(!actualbalance), 0, !actualbalance), CfMt)
                            cutbal = True
                        Else
                            actualB = actualB + Format(IIf(IsNull(!Amount), 0, !Amount), CfMt)
                        End If
                    Else
                        balance = balance - IIf(IsNull(!Amount), 0, !Amount)
                        'AvalBalance = AvalBalance - IIf(IsNull(!Amount), 0, !Amount)
                        If !transdate >= availopdate Then
                            AvalBalance = AvalBalance - IIf(IsNull(!Amount), 0, !Amount)
                            End If
                        li.SubItems(3) = Format(IIf(IsNull(!Amount), 0, !Amount), CfMt)
                        li.SubItems(2) = "0.00"
                        If cutbal = False Then
                            actualB = actualB - Format(IIf(IsNull(!actualbalance), 0, !actualbalance), CfMt)
                            cutbal = True
                        Else
                            actualB = actualB - Format(IIf(IsNull(!Amount), 0, !Amount), CfMt)
                        End If
                    End If
                    Case "ADVANCE ACCOUNT"
                    If UCase(!TransType) = "DR" Then
                        balance = balance + IIf(IsNull(!Amount), 0, !Amount)
                        'AvalBalance = AvalBalance + IIf(IsNull(!Amount), 0, !Amount)
                        If !transdate >= availopdate Then
                            AvalBalance = AvalBalance + IIf(IsNull(!Amount), 0, !Amount)
                            End If
                        li.SubItems(2) = Format(IIf(IsNull(!Amount), 0, !Amount), CfMt)
                        li.SubItems(3) = "0.00"
                        If cutbal = False Then
                            actualB = actualB + Format(IIf(IsNull(!actualbalance), 0, !actualbalance), CfMt)
                            cutbal = True
                        Else
                            actualB = actualB + Format(IIf(IsNull(!Amount), 0, !Amount), CfMt)
                        End If
                    Else
                        balance = balance - IIf(IsNull(!Amount), 0, !Amount)
                        'AvalBalance = AvalBalance - IIf(IsNull(!Amount), 0, !Amount)
                        If !transdate >= availopdate Then
                            AvalBalance = AvalBalance - IIf(IsNull(!Amount), 0, !Amount)
                            End If
                        li.SubItems(3) = Format(IIf(IsNull(!Amount), 0, !Amount), CfMt)
                        li.SubItems(2) = "0.00"
                        If cutbal = False Then
                            actualB = actualB - Format(IIf(IsNull(!actualbalance), 0, !actualbalance), CfMt)
                            cutbal = True
                        Else
                            actualB = actualB - Format(IIf(IsNull(!Amount), 0, !Amount), CfMt)
                        End If
                    End If
                    Case "OKOA ADVANCE ACCOUNT"
                    If UCase(!TransType) = "DR" Then
                        balance = balance + IIf(IsNull(!Amount), 0, !Amount)
                        'AvalBalance = AvalBalance + IIf(IsNull(!Amount), 0, !Amount)
                        If !transdate >= availopdate Then
                            AvalBalance = AvalBalance + IIf(IsNull(!Amount), 0, !Amount)
                            End If
                        li.SubItems(2) = Format(IIf(IsNull(!Amount), 0, !Amount), CfMt)
                        li.SubItems(3) = "0.00"
                        If cutbal = False Then
                            actualB = actualB + Format(IIf(IsNull(!actualbalance), 0, !actualbalance), CfMt)
                            cutbal = True
                        Else
                            actualB = actualB + Format(IIf(IsNull(!Amount), 0, !Amount), CfMt)
                        End If
                    Else
                        balance = balance - IIf(IsNull(!Amount), 0, !Amount)
                        'AvalBalance = AvalBalance - IIf(IsNull(!Amount), 0, !Amount)
                        If !transdate >= availopdate Then
                            AvalBalance = AvalBalance - IIf(IsNull(!Amount), 0, !Amount)
                            End If
                        li.SubItems(3) = Format(IIf(IsNull(!Amount), 0, !Amount), CfMt)
                        li.SubItems(2) = "0.00"
                        If cutbal = False Then
                            actualB = actualB - Format(IIf(IsNull(!actualbalance), 0, !actualbalance), CfMt)
                            cutbal = True
                        Else
                            actualB = actualB - Format(IIf(IsNull(!Amount), 0, !Amount), CfMt)
                        End If
                    End If
                    Case Else
                    If UCase(!TransType) = "DR" Then
                        balance = balance - IIf(IsNull(!Amount), 0, !Amount)
                       
                        If !transdate >= availopdate Then
                            AvalBalance = AvalBalance - IIf(IsNull(!Amount), 0, !Amount)
                        End If
                        li.SubItems(2) = Format(IIf(IsNull(!Amount), 0, !Amount), CfMt)
                        li.SubItems(3) = "0.00"
                        If cutbal = False Then
                            actualB = actualB - Format(IIf(IsNull(!actualbalance), 0, !actualbalance), CfMt)
                            cutbal = True
                        Else
                            actualB = actualB - Format(IIf(IsNull(!Amount), 0, !Amount), CfMt)
                        End If
                    Else
                        If Left(UCase(!TransDescription), 10) = "CHEQUE DEP" Then
                        
                        Else
                            balance = balance + IIf(IsNull(!Amount), 0, !Amount)
                            If !transdate >= availopdate Then
                            AvalBalance = AvalBalance + IIf(IsNull(!Amount), 0, !Amount)
                            End If
                            If cutbal = False Then
                            actualB = actualB + Format(IIf(IsNull(!actualbalance), 0, !actualbalance), CfMt)
                            cutbal = True
                            Else
                                actualB = actualB + Format(IIf(IsNull(!Amount), 0, !Amount), CfMt)
                            End If
                        End If
                        li.SubItems(3) = Format(IIf(IsNull(!Amount), 0, !Amount), CfMt)
                        li.SubItems(2) = "0.00"
                    End If
                End Select
                Select Case Trim(Customer.AccTypeName)
                    Case "GENERAL LEDGER"
                    li.SubItems(4) = Format(balance, CfMt)
                    Case "ADVANCE ACCOUNT"
                    li.SubItems(4) = Format(balance, CfMt)
                    Case "SACCO B ADVANCE ACCOUNT"
                    li.SubItems(4) = Format(balance, CfMt)
                    Case "OKOA ADVANCE ACCOUNT"
                    li.SubItems(4) = Format(balance, CfMt)
                    Case Else
                    If !transdate < availopdate Then
                        li.SubItems(4) = Format(IIf(IsNull(!AvailableBalance), 0, !AvailableBalance), CfMt)
                    Else
                        
                        li.SubItems(4) = Format(AvalBalance, CfMt)
                    End If
                    AvalBalance = CDbl(li.SubItems(4))
                End Select
                li.SubItems(5) = Format(balance, CfMt)
                'li.SubItems(5) = Format(actualB, CfMt) '+ Format(IIf(IsNull(!Amount), 0, !Amount), CfMt)
                li.SubItems(6) = IIf(IsNull(!vno), "", !vno)
                li.SubItems(7) = IIf(IsNull(!auditid), "", !auditid)
                li.SubItems(8) = IIf(IsNull(!valuedate), IIf(IsNull(!transdate), "", !transdate), !valuedate)
                If !BranchCode = 0 Then
                li.SubItems(9) = "Litein"
                ElseIf !BranchCode = 1 Then
                li.SubItems(9) = "Sotik"
                ElseIf !BranchCode = 2 Then
                li.SubItems(9) = "Roret"
                ElseIf !BranchCode = 3 Then
                li.SubItems(9) = "Cheborgei"
                ElseIf !BranchCode = 4 Then
                li.SubItems(9) = "Chebirbelek"
                End If
                li.SubItems(10) = IIf(IsNull(!TransactionID), "", !TransactionID)
                gl_lastmove = Now()
                .MoveNext
            Wend
            End If
        End With
        If Trim(Customer.AccTypeName) = "GENERAL LEDGER" Then
            If Not Execute_Command("Update GLSETUP Set CurrentBal=" & balance & " where " _
            & "AccNo='" & SearchValue & "'", ErrorMessage) Then
                If ErrorMessage <> "" Then
                    MsgBox ErrorMessage, vbInformation, "Account Update"
                    ErrorMessage = ""
                End If
            End If
        Else
            If Not Execute_Command("Update CUB Set AvailaBleBalance=" & Customer.BookBalance & " where " _
            & "AccNo='" & SearchValue & "'", ErrorMessage) Then
                If ErrorMessage <> "" Then
                    MsgBox ErrorMessage, vbInformation, "Account Update"
                    ErrorMessage = ""
                End If
            End If
            If Not Execute_Command("Update CASHPROCEEDSMEMBERS Set Balance=" & AvalBalance _
            & " where MemberNo='" & SearchValue & "'", ErrorMessage) Then
                If ErrorMessage <> "" Then
                    MsgBox ErrorMessage, vbInformation, Me.Caption
                    ErrorMessage = ""
                End If
            End If
            lblavail = Format(Customer.BookBalance, CfMt)
            'Format(Customer.BookBalance, CfMt)
        End If
        'XXXXXXXXXXXXXXXXXXXXX Reload The Account Details XXXXXXXXXXXXXXXXXXXXXXXXXX
        Customer = Get_Member_Details(AccountNo, lblaccno, ErrorMessage)
        If Customer.AccountNo <> "" Then
            lblbookbalance = Format(Customer.WithdrawableAmount, CfMt)
        End If
        If Trim(Customer.AccTypeName) = "SAVINGS ACCOUNT" Then
            If DateDiff("d", Customer.LastWithdrawalDate, Get_Server_Date) < 7 Then
                FrequencyCharge = Get_Charge_Amount("006", ErrorMessage)
                lblFrequency = Format(FrequencyCharge, CfMt)
            Else
                lblFrequency = "0.00"
            End If
            WithdrawalCharge = Get_Charge_Amount("001", ErrorMessage, Customer.WithdrawableAmount)
            Exciseduty = Get_Charge_Amount("048", ErrorMessage, WithdrawalCharge)
            WithdrawableAmt = Format(Get_Withdrawable_Amount(Customer.WithdrawableAmount - _
            FrequencyCharge - Exciseduty, "001"), CfMt)
        End If
        If Trim(Customer.AccTypeName) = "CURRENT ACCOUNT" Then
            If Not Customer.IsStaff Then
                WithdrawalCharge = Get_Charge_Amount("001", ErrorMessage, Customer.WithdrawableAmount)
                WithdrawableAmt = Get_Withdrawable_Amount(Customer.WithdrawableAmount - Exciseduty, "001")
            End If
        End If
        If WithdrawableAmt = 0 Then
            WithdrawableAmt = Format(Customer.WithdrawableAmount, CfMt)
        End If
        lblbookbalance = Format(Customer.WithdrawableAmount, CfMt)
        If Trim(Customer.AccTypeName) = "SACCO B SAVINGS ACCOUNT" Then
        If DateDiff("d", Customer.LastWithdrawalDate, Get_Server_Date) < 7 Then
                FrequencyCharge = Get_Charge_Amount("006", ErrorMessage)
                lblFrequency = Format(FrequencyCharge, CfMt)
            Else
                lblFrequency = "0.00"
            End If
            WithdrawalCharge = Get_Charge_Amount("047", ErrorMessage, Customer.WithdrawableAmount)
            Exciseduty = Get_Charge_Amount("048", ErrorMessage, WithdrawalCharge)
            WithdrawableAmt = Format(Get_Withdrawable_Amount(Customer.WithdrawableAmount - _
            FrequencyCharge - Exciseduty, "047"), CfMt)
            txtwithrawableamount = WithdrawableAmt ' txtwithrawableamount - WithdrawalCharge - Exciseduty
        Else
        txtwithrawableamount = Format(WithdrawableAmt, CfMt)
        End If
        If Trim(Customer.AccTypeName) = "SAVINGS ACCOUNT" Then
            txtwithrawalcharges = Format(Get_Charge_Amount("001", ErrorMessage, _
            CDbl(txtwithrawableamount)) + FrequencyCharge, CfMt)
        ElseIf Trim(Customer.AccTypeName) = "CURRENT ACCOUNT" Then
            If Not Customer.IsStaff Then
                txtwithrawalcharges = Format(Get_Charge_Amount("001", ErrorMessage, _
                CDbl(txtwithrawableamount)), CfMt)
            Else
                txtwithrawalcharges = "0.00"
            End If
        End If
'        WithdrawalCharge1 = Get_Charge_Amount("001", ErrorMessage, CDbl(txtwithrawableamount))
        
        If Trim(Customer.AccTypeName) = "GENERAL LEDGER" Then
            lblavail = Format(balance, CfMt)
            txtwithrawableamount = lblavail
        Else
            lblavail = Format(Customer.BookBalance, CfMt)
            If txtwithrawalcharges = "" Then txtwithrawalcharges = "0"
            txtwithrawableamount = Format(CDbl(lblbookbalance) - CDbl(txtwithrawalcharges) - Exciseduty, CfMt)
        End If
        If txtwithrawalcharges = "" Then txtwithrawalcharges = 0
        If CDbl(txtwithrawableamount) < 0 Then
            txtwithrawableamount = "0.00"
        End If
        'lblFrequency = Format(FrequencyCharge, CfMt)
        'XXXXXXXXXXXXXXXX Get_Lumpsum_Charge XXXXXXXXXXXXX
        ComAmount = Get_Charge_Amount("002", ErrorMessage, lblbookbalance)
        If ComAmount > 0 Then
            lblstandingorderbalance = Format(ComAmount, CfMt)
            txtwithrawableamount = Format(CDbl(txtwithrawableamount) - ComAmount, CfMt)
        End If
        If CDbl(txtwithrawableamount) < 0 Then
            txtwithrawableamount = "0.00"
        End If
       
        
        txtwithrawalcharges = Format(CDbl(txtwithrawalcharges) - FrequencyCharge, CfMt)
        FrequencyCharge = 0
        WithdrawableAmt = 0
        WithdrawalCharge1 = 0
'        If Trim(Customer.AccTypeName) = "SACCO B SAVINGS ACCOUNT" Then
'            lblFrequency = "0.00"
'            WithdrawalCharge = Get_Charge_Amount("029", ErrorMessage, Customer.WithdrawableAmount)
'            WithdrawableAmt = lblavail - WithdrawalCharge
'            txtwithrawableamount = Format(WithdrawableAmt, CfMt)
'        End If
        If lvememtrans.ListItems.Count > 0 Then
            lvememtrans.SetFocus
            Set lvememtrans.SelectedItem = li
            lvememtrans.SelectedItem.EnsureVisible
        End If
    End If
    
    Editing = False
    Exit Sub
ErrHandler:
    MsgBox Err.Description, vbInformation, Me.Caption
    Editing = False
End Sub
Private Sub txtAccNo_Change()
    On Error GoTo SysError
    Dim rstype As New Recordset
    accno = txtAccno
    lbladvancebalance = "0.00"
    lblLoabBalance = "0.00"
    lblstandingorderbalance = "0.00"
    lblFrequency = "0.00"
    txtwithrawalcharges = "0.00"
    txtothercharges = "0.00"
    gl_lastmove = Now()
    If Not Editing Then
        If Len(accno) > 6 Then
        checkglacc accno
        If isgl = 1 Then
            MsgBox "This is a general ledger account", vbInformation
            Exit Sub
        Else
            Image1.Picture = Nothing
            Image2.Picture = Nothing
            
            Get_Member_Transactions AccountNo, accno
        End If
            gl_lastmove = Now()
'//check if the accno has been issued with a card
            Dim is1 As Integer
            sql = "select issued from cub where accno='" & txtAccno & "'"
            Set rs = oEasySacco.GetRecordset(sql)
            If Not rs.EOF Then
            is1 = IIf(IsNull(rs.Fields(0)), 0, rs.Fields(0))
            Else
            is1 = 0
            End If
            If is1 = 1 Then
            chkissued = vbChecked
            Else
            chkissued = vbUnchecked
            End If
        End If
    End If
    gl_lastmove = Now()
     Editing = False
    Exit Sub
SysError:
    MsgBox Err.Description, vbInformation, Me.Caption
    Editing = False
End Sub

Private Sub Update_Balances(accno As String)
    On Error GoTo SysError
    Dim ad As String
    Dim cnn As ADODB.connection, balance As Double, rsUpdate As New ADODB.Recordset
    Dim Updated As Boolean, j As Long, CustBalID As String
    Dim LastTransDate As Date, lastWithDate As Date
    Set cnn = New ADODB.connection
    With cnn
        If .State = adStateClosed Then
            .Open "KONOIN"
        End If
        accno = txtAccno
        balance = 0
        j = 0
        CustBalID = ""
        Updated = False
        Set rsUpdate = cnn.Execute("Select * From CUSTOMERBALANCE where AccNo='" & accno & _
        "' order by TransDate,CustomerBalanceID")
        With rsUpdate
            While Not .EOF
                j = j + 1
                CustBalID = !customerbalanceid
                LastTransDate = !transdate
                If !TransDescription = "Cash Withdrawal" Then
                    lastWithDate = !transdate
                End If
                If !transdate >= CDate("10/23/2008") Then
                    If Not Updated Then
                        balance = IIf(IsNull(!AvailableBalance), 0, !AvailableBalance)
                        Updated = True
                    Else
                        Select Case UCase(!TransType)
                            Case "CR"
                            If Left(!TransDescription, 10) <> "Cheque Dep" Then
                                balance = balance + IIf(IsNull(!Amount), 0, !Amount)
                            End If
                            cnn.Execute ("Exec Update_CustBal_Balance " & balance & ",'" _
                            & accno & "','" & CustBalID & "'")
                            Case "DR"
                            If !TransDescription <> "Bounced Cheque" Then
                                balance = balance - IIf(IsNull(!Amount), 0, !Amount)
                            End If
                            cnn.Execute ("Exec Update_CustBal_Balance " & balance & ",'" _
                            & accno & "','" & CustBalID & "'")
                        End Select
                    End If
                Else
                    balance = IIf(IsNull(!AvailableBalance), 0, !AvailableBalance)
                    Updated = True
                End If
                .MoveNext
            Wend
        End With
        .Execute ("Exec Upadte_AccBalance '" & accno & "'," & balance)
        lblavail = Format(balance, CfMt)

    If txtwithrawableamount = "" Then txtwithrawableamount = 0
    If txtwithrawalcharges = "" Then txtwithrawalcharges = 0
        lblbookbalance = Format(balance, CfMt) - 1020
        txtwithrawableamount = lblbookbalance - txtwithrawalcharges - 2
        If txtwithrawableamount <= 0 Then
        txtwithrawableamount = 0
        Else
        If txtwithrawableamount = "" Then txtwithrawableamount = 0
        txtwithrawableamount = txtwithrawableamount
        End If
        ad = Left(accno, Len(accno) - 7)
        If ad = "201" Then
        txtwithrawableamount = 0
       End If
    End With
    Exit Sub
SysError:
2    MsgBox Err.Description, vbInformation, Me.Caption
End Sub
Private Sub Txtaccno_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtAccNo_KeyPress(KeyAscii As Integer)
    KeyAscii = To_Upper_Case(KeyAscii)
End Sub

Private Sub txtIDNo_Change()
    On Error GoTo SysError
    gl_lastmove = Now()
    lbladvancebalance = "0.00"
    lblLoabBalance = "0.00"
    lblstandingorderbalance = "0.00"
    lblFrequency = "0.00"
    txtwithrawalcharges = "0.00"
    txtothercharges = "0.00"
    If Not Editing Then
        Get_Member_Transactions IDNO, txtidno
    End If
    gl_lastmove = Now()
     Editing = False
    Exit Sub
SysError:
    MsgBox Err.Description, vbInformation, Me.Caption
End Sub

Private Sub txtIDNo_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtIDNo_KeyPress(KeyAscii As Integer)
    KeyAscii = To_Upper_Case(KeyAscii)
End Sub

Private Sub txtpayno_Change() '//represent tea growers number
    On Error GoTo SysError
    lbladvancebalance = "0.00"
    lblLoabBalance = "0.00"
    lblstandingorderbalance = "0.00"
    lblFrequency = "0.00"
    txtwithrawalcharges = "0.00"
    txtothercharges = "0.00"
    gl_lastmove = Now()
    If Not Editing Then
        Get_Member_Transactions growerno, txtpayno
    End If
    gl_lastmove = Now()
     Editing = False
    Exit Sub
SysError:
    MsgBox Err.Description, vbInformation, Me.Caption
End Sub

Private Sub txtpayno_KeyPress(KeyAscii As Integer)
    KeyAscii = To_Upper_Case(KeyAscii)
End Sub

Private Sub txttgno_Change()
    On Error GoTo SysError
    lbladvancebalance = "0.00"
    lblLoabBalance = "0.00"
    lblstandingorderbalance = "0.00"
    lblFrequency = "0.00"
    txtwithrawalcharges = "0.00"
    txtothercharges = "0.00"
    gl_lastmove = Now()
    If Not Editing Then
        Get_Member_Transactions memberno, TXTTGNO
    End If
    gl_lastmove = Now()
    Editing = False
    Exit Sub
SysError:
    MsgBox Err.Description, vbInformation, Me.Caption
End Sub

Private Sub txttgno_KeyPress(KeyAscii As Integer)
    KeyAscii = To_Upper_Case(KeyAscii)
End Sub

Private Sub UserControl_show()
    'Lbldate = Date
    On Error Resume Next
    With lvememtrans
        .ColumnHeaders.Clear
        .ListItems.Clear
    End With
    With lvememtrans
        .ColumnHeaders.Clear
        .ColumnHeaders.Add 1, , "Transaction Date"
         .ColumnHeaders.Add 2, , "Voucher Number"
        .ColumnHeaders.Add 3, , "Dr"
        .ColumnHeaders.Add 4, , "Cr"
        .ColumnHeaders.Add 5, , "Available Balance"
        .ColumnHeaders.Add 5, , "Actual Balance"
        .ColumnHeaders.Add 6, , "Commissions"
       .ColumnHeaders.Add 7, , "Description"
    End With
    
   lblname = ""
     lblaccname = ""
     
     lblavail = ""
 
     lblname = ""
     TXTTGNO = ""
     
  
End Sub
'


