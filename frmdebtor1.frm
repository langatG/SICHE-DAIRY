VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmdebtor1 
   Caption         =   "DEBTORS SALES AND REGISTRY"
   ClientHeight    =   8835
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8835
   ScaleWidth      =   11160
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox CheckTranfer 
      Caption         =   "Inter Vehicle Transfer?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   5280
      TabIndex        =   104
      Top             =   1080
      Width           =   2775
   End
   Begin VB.CommandButton cmdNewVehc 
      Caption         =   "New Vehicle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8400
      TabIndex        =   100
      Top             =   8160
      Width           =   1455
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10200
      TabIndex        =   26
      Top             =   8160
      Width           =   855
   End
   Begin MSComCtl2.DTPicker txtdateenterered 
      Height          =   375
      Left            =   9240
      TabIndex        =   20
      Top             =   840
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   111083521
      CurrentDate     =   38814
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8775
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   15478
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      BackColor       =   32768
      ForeColor       =   192
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Britannic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "DEBTORS REGISTRATION"
      TabPicture(0)   =   "frmdebtor1.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(1)=   "cmdNew"
      Tab(0).Control(2)=   "cmdEdit"
      Tab(0).Control(3)=   "cmdSave"
      Tab(0).Control(4)=   "cmdnewvehicle"
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "POINT OF SALES "
      TabPicture(1)   =   "frmdebtor1.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label7"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label3"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label4"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label5"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label9"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label6"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "txtdcode"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label11"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label12"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Label13"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "lbldrstock"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "lblcrvehicle"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Label16"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Label10"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "lbltotal"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "lbltotalkg"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Label34"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "Label35"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "ListView1"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "cmdnewsearch"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "Picture1"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "txtRefNo"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "cboVehicle"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "cboNames"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "txtDispatch"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "txtamountp"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "txtIntake"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "chkpai"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "txtamount"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "cmdnew3"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "cmdsave3"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "cmdstatement"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).Control(34)=   "txtremarks"
      Tab(1).Control(34).Enabled=   0   'False
      Tab(1).Control(35)=   "fra1"
      Tab(1).Control(35).Enabled=   0   'False
      Tab(1).Control(36)=   "chkdelete"
      Tab(1).Control(36).Enabled=   0   'False
      Tab(1).Control(37)=   "Command2"
      Tab(1).Control(37).Enabled=   0   'False
      Tab(1).Control(38)=   "cmdshort"
      Tab(1).Control(38).Enabled=   0   'False
      Tab(1).Control(39)=   "chkoutletre"
      Tab(1).Control(39).Enabled=   0   'False
      Tab(1).Control(40)=   "cmddedproce"
      Tab(1).Control(40).Enabled=   0   'False
      Tab(1).ControlCount=   41
      TabCaption(2)   =   "ASSIGN VEHICLE"
      TabPicture(2)   =   "frmdebtor1.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdreport"
      Tab(2).Control(1)=   "cmdprce"
      Tab(2).Control(2)=   "cboGari2"
      Tab(2).Control(3)=   "cboGari"
      Tab(2).Control(4)=   "cmdAssign"
      Tab(2).Control(5)=   "cmdActive"
      Tab(2).Control(6)=   "ListViewG"
      Tab(2).Control(7)=   "prgStatus"
      Tab(2).Control(8)=   "Label32"
      Tab(2).Control(9)=   "Label31"
      Tab(2).Control(10)=   "Label30"
      Tab(2).Control(11)=   "Label28"
      Tab(2).Control(12)=   "Label29"
      Tab(2).ControlCount=   13
      Begin VB.CommandButton cmddedproce 
         Caption         =   "Vehicle Customers Balance"
         Height          =   495
         Left            =   1560
         TabIndex        =   103
         Top             =   8040
         Width           =   2295
      End
      Begin VB.CheckBox chkoutletre 
         Caption         =   "Disputch for Outlet.?"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   5160
         TabIndex        =   101
         Top             =   720
         Width           =   2775
      End
      Begin VB.CommandButton cmdreport 
         Caption         =   "Print Report"
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
         Height          =   615
         Left            =   -65760
         TabIndex        =   93
         Top             =   2880
         Width           =   1455
      End
      Begin VB.CommandButton cmdshort 
         Caption         =   "Expenses form"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1800
         TabIndex        =   92
         Top             =   4560
         Width           =   1815
      End
      Begin VB.CommandButton cmdprce 
         Caption         =   "Vehicle and Expenses Process"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -68160
         TabIndex        =   91
         Top             =   2880
         Width           =   2175
      End
      Begin VB.ComboBox cboGari2 
         Height          =   315
         Left            =   -72600
         TabIndex        =   86
         Top             =   1920
         Width           =   2775
      End
      Begin VB.ComboBox cboGari 
         Height          =   315
         Left            =   -72600
         TabIndex        =   85
         Top             =   1080
         Width           =   2775
      End
      Begin VB.CommandButton cmdAssign 
         Caption         =   "Assign"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -72720
         TabIndex        =   83
         Top             =   3360
         Width           =   1215
      End
      Begin VB.CommandButton cmdActive 
         Caption         =   "In Activate"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -71160
         TabIndex        =   82
         Top             =   3360
         Width           =   1575
      End
      Begin VB.CommandButton cmdnewvehicle 
         Caption         =   "Debtors List"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -70080
         TabIndex        =   80
         Top             =   8160
         Width           =   1935
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Kiarie Rports"
         Height          =   495
         Left            =   6240
         TabIndex        =   79
         Top             =   8100
         Width           =   1935
      End
      Begin VB.CheckBox chkdelete 
         Caption         =   "Remove intake kgs"
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
         Height          =   225
         Left            =   7560
         TabIndex        =   76
         Top             =   2700
         Width           =   2655
      End
      Begin VB.Frame fra1 
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   6240
         TabIndex        =   65
         Top             =   3540
         Width           =   4695
         Begin VB.PictureBox Picture2 
            Height          =   255
            Left            =   1320
            Picture         =   "frmdebtor1.frx":0054
            ScaleHeight     =   195
            ScaleWidth      =   195
            TabIndex        =   69
            Top             =   240
            Width           =   255
         End
         Begin VB.PictureBox Picture3 
            Height          =   255
            Left            =   1320
            Picture         =   "frmdebtor1.frx":091E
            ScaleHeight     =   195
            ScaleWidth      =   195
            TabIndex        =   68
            Top             =   720
            Width           =   255
         End
         Begin VB.TextBox txtdracc 
            Height          =   375
            Left            =   1680
            TabIndex        =   67
            Top             =   120
            Width           =   2895
         End
         Begin VB.TextBox txtcracc 
            Height          =   375
            Left            =   1680
            TabIndex        =   66
            Top             =   600
            Width           =   2895
         End
         Begin VB.Label lbldracc 
            BackColor       =   &H8000000E&
            Height          =   375
            Left            =   120
            TabIndex        =   71
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label lblcracc 
            BackColor       =   &H8000000E&
            Height          =   375
            Left            =   120
            TabIndex        =   70
            Top             =   600
            Width           =   1095
         End
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -71520
         TabIndex        =   59
         Top             =   8160
         Width           =   735
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Edit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -73080
         TabIndex        =   58
         Top             =   8160
         Width           =   735
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "New"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74760
         TabIndex        =   57
         Top             =   8160
         Width           =   735
      End
      Begin VB.Frame Frame1 
         Caption         =   "Debtors Details"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7455
         Left            =   -74880
         TabIndex        =   35
         Top             =   660
         Width           =   10935
         Begin VB.ComboBox txtVehicle 
            Height          =   315
            Left            =   5040
            TabIndex        =   81
            Top             =   2160
            Width           =   2895
         End
         Begin VB.TextBox txtsta 
            Height          =   405
            Left            =   9480
            TabIndex        =   77
            Top             =   2040
            Width           =   615
         End
         Begin VB.CheckBox chkActive 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Active"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   8760
            TabIndex        =   73
            Top             =   1440
            Value           =   1  'Checked
            Width           =   975
         End
         Begin VB.CommandButton cmdSearch 
            Caption         =   "..."
            Height          =   285
            Left            =   1560
            TabIndex        =   72
            Top             =   3240
            Width           =   300
         End
         Begin VB.TextBox txtDrAccNo 
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
            Height          =   315
            Left            =   1860
            TabIndex        =   64
            Top             =   3240
            Width           =   1080
         End
         Begin VB.TextBox lblDrAccName 
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
            Height          =   300
            Left            =   3135
            TabIndex        =   63
            Top             =   3240
            Width           =   3225
         End
         Begin VB.TextBox txtCrAccNo 
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
            Height          =   315
            Left            =   1875
            TabIndex        =   62
            Top             =   3600
            Width           =   1080
         End
         Begin VB.CommandButton Command1 
            Caption         =   "..."
            Height          =   285
            Left            =   1560
            TabIndex        =   61
            Top             =   3600
            Width           =   315
         End
         Begin VB.TextBox txtCrAccName 
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
            Height          =   300
            Left            =   3150
            TabIndex        =   60
            Top             =   3600
            Width           =   3225
         End
         Begin VB.TextBox txtEmail 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   960
            TabIndex        =   56
            Top             =   1440
            Width           =   2175
         End
         Begin VB.TextBox txtPrice 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   960
            TabIndex        =   55
            Text            =   "0.00"
            Top             =   2160
            Width           =   1095
         End
         Begin VB.TextBox txtPAddress 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5040
            TabIndex        =   54
            Top             =   1320
            Width           =   2895
         End
         Begin VB.TextBox txtTown 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5040
            TabIndex        =   53
            Top             =   1680
            Width           =   2895
         End
         Begin VB.TextBox txtPhone 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5040
            TabIndex        =   52
            Top             =   840
            Width           =   2895
         End
         Begin VB.TextBox txtId 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   960
            TabIndex        =   51
            Top             =   840
            Width           =   2295
         End
         Begin VB.TextBox txtNames 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4320
            TabIndex        =   49
            Top             =   360
            Width           =   3615
         End
         Begin VB.PictureBox Picture5 
            Height          =   255
            Left            =   3000
            Picture         =   "frmdebtor1.frx":11E8
            ScaleHeight     =   195
            ScaleWidth      =   195
            TabIndex        =   48
            Top             =   360
            Width           =   255
         End
         Begin VB.TextBox txtTCode 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1680
            TabIndex        =   47
            Top             =   360
            Width           =   1335
         End
         Begin VB.Frame Frame2 
            Caption         =   "GL Ledgers"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1575
            Left            =   120
            TabIndex        =   44
            Top             =   2760
            Width           =   10335
            Begin VB.Label Label26 
               Caption         =   "Cr"
               BeginProperty Font 
                  Name            =   "Century"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   46
               Top             =   840
               Width           =   1215
            End
            Begin VB.Label Label25 
               Caption         =   "Dr"
               BeginProperty Font 
                  Name            =   "Century"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   45
               Top             =   480
               Width           =   1335
            End
         End
         Begin MSComctlLib.ListView ListView8 
            Height          =   2895
            Left            =   120
            TabIndex        =   75
            Top             =   4440
            Width           =   10695
            _ExtentX        =   18865
            _ExtentY        =   5106
            View            =   3
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            AllowReorder    =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            HotTracking     =   -1  'True
            HoverSelection  =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   65280
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
         Begin VB.Label Label27 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Multiple:"
            BeginProperty Font 
               Name            =   "Century"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   8400
            TabIndex        =   78
            Top             =   2160
            Width           =   945
         End
         Begin VB.Label Label24 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Vehicle No:"
            BeginProperty Font 
               Name            =   "Century"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   3720
            TabIndex        =   74
            Top             =   2160
            Width           =   1185
         End
         Begin VB.Label Label23 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Date:"
            BeginProperty Font 
               Name            =   "Century"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   8280
            TabIndex        =   50
            Top             =   600
            Width           =   555
         End
         Begin VB.Label Label22 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Town:"
            BeginProperty Font 
               Name            =   "Century"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   4080
            TabIndex        =   43
            Top             =   1680
            Width           =   630
         End
         Begin VB.Label Label21 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Price:"
            BeginProperty Font 
               Name            =   "Century"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   240
            TabIndex        =   42
            Top             =   2160
            Width           =   615
         End
         Begin VB.Label Label20 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Email:"
            BeginProperty Font 
               Name            =   "Century"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   240
            TabIndex        =   41
            Top             =   1440
            Width           =   690
         End
         Begin VB.Label Label19 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Phone No:"
            BeginProperty Font 
               Name            =   "Century"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   3480
            TabIndex        =   40
            Top             =   840
            Width           =   1065
         End
         Begin VB.Label Label18 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Name:"
            BeginProperty Font 
               Name            =   "Century"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   3480
            TabIndex        =   39
            Top             =   360
            Width           =   675
         End
         Begin VB.Label Label17 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Postal Address:"
            BeginProperty Font 
               Name            =   "Century"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   3360
            TabIndex        =   38
            Top             =   1320
            Width           =   1605
         End
         Begin VB.Label Label15 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "ID No:"
            BeginProperty Font 
               Name            =   "Century"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   120
            TabIndex        =   37
            Top             =   840
            Width           =   690
         End
         Begin VB.Label Label14 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Debtors Code:"
            BeginProperty Font 
               Name            =   "Century"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   120
            TabIndex        =   36
            Top             =   360
            Width           =   1425
         End
      End
      Begin VB.TextBox txtremarks 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1800
         TabIndex        =   34
         Top             =   3660
         Width           =   2655
      End
      Begin VB.CommandButton cmdstatement 
         Caption         =   "Debtors Statements"
         Height          =   495
         Left            =   3960
         TabIndex        =   32
         Top             =   8100
         Width           =   1935
      End
      Begin VB.CommandButton cmdsave3 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4080
         TabIndex        =   31
         Top             =   4560
         Width           =   975
      End
      Begin VB.CommandButton cmdnew3 
         Caption         =   "New"
         Height          =   495
         Left            =   360
         TabIndex        =   30
         Top             =   8100
         Width           =   1095
      End
      Begin VB.TextBox txtamount 
         Alignment       =   1  'Right Justify
         Height          =   495
         Left            =   1800
         TabIndex        =   28
         Top             =   3060
         Width           =   1095
      End
      Begin VB.CheckBox chkpai 
         Caption         =   "Make Payments without milk"
         BeginProperty Font 
            Name            =   "Century"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   255
         Left            =   4080
         TabIndex        =   27
         Top             =   3180
         Width           =   3855
      End
      Begin VB.TextBox txtIntake 
         Height          =   375
         Left            =   9000
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   3060
         Width           =   1335
      End
      Begin VB.TextBox txtamountp 
         Height          =   375
         Left            =   5400
         TabIndex        =   11
         Top             =   2580
         Width           =   1215
      End
      Begin VB.TextBox txtDispatch 
         Height          =   375
         Left            =   1800
         TabIndex        =   10
         Top             =   2580
         Width           =   1335
      End
      Begin VB.ComboBox cboNames 
         Height          =   315
         Left            =   1800
         TabIndex        =   7
         Top             =   1980
         Width           =   2055
      End
      Begin VB.ComboBox cboVehicle 
         Height          =   315
         Left            =   1800
         TabIndex        =   6
         Top             =   1500
         Width           =   2055
      End
      Begin VB.TextBox txtRefNo 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   1800
         TabIndex        =   3
         Top             =   900
         Width           =   2055
      End
      Begin VB.PictureBox Picture1 
         Height          =   285
         Left            =   3840
         Picture         =   "frmdebtor1.frx":14AA
         ScaleHeight     =   225
         ScaleWidth      =   195
         TabIndex        =   2
         Top             =   900
         Width           =   255
      End
      Begin VB.CommandButton cmdnewsearch 
         Caption         =   "New "
         Height          =   285
         Left            =   4080
         TabIndex        =   1
         Top             =   900
         Width           =   615
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2895
         Left            =   240
         TabIndex        =   5
         Top             =   5100
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   5106
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   65280
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListViewG 
         Height          =   3495
         Left            =   -74880
         TabIndex        =   88
         Top             =   4560
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   6165
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   65280
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ProgressBar prgStatus 
         Height          =   255
         Left            =   -68160
         TabIndex        =   102
         Top             =   2640
         Visible         =   0   'False
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label Label35 
         Caption         =   "Vehicle kgs"
         BeginProperty Font 
            Name            =   "Century"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9360
         TabIndex        =   99
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label34 
         Caption         =   "Vehicle Amount"
         BeginProperty Font 
            Name            =   "Century"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8400
         TabIndex        =   98
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label lbltotalkg 
         BackColor       =   &H00FFFFC0&
         Caption         =   "k"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   9480
         TabIndex        =   96
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label lbltotal 
         BackColor       =   &H00FFFFC0&
         Caption         =   "k"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   8400
         TabIndex        =   95
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label Label32 
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Left            =   -66840
         TabIndex        =   94
         Top             =   2040
         Width           =   2535
      End
      Begin VB.Label Label31 
         BackColor       =   &H00FFC0FF&
         Height          =   375
         Left            =   -68640
         TabIndex        =   90
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label30 
         BackColor       =   &H00FFC0FF&
         Height          =   375
         Left            =   -68640
         TabIndex        =   89
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label28 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Expense Ledger:"
         BeginProperty Font 
            Name            =   "Century"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   -74640
         TabIndex        =   87
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label29 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Vehicle No:"
         BeginProperty Font 
            Name            =   "Century"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   -74640
         TabIndex        =   84
         Top             =   1080
         Width           =   1185
      End
      Begin VB.Label Label10 
         Caption         =   "Remarks."
         BeginProperty Font 
            Name            =   "Century"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   33
         Top             =   3660
         Width           =   1455
      End
      Begin VB.Label Label16 
         Caption         =   "Paid Amount:"
         BeginProperty Font 
            Name            =   "Century"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   3180
         Width           =   1335
      End
      Begin VB.Label lblcrvehicle 
         BackColor       =   &H00FFFFC0&
         Height          =   255
         Left            =   5400
         TabIndex        =   25
         Top             =   1860
         Width           =   1095
      End
      Begin VB.Label lbldrstock 
         BackColor       =   &H00FFFFC0&
         Height          =   255
         Left            =   5400
         TabIndex        =   24
         Top             =   1500
         Width           =   1095
      End
      Begin VB.Label Label13 
         Caption         =   "Cr"
         BeginProperty Font 
            Name            =   "Century"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5040
         TabIndex        =   23
         Top             =   1860
         Width           =   1215
      End
      Begin VB.Label Label12 
         Caption         =   "Dr"
         BeginProperty Font 
            Name            =   "Century"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5040
         TabIndex        =   22
         Top             =   1500
         Width           =   1335
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Date."
         BeginProperty Font 
            Name            =   "Century"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   8160
         TabIndex        =   21
         Top             =   1140
         Width           =   555
      End
      Begin VB.Label txtdcode 
         BackColor       =   &H00FFFFC0&
         Height          =   255
         Left            =   3840
         TabIndex        =   19
         Top             =   1980
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Dr"
         BeginProperty Font 
            Name            =   "Century"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5400
         TabIndex        =   17
         Top             =   3660
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "Cr"
         BeginProperty Font 
            Name            =   "Century"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5400
         TabIndex        =   16
         Top             =   4260
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Intake:"
         BeginProperty Font 
            Name            =   "Century"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8040
         TabIndex        =   14
         Top             =   3180
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Amount Payable."
         BeginProperty Font 
            Name            =   "Century"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3360
         TabIndex        =   13
         Top             =   2580
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Dispatch."
         BeginProperty Font 
            Name            =   "Century"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   2580
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Debtor Name."
         BeginProperty Font 
            Name            =   "Century"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   1980
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Vehicle No."
         BeginProperty Font 
            Name            =   "Century"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   1380
         Width           =   1455
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Reference No."
         BeginProperty Font 
            Name            =   "Century"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   120
         TabIndex        =   4
         Top             =   900
         Width           =   1425
      End
   End
   Begin VB.Label Label33 
      Caption         =   "Label33"
      Height          =   495
      Left            =   5040
      TabIndex        =   97
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "Label8"
      Height          =   495
      Left            =   5520
      TabIndex        =   18
      Top             =   4440
      Width           =   1215
   End
End
Attribute VB_Name = "frmdebtor1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Price As Currency
Dim newa As Integer

Private Sub cboGari_Click()
'NAMES1
NAMES3
Set rst = oSaccoMaster.GetRecordset("select distinct(AccCr) from d_Debtors where Locations ='" & cboGari & "'")
If Not rst.EOF Then
 Label31 = rst.Fields("AccCr")
End If
cboGari2.SetFocus
End Sub
Private Sub NAMES1()
'Private Sub SSTab1_DblClick()
    cboGari.Clear
    Set rst = New Recordset
    Dim cn As Connection
    Set cn = New ADODB.Connection
    Provider = "MAZIWA"
    'Provider = cnn
    cn.Open Provider, "atm", "atm"
    Set rst = New Recordset
    sql = "Select distinct(Locations) from   d_Debtors order by Locations"
    'Select distinct(Locations) from   d_Debtors
    rst.Open sql, cn, adOpenKeyset, adLockOptimistic
    While Not rst.EOF
    cboGari.AddItem rst.Fields(0)
    rst.MoveNext
    Wend
End Sub
Private Sub NAMES3()
'Private Sub SSTab1_DblClick()
    cboGari2.Clear
    Set rst = New Recordset
    Dim cn As Connection
    Set cn = New ADODB.Connection
    Provider = "MAZIWA"
    'Provider = cnn
    cn.Open Provider, "atm", "atm"
    Set rst = New Recordset
    sql = "Select GlAccName, AccNo from  GLSETUP WHERE  (GlAccName LIKE 'K%') order by GlAccName"
    'Select distinct(Locations) from   d_Debtors
    rst.Open sql, cn, adOpenKeyset, adLockOptimistic
    While Not rst.EOF
    cboGari2.AddItem rst.Fields(0)
    rst.MoveNext
    Wend
End Sub

Private Sub cboGari2_Click()
Set rst = oSaccoMaster.GetRecordset("select distinct(AccNo) from GLSETUP where GlAccName ='" & cboGari2 & "'")
If Not rst.EOF Then
 Label30 = rst.Fields("AccNo")
End If
End Sub

Private Sub cboNames_Click()
'NAMES
If cboVehicle = "" Then
    MsgBox "Please select the Vehicle Number."
        cboVehicle.SetFocus
    Exit Sub
End If
loadb
'txtDispatch.SetFocus
End Sub
Private Sub loadb()
   Set rst = New ADODB.Recordset
    Set rs = New ADODB.Recordset
   If chkoutletre = vbChecked Then
   ''outlet
      sql = "Select Branch from d_Outlet where p_name='" & cboNames & "'"
      Set rs = oSaccoMaster.GetRecordset(sql)
      If Not rs.EOF Then
        sql = "Select Dr as DCode,Dr as AccDr , Cr as AccCr from d_Outletbranch where BName1='" & rs!Branch & "'"
      End If
    Else
      sql = "select DCode,AccDr, AccCr from d_Debtors where DName ='" & cboNames & "' and Locations ='" & cboVehicle & "'"
    End If

    'If rs.EOF Then
    Set rst = oSaccoMaster.GetRecordset(sql)
    If Not rst.EOF Then
        txtdcode = rst.Fields("DCode")
        lbldrstock = rst.Fields("AccDr")
        lblcrvehicle = rst.Fields("AccCr")
    End If
Debtorsgl
End Sub

Private Sub cboVehicle_Click()
NAMES
loadBranchesTypes
lbltotal.Visible = True
lbltotalkg.Visible = True
Label34.Visible = True
Label35.Visible = True
'cboNames_Click
'SSTab1_DblClick
    'cboVehicle.Clear
'cboNames.SetFocus
chkoutletre.Visible = True
End Sub

Private Sub chkdelete_Click()
If chkdelete = 1 Then
Else
chkdelete.value = 0
End If
End Sub

Private Sub chkoutletre_Click()
If chkoutletre = vbChecked Then
    chkoutletre.value = 1
    outletre
Else
    chkoutletre.value = 0
    NAMES
End If
End Sub
Private Sub outletre()
    cboNames.Clear
    Set rst = New Recordset
    Dim cn As Connection
    Set cn = New ADODB.Connection
    Provider = "MAZIWA"
    cn.Open Provider, "atm", "atm"
    Set rst = New Recordset
    sql = "Select distinct(p_name) from d_Outlet where p_name like '%Milk%' order by p_name"
    rst.Open sql, cn, adOpenKeyset, adLockOptimistic
    While Not rst.EOF
    cboNames.AddItem rst.Fields(0)
    rst.MoveNext
    Wend
End Sub

Private Sub chkpai_Click()
If chkpai = 1 Then
chkpai.value = 1
Else

chkpai.value = 0
End If

End Sub
Private Sub cmdActive_Click()
On Error GoTo ErrorHandler

If cboGari = "" Then
 MsgBox "Please enter the Vehicle code", vbInformation
 cboGari.SetFocus
Exit Sub
End If
If cboGari2 = "" Then
 MsgBox "Please enter the Expense Ledger", vbInformation
 cboGari2.SetFocus
Exit Sub
End If

''Set cn = New ADODB.Connection
''sql = "SELECT  startdate FROM d_Transport WHERE  (Sno = " & txtSNo & ") AND (Trans_Code = '" & txtTCode & "')"
''Set rs = oSaccoMaster.GetRecordset(sql)
''If Not rs.EOF Then
''If rs.Fields("StartDate") = DTPDRemoved Then
''oSaccoMaster.ExecuteThis ("SET dateformat DMY delete FROM d_Transport where SNO= " & txtSNo & " and Trans_Code= '" & txtTCode & "' AND StartDate= '" & DTPDRemoved & "'")
''MsgBox "Record successively updated "
''End If
''End If
Set rs = New ADODB.Recordset
Set cn = New ADODB.Connection
sql = "set dateformat dmy select Active from  d_AssignmentVehicle where ExpeLedger='" & Label30 & "' and ExpenseAcc='" & cboGari2 & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
If rs.Fields(0) = "True" Then
'Set cn = New ADODB.Connection
 sql = ""
 sql = "SET dateformat DMY Update  d_AssignmentVehicle SET Active= '0', Date='" & txtdateenterered & "'where Vehicle ='" & cboGari & "' and ExpenseAcc='" & cboGari2 & "'"
oSaccoMaster.ExecuteThis (sql)
Else
 sql = ""
 sql = "SET dateformat DMY Update  d_AssignmentVehicle SET Active= '1', Date='" & txtdateenterered & "'where Vehicle ='" & cboGari & "' and ExpenseAcc='" & cboGari2 & "'"
oSaccoMaster.ExecuteThis (sql)
End If
'loadTransportAssignments
loadAssignments
If cmdActive.Caption = "Activate" Then
 cmdActive.Caption = "In Activate"
Else
 cmdActive.Caption = "Activate"
End If

MsgBox "Records successively updated."
cmdActive.Enabled = False
Label30 = ""
Label31 = ""
cboGari = ""
cboGari2 = ""
Exit Sub
ErrorHandler:
MsgBox err.description
End Sub
Private Sub cmdAssign_Click()
On Error GoTo ErrorHandler

If cboGari = "" Then
 MsgBox "Please enter the Vehicle code", vbInformation
 cboGari.SetFocus
Exit Sub
End If
If cboGari2 = "" Then
 MsgBox "Please enter the Expense Ledger", vbInformation
 cboGari2.SetFocus
Exit Sub
End If

Set rs = New ADODB.Recordset
Set cn = New ADODB.Connection
sql = "set dateformat dmy select AccnoV, Vehicle, ExpenseAcc, ExpeLedger, Active,Date, UserID from  d_AssignmentVehicle where active=1 and ExpeLedger='" & Label30 & "' and ExpenseAcc='" & cboGari2 & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
  MsgBox "This Ledger Number has been assigned to Vehicle code : " & rs.Fields("Vehicle") & ""
  Exit Sub
Else
  sql = ""
  sql = "set dateformat dmy insert into  d_AssignmentVehicle( AccnoV, Vehicle, ExpenseAcc, ExpeLedger, Active, Date, UserID)"
  sql = sql & "  values('" & Label31 & "','" & cboGari & "','" & cboGari2 & "','" & Label30 & "','1','" & txtdateenterered & "','" & User & "')"
  oSaccoMaster.ExecuteThis (sql)
End If
loadAssignments
MsgBox "Records successively updated."
Label30 = ""
Label31 = ""
cboGari = ""
cboGari2 = ""
Exit Sub
ErrorHandler:
MsgBox err.description
End Sub
Public Sub loadAssignments()
    With ListViewG
        .ListItems.Clear
        .ColumnHeaders.Clear
    End With
    Set rs2 = CreateObject("adodb.recordset")
    sql = ""
    sql = "Select  AccnoV, Vehicle, ExpenseAcc, ExpeLedger,Date,Active from d_AssignmentVehicle "
   'sql = "set dateformat dmy SELECT d.RefNo,m.DName, d.DispDate, d.DispQnty,d.Amount,d.PaidAmount FROM d_MilkControl AS d INNER JOIN d_Debtors AS m ON d.DCode = m.DCode WHERE     (DispDate = '" & txtdateenterered & "') and vehicleno='" & cboVehicle & "' order by RefNo desc"
    Set rs = oSaccoMaster.GetRecordset(sql)
    
    Set rs2 = CreateObject("adodb.recordset")
    
    Set clsClass = New cdbase
    
    Provider = clsClass.OpenCon
    
    Set cn = CreateObject("adodb.connection")
    
   cn.Open Provider, "atm", "atm"
    
    rs2.Open sql, cn
    
    With ListViewG
        .ColumnHeaders.Add , , "V.Accno"
        .ColumnHeaders.Add , , "Vehicle"
        '.ColumnHeaders.Add , , "Date"
        .ColumnHeaders.Add , , "ExpenseAcc"
        .ColumnHeaders.Add , , "ExpeLedger"
        .ColumnHeaders.Add , , "Assign Date"
        .ColumnHeaders.Add , , "Active"
      While Not rs2.EOF
        Set li = .ListItems.Add(, , Trim(rs2.Fields("AccnoV")))
            li.ListSubItems.Add , , Trim(rs2.Fields("Vehicle"))
            li.ListSubItems.Add , , Trim(rs2.Fields("ExpenseAcc"))
            li.ListSubItems.Add , , Trim(rs2.Fields("ExpeLedger"))
            li.ListSubItems.Add , , Trim(rs2.Fields("Date"))
            li.ListSubItems.Add , , Trim(rs2.Fields("Active"))
        rs2.MoveNext
      Wend
        
    End With
    
    rs2.Close
    
    Set rs2 = Nothing
    
ListViewG.View = lvwReport
End Sub

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmddedproce_Click()
frmdebtor1milk.Show vbModal
End Sub

Private Sub cmdedit_Click()
newa = 0
txtVehicle.Locked = False
txtEMail.Locked = False
txtId.Locked = False
txtNames.Locked = False
txtPAddress.Locked = False
txtPhone.Locked = False
'txtsubsidy.Locked = False
txtTCode.Locked = False
txtTown.Locked = False
'cboBBranch.Locked = False
'cboBName.Locked = False
'cbolocation.Locked = False
cmdsave.Enabled = True
End Sub
Private Sub cmdNew_Click()
newa = 1
txtVehicle = ""
txtEMail = ""
txtId = ""
txtNames = ""
txtPAddress = ""
txtPhone = ""
txtCrAccName = ""
txtTCode = ""
txtTown = ""
txtDrAccNo = ""
txtCrAccNo = ""
lblDrAccName = ""
'cbobranch.Text = ""
txtprice = "0.00"

txtVehicle.Locked = False
txtEMail.Locked = False
txtId.Locked = False
txtNames.Locked = False
txtPAddress.Locked = False
txtPhone.Locked = False
'txtsubsidy.Locked = False
txtTCode.Locked = False
txtTown.Locked = False
'cboBBranch.Locked = False
'cboBName.Locked = False
'cbolocation.Locked = False
cmdedit.Enabled = False
'cmdSave.Enabled = False
cmdsave.Enabled = True
End Sub
Private Sub cmdnew3_Click()
    txtDispatch.Locked = False
    txtIntake.Locked = True
    txtDispatch = ""
    cboNames = ""
    txtRemarks = "CASH"
    'cboVehicle = ""
    txtamountp = ""
    txtAmount = ""
'    lbltotal.Visible = False
'    lbltotalkg.Visible = False
'    Label34.Visible = False
'    Label35.Visible = False
    chkpai.value = vbUnchecked
    chkdelete.value = vbUnchecked
    cmdnew3.Enabled = False
    cmdsave3.Enabled = True
    cmdnewsearch_Click
End Sub
Public Sub loadReg()
    
    With ListView8
    
        .ListItems.Clear
        
        .ColumnHeaders.Clear
    
    End With

    Set rs2 = CreateObject("adodb.recordset")
    
    sql = "Select DCode,DName,Locations,price,AccDr,AccCr from d_Debtors order by DName"
'    sql = ""
'    sql = "set dateformat dmy SELECT d.RefNo,m.DName, d.DispDate, d.DispQnty,d.Amount,d.PaidAmount FROM d_MilkControl AS d INNER JOIN d_Debtors AS m ON d.DCode = m.DCode WHERE     (DispDate = '" & txtdateenterered & "') and vehicleno='" & cboVehicle & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    
    Set rs2 = CreateObject("adodb.recordset")
    
    Set clsClass = New cdbase
    
    Provider = clsClass.OpenCon
    
    Set cn = CreateObject("adodb.connection")
    
   cn.Open Provider, "atm", "atm"
    
    rs2.Open sql, cn
    
    With ListView8
        
        .ColumnHeaders.Add , , "Debtor Code"
        .ColumnHeaders.Add , , "Name"
        .ColumnHeaders.Add , , "Vehicle"
        .ColumnHeaders.Add , , "Price"
        .ColumnHeaders.Add , , "Dr"
        .ColumnHeaders.Add , , "Cr"
        While Not rs2.EOF
        
            Set li = .ListItems.Add(, , Trim(rs2.Fields("DCode")))
            
            li.ListSubItems.Add , , Trim(rs2.Fields("DName"))
            li.ListSubItems.Add , , Trim(rs2.Fields("Locations"))
            li.ListSubItems.Add , , Trim(rs2.Fields("price"))
            li.ListSubItems.Add , , Trim(rs2.Fields("AccDr"))
            li.ListSubItems.Add , , Trim(rs2.Fields("AccCr"))
            rs2.MoveNext
        
        Wend
        
    End With
    
    rs2.Close
    
    Set rs2 = Nothing
    
ListView8.View = lvwReport

End Sub

Private Sub cmdnewsearch_Click()
Dim rsr As New ADODB.Recordset
Dim rsg As New ADODB.Recordset
Dim I As Object
Dim Mylength As Integer
Dim sa As Double
'//if this record is new then look for receipts no

''//clear all textboxes
'mysql = ""
'mysql = "set dateformat dmy select GenerateReceiptno from param"
sql = ""
sql = "set dateformat dmy select GenerateReceiptno from param"
Set rsg = oSaccoMaster.GetRecordset(sql)
If Not rsg.EOF Then
    ''''check check
    If rsg!GenerateReceiptno = True Then
    
        sql = ""
        sql = "select ReceiptNo from Receiptno where receiptno like 'RF-%' order by Receipthnoid desc"
        
        Set rsr = oSaccoMaster.GetRecordset(sql)
        If Not rsr.EOF Then
        If rsr!ReceiptNo < "RF-30000" Then
         sql = ""
         sql = "delete Receiptno where receiptno like 'RF-%' "
         oSaccoMaster.ExecuteThis (sql)
        End If
        End If
        sql = ""
        sql = "select * from Receiptno where receiptno like 'RF-%' order by Receipthnoid desc"
        Set rsr = oSaccoMaster.GetRecordset(sql)
        If Not rsr.EOF Then
            Mylength = CInt(Mid(rsr!ReceiptNo, 5, 10))
            'Mylength = Mylength + 1
            txtRefNo = Padding(Mylength + 1)
            txtRefNo = "RF-" & txtRefNo
        Else
            Mylength = 1
            txtRefNo = "RF-" & Padding(1)
        End If
Else
    ''//receiptno  will be keyed in
End If
mysql = "set dateformat dmy Insert into Receiptno(Receiptno,Auditdate,auditid)values('" & txtRefNo & "','" & txtdateenterered & "','" & User & "')"
oSaccoMaster.ExecuteThis (mysql)
End If
End Sub

Private Sub cmdNewVehc_Click()
frmVehicleReg.Show vbModal
End Sub

Private Sub cmdnewvehicle_Click()
    reportname = "debtorslist.rpt"
    Show_Sales_Crystal_Report "", reportname, ""
End Sub
Private Sub cmdprce_Click()
On Error GoTo ErrorHandler
Startdate = DateSerial(Year(txtdateenterered), month(txtdateenterered), 1)
Enddate = DateSerial(Year(txtdateenterered), month(txtdateenterered) + 1, 1 - 1)
'If cboGari = "" Then
' MsgBox "Please enter the Vehicle code", vbInformation
' cboGari.SetFocus
'Exit Sub
'End If
Label32.Visible = True
sql = ""
sql = "set dateformat dmy delete from d_Debtorsparchases2 where Date>='" & Startdate & "' and Date<='" & Enddate & "'"
cn.Execute sql
'AccnoV
sql = ""
sql = "set dateformat dmy Select count(distinct(Vehicle)) as j  from   d_AssignmentVehicle where Active = '" & True & "' "
Set rsj = cn.Execute(sql)
C = rsj.Fields(0)
Dim a As Double
a = rsj.Fields(0)
prgStatus.max = 100
prgStatus.Min = 0
I = 0
prgStatus.Visible = True
sql = ""
sql = "set dateformat dmy Select distinct(Vehicle) as j  from   d_AssignmentVehicle where Active = '" & True & "' "
Set rsd = cn.Execute(sql)
Do While Not C <= 0
I = I + 1
prgStatus = Round((I / a) * 100, 0)
If Not rsd.EOF Then
l = rsd.Fields(0)
sql = ""
sql = "set dateformat dmy Select count(DCode) as j  from   d_MilkControl where vehicleno='" & l & "'and DispDate>= '" & Startdate & "' And DispDate<='" & Enddate & "'  "
Set rs = cn.Execute(sql)
j = rs.Fields(0)
  sql = ""
  sql = "set dateformat dmy Select DCode as y ,DispDate  from   d_MilkControl where vehicleno='" & l & "'and DispDate>= '" & Startdate & "' And DispDate<='" & Enddate & "'  "
  Set rsg = cn.Execute(sql)
  Do While Not j <= 0
  Dim M As Integer
  Label32 = "Please wait as we process"

    sql = ""
    sql = "set dateformat dmy Select d.DispDate,m.DName,d.DispQnty,d.Amount,d.PaidAmount from   d_MilkControl AS d INNER JOIN d_Debtors AS m ON d.DCode = m.DCode where d.DCode='" & rsg.Fields(0) & "'and d.DispDate='" & rsg.Fields(1) & "' "
    Set rst = cn.Execute(sql)
    If Not rst.EOF Then
     Dim bal As Double
     ''''''check balance
     sql = ""
     sql = "set dateformat dmy Select isnull(sum(Amount),0) as je,isnull(sum(PaidAmount),0) as ye  from d_MilkControl where DCode='" & rsg.Fields(0) & "' and DispDate>= '" & Startdate & "' And DispDate<='" & rsg.Fields(1) & "'  "
     Set rsbal = cn.Execute(sql)
     bal = rsbal.Fields(1) - rsbal.Fields(0)
     ''''''end
     
    sql = ""
    sql = "set dateformat dmy insert into  d_Debtorsparchases2( Date, Remarks, kgs, Amount, [Paid Amount],Balance, Expenses,Vehicle)"
    sql = sql & "  values('" & rsg.Fields(1) & "','" & rst.Fields(1) & "','" & rst.Fields(2) & "','" & rst.Fields(3) & "','" & rst.Fields(4) & "','" & bal & "','0','" & l & "')"
    cn.Execute sql
    rst.MoveNext
    End If
     
    rsg.MoveNext
  j = j - 1
'  If j = 1 Then
'   MsgBox "1"
'   End If
 Loop
 rsd.MoveNext
 End If
 C = C - 1
Loop
vehiclepro
Label32.Visible = False
 MsgBox "Completed succesfully ", vbInformation
 
 If cmdreport.Caption = "Print Report" Then
 cmdreport.Caption = "Print Report"
Else
 cmdreport.Caption = "Print Report"
End If
cmdreport.Enabled = True
Exit Sub
ErrorHandler:
MsgBox err.description
End Sub
Private Sub vehiclepro()
Startdate = DateSerial(Year(txtdateenterered), month(txtdateenterered), 1)
Enddate = DateSerial(Year(txtdateenterered), month(txtdateenterered) + 1, 1 - 1)
 sql = ""
 sql = "set dateformat dmy Select count(distinct(Vehicle)) as j  from   d_AssignmentVehicle where Active = '" & True & "' "
 Set rsj = cn.Execute(sql)
 C = rsj.Fields(0)
 sql = ""
 sql = "set dateformat dmy Select distinct(Vehicle) as j  from   d_AssignmentVehicle where Active = '" & True & "' "
 Set rsr = cn.Execute(sql)
 Do While Not C <= 0
 l = rsr.Fields(0)
    
    sql = ""
    sql = "set dateformat dmy SELECT ExpenseAcc, ExpeLedger FROM d_AssignmentVehicle where Vehicle='" & l & "'"
    Set rsd = cn.Execute(sql)
  sql = ""
  sql = "set dateformat dmy Select count(CrAccNo) as x from GLTRANSACTIONS where CrAccNo ='" & rsd.Fields(1) & "'and TransDate>= '" & Startdate & "' And TransDate<='" & Enddate & "'  "
  Set rs = cn.Execute(sql)
  X = rs.Fields(0)
  sql = ""
  sql = "set dateformat dmy Select CrAccNo,TransDate  from   GLTRANSACTIONS where CrAccNo='" & rsd.Fields(1) & "'and TransDate>= '" & Startdate & "' And TransDate<='" & Enddate & "'  "
  Set rsk = cn.Execute(sql)
  Do While Not X <= 0
  Label32 = X
  Label32 = "Plase wait " & Label32 & ""
   sql = ""
   sql = "set dateformat dmy Select TransDate,Amount,TransDescript from GLTRANSACTIONS where CrAccNo='" & rsk.Fields(0) & "'and TransDate ='" & rsk.Fields(1) & "'  "
  'sql = "set dateformat dmy SELECT d.RefNo,m.DName, d.DispDate, d.DispQnty,d.Amount,d.PaidAmount FROM d_MilkControl AS d INNER JOIN d_Debtors AS m ON d.DCode = m.DCode
   Set rst = cn.Execute(sql)
    sql = ""
    sql = "set dateformat dmy insert into  d_Debtorsparchases2( Date, Remarks, kgs, Amount, [Paid Amount],Balance, Expenses,Vehicle)"
    sql = sql & "  values('" & rsk.Fields(1) & "','" & rst.Fields(2) & "','0','0','0','0','" & rst.Fields(1) & "','" & l & "')"
    cn.Execute sql
    rsk.MoveNext
  X = X - 1
 Loop
 rsr.MoveNext
 C = C - 1
Loop
Exit Sub
End Sub

Private Sub cmdreport_Click()
    reportname = "incomevsevehicle.rpt"
    Show_Sales_Crystal_Report "", reportname, ""
    'cmdreport.Enabled = False
End Sub

Private Sub cmdsave_Click()
Dim Active As String
On Error GoTo ErrorHandler

If txtTCode = "" Then
MsgBox "Please enter the Debtor code ", vbInformation, "Missing Information"
txtTCode.SetFocus
Exit Sub
End If

If chkActive.value = vbChecked Then
    Active = "1"
Else
    Active = "0"
End If
'sql = ""
'sql = "set dateformat dmy SELECT * From d_Debtors where DCode ='" & txtTCode & "'"
'Set rs = oSaccoMaster.GetRecordset(sql)
'If rs.EOF Then
'      Set cn = New ADODB.Connection
'    sql = ""
'    sql = "SET dateformat DMY Update  d_Debtors SET DNmame= '" & txtNames & "',CertNo='" & txtId & "',Locations='" & cboLocation & "',TregDate='" & DTPRegDate & "',email='" & txtEmail & "',Phoneno='" & txtPhone & "',Town='" & txtTown & "',Address='" & txtPAddress & "',price=" & CCur(txtPrice) & ",Active=" & Active & ",AccDr='" & txtDrAccNo & "',AccCr='" & txtCrAccNo & "' where DCode='" & txtTCode & "'"
'    oSaccoMaster.ExecuteThis (sql)
'
'    Else
'     MsgBox "Debtor code already exist, Use a different code ", vbInformation, "Missing Information"
'   Exit Sub
'  Exit Sub
  
 If newa = 1 Then
    Set cn = New ADODB.Connection
    sql = ""
    sql = "d_sp_Debtors '" & txtTCode & "','" & txtNames & "','" & txtId & "','" & txtVehicle & "','" & txtdateenterered & "','" & txtEMail & "','" & txtPhone & "','" & txtTown & "','" & txtPAddress & "'," & CCur(txtprice) & "," & CCur(txtsubsidy) & ",'" & txtAccno & "','" & cboBName & "'," & Active & ",'" & cboBBranch & "','" & cbobranch & "','" & User & "','" & txtDrAccNo & "','" & txtCrAccNo & "','" & txtcessrate & "','" & txtcessdebit & "','" & txtsta & "','" & cessapp & "'"
    oSaccoMaster.ExecuteThis (sql)
   Else
    Set cn = New ADODB.Connection
    sql = ""
    sql = "SET dateformat DMY Update  d_Debtors SET DName= '" & txtNames & "',CertNo='" & txtId & "',Locations='" & txtVehicle & "',TregDate='" & txtdateenterered & "',email='" & txtEMail & "',Phoneno='" & txtPhone & "',Town='" & txtTown & "',Address='" & txtPAddress & "',price=" & CCur(txtprice) & ",Active=" & Active & ",AccDr='" & txtDrAccNo & "',AccCr='" & txtCrAccNo & "',crcess='" & txtsta & "' where DCode='" & txtTCode & "'"
    oSaccoMaster.ExecuteThis (sql)
 End If
cmdNew_Click
cmdsave.Enabled = False
MsgBox "Records successively updated."
loadReg
Exit Sub
ErrorHandler:
MsgBox err.description

End Sub
Private Sub cmdsave3_Click()
On Error GoTo ErrorHandler

        Startdate = DateSerial(Year(txtdateenterered), month(txtdateenterered), 1)
        Enddate = DateSerial(Year(txtdateenterered), month(txtdateenterered) + 1, 1 - 1)

If chkoutletre.value = 0 Then
  If txtdcode = "" Then
   MsgBox "Debtors code cannot be blank; input an existing one", vbCritical
   cboNames.SetFocus
  Exit Sub
  End If

  If txtAmount > 0 Then
   If txtRemarks = "" Then
     MsgBox "Please enter the Remarks if Cash or Paybill."
     txtRemarks.SetFocus
    Exit Sub
   End If
  End If
  Dim transfer As Integer
  
  If CheckTranfer = vbChecked Then
   transfer = 1
  Else
    transfer = 0
  End If
  

 If chkpai = 0 Then
      If txtdcode = "" Then
      MsgBox "Debtors code cannot be blank; input an existing one", vbCritical
      cboNames.SetFocus
      Exit Sub
      End If
      
     If txtDispatch = "" Then
        MsgBox "Please enter the dispatch quantity."
            txtDispatch.SetFocus
        Exit Sub
     End If
    
     If txtIntake = "" Then
        MsgBox "Please enter the intake quantity."
            txtIntake.SetFocus
        Exit Sub
     End If

    If txtRefNo = "" Then
       MsgBox "Please enter the reference number."
           txtRefNo.SetFocus
       Exit Sub
    End If
    '/////check if it is to delete
    If chkdelete = 0 Then
 
        '//check if the dispatch is greater than the dipping
            If CDbl(txtIntake) < CDbl(txtDispatch) Then 'raiise an alarm
             MsgBox "You cannot take more than what you have in the tank", vbCritical
            Exit Sub
            End If
          Dim Debit As String
          Dim Credit As String
        
            sql = ""
            sql = "SET dateformat dmy SELECT * FROM  d_MilkControl  WHERE     DispDate = '" & txtdateenterered & "' and DispQnty = '" & txtDispatch & "'and dcode = '" & txtdcode & "'"
            Set rs = oSaccoMaster.GetRecordset(sql)
            If Not rs.EOF Then
              MsgBox "You have already dispatch for that day", vbInformation
            Exit Sub
            End If
          Dim Y As String
          Y = cboNames
          Debit = lbldrstock
        
          Credit = lblcrvehicle
        
            If Not Save_GLTRANSACTION(Format(txtdateenterered, "dd/mm/yyyy"), (CCur(Price) * CCur(txtDispatch)), Debit, Credit, Y, txtRefNo, User, ErrorMessage, "Milk Sales", 1, 1, txtRefNo, transactionNo, "", "", 0) Then
                    If ErrorMessage <> "" Then
                        MsgBox ErrorMessage, vbInformation, Me.Caption
                        ErrorMessage = ""
                    End If
            End If
            '''' behind the scene
            sql = "SET dateformat dmy SELECT AVG(distinct PPU) AS AveragePrice FROM d_Milkintake where TransDate='" & txtdateenterered & "'"
            Set rsg = oSaccoMaster.GetRecordset(sql)
            sql = "select * from GLSetDefaultGls Where Affect='Join'"
            Set rst = oSaccoMaster.GetRecordset(sql)
            
            If Not Save_GLTRANSACTION(Format(txtdateenterered, "dd/mm/yyyy"), (CCur(rsg!AveragePrice) * CCur(txtDispatch)), rst!dr, rst!cr, Y, txtRefNo, User, ErrorMessage, "Milk Sales", 1, 1, txtRefNo, transactionNo, "", "", 0) Then
                    If ErrorMessage <> "" Then
                        MsgBox ErrorMessage, vbInformation, Me.Caption
                        ErrorMessage = ""
                    End If
            End If
            '''' end of behind the scene
            
            
   Else
           Y = cboNames
           Debit = lbldrstock
        
           Credit = lblcrvehicle
        
           If Not Save_GLTRANSACTION(Format(txtdateenterered, "dd/mm/yyyy"), (CCur(Price) * CCur(txtDispatch)), Credit, Debit, Y, txtRefNo, User, ErrorMessage, "Milk Sales Remove", 1, 1, txtRefNo, transactionNo, "", "", 0) Then
                    If ErrorMessage <> "" Then
                        MsgBox ErrorMessage, vbInformation, Me.Caption
                        ErrorMessage = ""
                    End If
            End If
            
                '''' behind the scene
            sql = "SET dateformat dmy SELECT AVG(distinct PPU) AS AveragePrice FROM d_Milkintake where TransDate='" & txtdateenterered & "'"
            Set rsg = oSaccoMaster.GetRecordset(sql)
            sql = "select * from GLSetDefaultGls Where Affect='Join'"
            Set rst = oSaccoMaster.GetRecordset(sql)
            
            If Not Save_GLTRANSACTION(Format(txtdateenterered, "dd/mm/yyyy"), (CCur(rsg!AveragePrice) * CCur(txtDispatch)), rst!cr, rst!dr, Y, txtRefNo, User, ErrorMessage, "Milk Sales Remove", 1, 1, txtRefNo, transactionNo, "", "", 0) Then
                    If ErrorMessage <> "" Then
                        MsgBox ErrorMessage, vbInformation, Me.Caption
                        ErrorMessage = ""
                    End If
            End If
            '''' end of behind the scene

   End If
'/////////////end of checking if it is to delete
    

Else
' If txtdcode = "" Then
'  MsgBox "Debtors code cannot be blank; input an existing one", vbCritical
'  cboNames.SetFocus
' Exit Sub
' End If
'
'  If txtremarks = "" Then
'    MsgBox "Please enter the Remarks if Cash or Paybill."
'        txtremarks.SetFocus
'    Exit Sub
' End If
End If
        '/////check if it is to delete
    If chkdelete = 0 Then
        ''...................insert the amount to debtor if available................................
           Dim Amount1 As Integer
           Set rs = New ADODB.Recordset
           sql = ""
           sql = "SET dateformat dmy Select Amount,PaidAmount  from d_MilkControl  where DCode ='" & txtdcode & "' and DispDate='" & txtdateenterered.value & "'"
           Set rs = oSaccoMaster.GetRecordset(sql)
    
           If rs.EOF Then
    '         sql = ""
    '         sql = "d_sp_MilkControl  '" & txtdateenterered & "'," & txtDispatch & ",'0'," & txtIntake & ",'0'," & Price & ",'" & txtRefNo & "','" & Credit & "','" & Debit & "','" & User & "','" & txtdcode & "','" & cboVehicle & "','" & txtamountp & "','" & txtamount & "'"
    '         oSaccoMaster.ExecuteThis (sql)
            sql = ""
            sql = "set dateformat dmy insert into  d_MilkControl(DispDate, DispQnty, DipQnty, InQnty, Variance, Price, RefNo, DebitAcc, CreditAcc, AuditId, Auditdatetime, DCode, vehicleno, Amount, PaidAmount,Transfer) values('" & txtdateenterered & "'," & txtDispatch & ",'0'," & txtIntake & ",'0'," & Price & ",'" & txtRefNo & "','" & Credit & "','" & Debit & "','" & User & "','" & Now & "','" & txtdcode & "','" & cboVehicle & "','" & txtamountp & "','" & txtAmount & "','" & transfer & "')"
            oSaccoMaster.ExecuteThis (sql)
           Else
             sql = ""
             sql = "set dateformat DMY update d_MilkControl set PaidAmount=" & rs.Fields("PaidAmount") + txtAmount & " where DCode ='" & txtdcode & "' and DispDate='" & txtdateenterered.value & "' "
             oSaccoMaster.ExecuteThis (sql)
           End If
         'Else
         'End If
    
        '''..................end of debtor...........................................................
    '******************* *********insert to gl
        'txtamount = 0
          If txtAmount > 0 Then
                If txtRemarks = "" Then
                  MsgBox "Please enter the Remarks if Cash or Paybill."
                  txtRemarks.SetFocus
                 Exit Sub
                End If
             Dim E As String
             E = txtRemarks
            sql = ""
            sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,AuditTime,auditid,cash,doc_posted) values"
            sql = sql & "('" & txtdateenterered & "'," & txtAmount & ",'" & lbldracc & "','" & lblcracc & "','" & cboNames & "','' ,'" & E & "-MILK PAYMENTS','" & Now & "','" & User & "',0,0)"
            oSaccoMaster.ExecuteThis (sql)
          Else
           'Exit Sub
            If txtAmount < 0 Then
             Dim mat As Integer
             mat = txtAmount * -1
             sql = ""
             sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,AuditTime,auditid,cash,doc_posted) values"
             sql = sql & "('" & txtdateenterered & "'," & mat & ",'" & lblcracc & "','" & lbldracc & "','" & cboNames & "','' ,'Reversal-MILK PAYMENTS','" & Now & "','" & User & "',0,0)"
             oSaccoMaster.ExecuteThis (sql)
            End If
          End If
    '****************************end of gl
    
    Else
    

    
        sql = ""
        sql = "delete from d_MilkControl where DCode ='" & txtdcode & "' and DispDate='" & txtdateenterered.value & "' "
        oSaccoMaster.ExecuteThis (sql)
        
        sql = ""
        sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,AuditTime,auditid,cash,doc_posted) values"
        sql = sql & " ('" & txtdateenterered & "'," & txtAmount & ",'" & lblcracc & "','" & lbldracc & "','" & cboNames & "','' ,'" & E & "-MILK PAYMENTS Remove','" & Now & "','" & User & "',0,0)"
        oSaccoMaster.ExecuteThis (sql)
       
       '''' behind the scene
        sql = "SET dateformat dmy SELECT AVG(distinct PPU) AS AveragePrice FROM d_Milkintake where TransDate='" & txtdateenterered & "'"
        Set rsg = oSaccoMaster.GetRecordset(sql)
        sql = "select * from GLSetDefaultGls Where Affect='Join'"
        Set rst = oSaccoMaster.GetRecordset(sql)
        
        If Not Save_GLTRANSACTION(Format(txtdateenterered, "dd/mm/yyyy"), (CCur(rsg!AveragePrice) * CCur(txtDispatch)), rst!cr, rst!dr, Y, txtRefNo, User, ErrorMessage, "Milk PAYMENTS Remove", 1, 1, txtRefNo, transactionNo, "", "", 0) Then
                If ErrorMessage <> "" Then
                    MsgBox ErrorMessage, vbInformation, Me.Caption
                    ErrorMessage = ""
                End If
        End If
      '''' end of behind the scene
       
       sql = "set dateformat dmy select Vehicle, Date, Kgs, Customer from d_OutletVehicle where Vehicle ='" & cboVehicle & "' and Date='" & txtdateenterered.value & "'"
       Set rst = New ADODB.Recordset
       rst.Open sql, cn
       If Not rst.EOF Then
         sql = ""
         sql = "set dateformat DMY update d_OutletVehicle set Kgs='" & rst.Fields("Kgs") - txtDispatch.Text & "' where Vehicle ='" & cboVehicle & "' and Date='" & txtdateenterered.value & "'"
         cn.Execute sql
       End If
       
  End If
  
  '//////end of checking if it is to delete
         Dim DName As String
          Set rs = New ADODB.Recordset
          sql = "SELECT DName from d_Debtors where DCode='" & txtdcode & "'"
          Set rs = oSaccoMaster.GetRecordset(sql)
          If Not rs.EOF Then
          DName = rs!DName
          End If


Else
''''''''' milk for outlets '''''''''''''''''''

        If txtAmount > 0 Then
             MsgBox "Please make payment from the outlet form."
             Exit Sub
        End If
        
                '/////check if it is to delete
   If chkdelete = 0 Then


          Provider = "MAZIWA"
          Set cn = New ADODB.Connection
         cn.Open Provider, "atm", "atm"
          Set rst = New ADODB.Recordset
          sql = ""
          sql = "set dateformat dmy select p_code, p_name, Date_Entered, Qin, Qout, o_bal, user_id, Wprice, Rprice, Branch from d_Outlet where p_name='" & cboNames & "' "
          rst.Open sql, cn
          'If Not rst.EOF Then
          '''''''check if the record exist'''''''''''''
              sql = ""
              sql = "set dateformat dmy select P_CODE,qout,Qin,Rprice from d_Outlet where p_name='" & cboNames & "'and Date_Entered='" & txtdateenterered.value & "'"
              Set rs = New ADODB.Recordset
              rs.Open sql, cn
             If rs.EOF Then
            '// insert into ag_products
                If txtSERIALNO = "" Then txtSERIALNO = 0
                 sql = ""
                 sql = "set dateformat dmy insert into  d_Outlet( p_code, p_name, Date_Entered, Qin, Qout, o_bal, user_id, Wprice, Rprice, Branch)"
                 sql = sql & "  values('" & rst.Fields(0) & "','" & cboNames.Text & "','" & txtdateenterered.value & "'," & txtDispatch.Text & "," & txtDispatch.Text & "," & txtDispatch.Text & ",'" & User & "'," & rst.Fields(7) & "," & rst.Fields(8) & ",'" & rst.Fields(9) & "')"
                 'sql = sql & "  values('" & rst.Fields(0) & "','" & cboNames.Text & "','" & txtdateenterered.value & "'," & txtDispatch.Text & "," & txtDispatch.Text + rst.Fields(4) & "," & txtDispatch.Text & ",'" & User & "'," & rst.Fields(7) & "," & rst.Fields(8) & ",'" & rst.Fields(9) & "')"
                 cn.Execute sql
                 
             Else
                
                sql = "set dateformat DMY update d_Outlet set p_name='" & cboNames.Text & "',Qin=" & rs.Fields("Qin") + txtDispatch.Text & ",Qout=" & txtDispatch.Text + rs.Fields("qout") & ",o_bal=" & rs.Fields("qout") + txtDispatch.Text & ",Date_Entered='" & txtdateenterered.value & "' where p_code='" & rst.Fields(0) & "' and branch='" & rst.Fields(9) & "'and Date_Entered='" & txtdateenterered.value & "'"
                cn.Execute sql
                
             End If
             
                sql = "set dateformat dmy select Date_Entered, p_name, Quantity, OutletName from d_Outletstock where Date_Entered='" & txtdateenterered.value & "' and p_name='" & cboNames & "'AND OutletName ='" & rst.Fields(9) & "'"
                Set rss = New ADODB.Recordset
                rss.Open sql, cn
                If rss.EOF Then
                    sql = ""
                    sql = "set dateformat dmy insert into  d_Outletstock(Date_Entered,p_name, Quantity, OutletName)"
                    sql = sql & "  values('" & txtdateenterered.value & "','" & rst.Fields(1) & "'," & txtDispatch.Text & ",'" & rst.Fields(9) & "')"
                    cn.Execute sql
                Else
                    sql = "set dateformat DMY update d_Outletstock set Quantity =" & rss.Fields(2) + txtDispatch.Text & " where p_name='" & cboNames & "'and Date_Entered='" & txtdateenterered.value & "' and OutletName='" & rst.Fields(9) & "'"
                    cn.Execute sql
                End If
            
                  '//XXXXXXXXXXXXXXX
            '********** credit agent ledger sale and stock
                Dim kgs, totalkgss As Double
                sql = "set dateformat dmy select P_CODE,qout,Qin,Rprice from d_Outlet where p_name='" & cboNames & "'and Date_Entered='" & txtdateenterered.value & "'"
                Set rsh = oSaccoMaster.GetRecordset(sql)
                kgs = CCur(txtDispatch) * CCur(rsh!Rprice)
                sql = ""
                sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,auditid,cash,doc_posted) values"
                sql = sql & "('" & txtdateenterered & "','" & kgs & "','" & lbldrstock & "','" & lblcrvehicle & "','','" & cboVehicle & "' ,'SALES ON- " & "" & cboNames & "','" & User & "',0,0)"
                oSaccoMaster.ExecuteThis (sql)
            'XXXXXXXXXXXXXXXXXXXXXX
        
                 '''' behind the scene
                 'Dim totalkgs As Double
                    sql = "SET dateformat dmy SELECT isnull(AVG(distinct PPU),0) AS AveragePrice FROM d_Milkintake where TransDate='" & txtdateenterered & "'"
                    Set rsg = oSaccoMaster.GetRecordset(sql)
                    sql = "select * from GLSetDefaultGls Where Affect='Join'"
                    Set rst = oSaccoMaster.GetRecordset(sql)
                    totalkgss = CCur(rsg!AveragePrice) * CCur(txtDispatch)
        
                    sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,auditid,cash,doc_posted) values"
                    sql = sql & "('" & txtdateenterered & "','" & totalkgss & "','" & rst!dr & "', '" & rst!cr & "','','" & cboVehicle & "' ,'SALES ON- " & "" & cboNames & "','" & User & "','0','0')"
                    oSaccoMaster.ExecuteThis (sql)
            '''''''end of code behind''''''''''''''''
                    
                
                ''''* to vehicle table
                  'If chkcustomer = 1 Then
                   Provider = "MAZIWA"
                   Set cn = New ADODB.Connection
                  cn.Open Provider, "atm", "atm"
                   'Set rs = New ADODB.Recordset
                   sql = "set dateformat dmy select Vehicle, Date, Kgs, Customer from d_OutletVehicle where Vehicle ='" & cboVehicle & "' and Date='" & txtdateenterered.value & "'"
                   Set rst = New ADODB.Recordset
                   rst.Open sql, cn
                    If rst.EOF Then
                    sql = ""
                    sql = "set dateformat dmy insert into  d_OutletVehicle(Vehicle, Date, Kgs, Customer)"
                    sql = sql & "  values('" & cboVehicle & "','" & txtdateenterered.value & "'," & txtDispatch.Text & ",'null')"
                    cn.Execute sql
                   Else
                    sql = ""
                    sql = "set dateformat DMY update d_OutletVehicle set Kgs=" & txtDispatch.Text + rst.Fields(2) & " where Vehicle ='" & cboVehicle & "' and Date='" & txtdateenterered.value & "'"
                    cn.Execute sql
                   End If
                '''''check if already dispatch to vehicle
                    sql = ""
                    sql = "select * from d_OutletDispatch where Vehicle= '" & cboVehicle & "' AND Date ='" & txtdateenterered.value & "'"
                    Set rsinstock = oSaccoMaster.GetRecordset(sql)
                    If rsinstock.EOF Then
                        sql = ""
                        sql = "set dateformat dmy insert into  d_OutletDispatch(Date, Vehicle, OutletName, Quantity)"
                        sql = sql & "  values('" & txtdateenterered.value & "','" & cboVehicle & "','Null','" & txtDispatch.Text & "')"
                        cn.Execute sql
                    Else
                        sql = ""
                        sql = "set dateformat DMY Update d_OutletDispatch SET Vehicle= '" & cboVehicle & "', Date='" & txtdateenterered.value & "',Quantity='" & txtDispatch.Text + rsinstock.Fields(4) & "' WHERE Vehicle= '" & cboVehicle & "' and Date='" & txtdateenterered.value & "'"
                        cn.Execute sql
                    End If
                '''''end
        
     Else
     
               '''''''check if the record exist'''''''''''''
                Set rst = New ADODB.Recordset
                sql = ""
                sql = "set dateformat dmy select p_code, p_name, Date_Entered, Qin, Qout, o_bal, user_id, Wprice, Rprice, Branch from d_Outlet where p_name='" & cboNames & "' "
                rst.Open sql, cn
                
              sql = ""
              sql = "set dateformat dmy select P_CODE,qout,Qin,Rprice from d_Outlet where p_name='" & cboNames & "'and Date_Entered='" & txtdateenterered.value & "'"
              Set rs = New ADODB.Recordset
              rs.Open sql, cn
             If Not rs.EOF Then
                sql = "set dateformat DMY update d_Outlet set p_name='" & cboNames.Text & "',Qin=" & rs.Fields("Qin") - txtDispatch.Text & ",Qout=" & rs.Fields("qout") - txtDispatch.Text & ",o_bal=" & rs.Fields("qout") - txtDispatch.Text & ",Date_Entered='" & txtdateenterered.value & "' where p_code='" & rst.Fields(0) & "' and branch='" & rst.Fields(9) & "'and Date_Entered='" & txtdateenterered.value & "'"
                cn.Execute sql
                
             End If
             
                sql = "set dateformat dmy select Date_Entered, p_name, Quantity, OutletName from d_Outletstock where Date_Entered='" & txtdateenterered.value & "' and p_name='" & cboNames & "'AND OutletName ='" & rst.Fields(9) & "'"
                Set rss = New ADODB.Recordset
                rss.Open sql, cn
                If Not rss.EOF Then
                    sql = "set dateformat DMY update d_Outletstock set Quantity =" & rss.Fields(2) - txtDispatch.Text & " where p_name='" & cboNames & "'and Date_Entered='" & txtdateenterered.value & "' and OutletName='" & rst.Fields(9) & "'"
                    cn.Execute sql
                End If
                
                
                     '//XXXXXXXXXXXXXXX
            '********** credit agent ledger sale and stock
                sql = "set dateformat dmy select P_CODE,qout,Qin,Rprice from d_Outlet where p_name='" & cboNames & "'and Date_Entered='" & txtdateenterered.value & "'"
                Set rsh = oSaccoMaster.GetRecordset(sql)
                kgs = CCur(txtDispatch) * CCur(rsh!Rprice)
                sql = ""
                sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,auditid,cash,doc_posted) values"
                sql = sql & "('" & txtdateenterered & "','" & kgs & "','" & lblcrvehicle & "','" & lbldrstock & "','','" & cboVehicle & "' ,'SALES ON- " & "" & cboNames & "','" & User & "',0,0)"
                oSaccoMaster.ExecuteThis (sql)
            'XXXXXXXXXXXXXXXXXXXXXX
        
                 '''' behind the scene
                 'Dim totalkgs As Double
                    sql = "SET dateformat dmy SELECT isnull(AVG(distinct PPU),0) AS AveragePrice FROM d_Milkintake where TransDate='" & txtdateenterered & "'"
                    Set rsg = oSaccoMaster.GetRecordset(sql)
                    sql = "select * from GLSetDefaultGls Where Affect='Join'"
                    Set rst = oSaccoMaster.GetRecordset(sql)
                    totalkgss = CCur(rsg!AveragePrice) * CCur(txtDispatch)
        
                    sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,auditid,cash,doc_posted) values"
                    sql = sql & "('" & txtdateenterered & "','" & totalkgss & "','" & rst!cr & "', '" & rst!dr & "','','" & cboVehicle & "' ,'SALES ON- " & "" & cboNames & "','" & User & "','0','0')"
                    oSaccoMaster.ExecuteThis (sql)
            '''''''end of code behind''''''''''''''''
            
                     ''''* to vehicle table
                  'If chkcustomer = 1 Then
                   Provider = "MAZIWA"
                   Set cn = New ADODB.Connection
                  cn.Open Provider, "atm", "atm"
                   'Set rs = New ADODB.Recordset
                   sql = "set dateformat dmy select Vehicle, Date, Kgs, Customer from d_OutletVehicle where Vehicle ='" & cboVehicle & "' and Date='" & txtdateenterered.value & "'"
                   Set rst = New ADODB.Recordset
                   rst.Open sql, cn
                    If Not rst.EOF Then
                        sql = ""
                        sql = "set dateformat DMY update d_OutletVehicle set Kgs=" & txtDispatch.Text - rst.Fields(2) & " where Vehicle ='" & cboVehicle & "' and Date='" & txtdateenterered.value & "'"
                        cn.Execute sql
                   End If
                '''''check if already dispatch to vehicle
                    sql = ""
                    sql = "select * from d_OutletDispatch where Vehicle= '" & cboVehicle & "' AND Date ='" & txtdateenterered.value & "'"
                    Set rsinstock = oSaccoMaster.GetRecordset(sql)
                    If Not rsinstock.EOF Then
                        sql = ""
                        sql = "set dateformat DMY Update d_OutletDispatch SET Vehicle= '" & cboVehicle & "', Date='" & txtdateenterered.value & "',Quantity='" & txtDispatch.Text - rsinstock.Fields(4) & "' WHERE Vehicle= '" & cboVehicle & "' and Date='" & txtdateenterered.value & "'"
                        cn.Execute sql
                    End If
                '''''end
            
            
     
     End If
        
End If
'..............END OF  DAILY INTAKE INSERT FOR DEBTORS ONLY.........................
'mysql = "set dateformat dmy Insert into Receiptno(Receiptno,Auditdate,auditid)values('" & txtRefNo & "','" & Format(Get_Server_Date, "dd/MM/yyyy") & "','" & User & "')"
'oSaccoMaster.ExecuteThis (mysql)
If chkPrint = vbChecked Then
    
If chkPrint = vbChecked Then
    
'/*Print out
 Dim fso, chkPrinter, txtFile
    Dim ttt
     Dim escFeedAndCut As String
     Dim escNewLine As String
     Dim escUnerLineON As String
     Dim escUnerLineOnX2 As String
     Dim escUnerLineOff As String
     Dim escBoldOn As String
     Dim escBoldOff As String
     Dim escNegativeOn As String
     Dim escNegativeOff As String
     Dim esc8CpiOn As String
     Dim esc8CPiOff As String
     Dim esc16Cpi As String
     Dim esc20Cpi As String
     Dim escAlignLeft As String
     Dim escAlignCenter As String
     Dim escAlignRight As String
    
     
        escNewLine = Chr(10) '//New Line (LF Line Feed)
        escUnerLineON = Chr(27) + Chr(45) + Chr(1) '//Unerline on
        escUnerLineOnX2 = Chr(27) + Chr(45) + Chr(1) '//Unerline on X2
        escUnerLineOff = Chr(27) + Chr(45) + Chr(0) '//unerline off
        escBoldOn = Chr(27) + Chr(69) + Chr(1) '//Bold on
        escBoldOff = Chr(27) + Chr(69) + Chr(0) '//Bold off
        escNegativeOn = Chr(29) + Chr(66) + Chr(1) '//White on Black on
        escNegativeOff = Chr(29) + Chr(66) + Chr(0) '//white on
        esc8CpiOn = Chr(29) + Chr(33) + Chr(16) '//Font Size X2 on
        esc8CPiOff = Chr(29) + Chr(33) + Chr(0) '//Font size X2 off
        esc16Cpi = Chr(27) + Chr(77) + Chr(48) '//Font A - Normal Size
        esc20Cpi = Chr(27) + Chr(77) + Chr(49) '//Font B - Small Font
        escAlignLeft = Chr(27) + Chr(97) + Chr(48) '//Align text to the left
        escAlignCenter = Chr(27) + Chr(97) + Chr(49) '//Align text to the center
        escAlignRight = Chr(27) + Chr(97) + Chr(50) '//Align text to the right
        escFeedAndCut = Chr(29) + Chr(86) + Chr(65) '//Partial cut and feed
       'cdgPrint.Filter = "*.csv|*.txt"
        'cdgPrint.ShowSave
        Dim PORT As String
   '     PORT = ports
        'ttt = "LPT1" 'LPT1
        ttt = ports
        'ttt = "D:\PROJECTS\FOSA\DAILY" & Date & ""
        Set fso = CreateObject("Scripting.FileSystemObject")
        'Set chkPrinter = fso.GetFile(ttt)
        
    Set txtFile = fso.CreateTextFile(ttt, True)
    txtFile.WriteLine "         " & cname & ""
    txtFile.WriteLine "         Address :" & paddress & ""
    txtFile.WriteLine "         Phone :" & Phone & ""
    txtFile.WriteLine "         Email :" & Email & ""
    'txtfile.WriteLine " " & txtSNo
    
    txtFile.WriteLine "          Delivery Note"
    txtFile.WriteLine "**********************************************"
        
    Set rs2 = New ADODB.Recordset
    sql = "d_sp_ReceiptNumber"
    Set rs2 = oSaccoMaster.GetRecordset(sql)
    
    Dim RNumber As String
    'RNumber = rs2.Fields(0)
    If Not IsNull(rs2.Fields(0)) Then RNumber = rs2.Fields(0)
    'Else
    'RNumber = "0"
    'End If
    
    txtFile.WriteLine "CsNO :" & txtRefNo
    txtFile.WriteLine "To :" & lblDebtors
   txtFile.WriteLine " *********************************************************************"
    txtFile.WriteLine "DESCRIPTION " & vbTab & "" & vbTab & "value"
    sql = "SELECT     d.DCode, d.DName, SUM(m.DispQnty) AS quantity FROM         d_Debtors AS d INNER JOIN d_MilkControl AS m ON d.DCode = m.dcode WHERE     (DispDate = '" & txtdateenterered & "') GROUP BY d.DCode, d.DName"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If Not rs.EOF Then

    txtFile.WriteLine "Milk supplied :" & vbTab & "" & vbTab & " " & rs!Quantity & ""
    txtFile.WriteLine "Amount Payable :" & vbTab & "  " & txtamountp
    txtFile.WriteLine "Receipt Number :" & vbTab & "  " & txtRefNo
    txtFile.WriteLine "Dispatched by :" & vbTab & " " & username & ""
    
    txtFile.WriteLine "---------------------------------------"
    End If
'    txtFile.WriteLine "Receipt Number :" & RNumber
'    txtFile.WriteLine "TRANSPORTER :" & TRANSPORTER
    txtFile.WriteLine "Vehicle No :" & cboVehicle
    txtFile.WriteLine "Received by :" & txtreceiveby
    txtFile.WriteLine "---------------------------------------"
    txtFile.WriteLine "     Date :" & Format(txtdateenterered, "dd/mm/yyyy") & " ,Time : " & Format(Time, "hh:mm:ss AM/PM")
    txtFile.WriteLine "" & motto & ""
    txtFile.WriteLine "---------------------------------------"
    'If chkComment.value = vbChecked Then
        'txtFile.WriteLine txtComment
        txtFile.WriteLine "---------------------------------------"
        txtFile.WriteLine "********POWERED BY EASYMA***************"
    'End If
    txtFile.WriteLine escFeedAndCut
    
 txtFile.Close
 Reset
End If
End If

MsgBox "Records saved successifully."
loadBranchesTypes
txtdcode = ""
txtDispatch = ""
'txtIntake = ""
txtAmount = ""
txtRefNo = ""
txtamountp = ""
txtRemarks = ""
txtdracc = ""
txtcracc = ""
lbldracc = ""
lblcracc = ""
    'ListView2.Visible = False
    chkpai.value = vbUnchecked
    'chkPay.value = vbUnchecked
    cmdnew3.Enabled = True
    cmdsave3.Enabled = True
    CheckTranfer.value = vbUnchecked
   ' cmdEdit.Enabled = False
    'SSTab1_DblClick
    cmdnew3_Click
    cmdnewsearch_Click
    
 Exit Sub
ErrorHandler:
MsgBox err.description

End Sub

Private Sub cmdSearch_Click()
frmSearchGLAccounts.Show vbModal, Me
    If Continue Then
        If SearchValue <> "" Then
            txtDrAccNo = SearchValue
            SearchValue = ""
        End If
    End If
End Sub

Private Sub Combo1_Change()

End Sub
Private Sub fghj()
'Private Sub SSTab1_DblClick()
    txtVehicle.Clear
    Set rstm = New Recordset
    Dim cn As Connection
    Set cn = New ADODB.Connection
    Provider = "MAZIWA"
    'Provider = cnn
    cn.Open Provider, "atm", "atm"
    Set rstm = New Recordset
    sql = "Select distinct(Vehicle) from   d_VehicleTill where Vehicle not like'%PLANT%' order by Vehicle"
    'Select distinct(Locations) from   d_Debtors
    rstm.Open sql, cn, adOpenKeyset, adLockOptimistic
    While Not rstm.EOF
    txtVehicle.AddItem rstm.Fields(0)
    rstm.MoveNext
    Wend
End Sub

Private Sub cmdshort_Click()
frmNominals.Show vbModal
End Sub

Private Sub Command2_Click()
    reportname = "Kiarie reports.rpt"
    Show_Sales_Crystal_Report "", reportname, ""
End Sub

Private Sub Command3_Click()

End Sub

Private Sub lblcracc_Change()
    Set rst = oSaccoMaster.GetRecordset("select glaccname from glsetup where accno='" & lblcracc & "'")
    If Not rst.EOF Then
    txtcracc = rst.Fields("glaccname")
    End If
End Sub

Private Sub lbldracc_Change()
    Set rst = oSaccoMaster.GetRecordset("select glaccname from glsetup where accno='" & lbldracc & "'")
    If Not rst.EOF Then
    txtdracc = rst.Fields("glaccname")
    End If
End Sub
Private Sub lbltotal_Change()
lbltotal.Caption = Format(lbltotal.Caption, "#,##0.00")
End Sub

Private Sub lbltotalkg_Change()
lbltotalkg.Caption = Format(lbltotalkg.Caption, "#,##0.0")
End Sub

Private Sub listview1_DblClick()
'cboNames = ListView1.SelectedItem
cboNames = ListView1.SelectedItem.SubItems(1)
loadb
txtDispatch = ListView1.SelectedItem.SubItems(3)
txtamountp = ListView1.SelectedItem.SubItems(4)
txtAmount = ListView1.SelectedItem.SubItems(5)
End Sub
'Private Sub cboNames_Validate(Cancel As Boolean)
'Dim a As Boolean, b As Integer
'Set rs = New ADODB.Recordset
'sql = ""
'sql = "set dateformat dmy select DispQnty,DCode,Amount, PaidAmount from d_MilkControl where DCode= '" & ListView1.SelectedItem.SubItems(1) & "' AND DispDate ='" & txtdateenterered.value & "'"
'    Set rs = oSaccoMaster.GetRecordset(sql)
'Set rs = oSaccoMaster.GetRecordset(sql)
'If Not rs.EOF Then
'If Not IsNull(rs.Fields(0)) Then txtDispatch = rs.Fields(0)
'If Not IsNull(rs.Fields(1)) Then txtdcode = rs.Fields(1)
'If Not IsNull(rs.Fields(2)) Then txtamountp = rs.Fields(2)
'If Not IsNull(rs.Fields(3)) Then txtamount = rs.Fields(3)
'End If
'End Sub

Private Sub ListViewG_DblClick()
Label31 = ListViewG.SelectedItem
cboGari = ListViewG.SelectedItem.SubItems(1)
cboGari2 = ListViewG.SelectedItem.SubItems(2)
Label30 = ListViewG.SelectedItem.SubItems(3)
 cmdActive.Enabled = "True"

End Sub

Private Sub txtDrAccNo_Change()
    On Error GoTo SysError
    Dim Account As Acc_Details
    Editing = True
    Account = Get_Acc_Details(txtDrAccNo, ErrorMessage)
    If Account.ACCNO <> "" Then
        lblDrAccName = Account.AccName
    Else
        If ErrorMessage <> "" Then
            MsgBox ErrorMessage, vbInformation, Me.Caption
            ErrorMessage = ""
        End If
        lblDrAccName = ""
    End If
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption

End Sub
Private Sub txtCrAccNo_Change()
 On Error GoTo SysError
    Dim Account As Acc_Details
        
        Editing = True
    Account = Get_Acc_Details(txtCrAccNo, ErrorMessage)
    If Account.ACCNO <> "" Then
        txtCrAccName = Account.AccName
    Else
        If ErrorMessage <> "" Then
            MsgBox ErrorMessage, vbInformation, Me.Caption
            ErrorMessage = ""
        End If
        txtCrAccName = ""
    End If
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub cmdstatement_Click()
   'reportname = "milkstatement.rpt"
    reportname = "d_DebtorsInvoice.rpt"
    Show_Sales_Crystal_Report "", reportname, ""
    'Show_Sales_Crystal_Report STRFORMULA, reportname, ""
End Sub
Public Sub loadBranchesTypes()
Startdate = DateSerial(Year(txtdateenterered), month(txtdateenterered), 1)
Enddate = DateSerial(Year(txtdateenterered), month(txtdateenterered) + 1, 1 - 1)

    With ListView1
    
        .ListItems.Clear
        
        .ColumnHeaders.Clear
    
    End With

    Set rs2 = CreateObject("adodb.recordset")
    'sql = "Select RefNo,DispDate, DispQnty, Amount, PaidAmount from d_MilkControl where DispDate='" & txtdateenterered & "'"
    sql = ""
    sql = "set dateformat dmy SELECT d.RefNo,m.DName, d.DispDate, d.DispQnty,d.Amount,d.PaidAmount,d.DCode FROM d_MilkControl AS d INNER JOIN d_Debtors AS m ON d.DCode = m.DCode WHERE     (DispDate = '" & txtdateenterered & "') and vehicleno='" & cboVehicle & "' order by RefNo desc"
    Set rs = oSaccoMaster.GetRecordset(sql)
    
    Set rs2 = CreateObject("adodb.recordset")
    
    Set clsClass = New cdbase
    
    Provider = clsClass.OpenCon
    
    Set cn = CreateObject("adodb.connection")
    
   cn.Open Provider, "atm", "atm"
    
    rs2.Open sql, cn
    
    With ListView1
        
        .ColumnHeaders.Add , , "Receipt"
        .ColumnHeaders.Add , , "Name"
        .ColumnHeaders.Add , , "Date"
        .ColumnHeaders.Add , , "Kgs"
        .ColumnHeaders.Add , , "Amount"
        .ColumnHeaders.Add , , "Paid Amount"
        .ColumnHeaders.Add , , "Balance"
        While Not rs2.EOF
        
            Set li = .ListItems.Add(, , Trim(rs2.Fields("RefNo")))
            
            li.ListSubItems.Add , , Trim(rs2.Fields("DName"))
            li.ListSubItems.Add , , Trim(rs2.Fields("DispDate"))
            li.ListSubItems.Add , , Trim(rs2.Fields("DispQnty"))
            li.ListSubItems.Add , , Trim(rs2.Fields("Amount"))
            li.ListSubItems.Add , , Trim(rs2.Fields("PaidAmount"))
            Dim balance As Double
            If rs2.Fields("DCode") = "BD060W09" Then
             MsgBox "Records of this vehicle " & txtVehicle & " does not exist.Please fill the ledgers manually"
            End If
             sql = ""
             sql = "set dateformat dmy SELECT sum(Amount), sum(PaidAmount) FROM d_MilkControl WHERE (DispDate>='" & Startdate & "' And DispDate <= '" & Enddate & "') and DCode='" & rs2.Fields("DCode") & "' "
             Set rs = oSaccoMaster.GetRecordset(sql)
             If Not rs.EOF Then
              balance = rs.Fields(0) - rs.Fields(1)
             Else
              balance = "0"
             End If
            li.ListSubItems.Add , , (balance)
            rs2.MoveNext
        
        Wend
    End With

    rs2.Close
    
    Set rs2 = Nothing
    loadOutlet
    ComputeTotal
    
ListView1.View = lvwReport

End Sub
Public Sub loadOutlet()
    
    Set rs3 = CreateObject("adodb.recordset")
    'sql = "Select RefNo,DispDate, DispQnty, Amount, PaidAmount from d_MilkControl where DispDate='" & txtdateenterered & "'"
    sql = ""
    sql = "set dateformat dmy Select p_code, p_name, Date_Entered, Qin from d_Outlet where Date_Entered='" & txtdateenterered & "'   "
    'sql = "set dateformat dmy SELECT d.RefNo,m.DName, d.DispDate, d.DispQnty,d.Amount,d.PaidAmount FROM d_MilkControl AS d INNER JOIN d_Debtors AS m ON d.DCode = m.DCode WHERE     (DispDate = '" & txtdateenterered & "') and vehicleno='" & cboVehicle & "' order by RefNo desc"
    Set rs = oSaccoMaster.GetRecordset(sql)
    
    Set rs3 = CreateObject("adodb.recordset")
    
    Set clsClass = New cdbase
    
    Provider = clsClass.OpenCon
    
    Set cn = CreateObject("adodb.connection")
    
   cn.Open Provider, "atm", "atm"
    
    rs3.Open sql, cn
    
    With ListView1
        
        '.ColumnHeaders.Add , , "Receipt"
        '.ColumnHeaders.Add , , "Name"
        '.ColumnHeaders.Add , , "Date"
        '.ColumnHeaders.Add , , "Kgs"
        '.ColumnHeaders.Add , , "Amount"
        '.ColumnHeaders.Add , , "Paid Amount"
        While Not rs3.EOF
        
            Set li = .ListItems.Add(, , Trim(rs3.Fields("p_code")))
            
            li.ListSubItems.Add , , Trim(rs3.Fields("p_name"))
            li.ListSubItems.Add , , Trim(rs3.Fields("Date_Entered"))
            li.ListSubItems.Add , , Trim(rs3.Fields("Qin"))
            'li.ListSubItems.Add , , Trim(rs3.Fields("0"))
            'li.ListSubItems.Add , , Trim(rs3.Fields("0"))
            rs3.MoveNext
        
        Wend
    End With
    rs3.Close
    Set rs3 = Nothing
ListView1.View = lvwReport

End Sub
Private Sub ComputeTotal()
  If cboVehicle <> "" Then
        sql = ""
        sql = "set dateformat dmy SELECT sum(PaidAmount) FROM d_MilkControl WHERE (DispDate = '" & txtdateenterered & "') and vehicleno='" & cboVehicle & "'"
        Set rs = oSaccoMaster.GetRecordset(sql)
        If Not rs.EOF Then
         If rs.Fields(0) <> "" Then
          lbltotal = rs.Fields(0)
         Else
        lbltotal = "0"
         End If
        Else
        lbltotal = "0"
        End If
        
         sql = ""
         sql = "set dateformat dmy SELECT sum(DispQnty) FROM d_MilkControl WHERE  (DispDate = '" & txtdateenterered & "') and vehicleno='" & cboVehicle & "'"
         Set rst = oSaccoMaster.GetRecordset(sql)
        If Not rst.EOF Then
         If rst.Fields(0) <> "" Then
            lbltotalkg = rst.Fields(0)
          Else
            lbltotalkg = "0"
          End If
        Else
         lbltotalkg = "0"
        End If
        
         sql = ""
         sql = "set dateformat dmy SELECT isnull(sum(Kgs),0) FROM d_OutletVehicle WHERE  (Date = '" & txtdateenterered & "') and vehicle='" & cboVehicle & "'"
         Set rsg = oSaccoMaster.GetRecordset(sql)
        If Not rsg.EOF Then
         If rsg.Fields(0) <> "" Then
            lbltotalkg = rsg.Fields(0) + lbltotalkg
            '+ lbltotalkg
          Else
           lbltotalkg = "0"
          End If
        Else
         lbltotalkg = "0"
        End If
 Else
    lbltotal = "0"
 End If
End Sub
Private Sub Command1_Click()
frmSearchGLAccounts.Show vbModal, Me
    If Continue Then
        If SearchValue <> "" Then
            txtCrAccNo = SearchValue
            SearchValue = ""
        End If
    End If
End Sub

Private Sub Form_Load()
txtdateenterered = Format(Get_Server_Date, "dd/mm/yyyy")
txtdateenterered.MaxDate = Format(Get_Server_Date, "dd/mm/yyyy")

'DTPMilkDate = Format(Get_Server_Date, "dd/mm/yyyy")
'DTPMilkDate.MaxDate = Format(Get_Server_Date, "dd/mm/yyyy")
'DTPcomplaintperiod = DTPMilkDate
txtDispatch = 0
txtsta = 0
txtAmount = 0
cmdreport.Enabled = True
lbltotal.Visible = False
lbltotalkg.Visible = False
Label34.Visible = False
Label35.Visible = False
chkoutletre.value = 0
chkoutletre.Visible = False
SSTab1_DblClick
loadBranchesTypes
NAMES
loadReg
fghj
 Label32.Visible = False
txtdateenterered_Change
End Sub

Private Sub ListView8_dbclick()

End Sub



Private Sub ListView8_DblClick()
txtTCode.Text = ListView8.SelectedItem
txtTCode_Validate True
End Sub

Private Sub Picture1_Click()
Me.MousePointer = vbHourglass
         frmSearchMilkControl.Show vbModal
        txtRefNo = sel
        txtRefNo_Validate True
        Me.MousePointer = 0
End Sub

Private Sub Picture2_Click()
Me.MousePointer = vbHourglass
        frmSearchGLAccounts.Show vbModal, Me
    If Continue Then
        If SearchValue <> "" Then
            lbldracc = SearchValue
            SearchValue = ""
        End If
    End If
    Me.MousePointer = 0
End Sub

Private Sub Picture3_Click()
Me.MousePointer = vbHourglass
        frmSearchGLAccounts.Show vbModal, Me
    If Continue Then
        If SearchValue <> "" Then
            lblcracc = SearchValue
            SearchValue = ""
        End If
    End If
    Me.MousePointer = 0
End Sub

Private Sub Picture4_Click()
Me.MousePointer = vbHourglass
        frmSearchGLAccounts.Show vbModal, Me
    If Continue Then
        If SearchValue <> "" Then
            lblcracc = SearchValue
            SearchValue = ""
        End If
    End If
    Me.MousePointer = 0
End Sub

Private Sub Picture5_Click()
Me.MousePointer = vbHourglass
         frmSearchDebtors.Show vbModal
        txtTCode = sel
        txtTCode_Validate True
        Me.MousePointer = 0
End Sub

Private Sub txtEmail_Change()
'txtPAddress.SetFocus
End Sub

Private Sub txtId_Change()
'txtPhone.SetFocus
End Sub

Private Sub txtIntake_Change()
'cmdsave3.SetFocus
End Sub

Private Sub txtNames_Change()
'txtId.SetFocus
End Sub

Private Sub txtPAddress_Change()
'txtTown.SetFocus
End Sub

Private Sub txtPhone_Change()
'txtEmail.SetFocus
End Sub

Private Sub txtprice_Change()
If Trim(txtprice) = "0.00" Then
txtprice = ""
End If
'txtVehicle.SetFocus
End Sub

Private Sub txtTCode_Validate(Cancel As Boolean)
Dim a As Boolean, b As Integer
Set rs = New ADODB.Recordset
sql = "d_sp_Selectdebtors '" & txtTCode & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
If Not IsNull(rs.Fields(0)) Then txtNames = rs.Fields(0)
If Not IsNull(rs.Fields(1)) Then txtId = rs.Fields(1)
'If Not IsNull(rs.Fields(2)) Then cbolocation = rs.Fields(2)
If Not IsNull(rs.Fields(3)) Then txtdateenterered = rs.Fields(3)
If Not IsNull(rs.Fields(4)) Then txtEMail = rs.Fields(4)
If Not IsNull(rs.Fields(5)) Then txtPhone = rs.Fields(5)
If Not IsNull(rs.Fields(6)) Then txtTown = rs.Fields(6)
If Not IsNull(rs.Fields(7)) Then txtPAddress = rs.Fields(7)
'If Not IsNull(rs.Fields(8)) Then txtsubsidy = Format(rs.Fields(8), "#0.00")
If Not IsNull(rs.Fields(9)) Then txtVehicle = rs.Fields(2)
'If Not IsNull(rs.Fields(10)) Then cboBName = rs.Fields(10)
'If Not IsNull(rs.Fields(11)) Then cboBBranch = rs.Fields(11)
If Not IsNull(rs.Fields(12)) Then a = rs.Fields(12)
'If Not IsNull(rs.Fields(13)) Then cboBranch = rs.Fields(13)
If Not IsNull(rs.Fields(14)) Then txtprice = Format(rs.Fields(14), "#0.00")
If Not IsNull(rs.Fields(15)) Then txtDrAccNo = rs.Fields(15)
If Not IsNull(rs.Fields(16)) Then txtCrAccNo = rs.Fields(16)
'If Not IsNull(rs.Fields(17)) Then txtcessrate = rs.Fields(17)
'If Not IsNull(rs.Fields(18)) Then txtcessdebit = rs.Fields(18)
If Not IsNull(rs.Fields(19)) Then txtsta = rs.Fields(19)

If a = True Then
chkActive = vbChecked
Else
chkActive = vbUnchecked
End If
cmdedit.Enabled = True
cmdsave.Enabled = False
End If
End Sub

Private Sub SSTab1_DblClick()
'Private Sub SSTab1_DblClick()
    cboVehicle.Clear
    Set rst = New Recordset
    Dim cn As Connection
    Set cn = New ADODB.Connection
    Provider = "MAZIWA"
    'Provider = cnn', "atm", "atm"
    cn.Open Provider, "atm", "atm"
    Set rst = New Recordset
    sql = "Select distinct(Locations) from   d_Debtors order by Locations"
    'Select distinct(Locations) from   d_Debtors
    rst.Open sql, cn, adOpenKeyset, adLockOptimistic
    While Not rst.EOF
    cboVehicle.AddItem rst.Fields(0)
    rst.MoveNext
    Wend
NAMES
'End Sub
End Sub
Private Sub NAMES()
'Private Sub SSTab1_DblClick()
    cboNames.Clear
    Set rst = New Recordset
    Dim cn As Connection
    Set cn = New ADODB.Connection
    Provider = "MAZIWA"
    'Provider = cnn
    cn.Open Provider, "atm", "atm"
    Set rst = New Recordset
    sql = "Select DName from   d_Debtors where Locations='" & cboVehicle & "' order by DName"
    'Select distinct(Locations) from   d_Debtors
    rst.Open sql, cn, adOpenKeyset, adLockOptimistic
    While Not rst.EOF
    cboNames.AddItem rst.Fields(0)
    rst.MoveNext
    Wend
End Sub

Private Sub Text3_Change()

End Sub

Private Sub txtAmount_Change()
If txtAmount = "" Then
 txtAmount = 0
 'txtIntake.SetFocus
 Exit Sub
End If

End Sub

Private Sub txtdateenterered_Change()
    Set rs = New ADODB.Recordset
    sql = "d_sp_DailyTotal '" & txtdateenterered & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If Not IsNull(rs.Fields(0)) Then
    txtIntake = Format(rs.Fields(0), "#0.00")
    Else
    txtIntake = "0"
    End If
    loadBranchesTypes
    loadReg
    NAMES1
    loadAssignments
End Sub

Private Sub Debtorsgl()
   lblcracc = "A004"
    Set rsd = New Recordset
    Dim cn As Connection
    Set cn = New ADODB.Connection
    Provider = "MAZIWA"
    'Provider = cnn
    cn.Open Provider, "atm", "atm"
    Set rsd = New Recordset
    sql = "Select  AccCr from  d_Debtors where DCode='" & txtdcode & "' and Locations='" & cboVehicle & "'"
    'Select distinct(Locations) from   d_Debtors
    rsd.Open sql, cn, adOpenKeyset, adLockOptimistic
    While Not rsd.EOF
    lbldracc = rsd.Fields(0)
    rsd.MoveNext
    Wend
    End Sub

Private Sub txtRefNo_Validate(Cancel As Boolean)
On Error GoTo ErrorHandler
'SELECT TOP 300 DispDate,RefNo,dcode,DispQnty,Price,InQnty,Variance FROM dbo.d_MilkControl"
If Trim(txtRefNo) = "" Then
Exit Sub
End If
 Set rs = oSaccoMaster.GetRecordset("SELECT DispDate,dcode,DispQnty,Price,InQnty,sum(Variance) FROM d_MilkControl WHERE RefNo = '" & txtRefNo & "'")
 
'txtdcode_Validate True
Exit Sub
ErrorHandler:
MsgBox err.description
End Sub
Private Sub txtdcode_Validate(Cancel As Boolean)
Set rs = oSaccoMaster.GetRecordset("SELECT dname,Price,accdr,acccr,drcess,crcess,capp,crate FROM d_Debtors WHERE DCode = '" & txtdcode & "'")
If rs.RecordCount > 0 Then
If Not IsNull(rs.Fields(1)) Then Price = rs.Fields(1)
If Not IsNull(rs.Fields(0)) Then lblDebtors = rs.Fields(0)
If Not IsNull(rs.Fields(2)) Then Label10 = rs.Fields(2)
If Not IsNull(rs.Fields(3)) Then Label11 = rs.Fields(3)
If Not IsNull(rs.Fields(4)) Then cessdr = rs.Fields(4)
If Not IsNull(rs.Fields(5)) Then cesscr = rs.Fields(5)
If Not IsNull(rs.Fields(6)) Then capp = Abs(rs.Fields(6))
If Not IsNull(rs.Fields(7)) Then crate = rs.Fields(7)
txtamountp = txtDispatch * rs.Fields(1)
Else
'lblDebtors = ""
End If
End Sub
Private Sub txtDispatch_Change()
On Error GoTo ErrorHandler
'txtDipping = txtDispatch
If txtDispatch = "" Then
txtDispatch = "0"
End If
'**************PRICE***************'
If chkoutletre.value = 0 Then
Set rs = oSaccoMaster.GetRecordset("SELECT dname,Price,accdr,acccr,drcess,crcess,capp,crate FROM d_Debtors WHERE DCode = '" & txtdcode & "'")
If rs.RecordCount > 0 Then
If Not IsNull(rs.Fields(1)) Then Price = rs.Fields(1)
txtamountp = txtDispatch * rs.Fields(1)
'txtamount.SetFocus
Else
'lblDebtors = ""
End If
Else
'txtamountp = txtDispatch * rs.Fields(1)
End If
'****************END********************'

Exit Sub
ErrorHandler:
MsgBox err.description
End Sub

Private Sub txtRefNo_Change()
On Error GoTo ErrorHandler
'SELECT TOP 300 DispDate,RefNo,dcode,DispQnty,Price,InQnty,Variance FROM dbo.d_MilkControl"
If Trim(txtRefNo) = "" Then
Exit Sub
End If
' Set rs = oSaccoMaster.GetRecordset("SELECT DispDate,dcode,DispQnty,Price,InQnty,sum(Variance) FROM d_MilkControl WHERE RefNo = '" & txtRefNo & "'")
'
' If rs.RecordCount > 0 Then
'    DTPDispatchDate = rs.Fields(0)
'    txtDispatch = rs.Fields(2)
'    'txtDipping = txtDispatch
'    txtIntake = rs.Fields(4)
'    'txtVariance = rs.Fields(5)
'    Label10 = rs.Fields(1)
'
'    cmdEdit.Enabled = True
'Else
'    cmdEdit.Enabled = False
'
'End If
'txtdcode_Validate True
Exit Sub
ErrorHandler:
MsgBox err.description
End Sub

Private Sub txtTown_Change()
txtprice.SetFocus
End Sub

Private Sub txtVehicle_Click()
'fghj
emb
End Sub
Private Sub txtVehicle_Change()
'fghj
emb
End Sub
Private Sub emb()
Provider = "MAZIWA"
Set cn = New ADODB.Connection
cn.Open Provider, "atm", "atm"
Set rst = New ADODB.Recordset
rst.Open sql, cn
'If rs.EOF Then
Set rst = oSaccoMaster.GetRecordset("select DCode,AccDr, AccCr from d_Debtors where Locations ='" & txtVehicle & "'")
If Not rst.EOF Then
'txtdcode = rst.Fields("DCode")
txtDrAccNo = rst.Fields("AccDr")
txtCrAccNo = rst.Fields("AccCr")
Else
MsgBox "Records of this vehicle " & txtVehicle & " does not exist.Please fill the ledgers manually"
End If
End Sub
