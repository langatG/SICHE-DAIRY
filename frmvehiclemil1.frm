VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmvehiclemil1 
   Caption         =   "INDIVIDUAL REPORT"
   ClientHeight    =   3465
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6555
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   6555
   StartUpPosition =   3  'Windows Default
   Begin MSComCtl2.DTPicker DTPickerFrom 
      Height          =   375
      Left            =   2400
      TabIndex        =   8
      Top             =   240
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      Format          =   264962049
      CurrentDate     =   44498
   End
   Begin VB.CommandButton frmprinty 
      Caption         =   "Print Individual Report"
      Height          =   495
      Left            =   2760
      TabIndex        =   5
      Top             =   2640
      Width           =   1815
   End
   Begin VB.ComboBox cbovb 
      Height          =   315
      Left            =   2400
      TabIndex        =   0
      Top             =   1440
      Width           =   2175
   End
   Begin MSComCtl2.DTPicker DTPMilkDate 
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   840
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      MouseIcon       =   "frmvehiclemil1.frx":0000
      CalendarBackColor=   8454016
      Format          =   264962049
      CurrentDate     =   40095
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   3210
      Width           =   6555
      _ExtentX        =   11562
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   2645
            MinWidth        =   2645
            Text            =   "USER : Birgen Gideon K."
            TextSave        =   "USER : Birgen Gideon K."
            Object.ToolTipText     =   "EASYMA User"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2470
            MinWidth        =   2470
            Picture         =   "frmvehiclemil1.frx":08DA
            Text            =   "DATE : 07/12/2009"
            TextSave        =   "DATE : 07/12/2009"
            Object.ToolTipText     =   "Today's Date"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "5:33 PM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Enabled         =   0   'False
            TextSave        =   "NUM"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ProgressBar prgStatus 
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   2040
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Milk Date From:"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Vehicle No:"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   0
      TabIndex        =   3
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Milk Date To:"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   2175
   End
End
Attribute VB_Name = "frmvehiclemil1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbovb_Change()
calc
End Sub
Private Sub cbovb_Click()
calc
End Sub
Private Sub calc()
Dim DTPfrom, DTPto As Date
DTPfrom = DateSerial(Year(DTPickerFrom), month(DTPickerFrom), 1)
DTPto = DateSerial(Year(DTPMilkDate), month(DTPMilkDate) + 1, 1 - 1)
sql = "delete from d_MilkVehiclecalc"
oSaccoMaster.Execute (sql)
Set rs = New Recordset

sql = ""
sql = "set dateformat dmy SELECT    isnull(count( Vehicle),0) From d_MilkVehicle  WHERE Date>='" & DTPfrom & "' and Date<='" & DTPto & "' and Vehicle='" & cbovb & "'  "
Set rsy = cn.Execute(sql)
Dim b As Double
b = rsy.Fields(0)

prgStatus.max = 100
prgStatus.Min = 0
I = 0
sql = ""
sql = "set dateformat dmy select Vehicle, Quantity, Actual, Varriance, Date from d_MilkVehicle where Date>='" & DTPfrom & "' and Date<='" & DTPto & "' and Vehicle='" & cbovb & "' order by Date asc"
Set rs = oSaccoMaster.GetRecordset(sql)
While Not rs.EOF
   I = I + 1
prgStatus = Round((I / b) * 100, 0)

 sql = ""
 sql = "set dateformat dmy insert into d_MilkVehiclecalc( Vehicle, Quantity, Actual, Varriance, Date)values('" & rs.Fields(0) & "','" & rs.Fields(1) & "','" & rs.Fields(2) & "','" & rs.Fields(3) & "','" & rs.Fields(4) & "')"
 oSaccoMaster.Execute (sql)
 rs.MoveNext
Wend
MsgBox "Data ready to be printed", vbExclamation
prgStatus = 0
Exit Sub
End Sub
Private Sub Form_Load()
DTPMilkDate = Format(Get_Server_Date, "dd/mm/yyyy")
DTPMilkDate.MaxDate = Format(Get_Server_Date, "dd/mm/yyyy")
'DTPcomplaintperiod1 = DTPMilkDate
With StatusBar1.Panels
    .Item(1).Text = "USER : " & username
    .Item(2).Text = "DATE : " & Format(Get_Server_Date, "dd/mm/yyyy")

End With
    NAMES3
End Sub
Private Sub NAMES3()
'Private Sub SSTab1_DblClick()
    cbovb.Clear
    Set rst = New Recordset
    Dim cn As Connection
    Set cn = New ADODB.Connection
    Provider = "MAZIWA"
    'Provider = cnn
    cn.Open Provider, "atm", "atm"
    Set rst = New Recordset
    'sql = "Select distinct(Locations) from   d_Debtors order by Locations"
     sql = "Select distinct(Vehicle) from d_VehicleTill WHERE Vehicle not like'%heh%' order by Vehicle"
    'Select distinct(Locations) from   d_Debtors
    rst.Open sql, cn, adOpenKeyset, adLockOptimistic
    While Not rst.EOF
    cbovb.AddItem rst.Fields(0)
    rst.MoveNext
    Wend
End Sub

Private Sub frmprinty_Click()
reportname = "d_vehicledeliveryper1.rpt"
Show_Sales_Crystal_Report "", reportname, ""
End Sub
