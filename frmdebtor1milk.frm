VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmdebtor1milk 
   Caption         =   "Vehicle Milk Sales "
   ClientHeight    =   8400
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9060
   LinkTopic       =   "Form1"
   ScaleHeight     =   8400
   ScaleWidth      =   9060
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdreport 
      Caption         =   "Print Report"
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
      Left            =   2520
      TabIndex        =   3
      Top             =   1320
      Width           =   2175
   End
   Begin VB.CommandButton cmdload 
      Caption         =   "Load Transactions"
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
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   1935
   End
   Begin VB.ComboBox cbomilkve 
      Height          =   315
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   2655
   End
   Begin MSComctlLib.ProgressBar prgStatus 
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   6375
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   11245
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
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1185
   End
End
Attribute VB_Name = "frmdebtor1milk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdreport_Click()
    reportname = "d_debtorsacrude.rpt"
    Show_Sales_Crystal_Report "", reportname, ""
End Sub

Private Sub cmdLoad_Click()
Dim datefrom, datetoday As Date
 If cbomilkve = "" Then
  MsgBox "Please select the Vehicle Number"
 Exit Sub
 End If
 Set rss = oSaccoMaster.GetRecordset("delete from d_Debtorenquiry")
 sql = ""
 sql = "Select distinct(DCode) from d_MilkControl where vehicleno='" & cbomilkve & "' order by DCode"
 Set rss = oSaccoMaster.GetRecordset(sql)
 If rss.EOF Then
   MsgBox "No records for this vehicle " & cbomilkve & ""
   Exit Sub
  End If
  Dim a As Double
  sql = "Select count(distinct(DCode)) from d_MilkControl where vehicleno='" & cbomilkve & "'"
  Set rsv = oSaccoMaster.GetRecordset(sql)
  a = rsv.Fields(0)
  prgStatus.Visible = True
  prgStatus.max = 100
  prgStatus.Min = 0
  I = 0
    While Not rss.EOF
        I = I + 1
        prgStatus = Round((I / a) * 100, 0)
        
      sql = "Select DCode, DName from d_Debtors where DCode='" & rss.Fields(0) & "'"
      Set rsk = oSaccoMaster.GetRecordset(sql)
      Dim name, code, vehicle As String
      Dim sum, Resum As Double
      name = rsk.Fields("DName")
      code = rsk.Fields("DCode")
      vehicle = cbomilkve
      sum = 0
      Resum = 0
       sql = "Select DispDate from d_MilkControl where DCode='" & rss.Fields(0) & "' and vehicleno='" & cbomilkve & "' order by DispDate asc"
       Set rst = oSaccoMaster.GetRecordset(sql)
       While Not rst.EOF
       datefrom = rst.Fields(0)
       'datetoday = Format(Get_Server_Date, "MMMM yyyy")
        'While Not datefrom > datetoday
         Startdate = DateSerial(Year(datefrom), Month(datefrom), 1)
         Enddate = DateSerial(Year(datefrom), Month(datefrom) + 1, 1 - 1)
         sql = ""
         Set rsd = oSaccoMaster.GetRecordset("d_sp_MilkControlDebtor '" & rss.Fields(0) & "','" & datefrom & "','" & datefrom & "','" & cbomilkve & "'")
         
         sql = ""
         Set rsv = oSaccoMaster.GetRecordset("d_sp_MilkControlDebtorsum '" & rss.Fields(0) & "','" & datefrom & "','" & cbomilkve & "'")
         
         Resum = (rsv.Fields(2) - rsv.Fields(1))
         sum = rsd.Fields(2) - rsd.Fields(1)
         ''''insert
         sql = ""
         sql = "set dateformat dmy insert into d_Debtorenquiry(DCode, Name, Date, Quantity, Amount, Paid, Balance, Vehicle,Recurr) Values('" & code & "','" & name & "','" & datefrom & "','" & rsd.Fields("qun") & "','" & rsd.Fields("amt") & "','" & rsd.Fields("pay") & "','" & sum & "','" & vehicle & "','" & Resum & "')"
         Set rsg = oSaccoMaster.GetRecordset(sql)
         'datefrom = Enddate + 1
         Resum = 0
         rst.MoveNext
        Wend
       'End If
     rss.MoveNext
    Wend
    MsgBox "Loading completed successfully"
    loadBranchesTypes
End Sub
Public Sub loadBranchesTypes()
    
    With ListView1
    
        .ListItems.Clear
        
        .ColumnHeaders.Clear
    
    End With

    Set rs2 = CreateObject("adodb.recordset")
    
    sql = "set dateformat dmy Select DCode, Name, Date, Quantity, Amount, Paid, Balance, Vehicle,Recurr from d_Debtorenquiry where Vehicle='" & cbomilkve & "' order by Date DESC"
    Set rs = oSaccoMaster.GetRecordset(sql)
    
    Set rs2 = CreateObject("adodb.recordset")
    
    Set clsClass = New cdbase
    
    Provider = clsClass.OpenCon
    
    Set cn = CreateObject("adodb.connection")
    
   cn.Open Provider, "atm", "atm"
    
    rs2.Open sql, cn
    
    With ListView1
        
        .ColumnHeaders.Add , , "DCode"
        .ColumnHeaders.Add , , "Name"
        .ColumnHeaders.Add , , "Date"
        .ColumnHeaders.Add , , "Quantity"
        .ColumnHeaders.Add , , "Amount"
        .ColumnHeaders.Add , , "Paid"
        .ColumnHeaders.Add , , "Balance"
        .ColumnHeaders.Add , , "Vehicle"
        .ColumnHeaders.Add , , "Accrude Amount"
        While Not rs2.EOF
        
            Set li = .ListItems.Add(, , Trim(rs2.Fields("DCode")))
            
            li.ListSubItems.Add , , Trim(rs2.Fields("Name"))
            li.ListSubItems.Add , , Trim(rs2.Fields("Date"))
            li.ListSubItems.Add , , Trim(rs2.Fields("Quantity"))
            li.ListSubItems.Add , , Trim(rs2.Fields("Amount"))
            li.ListSubItems.Add , , Trim(rs2.Fields("Paid"))
            li.ListSubItems.Add , , Trim(rs2.Fields("Balance"))
            li.ListSubItems.Add , , Trim(rs2.Fields("Vehicle"))
            li.ListSubItems.Add , , Trim(rs2.Fields("Recurr"))
            rs2.MoveNext
        
        Wend
        
    End With
    
    rs2.Close
    
    Set rs2 = Nothing
    
ListView1.View = lvwReport

End Sub
Private Sub Form_Load()
NAMES
End Sub
Private Sub NAMES()
'Private Sub SSTab1_DblClick()
    cbomilkve.Clear
    Set rst = New Recordset
    Dim cn As Connection
    Set cn = New ADODB.Connection
    Provider = "MAZIWA"
    'Provider = cnn
    cn.Open Provider, "atm", "atm"
    Set rst = New Recordset
    sql = "Select distinct(vehicleno) from d_MilkControl order by vehicleno"
    'Select distinct(Locations) from   d_Debtors
    rst.Open sql, cn, adOpenKeyset, adLockOptimistic
    While Not rst.EOF
    cbomilkve.AddItem rst.Fields(0)
    rst.MoveNext
    Wend
End Sub
