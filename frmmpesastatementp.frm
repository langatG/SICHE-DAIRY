VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmmpesastatementp 
   Caption         =   "Payment Statement"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4875
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4875
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdclose 
      Caption         =   "Close"
      Height          =   375
      Left            =   3960
      TabIndex        =   10
      Top             =   2520
      Width           =   735
   End
   Begin VB.CommandButton cmdreport 
      Caption         =   "Report"
      Height          =   375
      Left            =   2400
      TabIndex        =   9
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmdload 
      Caption         =   "Load"
      Height          =   375
      Left            =   1200
      TabIndex        =   8
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton cmdnew 
      Caption         =   "New"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   2520
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      Begin VB.ComboBox cmdpaymentMode 
         Height          =   315
         Left            =   1920
         TabIndex        =   6
         Top             =   1320
         Width           =   1575
      End
      Begin VB.ComboBox cmdBranch 
         Height          =   315
         Left            =   1920
         TabIndex        =   5
         Top             =   840
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1920
         TabIndex        =   4
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   112852993
         CurrentDate     =   44744
      End
      Begin VB.Label Label3 
         Caption         =   "Payment Mode:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Branch:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Date:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmmpesastatementp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmdload_Click()
On Error GoTo ErrorHandler
Dim Startdate, Enddate As Date
Startdate = DateSerial(Year(DTPicker1), month(DTPicker1), 1)
Enddate = DateSerial(Year(DTPicker1), month(DTPicker1) + 1, 1 - 1)
If cmdBranch = "" Then
 MsgBox "PLEASE SELECT THE BRANCH."
 cmdBranch.SetFocus
 Exit Sub
End If
If cmdpaymentMode = "" Then
   MsgBox "PLEASE SELECT THE PAYMENT MODE."
   cmdpaymentMode.SetFocus
 Exit Sub
End If
sql = "set dateformat dmy delete d_PayrollEnquiry where Branch='" & cmdBranch & "' and Mode='" & cmdpaymentMode & "' and Date>='" & Startdate & "' and Date<='" & Enddate & "'"
oSaccoMaster.GetRecordset (sql)

sql = ""
sql = "set dateformat dmy SELECT SNo,Names From d_Suppliers where Location='" & cmdBranch & "'"
Set rss = oSaccoMaster.GetRecordset(sql)
While Not rss.EOF

'sql = "set dateformat dmy SELECT TDeductions, KgsSupplied, GPay, NPay,EndofPeriod From d_PayrollCopy where SNo='" & rss.Fields(0) & "' and sortstatus='" & cmdpaymentMode & "' and EndofPeriod>='" & Startdate & "' and EndofPeriod<='" & Enddate & "' "

 sql = "set dateformat dmy SELECT TDeductions, KgsSupplied, GPay, NPay,EndofPeriod From d_PayrollCopy where SNo='" & rss.Fields(0) & "' and sortstatus='" & cmdpaymentMode & "' and EndofPeriod>='" & Startdate & "' and EndofPeriod<='" & Enddate & "' "
 Set rs = oSaccoMaster.GetRecordset(sql)
 While Not rs.EOF
 ''insert the new table
   sql = "INSERT INTO d_PayrollEnquiry(sno,name,Date,kgs, Gross,Deductions,Netpay,Branch,Mode)"
   sql = sql & "Values ('" & rss.Fields(0) & "','" & rss.Fields(1) & "','" & rs.Fields(4) & "','" & rs.Fields(1) & "','" & rs.Fields(2) & "','" & rs.Fields(0) & "','" & rs.Fields(3) & "','" & cmdBranch & "','" & cmdpaymentMode & "')"
   oSaccoMaster.GetRecordset (sql)
  rs.MoveNext
 Wend
 
 rss.MoveNext
Wend

 MsgBox "Records loaded successfuly"
Exit Sub
ErrorHandler:
MsgBox err.description
End Sub

Private Sub cmdnew_Click()
cmdBranch = ""
cmdpaymentMode = ""
loadBranches
loadPayment
End Sub

Private Sub cmdreport_Click()
    reportname = "midmonthmpesastatement.rpt"
    Show_Sales_Crystal_Report "", reportname, ""
End Sub

Private Sub Form_Load()
DTPicker1 = Format(Get_Server_Date, "dd/mm/yyyy")
loadBranches
loadPayment
End Sub
Private Sub loadBranches()
    Set rst = New Recordset
    Dim cn As Connection
    Set cn = New ADODB.Connection
    Provider = "MAZIWA"
    cn.Open Provider, "atm", "atm"
    Set rst = New Recordset
    cmdBranch.Clear
    sql = "Select Bname from   d_Branch order by Bname"
    rst.Open sql, cn, adOpenKeyset, adLockOptimistic
    While Not rst.EOF
    cmdBranch.AddItem rst.Fields(0)
    rst.MoveNext
    Wend
End Sub
Private Sub loadPayment()
    Set rst = New Recordset
    Dim cn As Connection
    Set cn = New ADODB.Connection
    Provider = "MAZIWA"
    cn.Open Provider, "atm", "atm"
    Set rst = New Recordset
    cmdpaymentMode.Clear
    sql = "Select distinct sortstatus from d_PayrollCopy where sortstatus IS NOT NULL order by sortstatus"
    Set rst = oSaccoMaster.GetRecordset(sql)
    'rst.Open sql, cn, adOpenKeyset, adLockOptimistic
    While Not rst.EOF
    cmdpaymentMode.AddItem rst.Fields(0)
    rst.MoveNext
    Wend
End Sub
