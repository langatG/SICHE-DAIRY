VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmstockbalance 
   Caption         =   "Stock Balance Processing"
   ClientHeight    =   3720
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7665
   LinkTopic       =   "Form1"
   ScaleHeight     =   3720
   ScaleWidth      =   7665
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Periodic stock Balance"
      Height          =   375
      Left            =   4800
      TabIndex        =   9
      Top             =   2880
      Width           =   2415
   End
   Begin MSComCtl2.DTPicker DTPS 
      Height          =   375
      Left            =   2160
      TabIndex        =   8
      Top             =   720
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      Format          =   115605505
      CurrentDate     =   42913
   End
   Begin MSComCtl2.DTPicker DTPend 
      Height          =   375
      Left            =   4560
      TabIndex        =   6
      Top             =   1920
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      Format          =   115605505
      CurrentDate     =   42910
   End
   Begin MSComCtl2.DTPicker DTPstart 
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   1920
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      Format          =   115605505
      CurrentDate     =   42910
   End
   Begin MSComCtl2.DTPicker DTPdate 
      Height          =   375
      Left            =   5040
      TabIndex        =   2
      Top             =   720
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      Format          =   115605505
      CurrentDate     =   42910
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Process Balance"
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "To date"
      Height          =   375
      Left            =   4920
      TabIndex        =   7
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "From date"
      Height          =   255
      Left            =   2160
      TabIndex        =   5
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Sales Period"
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label lblstartdate 
      AutoSize        =   -1  'True
      Caption         =   "As at Date"
      Height          =   195
      Left            =   720
      TabIndex        =   1
      Top             =   840
      Width           =   750
   End
End
Attribute VB_Name = "frmstockbalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim lastdate As Date
Dim lastdateofsale As Date
Dim pcode As String
Dim Quant As Double
Dim q As Double
Dim pp As Double
Dim ss As Double
Dim dy As Integer
Dim grade As String
Dim rsd As New ADODB.Recordset
Dim rsds As New ADODB.Recordset
Dim rsdD As New ADODB.Recordset
Dim rsbal As New ADODB.Recordset
Dim ball As Double
sql = ""
sql = "DELETE FROM ag_sales"
sql = ""
sql = "DELETE FROM ag_sales1"
Set rs = oSaccoMaster.GetRecordset(sql)
sql = ""
'sql = "set dateformat dmy SELECT     r.P_code,p.p_name, SUM(Qua) AS Quantity From ag_Receipts r inner join ag_stockbalance p on r.p_code=p.p_code WHERE   r.T_Date >= '" & DTPstdate & "' and r.T_Date <= '" & DTPedate & "'  GROUP BY r.P_code, p.productname ORDER BY r.P_code asc"
sql = "set dateformat dmy SELECT    SUM(openningstock) AS Quantity, p_code, productname From ag_stockbalance WHERE     transdate >= '" & DTPS & "' and  transdate <= '" & DTPdate & "' GROUP BY p_code, productname"
'sql = "set dateformat dmy SELECT    SUM(openningstock) AS Quantity, p_code, productname From ag_stockbalance WHERE    transdate <= '" & DTPdate & "' GROUP BY p_code, productname"

'sql = "set dateformat dmy SELECT distinct p_code ,openningstock,ProductName From ag_stockbalance WHERE     transdate >= '" & DTPS & "' and  transdate <= '" & DTPdate & "' ORDER BY p_code"
'sql = "set dateformat dmy SELECT TOP(1)    stockbalance, p_code, productname From ag_stockbalance WHERE     transdate >= '" & DTPS & "' and  transdate <= '" & DTPdate & "'"

Set rs = oSaccoMaster.GetRecordset(sql)
While Not rs.EOF
pcode = rs!p_code
quantity = rs!quantity
pname = rs!ProductName
'pp = rs!Pprice
'ss = rs!sprice
'If rs!p_code = 52 Then
'MsgBox "here"
'End If

'select pcode,ldate,dy,auditdate,audit,grade from ag_paging
sql = ""
sql = "set dateformat dmy insert into  ag_sales (pcode,pname,Quantity,Bprice,sprice)"
sql = sql & "values('" & pcode & "','" & pname & "','" & quantity & "','" & pp & "','" & sp & "') "
oSaccoMaster.ExecuteThis (sql)

rs.MoveNext
Wend


sql = ""

sql = "SELECT  Pcode, pname, Quantity, Sales, Bal FROM         ag_sales"
Set rsds = oSaccoMaster.GetRecordset(sql)
While Not rsds.EOF

pcodes = rsds!pcode
sql = "set dateformat dmy SELECT     SUM(Qua) AS quon, P_code From ag_Receipts WHERE T_Date >= '" & DTPstart & "' and T_Date <= '" & DTPend & "' and p_code='" & pcodes & "' GROUP BY P_code"
Set rsd = oSaccoMaster.GetRecordset(sql)
If Not rsd.EOF Then
quan = rsd!quon
Else
quan = 0
End If

'sql = "set dateformat dmy SELECT     SUM(Qua) AS quon, P_code From ag_Receipts WHERE T_Date >= '" & DTPstart & "' and T_Date <= '" & DTPend & "' and p_code='" & pcodes & "' GROUP BY P_code"
sql = "SELECT     p_code, p_name, pprice, sprice From ag_Products where  p_code='" & pcodes & "' "
Set rsdD = oSaccoMaster.GetRecordset(sql)
If Not rsdD.EOF Then
ss = rsdD!sprice
pp = rsdD!Pprice
Else
ss = 0
pp = 0
End If
'sql = "set dateformat dmy SELECT     SUM(changeinstock) AS Quantity, p_code, productname From ag_stockbalance  WHERE transdate >= '" & DTPstart & "' and transdate <= '" & DTPend & "' and p_code='" & pcodes & "'GROUP BY p_code, productname"
'Set rsbal = oSaccoMaster.GetRecordset(sql)
'If Not rsbal.EOF Then
'q = rsbal!quantity
'Else
'q = 0
'End If
'pcode = rs!P_code
sql = "set dateformat dmy SELECT top 1 max(R_id) as R_id,s_bal  From ag_Receipts  WHERE t_date  <= '" & DTPdate & "' and   p_code='" & pcode & "'group by  S_Bal"
Set rsbal = oSaccoMaster.GetRecordset(sql)
If Not rsbal.EOF Then
'q = rsbal!quantity
ball = rsbal!s_bal
rsbal.MoveNext
'q = 0
End If
'pname = rs!ProductName
'select pcode,ldate,dy,auditdate,audit,grade from ag_paging
'sql = ""

sql = ""
sql = "update ag_sales1 set  Bal= '" & ball & "' where pcode='" & pcode & "' "
oSaccoMaster.ExecuteThis (sql)
rsds.MoveNext
'rsbal.MoveNext
Wend
'//give him the report here
'agrovetagingreport

MsgBox "Records successfully done", vbInformation
reportname = "cummulative sales.rpt"

 Show_Sales_Crystal_Report STRFORMULA, reportname, ""
'//we look for receipts tables
'//get the number of days
'/// insert into the number of days
'//give us a report

End Sub

Private Sub Command2_Click()
'Dim lastdate As Date
'Dim lastdateofsale As Date
'Dim pcode As String
'Dim Quant As Double
'Dim q As Double
'Dim pp As Double
'Dim ss As Double
'Dim dy As Integer
'Dim grade As String
'Dim rsd As New ADODB.Recordset
'Dim rsds As New ADODB.Recordset
'Dim rsdD As New ADODB.Recordset
'Dim rsbal As New ADODB.Recordset
'Dim balz As Double
'Dim pcodd As Double
'Dim bp As Double
'Dim sp As Double
'sql = ""
'sql = "DELETE FROM ag_sales1"
'Set rs = oSaccoMaster.GetRecordset(sql)
'sql = ""
'sql = "set dateformat dmy SELECT  *   From ag_products order by  p_code"
'
'Set rs = oSaccoMaster.GetRecordset(sql)
'While Not rs.EOF
'pcode = rs!p_code
'quantity = rs!o_bal
'pname = rs!p_name
'bp = rs!Pprice
'sp = rs!sprice
'sql = ""
'sql = "set dateformat dmy insert into  ag_sales1 (pcode,pname,Quantity,Bprice,sprice,Bal)"
'sql = sql & "values('" & pcode & "','" & pname & "','" & quantity & "','" & bp & "','" & sp & "','0') "
'
'oSaccoMaster.ExecuteThis (sql)
'oSaccoMaster.ExecuteThis = "set dateformat dmy SELECT top 1 max(R_id) as R_id,s_bal  From ag_Receipts  WHERE t_date  <= '" & DTPdate & "' and   p_code='" & pcode & "'group by  S_Bal"
'oSaccoMaster.ExecuteThis = "update ag_sales1 set  Bal= '" & s_bal & "' where pcode='" & pcode & "' "
'rs.MoveNext
'Wend
End Sub

