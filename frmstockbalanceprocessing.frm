VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmstockbalance 
   Caption         =   "Stock Balance Processing"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5370
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   5370
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Process Balance"
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   2040
      Width           =   1575
   End
   Begin MSComCtl2.DTPicker DTPedate 
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   840
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      _Version        =   393216
      Format          =   123731969
      CurrentDate     =   42680
   End
   Begin VB.Label lbledate 
      Caption         =   "End date"
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   840
      Width           =   1095
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
Dim dy As Integer
Dim grade As String
Dim curamt As Double
Dim pprice As Double
Dim sprice As Double

Dim rsd As New ADODB.Recordset
sql = ""
sql = "DELETE FROM ag_sales"
Set rs = oSaccoMaster.GetRecordset(sql)
sql = ""
sql = "set dateformat dmy SELECT     SUM(Qin) AS Quantity, p_code,p_name,branch From ag_Products3 WHERE   Date_Entered <= '" & DTPedate & "' GROUP BY p_code,p_name,branch"
'sql = "set dateformat dmy SELECT     SUM(Qin) AS Quantity, p_code,p_name,branch,pprice, sprice From ag_Products3 WHERE   Date_Entered <= '" & DTPedate & "' GROUP BY p_code,p_name,branch,pprice, sprice"
'sql = "set dateformat dmy SELECT     r.P_code,p.p_name, SUM(Qua) AS Quantity From ag_Receipts r inner join ag_products p on r.p_code=p.p_code WHERE   r.T_Date >= '" & DTPstdate & "' and r.T_Date <= '" & DTPedate & "'  GROUP BY r.P_code, p.P_name ORDER BY r.P_code asc"
'sql = "set dateformat dmy SELECT     r.P_code,p.productname, SUM(changeinstock) AS Quantity From ag_stockbalance p inner join ag_Receipts r on r.p_code=p.p_code WHERE   r.T_Date >= '" & DTPstdate & "' and r.T_Date <= '" & DTPedate & "'  GROUP BY r.P_code, p.productname ORDER BY r.P_code asc"

Set rs = oSaccoMaster.GetRecordset(sql)
While Not rs.EOF
pcode = rs!p_code
Quantity = IIf(IsNull(rs!Quantity), 0, rs!Quantity)
pname = rs!p_name
Branch = rs!Branch
'pprice = rs!pprice
'sprice = rs!sprice

sql = "set dateformat dmy SELECT     SUM(Qua) AS Qty, p_code,Branch From ag_Receipts WHERE   T_Date <= '" & DTPedate & "' and p_code='" & rs!p_code & "'and branch='" & rs!Branch & "'  GROUP BY p_code,Branch"
Set rsd = oSaccoMaster.GetRecordset(sql)
If Not rsd.EOF Then
curamt = IIf(IsNull(rs!Quantity), 0, rs!Quantity) - IIf(IsNull(rsd!qty), 0, rsd!qty)
Else
curamt = IIf(IsNull(rs!Quantity), 0, rs!Quantity)
End If
'curamt = IIf(IsNull(rs!Quantity), 0, rs!Quantity) - IIf(IsNull(rsd!qty), 0, rsd!qty)
'select pcode,ldate,dy,auditdate,audit,grade from ag_paging
'sql = ""
sql = "set dateformat dmy insert into  ag_sales (pcode,pname,Quantity,branch)"
sql = sql & "values('" & pcode & "','" & pname & "','" & curamt & "','" & Branch & "') "
oSaccoMaster.ExecuteThis (sql)


rs.MoveNext
Wend
MsgBox "Records successfully done", vbInformation

'//give him the report here
'agrovetagingreport
reportname = "stobal.rpt"
'reportname = "evans.rpt"

 
 Show_Sales_Crystal_Report STRFORMULA, reportname, ""
'//we look for receipts tables
'//get the number of days
'/// insert into the number of days
'//give us a report

End Sub

Private Sub Form_Load()
DTPedate = Get_Server_Date
End Sub
