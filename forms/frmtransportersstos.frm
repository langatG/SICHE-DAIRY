VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmtransportersstos 
   Caption         =   "Transporters Standing order"
   ClientHeight    =   4320
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9765
   LinkTopic       =   "Form1"
   ScaleHeight     =   4320
   ScaleWidth      =   9765
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtmaximumamount 
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """Kshs ""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   0
      EndProperty
      Height          =   285
      Left            =   4080
      TabIndex        =   24
      Text            =   "0"
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CommandButton cmdpostall 
      BackColor       =   &H00FFFF80&
      Caption         =   "Post All Suppliers"
      Height          =   375
      Left            =   6240
      TabIndex        =   23
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFF80&
      Caption         =   "Stop Standing Order"
      Height          =   375
      Left            =   4560
      TabIndex        =   22
      Top             =   3360
      Width           =   1695
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Top             =   3360
      Width           =   855
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   2880
      TabIndex        =   9
      Top             =   3360
      Width           =   855
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   2040
      TabIndex        =   8
      Top             =   3360
      Width           =   855
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   375
      Left            =   1200
      TabIndex        =   7
      Top             =   3360
      Width           =   855
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   3720
      TabIndex        =   6
      Top             =   3360
      Width           =   855
   End
   Begin VB.ComboBox cboDeductionType 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "frmtransportersstos.frx":0000
      Left            =   240
      List            =   "frmtransportersstos.frx":0019
      TabIndex        =   5
      Top             =   1680
      Width           =   1815
   End
   Begin VB.TextBox txtTNames 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2160
      TabIndex        =   4
      Top             =   840
      Width           =   5775
   End
   Begin VB.TextBox txtAmount 
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """Kshs ""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   0
      EndProperty
      Height          =   285
      Left            =   2280
      TabIndex        =   3
      Text            =   "0"
      Top             =   1680
      Width           =   1575
   End
   Begin VB.TextBox txtTCode 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   1575
   End
   Begin VB.PictureBox Picture5 
      Height          =   255
      Left            =   1800
      Picture         =   "frmtransportersstos.frx":0056
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   1
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox txtremarks 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   2760
      Width           =   4815
   End
   Begin MSComCtl2.DTPicker DTPEndDate 
      Height          =   255
      Left            =   7800
      TabIndex        =   11
      Top             =   1680
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      _Version        =   393216
      Format          =   122617857
      CurrentDate     =   40096
   End
   Begin MSComCtl2.DTPicker DTPStartDate 
      Height          =   255
      Left            =   5760
      TabIndex        =   12
      Top             =   1680
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   450
      _Version        =   393216
      Format          =   122617857
      CurrentDate     =   40096
   End
   Begin MSComCtl2.DTPicker DTPDDate 
      Height          =   255
      Left            =   6120
      TabIndex        =   13
      Top             =   360
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      _Version        =   393216
      Format          =   122617857
      CurrentDate     =   40096
   End
   Begin VB.Label Label9 
      BackColor       =   &H80000004&
      Caption         =   "Maximum Amount"
      Height          =   255
      Left            =   4080
      TabIndex        =   25
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "End Date"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   7800
      TabIndex        =   21
      Top             =   1320
      Width           =   675
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Start Date"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5760
      TabIndex        =   20
      Top             =   1320
      Width           =   720
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Type of standing order"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   19
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Name"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2160
      TabIndex        =   18
      Top             =   480
      Width           =   420
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Date of standing order"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6120
      TabIndex        =   17
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Transporter Code"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   16
      Top             =   480
      Width           =   1230
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Amount"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2280
      TabIndex        =   15
      Top             =   1320
      Width           =   540
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2760
      Width           =   975
   End
End
Attribute VB_Name = "frmtransportersstos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmdNew_Click()
txtamount.Locked = False
txtTCode.Locked = False


txtamount = ""
txtTCode = ""

cboDeductionType.Locked = False

cboDeductionType = ""
cmdnew.Enabled = False
cmdsave.Enabled = True
End Sub

Private Sub cmdsave_Click()
On Error GoTo ErrorHandler
Enddate = DateSerial(year(DTPDDate), month(DTPDDate) + 1, 1 - 1)
Set rs = oSaccoMaster.GetRecordset("d_sp_IsClosed '" & Enddate & "'")
Dim DESCR As String
If Not rs.EOF Then
    MsgBox "The period " & Enddate & " has been closed by " & rs.Fields(0)
    Exit Sub
End If

Set cn = New ADODB.Connection
''sql = "d_sp_TransDeduct '" & txtTCode & "','" & DTPDDate & "','" & cboDeductionType & "'," & txtamount & ",'" & dtpstartdate & "','" & dtpenddate & "','" & year(dtpenddate) & "','" & User & "','" & txtremarks & "'"
' sql = "set dateformat dmy insert into d_Transport_Deduc (TransCode, TDate_Deduc,[Description],Amount,StartDate,enddate,yyear,AuditID,remarks)"
'    sql = sql & " values ('" & txtTCode & "','" & DTPDDate & "','" & cboDeductionType & "'," & txtAmount & ",'" & DTPStartDate & "','" & DTPEndDate & "','" & year(DTPEndDate) & "','" & User & "','" & txtremarks & "')"
'
'oSaccoMaster.ExecuteThis (sql)
''Form_Load
'If Not Trim(txtremarks) = "" Then
'sql = "SELECT TOP 1 id From d_Transport_Deduc ORDER BY id DESC"
'Set rs2 = oSaccoMaster.GetRecordset(sql)
'
'If Not rs2.EOF And Not rs2.Fields(0) = "" Then
'oSaccoMaster.ExecuteThis ("UPDATE d_Transport_Deduc SET remarks= '" & txtremarks & "' WHERE Id = " & rs2.Fields(0))
'End If
'
'End If
'//check if it had another deductions of the same nature.
sql = ""
sql = "select description from d_transport_standingorder where TransCode='" & txtTCode & "' and description ='" & cboDeductionType & "'"
Set rst = oSaccoMaster.GetRecordset(sql)
If rst.EOF Then
'//Update deductions
    Set cn = New ADODB.Connection
    sql = "set dateformat dmy INSERT INTO d_transport_standingorder"
    sql = sql & "           (TransCode, Date_Deduc, Description, Amount, MaxAmount, Period, StartDate, EndDate, auditid,  yyear, Remarks)"
    sql = sql & "  VALUES     ('" & txtTCode & "','" & DTPDDate & "','" & cboDeductionType & "'," & txtamount & "," & txtmaximumamount & ",'" & Format(DTPDDate, "mmm-YYYY") & "','" & dtpStartDate & "','" & DTPEndDate & "','" & User & "'," & year(DTPEndDate) & ",'" & txtremarks & "')"
    oSaccoMaster.ExecuteThis (sql)
    
    'sql = "d_sp_SupplierDeduct " & txtSNo & ",'" & DTPDDeduction & "','" & DESCR & "'," & txtAmount & ",'" & DTPStartDate & "','" & DTPEndDate & "'," & Year(DTPEndDate) & ",'" & User & "','" & txtRemarks & "'"
    'oSaccoMaster.ExecuteThis (sql)
    Else
    MsgBox "The Deduction Code Has Been Defined for this Member", vbInformation, "Standing Order Set Up"
    Exit Sub
End If



txtamount = ""
txtTCode = ""
txtTCode_Validate True
txtTCode.SetFocus

MsgBox "Records successively updated."
Exit Sub
ErrorHandler:
MsgBox err.description

End Sub

Private Sub DTPDDate_Change()
dtpStartDate = DateSerial(year(DTPDDate), month(DTPDDate), 1)
DTPEndDate = DateSerial(year(DTPDDate), month(DTPDDate) + 1, 1 - 1)
End Sub

Private Sub Form_Load()
Dim myclass As cdbase
DTPDDate = Format(Get_Server_Date, "dd/mm/yyyy")
txtamount.Locked = True
txtTNames.Locked = True
txtTCode.Locked = True

'DTPDDate.MaxDate = DTPDDate
'DTPDDate.MinDate = DTPDDate


txtamount = ""
txtTNames = ""
txtTCode = ""

cboDeductionType.Locked = True

cboDeductionType = ""

cmdnew.Enabled = True
cmdDelete.Enabled = False
cmdEdit.Enabled = False
cmdsave.Enabled = False

    

    Set myclass = New cdbase

    Provider = myclass.OpenCon
    
    Set cn = CreateObject("adodb.connection")
    
   cn.Open Provider, "atm", "atm"

Set rs = CreateObject("adodb.recordset")
    
    rs.Open "SELECT Description FROM d_DCodes", cn
    
    If rs.EOF Then Exit Sub
    
    With rs
        
        While Not .EOF
         
         cboDeductionType.AddItem rs.Fields("Description")
         
         .MoveNext
        
        Wend
    
    End With
    
    DTPEndDate = DateSerial(year(DTPDDate), month(DTPDDate) + 1, 1 - 1)
    dtpStartDate = DateSerial(year(DTPDDate), month(DTPDDate), 1)
End Sub

Private Sub Picture5_Click()
        Me.MousePointer = vbHourglass
        frmSearchPTransporter.Show vbModal
        txtTCode = sel
        txtTCode_Validate True
        Me.MousePointer = 0
End Sub
Private Sub txtTCode_Validate(Cancel As Boolean)
Set rs = New ADODB.Recordset
sql = "d_sp_SelectTrans '" & txtTCode & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
If Not IsNull(rs.Fields(0)) Then txtTNames = rs.Fields(0)
Else
txtTNames = ""
End If
End Sub

