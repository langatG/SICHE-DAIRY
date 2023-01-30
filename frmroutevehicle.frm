VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmroutevehicle 
   Caption         =   "Route Collectors"
   ClientHeight    =   3840
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4785
   LinkTopic       =   "Form1"
   ScaleHeight     =   3840
   ScaleWidth      =   4785
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtBCode 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      TabIndex        =   6
      Top             =   1800
      Width           =   3375
   End
   Begin VB.TextBox txtBName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      TabIndex        =   5
      Top             =   2400
      Width           =   3375
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   3600
      TabIndex        =   0
      Top             =   2880
      Width           =   855
   End
   Begin MSComctlLib.ListView lvWBranch 
      Height          =   1695
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   2990
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Code"
      Height          =   195
      Left            =   480
      TabIndex        =   9
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Name"
      Height          =   195
      Left            =   480
      TabIndex        =   8
      Top             =   2400
      Width           =   420
   End
End
Attribute VB_Name = "frmroutevehicle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim myclass As cdbase
Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmddelete_Click()
Set myclass = New cdbase
Provider = myclass.OpenCon
Set cn = CreateObject("adodb.connection")
cn.Open Provider, "atm", "atm"
sql = "delete from d_RouteCollectors where Code='" & txtBCode & "'"

myclass.Delete sql
loadBranchesTypes
txtBCode = ""
txtBName = ""
End Sub

Private Sub cmdedit_Click()
txtBCode.Locked = False
txtBName.Locked = False
cmdNew.Enabled = False
cmdEdit.Enabled = False
End Sub

Private Sub cmdNew_Click()
txtBCode = ""
txtBName = ""
txtBCode.Locked = False
txtBName.Locked = False
cmdNew.Enabled = False
cmdEdit.Enabled = False
sql = ""
sql = "select count(Code) from d_RouteCollectors"
Set rs = oSaccoMaster.GetRecordset(sql)

If Not rs.EOF Then
txtBCode = rs.Fields(0) + 1
Else
txtBCode = 1
End If

End Sub

Public Sub loadBranchesTypes()
    
    With lvWBranch
    
        .ListItems.Clear
        
        .ColumnHeaders.Clear
    
    End With

    Set rs2 = CreateObject("adodb.recordset")
    
    sql = "Select * from d_RouteCollectors"
    
    Set rs2 = CreateObject("adodb.recordset")
    
    Set clsClass = New cdbase
    
    Provider = clsClass.OpenCon
    
    Set cn = CreateObject("adodb.connection")
    
   cn.Open Provider, "atm", "atm"
    
    rs2.Open sql, cn
    
    With lvWBranch
        
        .ColumnHeaders.Add , , "Code"
        .ColumnHeaders.Add , , "Name"
    
        While Not rs2.EOF
        
            Set li = .ListItems.Add(, , Trim(rs2.Fields("Code")))
            
            li.ListSubItems.Add , , Trim(rs2.Fields("Name"))
            
        
            
            rs2.MoveNext
        
        Wend
        
    End With
    
    rs2.Close
    
    Set rs2 = Nothing
    
lvWBranch.View = lvwReport

End Sub

Private Sub cmdsave_Click()
On Error GoTo ErrorHandler
If txtBCode = "" Then
MsgBox "Enter the Code", vbInformation
Exit Sub
End If
If txtBName = "" Then
MsgBox "Enter the Name", vbInformation
Exit Sub
End If
Set cn = New ADODB.Connection
sql = "d_sp_RouteCollectors '" & txtBCode & "','" & txtBName & "','" & User & "'"
oSaccoMaster.ExecuteThis (sql)
txtBCode = ""
txtBName = ""
txtBCode.Locked = True
txtBName.Locked = True
cmdNew.Enabled = True
cmdEdit.Enabled = False
'cmdsave.Enabled = True
loadBranchesTypes
MsgBox "Records successively updated."
Exit Sub
ErrorHandler:
MsgBox err.description
End Sub

Private Sub Form_Load()
txtBCode.Locked = True
txtBName.Locked = True
cmdDelete.Enabled = False
loadBranchesTypes
End Sub

Public Sub edit(selected As String)

Set myclass = New cdbase
Set cn = CreateObject("adodb.connection")
Provider = myclass.OpenCon
cn.Open Provider, "atm", "atm"
Set rs = CreateObject("adodb.recordset")
sql = "select * from d_RouteCollectors where Code='" & selected & "'"
rs.Open sql, cn
If Not rs.EOF Then
txtBCode = selected
txtBName = rs!name
End If
cmdDelete.Enabled = True

End Sub
Private Sub lvWBranch_DblClick()
cmdEdit.Enabled = True
edit lvWBranch.SelectedItem
End Sub

