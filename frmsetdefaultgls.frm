VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmsetdefaultgls 
   Caption         =   "Set Default Gls"
   ClientHeight    =   6465
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8070
   LinkTopic       =   "Form1"
   ScaleHeight     =   6465
   ScaleWidth      =   8070
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmddelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   4800
      TabIndex        =   15
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmdedit 
      Caption         =   "Edit"
      Height          =   375
      Left            =   2040
      TabIndex        =   14
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   375
      Left            =   840
      TabIndex        =   13
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "Save"
      Height          =   375
      Left            =   3240
      TabIndex        =   11
      Top             =   2520
      Width           =   1095
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
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7815
      Begin VB.ComboBox cbosuppliers 
         Height          =   315
         ItemData        =   "frmsetdefaultgls.frx":0000
         Left            =   2040
         List            =   "frmsetdefaultgls.frx":000D
         TabIndex        =   10
         Top             =   600
         Width           =   3255
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
         Left            =   2070
         TabIndex        =   8
         Top             =   1800
         Width           =   3225
      End
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   285
         Left            =   480
         TabIndex        =   7
         Top             =   1800
         Width           =   315
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
         Left            =   795
         TabIndex        =   6
         Top             =   1800
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
         Left            =   2055
         TabIndex        =   5
         Top             =   1320
         Width           =   3225
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
         Left            =   780
         TabIndex        =   4
         Top             =   1320
         Width           =   1080
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "..."
         Height          =   285
         Left            =   480
         TabIndex        =   3
         Top             =   1320
         Width           =   300
      End
      Begin VB.Label Label1 
         Caption         =   "Where To Affect"
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
         TabIndex        =   12
         Top             =   600
         Width           =   1815
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
         TabIndex        =   2
         Top             =   1320
         Width           =   1335
      End
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
         TabIndex        =   1
         Top             =   1920
         Width           =   1215
      End
   End
   Begin MSComctlLib.ListView ListView8 
      Height          =   3375
      Left            =   120
      TabIndex        =   9
      Top             =   3000
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   5953
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
End
Attribute VB_Name = "frmsetdefaultgls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public edit As Integer

Private Sub cmddelete_Click()
On Error GoTo SysError
    sql = "delete from GLSetDefaultGls Where Dr='" & txtDrAccNo & "' and Cr='" & txtCrAccNo & "' and Affect='" & cbosuppliers & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
MsgBox "Deleted Successfully"
loadReg
cbosuppliers = ""
txtDrAccNo = ""
txtCrAccNo = ""
 Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub cmdedit_Click()
edit = 1
End Sub

Private Sub cmdNew_Click()
edit = 0
cbosuppliers = ""
txtDrAccNo = ""
txtCrAccNo = ""
cmdedit.Enabled = False
End Sub

Private Sub cmdsave_Click()
On Error GoTo SysError
    'Purchases
    'Payables
    'Join
    If cbosuppliers = "" Then
        MsgBox "Please select Where to Affect..."
      cbosuppliers.SetFocus
       Exit Sub
    End If
    If txtDrAccNo = "" Then
        MsgBox "Please enter the DrAccNo..."
     txtDrAccNo.SetFocus
      Exit Sub
    End If
    If txtCrAccNo = "" Then
        MsgBox "Please enter CrAccNo..."
     txtCrAccNo.SetFocus
      Exit Sub
    End If
    If txtCrAccNo = txtDrAccNo Then
        MsgBox "CrAccNo cannot be the same with DrAccNo"
     txtCrAccNo.SetFocus
      Exit Sub
    End If
        sql = "select * from GLSetDefaultGls Where Affect='" & cbosuppliers & "'"
        Set rs = oSaccoMaster.GetRecordset(sql)
        If Not rs.EOF Then
            MsgBox "Where it Affect Already Exist"
        Exit Sub
        End If
    If edit = 0 Then
            sql = "select * from GLSetDefaultGls Where Dr='" & txtDrAccNo & "' and Cr='" & txtCrAccNo & "' and Affect='" & cbosuppliers & "'"
            Set rs = oSaccoMaster.GetRecordset(sql)
            If Not rs.EOF Then
                MsgBox "Entry Already Exist"
            Exit Sub
            End If
            
            sql = "insert into GLSetDefaultGls(Dr,Cr,Affect,Audituser, Audidatetime) values('" & txtDrAccNo & "','" & txtCrAccNo & "','" & cbosuppliers & "','" & User & "','" & Now & "')"
            Set rss = oSaccoMaster.GetRecordset(sql)
    Else
       Exit Sub
    End If
    MsgBox "Save Successfully"
    loadReg
    edit = 0
    cmdedit.Enabled = False
     Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
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
loadReg
End Sub
Public Sub loadReg()
    
    With ListView8
    
        .ListItems.Clear
        
        .ColumnHeaders.Clear
    
    End With

    Set rs2 = CreateObject("adodb.recordset")
    
    sql = "Select Dr, Cr, Affect from GLSetDefaultGls order by Audidatetime"
    Set rs = oSaccoMaster.GetRecordset(sql)
    
    Set rs2 = CreateObject("adodb.recordset")
    
    Set clsClass = New cdbase
    
    Provider = clsClass.OpenCon
    
    Set cn = CreateObject("adodb.connection")
    
   cn.Open Provider, "atm", "atm"
    
    rs2.Open sql, cn
    
    With ListView8
        
        .ColumnHeaders.Add , , "Dr"
        .ColumnHeaders.Add , , "Cr"
        .ColumnHeaders.Add , , "Affect"
        While Not rs2.EOF
        
            Set li = .ListItems.Add(, , Trim(rs2.Fields("Dr")))
            li.ListSubItems.Add , , Trim(rs2.Fields("Cr"))
            li.ListSubItems.Add , , Trim(rs2.Fields("Affect"))
            rs2.MoveNext
        
        Wend
        
    End With
    
    rs2.Close
    
    Set rs2 = Nothing
    
ListView8.View = lvwReport

End Sub

Private Sub ListView8_DblClick()
On Error GoTo SysError
    cmdedit.Enabled = True
    txtDrAccNo = ListView8.SelectedItem
    txtCrAccNo = ListView8.SelectedItem.SubItems(1)
    cbosuppliers = ListView8.SelectedItem.SubItems(2)
    cmdedit.Enabled = True
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
