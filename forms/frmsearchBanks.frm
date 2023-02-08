VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmsearchBanks 
   Caption         =   "Finder Banks"
   ClientHeight    =   4530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6945
   Icon            =   "frmsearchBanks.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4530
   ScaleWidth      =   6945
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   4455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6855
      Begin VB.TextBox txtValue 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2040
         TabIndex        =   10
         Top             =   360
         Width           =   2415
      End
      Begin VB.CommandButton cmdRef 
         Height          =   495
         Left            =   3840
         Picture         =   "frmsearchBanks.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Refresh"
         Top             =   3840
         Width           =   495
      End
      Begin VB.ComboBox cboField 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         ItemData        =   "frmsearchBanks.frx":074C
         Left            =   120
         List            =   "frmsearchBanks.frx":074E
         TabIndex        =   3
         Top             =   345
         Width           =   1695
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "SELECT"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5400
         TabIndex        =   2
         Top             =   3795
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancel 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3240
         Picture         =   "frmsearchBanks.frx":0750
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Cancel"
         Top             =   3840
         Width           =   495
      End
      Begin MSComctlLib.ListView lstSearch 
         Height          =   2535
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "Search Accounts Record"
         Top             =   1200
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   4471
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   16711680
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label Label1 
         Caption         =   "Search Field"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Records Found"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label lblRecords 
         Height          =   255
         Left            =   1800
         TabIndex        =   7
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Criteria"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   6
         Top             =   120
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmsearchBanks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SelectedDsn As String
'Dim CConnect As USACBOSA6.cdbase
Dim rst As Recordset
Dim rst1 As Recordset
Dim li As ListItem
Dim recordfound As String


Private Sub cboCrieria_Click()
      Call SRefresh
End Sub

Private Sub cboCrieria_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cboField_Change()
    txtValue_Change
End Sub

Private Sub cboField_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cmdCancel_Click()
    sel = ""
    frmsearchBanks.Visible = False
    
End Sub

Private Sub cmdFind_Click()
End Sub

Private Sub cmdRef_Click()
    Call SRefresh
End Sub

Private Sub cmdSelect_Click()
    sel = ""
        If lstSearch.ListItems.count > 0 Then
        sel = lstSearch.SelectedItem
        Me.Visible = False
        Continue = True
    Else
        MsgBox "No record selected.", vbExclamation
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    With frmsearchBanks.lstSearch
        .ListItems.Clear
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "Accno", 1500
        .ColumnHeaders.Add , , "Bank Code", 1500
        .ColumnHeaders.Add , , "Bank Name", 3000
        .ColumnHeaders.Add , , "Branch", 2800
        .View = lvwReport
        .Gridlines = True
    End With
    With frmsearchBanks.cboField
        .AddItem "Accno"
        .AddItem "BankCode"
        .AddItem "BankName"
        .AddItem "BranchName"
    End With
    cboField.Text = cboField.List(2)
    Me.Top = (Screen.Height - Height) / 2
    Me.Left = (Screen.Width - Width) / 1.4
    cmdFind_Click
End Sub
Sub Load(sql As String)
    lstSearch.ListItems.Clear
    Set rst = oSaccoMaster.GetRecordSet(sql)
    With rst
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                Set li = lstSearch.ListItems.Add(, , .Fields(0))
                li.SubItems(1) = .Fields(1) & ""
                li.SubItems(2) = .Fields(2) & ""
                li.SubItems(3) = .Fields(3) & ""
                .MoveNext
            Loop
        End If
        .Close
    End With

End Sub



Private Sub lstSearch_DblClick()
    If lstSearch.ListItems.count > 0 Then
        SearchValue = lstSearch.SelectedItem.Text
        Continue = True
    Else
        SearchValue = ""
        Continue = False
    End If
    Unload Me
End Sub


Public Sub SRefresh()
     cmdFind_Click
End Sub

Private Sub txtValue_Click()
    txtValue_Change
End Sub

Private Sub txtValue_keypress(KeyAscii As Integer)
    Select Case KeyAscii
      Case Asc("A") To Asc("Z")
      Case Asc("a") To Asc("z")
      Case Asc("0") To Asc("9")
      Case Asc("/")
      Case Asc("-")
      Case Asc("(")
      Case Asc(")")
      Case Asc(" ")
      Case Asc(".")
      'Case Asc("'")
      Case Is = 8
      
      Case Else
        Beep
        KeyAscii = 0
    End Select
End Sub
Private Sub txtValue_Change()
    On Error Resume Next
    If cboField.Text = "" Then
        MsgBox "The Selection Cretaria is not complete", vbExclamation
        Exit Sub
    End If
    
    'Biuild the select String
    If cboField.Text <> "" Then
        sql = "Select accno,bankCode,BankName,Branchname from banks where " & cboField.Text & " like '%" & txtValue.Text & "%' order by bankcode"
    Else
        sql = "Select accno,bankCode,BankName,Branchname from banks order by bankcode"
    End If
    
    Load sql

End Sub
