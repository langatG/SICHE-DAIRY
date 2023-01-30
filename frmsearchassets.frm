VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmsearchassets 
   Caption         =   "Finder  Assets"
   ClientHeight    =   4800
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7005
   Icon            =   "frmsearchassets.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4800
   ScaleWidth      =   7005
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   4695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6855
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
         Picture         =   "frmsearchassets.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Cancel"
         Top             =   4080
         Width           =   495
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
         TabIndex        =   8
         Top             =   4080
         Width           =   1335
      End
      Begin VB.Frame Frame2 
         Height          =   1215
         Left            =   3840
         TabIndex        =   4
         Top             =   120
         Width           =   2895
         Begin VB.TextBox txtTo 
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
            Height          =   315
            Left            =   120
            TabIndex        =   7
            Top             =   720
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.TextBox txtFrom 
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
            Height          =   315
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   2175
         End
         Begin VB.CommandButton cmdFind 
            Enabled         =   0   'False
            Height          =   375
            Left            =   2400
            Picture         =   "frmsearchassets.frx":0544
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.ComboBox cboCrieria 
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
         ItemData        =   "frmsearchassets.frx":0646
         Left            =   2280
         List            =   "frmsearchassets.frx":065F
         TabIndex        =   3
         Top             =   360
         Width           =   1455
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
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   2055
      End
      Begin VB.CommandButton cmdRef 
         Height          =   495
         Left            =   3840
         Picture         =   "frmsearchassets.frx":0683
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Refresh"
         Top             =   4080
         Width           =   495
      End
      Begin MSComctlLib.ListView lstSearch 
         Height          =   2535
         Left            =   120
         TabIndex        =   10
         Top             =   1440
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   4471
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
         Left            =   2280
         TabIndex        =   14
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label lblRecords 
         Height          =   255
         Left            =   1800
         TabIndex        =   13
         Top             =   1320
         Width           =   855
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
         TabIndex        =   12
         Top             =   1080
         Width           =   1575
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
         TabIndex        =   11
         Top             =   120
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmsearchassets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SelectedDsn As String
Dim CConnect As MAZIWA.CConnect
Dim Rst As Recordset
Dim Rst1 As Recordset
Dim li As ListItem
Dim recordfound As String
Dim cn As Connection


Private Sub cboCrieria_Click()
  If cboCrieria.Text = "Between" Then
    txtTo.Visible = True
  Else
    txtTo.Visible = False
  End If
End Sub

Private Sub cboCrieria_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cboField_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cmdcancel_Click()
    sel = ""
    frmsearchre.Visible = False
    
End Sub

Private Sub cmdFind_Click()
Dim Find As Long
Dim li As ListItem
Dim Field As String
Set cn = New Connection
Set Rst = New Recordset
Field = cboField.Text
'Find = Button.Index
lstSearch.ListItems.Clear
  
'If Find = 1 Then
'CConnect.cnnConnect

If Not cboField.Text = "" Then
    If Not cboCrieria.Text = "" Then
        If Not cboCrieria.Text = "Between" And Not cboCrieria.Text = "Like" Then
        If cboField.Text = "R_No" Then
        sql = "SELECT     * FROM         assets where " & cboField.Text & "" & cboCrieria.Text & "" & txtFrom.Text & ""
 
        Else
            sql = "SELECT     * FROM         assets Where " & cboField.Text & "" & cboCrieria.Text & "'" & txtFrom.Text & "'"
            End If
           CConnect.Openrs
            
            With rs
                If .RecordCount > 0 Then
                    .MoveFirst
                    Do While Not .EOF
 Set li = frmsearchassets.lstSearch.ListItems.Add(, , !AssetsNo)
                li.SubItems(1) = !assetserialno & ""
                li.SubItems(2) = !assetsname & ""
                li.SubItems(3) = !assettype & ""
                
                        .MoveNext
                    Loop
                End If
            End With
            
            Set rs = Nothing
            
        ElseIf cboCrieria.Text = "Like" Then
         'If cboField.Text = "R_No" Then
            sql = "SELECT     * FROM         assets order by assetsno"
            'Else
            'end if
           CConnect.Openrs
            
            With rs
                If .RecordCount > 0 Then
                    .MoveFirst
                    Do While Not .EOF
                     If cboField.Text = "R_No" Then
                        .Find "" & cboField.Text & " " & cboCrieria.Text & " " & txtFrom.Text & "%", , adSearchForward
                        Else
                         .Find "" & cboField.Text & " " & cboCrieria.Text & " '" & txtFrom.Text & "%'", , adSearchForward

                        End If
                        If Not .EOF Then
Set li = frmsearchassets.lstSearch.ListItems.Add(, , !AssetsNo)
                li.SubItems(1) = !assetserialno & ""
                li.SubItems(2) = !assetsname & ""
                 li.SubItems(3) = !assettype & ""
                            .MoveNext
                        End If
                        
                    Loop
                End If
            End With
            
            Set rs = Nothing
            
            
        Else
            If cboField.Text = "Amount" Then
                sql = "SELECT     * FROM         assets where " & cboField.Text & " " & cboCrieria & " " & txtFrom.Text & " And " & txtTo.Text & ""
               CConnect.Openrs
            Else
                sql = "SELECT     * FROM         assets  where " & cboField.Text & " " & cboCrieria & " '" & txtFrom.Text & "' And '" & txtTo.Text & "'"
               CConnect.Openrs
            End If
            
            With rs
                If .RecordCount > 0 Then
                    .MoveFirst
                    Do While Not .EOF
Set li = frmsearchassets.lstSearch.ListItems.Add(, , !AssetsNo)
                li.SubItems(1) = !assetserialno & ""
                li.SubItems(2) = !assetsname & ""
                 li.SubItems(3) = !assettype & ""
                        .MoveNext
                    Loop
                End If
            End With
            
            Set rs = Nothing
            
        End If
        
    Else
        MsgBox "Select the search criteria.", vbExclamation
    End If
Else
    MsgBox "Select the search field.", vbExclamation
End If

'Set cnnPayroll = Nothing


End Sub

Private Sub cmdRef_Click()
    Call SRefresh
    
End Sub

Private Sub cmdSelect_Click()
sel = ""
    If lstSearch.ListItems.Count > 0 Then
        sel = lstSearch.SelectedItem
        Me.Visible = False
       ' Me.Unload Me
    Else
        MsgBox "No record selected.", vbExclamation
    End If
Unload Me
End Sub

Private Sub Form_Load()


    'DSource = cnn

    With frmsearchassets.lstSearch
        .ListItems.Clear
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "Assets No", 1500
        .ColumnHeaders.Add , , "Assets Serial No", 2800
        .ColumnHeaders.Add , , "Asset Name", 2800
        .ColumnHeaders.Add , , "Asset Type", 2800
        .View = lvwReport
        .GridLines = True
    End With
    
    With frmsearchassets.cboField
        .AddItem "AssetsNo"
        .AddItem "Assetserialno"
        .AddItem "AssetsName"
        .AddItem "AssetType"
    End With
    Set CConnect = New MAZIWA.CConnect
    CConnect.cnnConnect
    sql = "SELECT     * FROM         assets order by assetsno"
    CConnect.Openrs
    
    With rs
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
Set li = frmsearchassets.lstSearch.ListItems.Add(, , !AssetsNo)
                li.SubItems(1) = !assetserialno & ""
                li.SubItems(2) = !assetsname & ""
                li.SubItems(3) = !assettype & ""
                .MoveNext
                
            Loop
            
        End If
        .Close
    End With
    
    Set rs = Nothing
    'Set cnnPayroll = Nothing
        
    cboCrieria.Text = cboCrieria.List(0)
    cboField.Text = cboField.List(0)
    'txtFrom.SetFocus
    
    Me.Top = (Screen.Height - Height) / 2
    Me.Left = (Screen.Width - Width) / 1.4
    
End Sub



Private Sub txtFrom_Change()
    If txtFrom.Text = "" Then
        cmdFind.Enabled = False
    Else
        cmdFind.Enabled = True
    End If
    
End Sub





Public Sub SRefresh()
lstSearch.ListItems.Clear

    CConnect.cnnConnect
   sql = "SELECT     * FROM         assets order by assetsno"
    CConnect.Openrs
    
       
    With rs
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
   Set li = frmsearchassets.lstSearch.ListItems.Add(, , !AssetsNo)
                li.SubItems(1) = !assetserialno & ""
                li.SubItems(2) = !assetsname & ""
                 li.SubItems(3) = !assettype & ""
                .MoveNext
                
            Loop
            
        End If
        .Close
    End With
    
    Set rs = Nothing
    'Set cnnPayroll = Nothing
    
    txtFrom.Text = ""
    txtTo.Text = ""
    cboCrieria.Text = "="
    cboField.Text = "R_No"
End Sub

Private Sub txtFrom_KeyPress(KeyAscii As Integer)
    If Len(Trim(txtFrom.Text)) > 20 Then
        Beep
        MsgBox "Can't enter more than 20 characters", vbExclamation
        KeyAscii = 8
    End If
  Select Case KeyAscii
    'Case Asc("vbBack")
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

Private Sub txtTo_KeyPress(KeyAscii As Integer)
    If Len(Trim(txtTo.Text)) > 20 Then
        Beep
        MsgBox "Can't enter more than 20 characters", vbExclamation
        KeyAscii = 8
    End If
  Select Case KeyAscii
    'Case Asc("vbBack")
    Case Asc("A") To Asc("Z")
    Case Asc("a") To Asc("z")
    Case Asc("0") To Asc("9")
    Case Asc("/")
    Case Asc("-")
    Case Asc("(")
    Case Asc(")")
    Case Asc(" ")
    Case Asc(".")
    Case Is = 8
    
    Case Else
    Beep
    KeyAscii = 0
  End Select
End Sub

