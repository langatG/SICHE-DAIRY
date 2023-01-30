VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmImportCollector 
   Caption         =   "Collector List"
   ClientHeight    =   6930
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9180
   LinkTopic       =   "Form1"
   ScaleHeight     =   6930
   ScaleWidth      =   9180
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cbochangecollector 
      ForeColor       =   &H008080FF&
      Height          =   315
      Left            =   1200
      TabIndex        =   16
      Top             =   1080
      Width           =   2895
   End
   Begin VB.CommandButton cmdLoadInterest 
      Caption         =   "Load Interest"
      Height          =   375
      Left            =   11400
      TabIndex        =   8
      Top             =   5160
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdProcess 
      Caption         =   "Process"
      Height          =   375
      Left            =   11400
      TabIndex        =   7
      Top             =   5520
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Update jan loan"
      Height          =   495
      Left            =   11160
      TabIndex        =   6
      Top             =   4560
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdPost 
      Caption         =   "Save"
      Height          =   375
      Left            =   8040
      TabIndex        =   5
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmdSelectAll 
      Caption         =   "Select All"
      Height          =   375
      Left            =   8040
      TabIndex        =   4
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "Import"
      Height          =   375
      Left            =   11760
      TabIndex        =   3
      Top             =   1800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdSearchFile 
      Caption         =   "Search"
      Height          =   375
      Left            =   11760
      TabIndex        =   2
      Top             =   1440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtFile 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9960
      TabIndex        =   1
      Top             =   840
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Close"
      Height          =   375
      Left            =   8040
      TabIndex        =   0
      Top             =   3360
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker DTPTransdate 
      Height          =   375
      Left            =   5280
      TabIndex        =   9
      Top             =   1080
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      _Version        =   393216
      Format          =   121634817
      CurrentDate     =   43885
   End
   Begin MSComctlLib.ListView lvwMemberDetails 
      Height          =   4815
      Left            =   240
      TabIndex        =   10
      Top             =   1920
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   8493
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin MSComctlLib.ProgressBar importationProgress 
      Height          =   255
      Left            =   8880
      TabIndex        =   11
      Top             =   1800
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7440
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Collector Change Form"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   18
      Top             =   120
      Width           =   7215
   End
   Begin VB.Label Label4 
      Caption         =   "Change To"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000004&
      Caption         =   "Date:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      TabIndex        =   15
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   495
      Left            =   3960
      TabIndex        =   14
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Height          =   375
      Left            =   7680
      TabIndex        =   13
      Top             =   1200
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label lbSourceFile 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000004&
      Caption         =   "Source File:"
      Height          =   255
      Left            =   10680
      TabIndex        =   12
      Top             =   360
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "frmImportCollector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim month As Integer
 Dim Year As Integer
 
Private Sub loadImport()
On Error GoTo ErrorHandler
Dim Graders As String
Set rs = oSaccoMaster.GetRecordset("set dateformat dmy select ID,SNo, TransDate, QSupplied,RouteCol from  d_Milkintake where TransDate = '" & DTPTransdate & "' order by auditdatetime desc")
'Set rs = oSaccoMaster.GetRecordset(sql)
lvwMemberDetails.ListItems.Clear
If rs.RecordCount > 0 Then

With lvwMemberDetails
    
       .ListItems.Clear
    
        .ColumnHeaders.Clear

  End With

    With lvwMemberDetails
        .ColumnHeaders.Add , , "ID"
        .ColumnHeaders.Add , , "SNo"
        .ColumnHeaders.Add , , "Date"
        .ColumnHeaders.Add , , "kgs"
        .ColumnHeaders.Add , , "Collector"
    
        While Not rs.EOF
            If Not IsNull(rs.Fields("ID")) Then
            Set li = .ListItems.Add(, , Trim(rs.Fields("ID")))
            End If
            If Not IsNull(rs.Fields("SNo")) Then
            li.ListSubItems.Add , , Trim(rs.Fields("SNo"))
            End If
            If Not IsNull(rs.Fields("TransDate")) Then
            li.ListSubItems.Add , , Trim(rs.Fields("TransDate"))
            End If
            If Not IsNull(rs.Fields("QSupplied")) Then
            li.ListSubItems.Add , , Trim(rs.Fields("QSupplied"))
            End If
            If Not IsNull(rs.Fields("RouteCol")) Then
            li.ListSubItems.Add , , Trim(rs.Fields("RouteCol"))
            End If
            
         rs.MoveNext
        Wend
        
    End With
    
    rs.Close
    
    Set rs = Nothing
Else
    MsgBox "No Records this Date", vbInformation, Me.Caption
End If

    Exit Sub
ErrorHandler:
MsgBox err.description
End Sub
Private Sub cmdPost_Click()
On Error GoTo ErrorHandler
 Dim totals As Double
 Dim collector As String
 Dim Id As String
 Dim yYear As Integer
'lvwMemberDetails
  If lvwMemberDetails.ListItems.Count < 1 Then
     MsgBox "  No Records to be posted", vbInformation
        Exit Sub
  End If
    
  If cbochangecollector = "" Then
     MsgBox "Select Collector Name", vbInformation
        Exit Sub
  End If

  MyPos = importationProgress.value
       
  For I = lvwMemberDetails.ListItems.Count To 1 Step -1
        If lvwMemberDetails.ListItems.Item(I).Checked = True Then
          MyPos = MyPos - 1
           'importationProgress.value = MyPos
            Set li = lvwMemberDetails.ListItems(I)
               
                Id = lvwMemberDetails.ListItems(I)
                sno = lvwMemberDetails.ListItems(I).SubItems(1)
                collector = cbochangecollector

        sql = ""
        sql = "set dateformat dmy  update d_Milkintake"
        sql = sql & " set RouteCol = '" & collector & "'"
        sql = sql & " where ID= '" & Id & "'"
        cn.Execute sql
          
           lvwMemberDetails.ListItems.Remove (lvwMemberDetails.ListItems(I).Index)
        End If
  Next I
        
         MsgBox "Records Updated Successfully", vbInformation
         cmdSelectAll.Caption = "Select All"
         cbochangecollector = ""
         loadImport
    Exit Sub
ErrorHandler:
MsgBox err.description
End Sub

Private Sub cmdSearchFile_Click()
With CommonDialog1
        .Filter = "Comma Seperated Values|*.csv"
        .ShowOpen
        If .FileName <> "" Then
            txtFile = .FileName
        Else
            MsgBox "No file selected", vbInformation, Me.Caption
            txtFile = ""
            Exit Sub
        End If
        .FileName = ""
    End With
End Sub

Private Sub cmdSelectAll_Click()
If lvwMemberDetails.ListItems.Count = 0 Then
  MsgBox "Please there is no item to select/Deselect", vbExclamation
Else
  If cmdSelectAll.Caption = "Select All" Then
       cmdSelectAll.Caption = "Deselect All"
       For I = 1 To lvwMemberDetails.ListItems.Count
             If lvwMemberDetails.ListItems.Item(I).Checked = False Then
                 lvwMemberDetails.ListItems.Item(I).Checked = True
             Else
                 lvwMemberDetails.ListItems.Item(I).Checked = False
             End If
       Next I
       
  ElseIf cmdSelectAll.Caption = "Deselect All" Then
        cmdSelectAll.Caption = "Select All"
        For I = 1 To lvwMemberDetails.ListItems.Count
             If lvwMemberDetails.ListItems.Item(I).Checked = False Then
                 lvwMemberDetails.ListItems.Item(I).Checked = True
             Else
                 lvwMemberDetails.ListItems.Item(I).Checked = False
             End If
       Next I
        
  End If
End If
End Sub
'get grader number
Private Function getGraderNo(subroute1 As String) As String
   Set rst1 = oSaccoMaster.GetRecordset("select LCode from d_Location where LName = '" & subroute1 & "' ")
    If Not rst1.EOF Then
        Set rst = oSaccoMaster.GetRecordset("select graderNo from d_GraderRoutesAllocation where locationId = '" & rst1(0) & "'")
          If Not rst.EOF Then
              getGraderNo = rst(0)
              Exit Function
          Else
              MsgBox "please Allocate grader to route : '" & subroute & "'", vbExclamation
          End If
    Else
            MsgBox "Please the Route is Not available", vbExclamation
    End If

End Function

'get grader number
Private Function getGraderNumber(subroute As String) As Integer
   Set rst1 = oSaccoMaster.GetRecordset("select  graderNo from d_Graders where Route = '" & subroute & "'")
     If Not rst1.EOF Then
         getGraderNumber = rst1(0)
     Else
           MsgBox "Please SubRoute Has not been allocated to SubRoute:>> '" & subroute & "'", vbExclamation
     End If
End Function


Private Sub Command1_Click()
  Dim rstww As New ADODB.Recordset
    Set rstww = oSaccoMaster.GetRecordset(" select m.mno , m.surNaeme,  m.First +' '+m.second as names  from  mweamem m ")
       If Not rstww.EOF Then
          While Not rstww.EOF
            oSaccoMaster.Execute " update members set Surname = '" & rstww(1) & "',OtherNames = '" & rstww(2) & "'  where MemberNo= '" & rstww(0) & "' "
              'Label1.Caption = rstww(0)
            rstww.MoveNext
             'Label1.Caption = rstww(0)
           Wend
       End If
      
     MsgBox " done"

End Sub

Private Sub Command2_Click()
   Dim rscub As New ADODB.Recordset
      'Set rscub = oSaccoMaster.GetRecordSet("SELECT      AccNo, Balance FROM         CUB WHERE     (AccountCode = 07)")
     Set rscub = oSaccoMaster.GetRecordset("SELECT     AccNo, Amount FROM         CUSTOMERBALANCE  WHERE     (AccNo LIKE '107%')")

       If Not rscub.EOF Then
          While Not rscub.EOF
          oSaccoMaster.Execute ("Update  CUB  set Balance='" & rscub(1) & "' where AccNo='" & rscub(0) & "' and AccountCode = 07")
         rscub.MoveNext
         Wend
       End If
   MsgBox "This is Good Place!"
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Command5_Click()

sql = ""
sql = "set dateformat dmy (SELECT     TOP (2) TransDate, Amount, DrAccNo, CrAccNo, DocumentNo, TransDescript, ChequeNo From GLTRANSACTIONS Where (auditid = 'mary'))"
cn.Execute sql
'Export

End Sub

Private Sub DTPTransdate_Click()
loadImport
End Sub
Private Sub dtpTransDate_change()
loadImport
End Sub

Private Sub Form_Load()
  DTPTransdate = Format(Get_Server_Date, "dd/mm/yyyy")
  collectors
  loadImport
End Sub
Private Sub collectors()
    Set rst = New Recordset
    sql = "Select Name from d_RouteCollectors order by Name"
    Set rst = oSaccoMaster.GetRecordset(sql)
    While Not rst.EOF
    cbochangecollector.AddItem rst.Fields(0)
    rst.MoveNext
    Wend
End Sub

Private Sub mnuPayoutReport_Click()
  STRFORMULA = ""
        reportname = "Payout List.rpt"
        Show_Sales_Crystal_Report STRFORMULA, reportname, CompanyName
End Sub




