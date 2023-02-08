Attribute VB_Name = "setdefaultgls"
Public Sub setdefaultgls(transdate As Date, Description As String)
        Dim E, Remark As String
        Dim amount As Double
        
        Startdate = DateSerial(Year(transdate), month(transdate), 1)
        Enddate = DateSerial(Year(transdate), month(transdate) + 1, 1 - 1)

            E = Format(transdate, "dd/mm/yyyy") & ""
            Remark = MILK + "& Description &"
            
        sql = "" 'Dr='" & txtDrAccNo & "' and Cr='" & txtCrAccNo & "' and
        sql = "select * from GLSetDefaultGls Where Affect='" & Description & "'"
        Set rs = oSaccoMaster.GetRecordset(sql)
        If Not rs.EOF Then
        
          If Description = "Purchases" Then
            sql = "set dateformat dmy select isnull(sum(PAmount),0) as amnt from  d_Milkintake Where TransDate='" & transdate & "'"
          ElseIf Description = "Payables" Then
            If transdate < Enddate Then
                sql = "set dateformat dmy select isnull(sum(Amount),0) as amnt from   d_supplier_deduc Where Date_Deduc='" & transdate & "' and Description='WEEKLY'"
            Else
            sql = "set dateformat dmy select isnull(sum(Amount),0) as midAmount from   d_supplier_deduc Where Date_Deduc>='" & Startdate & "' and Date_Deduc<='" & transdate & "' and Description='WEEKLY'"
            Set rsg = oSaccoMaster.GetRecordset(sql)
            
            sql = "set dateformat dmy select isnull(sum(PAmount),0) as amnt from  d_Milkintake Where TransDate>='" & Startdate & "' and TransDate<='" & transdate & "'"
            End If
          End If
          Set rst = oSaccoMaster.GetRecordset(sql)
          If rst!Amnt <> 0 Then
                amount = rst!Amnt
                If Description = "Payables" And transdate = Enddate Then
                 amount = amount - rsg!midAmount
                End If
            sql = "set dateformat dmy select * from  gltransactions Where transdate='" & transdate & "' and documentno='" & Description & "'"
            Set rss = oSaccoMaster.GetRecordset(sql)
            If Not rss.EOF Then
             
               
                sql = "set dateformat dmy update gltransactions set amount='" & amount & "' where transdate='" & transdate & "' and documentno='" & Description & "' and transdescript='" & E & "-" & Remark & "'"
               oSaccoMaster.ExecuteThis (sql)
             Else
               sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,AuditTime,auditid,cash,doc_posted) "
               sql = sql & " values ('" & transdate & "','" & amount & "','" & rs!dr & "','" & rs!cr & "','" & Description & "','' ,'" & E & "-" & Remark & "','" & Now & "','" & User & "',0,0)"
               oSaccoMaster.ExecuteThis (sql)
            End If
          End If
        End If
        
        Exit Sub
Capture:
        ErrorMessage = ErrorMessage
End Sub
