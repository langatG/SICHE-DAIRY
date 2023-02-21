Attribute VB_Name = "setdefaultgls"
Public Sub setdefaultgls(transdate As Date, description As String)
        Dim E, remark, Ddr, Ccr As String
        Dim amount As Double
        
        Startdate = DateSerial(Year(transdate), month(transdate), 1)
        Enddate = DateSerial(Year(transdate), month(transdate) + 1, 1 - 1)

            E = Format(transdate, "dd/mm/yyyy") & ""
            remark = "milk" + " " + description
            
        sql = "" 'Dr='" & txtDrAccNo & "' and Cr='" & txtCrAccNo & "' and
        sql = "select * from GLSetDefaultGls Where Affect='" & description & "'"
        Set rs = oSaccoMaster.GetRecordset(sql)
        If Not rs.EOF Then
          Ddr = rs!dr
          Ccr = rs!cr
          If description = "Purchases" Then
            sql = "set dateformat dmy select isnull(sum(PAmount),0) as amnt from  d_Milkintake Where TransDate='" & transdate & "'"
          ElseIf description = "Payables" Then
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
                If description = "Payables" And transdate = Enddate Then
                 amount = amount - rsg!midAmount
                End If
            sql = "set dateformat dmy select * from  gltransactions Where transdate='" & transdate & "' and documentno='" & description & "'"
            Set rss = oSaccoMaster.GetRecordset(sql)
            If Not rss.EOF Then
             
               
                sql = "set dateformat dmy update gltransactions set amount='" & amount & "' where transdate='" & transdate & "' and documentno='" & description & "' and transdescript='" & E & "-" & remark & "'"
               oSaccoMaster.ExecuteThis (sql)
             Else
               sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,AuditTime,auditid,cash,doc_posted) "
               sql = sql & " values ('" & transdate & "','" & amount & "','" & Ddr & "','" & Ccr & "','" & description & "','' ,'" & E & "- " & remark & "','" & Now & "','" & User & "',0,0)"
               oSaccoMaster.ExecuteThis (sql)
            End If
          End If
        End If
        
        Exit Sub
Capture:
        ErrorMessage = ErrorMessage
End Sub
Public Sub deductions(transdate As Date, description As String, sno As String, Remarks As String)
        Dim E, remark, Ddr, Ccr As String
        Dim amount As Double
        
        Startdate = DateSerial(Year(transdate), month(transdate), 1)
        Enddate = DateSerial(Year(transdate), month(transdate) + 1, 1 - 1)

            E = Format(transdate, "dd/mm/yyyy") & ""
            
        sql = "SELECT Dedaccno, Contraacc FROM d_DCodes where Description='" & description & "'"
        Set rs = oSaccoMaster.GetRecordset(sql)
        If Not rs.EOF Then
        
               sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,AuditTime,auditid,cash,doc_posted) "
               sql = sql & " values ('" & transdate & "','" & amount & "','" & Ddr & "','" & Ccr & "','" & description & "','' ,'" & E & "- " & Remarks & "','" & Now & "','" & User & "',0,0)"
               oSaccoMaster.ExecuteThis (sql)
        
        End If
        
        Exit Sub
Capture:
        ErrorMessage = ErrorMessage
End Sub
Public Sub defaultglsdebit(transdate As Date)
        Dim E, remark, Ddr, Ccr, description As String
        Dim amount As Double
        description = "STORE FARMERS PAYMENT"
        remark = "FARMERS PAYROLL"
        Startdate = DateSerial(Year(transdate), month(transdate), 1)
        Enddate = DateSerial(Year(transdate), month(transdate) + 1, 1 - 1)

            E = Format(transdate, "dd/mm/yyyy") & ""
        'EndofPeriod
        sql = "set dateformat dmy select isnull(sum(Agrovet),0)as Agrovet, isnull(sum(GPay),0)as GPay FROM d_Payroll where EndofPeriod='" & transdate & "'"
        Set rss = oSaccoMaster.GetRecordset(sql)
        If Not rss.EOF Then
            '''''''DEBIT CREDIT BANK AND STORE PURCHASES
            If rss!agrovet > 0 Then
                sql = ""
                sql = "set dateformat dmy select * from  gltransactions Where transdate='" & transdate & "' and documentno='" & description & "'"
                Set rs = oSaccoMaster.GetRecordset(sql)
                If rs.EOF Then
                     
                       sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,AuditTime,auditid,cash,doc_posted) "
                       sql = sql & " values ('" & transdate & "','" & rss!agrovet & "','C002','AG003','" & description & "','' ,'" & E & "- " & description & "','" & Now & "','" & User & "',0,0)"
                       oSaccoMaster.ExecuteThis (sql)
                Else
                     sql = "set dateformat dmy update gltransactions set amount='" & rss!agrovet & "' where transdate='" & transdate & "' and documentno='" & description & "' and transdescript='" & E & "- " & description & "'"
                       oSaccoMaster.ExecuteThis (sql)
                End If
            End If
            '''''''DEBIT ACCOUNT PAYABLES AND CREDIT BANK
            If rss!GPay > 0 Then
                sql = ""
                sql = "set dateformat dmy select * from  gltransactions Where transdate='" & transdate & "' and documentno='" & remark & "'"
                Set rs = oSaccoMaster.GetRecordset(sql)
                If rs.EOF Then
                     
                       sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,AuditTime,auditid,cash,doc_posted) "
                       sql = sql & " values ('" & transdate & "','" & rss!GPay & "','A001','C002','" & remark & "','' ,'" & E & "- " & remark & "','" & Now & "','" & User & "',0,0)"
                       oSaccoMaster.ExecuteThis (sql)
                Else
                     sql = "set dateformat dmy update gltransactions set amount='" & rss!GPay & "' where transdate='" & transdate & "' and documentno='" & remark & "' and transdescript='" & E & "- " & remark & "'"
                       oSaccoMaster.ExecuteThis (sql)
                End If
            End If
            
        End If
        
        Exit Sub
Capture:
        ErrorMessage = ErrorMessage
End Sub
