<%@ language=vbscript %>
<% option explicit %>

<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"--> 

<% 
	Dim YYYYMM, vRegUserid, strSql, objCmd, vBody, vProcedure
	Dim CardCo, CardNo, AuthDate, OpExpObj, outExp, supExp, vatExp, sevExp, authNo, vatType, bizNo, useScope, acc_nm, AssignedRow, AssignedRow2, AssignedRow3, returnValue
	Dim cardTmpSeq
	
	vRegUserid = session("ssBctId")
	vProcedure = "[db_partner].[dbo].[sp_Ten_OpExpDailyCard_Insert]"
	'vProcedure = "[db_partner].[dbo].[sp_Ten_OpExpDailyCard_Insert_test]"
	AssignedRow = 0
	AssignedRow2 = 0
	AssignedRow3 = 0
	
	IF vRegUserid = "" THEN
	  	Response.Write "<script language='javascript'>alert('���������� ������ �ֽ��ϴ�.��α��� �� �ٽ� �õ����ּ���');history.back(-1);</script>"
		dbget.close()
		Response.End
	END IF

'Response.Write "<script language='javascript'>alert('������');history.back(-1);</script>"
'Response.End
	
''	dbget.beginTrans
	'############################################################## [TMSDB].db_scm_Link.dbo.vw_reged_Card_AppList ��ȸ [1] ##############################################################
	strSql = "SELECT " & vbCrLf
	strSql = strSql & " 	isNull(b.CardCo,'') AS CardCo, A.CardNo, A.TransDate, A.TransTime, A.ApprTot, A.ApprAmt1, A.VAT1, A.TIPS1, A.ApprNo, A.MerchName, A.TaxType, A.MerchBizNo, A.AbroadNM, A.ACC_NM" & vbCrLf
	strSql = strSql & " FROM [TMSDB].[db_scm_Link].[dbo].[vw_reged_Card_AppList] AS A " & vbCrLf
	strSql = strSql & " 	LEFT JOIN [db_partner].[dbo].[tbl_OpExpPart] AS B ON A.CardNo = B.CardNo AND B.isUsing = 1 " & vbCrLf
	strSql = strSql & " 	LEFT JOIN [db_partner].[dbo].[tbl_OpExpDailyCard] AS C ON A.ApprNo = C.authno AND Convert(datetime,(A.TransDate + ' ' + A.TransTime)) = C.authdate " & vbCrLf
	strSql = strSql & " WHERE C.authno is Null AND A.TransDate >= '2012-07-01' AND A.TransDate >= DateAdd(m,-2,getdate()) " & vbCrLf
	strSql = strSql & " ORDER BY A.Seq ASC " & vbCrLf

    strSql = " SELECT " & vbCrLf
    strSql = strSql & "  	isNull(b.CardCo,'') AS CardCo, A.CardNo, A.TransDate, A.TransTime, A.ApprTot, A.ApprAmt, A.VAT, A.TIPS, A.ApprNo, A.MerchName, A.TaxType, A.MerchBizNo, A.AbroadNM, A.vsAcc_cd as vsAcc_cd" & vbCrLf
    strSql = strSql & " 	,A.Seq" & vbCrLf
    strSql = strSql & "  FROM [TMSDB].[db_scm_Link].[dbo].[vw_reged_Card_AppList_sERP] AS A " & vbCrLf
    strSql = strSql & "  	LEFT JOIN [db_partner].[dbo].[tbl_OpExpPart] AS B ON A.CardNo = B.CardNo AND B.isUsing = 1 " & vbCrLf
    strSql = strSql & "  	LEFT JOIN [db_partner].[dbo].[tbl_OpExpDailyCard] AS C ON A.Seq = C.cardTmpSeq" & vbCrLf
    strSql = strSql & "  WHERE C.opExpDailyCArdIdx is Null  AND A.TransDate >= DateAdd(m,-1,getdate()) " & vbCrLf
    strSql = strSql & "  ORDER BY A.Seq ASC " & vbCrLf

''rw strSql
''response.end

	rsget.Open strSql,dbget,1
	
	If Not rsget.Eof Or rsget.Bof Then
		
		'vBody = "<table border=1>"
		'vBody = vBody & "<tr><td>B.CardCo</td><td>A.CardNo</td><td>A.TransDate</td><td>A.TransTime</td><td>A.ApprTot</td><td>A.ApprAmt1</td>"
		'vBody = vBody & "<td>A.VAT1</td><td>A.TIPS1</td><td>A.ApprNo</td><td>A.MerchName</td><td>A.TaxType</td><td>A.MerchBizNo</td><td>A.AbroadNM</td><td>A.ACC_NM</td></tr>"
		
		'############################################################## Loop & ���� ���� ��� [2] ##############################################################
		Do Until rsget.Eof
			CardCo   = rsget("CardCo")
			CardNo   = rsget("CardNo")
			
			If rsget("TransDate") <> "��������" Then
				AuthDate = rsget("TransDate") & " " & rsget("TransTime")
				OpExpObj = ReplaceRequestSpecialChar(rsget("MerchName"))
			Else
				OpExpObj = OpExpObj&"(��������)"
			End If
			
			outExp   = rsget("ApprTot")
			supExp   = rsget("ApprAmt")
			vatExp   = rsget("VAT")
			sevExp   = rsget("TIPS")
			authNo   = rsget("ApprNo")
			vatType  = Trim(rsget("TaxType"))
			bizNo    = rsget("MerchBizNo")
	
			If Trim(rsget("AbroadNM")) ="����" Then
				useScope = 1
			ElseIf Trim(rsget("AbroadNM")) ="����" Then
				useScope = 2
			Else
				useScope = 0
			End If
			
			''acc_nm = rsget("ACC_NM")
			cardTmpSeq = rsget("seq")
			
			''response.write "{?= call " & vProcedure & "('"&YYYYMM&"','"&CardCo&"','"&CardNo&"','"&AuthDate&"','"&outExp&"','"&supExp&"','"&vatExp&"','"&sevExp&"', '"&authNo&"' ,'"&OpExpObj&"','"&vatType&"','"&bizNo&"',"&useScope&",'"&acc_nm&"','"&vRegUserid&"')}"
			
			IF  CardCo <> "" THEN
				'############################################################## [db_partner].[dbo].[tbl_OpExpDailyCard] �� INSERT [3] ##############################################################
				Set objCmd = Server.CreateObject("ADODB.COMMAND")
					With objCmd
					.ActiveConnection = dbget
					.CommandType = adCmdText
					.CommandText = "{?= call " & vProcedure & "('"&YYYYMM&"','"&CardCo&"','"&CardNo&"','"&AuthDate&"','"&outExp&"','"&supExp&"','"&vatExp&"','"&sevExp&"', '"&authNo&"' ,'"&OpExpObj&"','"&vatType&"','"&bizNo&"',"&useScope&",'','"&vRegUserid&"',"&cardTmpSeq&")}"
					.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
					.Execute, , adExecuteNoRecords
					End With
					returnValue = objCmd(0).Value
				Set objCmd = nothing
			
				IF returnValue = 1 THEN
					AssignedRow=AssignedRow+1
				ELSEIF returnValue = 2 THEN
					AssignedRow2 = AssignedRow2 + 1
				ELSEIF returnValue = 3 THEN
					AssignedRow3 = AssignedRow3 + 1
				END IF
			
				IF returnValue  = 0 THEN
				''	dbget.RollbackTrans
					rsget.close()
					dbget.close()
					Response.Write "<script language='javascript'>alert('������ó���� ������ �߻� �߽��ϴ�.');history.back(-1);</script>"
					Response.End
				END IF
				
				'vBody = vBody & "<tr><td>"&rsget("CardCo")&"</td><td>"&rsget("CardNo")&"</td><td>"&rsget("TransDate")&"</td><td>"&rsget("TransTime")&"</td><td>"&rsget("ApprTot")&"</td><td>"&rsget("ApprAmt1")&"</td>"
				'vBody = vBody & "<td>"&rsget("VAT1")&"</td><td>"&rsget("TIPS1")&"</td><td>"&rsget("ApprNo")&"</td><td>"&rsget("MerchName")&"</td><td>"&rsget("TaxType")&"</td><td>"&rsget("MerchBizNo")&"</td>"
				'vBody = vBody & "<td>"&rsget("AbroadNM")&"</td><td>"&rsget("ACC_NM")&"</td></tr>"
			END IF
		rsget.MoveNext
		Loop
		'vBody = vBody & "</table>"
	End If
	rsget.close()
	''dbget.CommitTrans

	vBody = "* " & AssignedRow & "�� ��ϵǰ�, " & AssignedRow2 & "���� �����Ͱ� ��ϵ��� �ʾҰ�, " & AssignedRow3 & "���� �ߺ��� �����Ͱ� ����."
	
'rw "* " & AssignedRow & "�� ��ϵǰ�, " & AssignedRow2 & "���� �����Ͱ� ��ϵ��� �ʾҰ�, " & AssignedRow3 & "���� �ߺ��� �����Ͱ� ����.<br><br>"
'rw "* ���� [db_partner].[dbo].[sp_Ten_OpExpDailyCard_Insert] ���ν������� insert ������ �ּ�ó���ϰ� ���� [db_partner].[dbo].[sp_Ten_OpExpDailyCard_Insert_test]�� �׽�Ʈ<br>"
'rw Replace(strSql,vbCrLf,"<br>")
'rw vBody
%>

<script language="javascript">
parent.document.getElementById("erpprocmessage").innerHTML = "<%=vBody%>";
parent.document.getElementById("reflashbutton").style.display = "block";
//alert("ó���Ϸ�Ǿ����ϴ�.");
</script>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->