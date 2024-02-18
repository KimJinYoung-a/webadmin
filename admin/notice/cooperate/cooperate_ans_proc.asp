<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  ��������
' History : ���ر� ����
'			2022.07.11 �ѿ�� ����(isms�����������ġ, ǥ���ڵ�κ���)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cooperate/chk_auth.asp"-->
<!-- #include virtual="/lib/classes/cooperate/cooperateCls.asp"-->
<!-- #include virtual="/lib/email/smslib.asp"-->

<%
	Dim strSql, vWorkerTemp, vWorkerViewTemp, vFileTemp, i, iCurrentpage
	Dim iDoc_Idx, iAns_Idx, sDoc_Id, sDoc_Content, sAns_Type, sDoc_SMS, sDoc_RegistId
	iDoc_Idx		= NullFillWith(requestCheckVar(Request("didx"),10),"")
	iAns_Idx		= NullFillWith(requestCheckVar(Request("aidx"),10),"")
	iCurrentpage 	= NullFillWith(requestCheckVar(Request("iC"),10),1)
	sDoc_Id			= session("ssBctId")
	sAns_Type		= "1"
	sDoc_Content	= html2db(Request("ans_content"))
	sDoc_SMS		= NullFillWith(Request("sms_send"),"")
	sDoc_RegistId	= NullFillWith(Request("registid"),"")
	
	if sDoc_Content <> "" and not(isnull(sDoc_Content)) then
		sDoc_Content = ReplaceBracket(sDoc_Content)
	end If
	If (checkNotValidHTML(sDoc_Content) = true) Then
		response.write "<script type='text/javascript'>alert('�������� ��� ���뿡�� Script �Ǵ� Action�� ����Ͻ� �� �����ϴ�.');history.back();</script>"
		response.End
	End If	
	
	If iAns_Idx = "" Then
	
		'####### �亯 ���� #######
		strSql = " INSERT INTO [db_partner].[dbo].tbl_cooperate_ans " & _
				 "		(doc_idx, id, ans_type, ans_content) " & _
				 "	VALUES " & _
				 "		('" & iDoc_Idx & "', '" & sDoc_Id & "', '" & sAns_Type & "', '" & sDoc_Content & "') " & _
				 " UPDATE [db_partner].[dbo].tbl_cooperate_document SET doc_ans_ox = 'o' WHERE doc_idx = '" & iDoc_Idx & "' "
		dbget.execute strSql
		
		'####### �α� ���� (insert:1, ������ �亯 �ۼ�:5) #######
		Call LogInsert(iDoc_Idx,"1","5")
		'####### �α� ���� #######
		
		
		'####### SMS ���� ######
		If sDoc_SMS = "o" Then
			'Call SendNormalSMS(fnGetMemberHp(sDoc_RegistId),"",""&session("ssBctCname")&"�Բ��� ������ �亯�� ����̽��ϴ�.(No." & iDoc_Idx & ")")
			'Call SendNormalSMS_LINK(fnGetMemberHp(sDoc_RegistId),"",""&session("ssBctCname")&"�Բ��� ������ �亯�� ����̽��ϴ�.(No." & iDoc_Idx & ")")

			dim docMsg
			docMsg = session("ssBctCname")&"�Բ��� ������ �亯�� ����̽��ϴ�.(No." & iDoc_Idx & ")" & vbCrLf
			docMsg = docMsg & "----------" & vbCrLf
			docMsg = docMsg & sDoc_Content
			docMsg = chrbyte(Trim(docMsg),2000,"Y")
			Call SendRadioWebHookMessage(fnGetMemberEmail(sDoc_RegistId),"admin","���� �˸�","�������� �亯",docMsg,"")

			'####### �α� ���� (insert:1, ������ �۾��ڿ��� SMS ����:8) #######
			Call LogInsert(iDoc_Idx,"1","8")
			'####### �α� ���� #######
		End If
		
	Else
	
		If requestCheckVar(Request("del"),1) = "o" Then
			
			'####### �亯 ���� #######
			strSql = " UPDATE [db_partner].[dbo].tbl_cooperate_ans SET " & _
					 "		ans_useyn = 'N' " & _
					 "	WHERE " & _
					 "		ans_idx = '" & iAns_Idx & "' "
			dbget.execute strSql
			
			'####### �α� ���� (delete:1, ������ �亯 ����:7) #######
			Call LogInsert(iDoc_Idx,"3","7")
			'####### �α� ���� #######
			
			
		Else
		
			'####### �亯 ���� #######
			strSql = " UPDATE [db_partner].[dbo].tbl_cooperate_ans SET " & _
					 "		ans_content = '" & sDoc_Content & "' " & _
					 "	WHERE " & _
					 "		ans_idx = '" & iAns_Idx & "' "
			dbget.execute strSql
			
			'####### �α� ���� (update:2, ������ �亯 ����:6) #######
			Call LogInsert(iDoc_Idx,"2","6")
			'####### �α� ���� #######
			
			
		End IF

	End If


	Response.Write "<script type='text/javascript'>alert('ó���Ǿ����ϴ�.');location.href='/admin/notice/cooperate/iframe_cooperate_ans.asp?didx="&iDoc_Idx&"&iC="&iCurrentpage&"&registid="&sDoc_RegistId&"';</script>"
	dbget.close()
	Response.End
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
