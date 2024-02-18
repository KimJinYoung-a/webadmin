<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  업무협조
' History : 강준구 생성
'			2022.07.11 한용민 수정(isms취약점보안조치, 표준코드로변경)
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
		response.write "<script type='text/javascript'>alert('업무협조 댓글 내용에는 Script 또는 Action을 사용하실 수 없습니다.');history.back();</script>"
		response.End
	End If	
	
	If iAns_Idx = "" Then
	
		'####### 답변 저장 #######
		strSql = " INSERT INTO [db_partner].[dbo].tbl_cooperate_ans " & _
				 "		(doc_idx, id, ans_type, ans_content) " & _
				 "	VALUES " & _
				 "		('" & iDoc_Idx & "', '" & sDoc_Id & "', '" & sAns_Type & "', '" & sDoc_Content & "') " & _
				 " UPDATE [db_partner].[dbo].tbl_cooperate_document SET doc_ans_ox = 'o' WHERE doc_idx = '" & iDoc_Idx & "' "
		dbget.execute strSql
		
		'####### 로그 저장 (insert:1, 협조문 답변 작성:5) #######
		Call LogInsert(iDoc_Idx,"1","5")
		'####### 로그 저장 #######
		
		
		'####### SMS 전송 ######
		If sDoc_SMS = "o" Then
			'Call SendNormalSMS(fnGetMemberHp(sDoc_RegistId),"",""&session("ssBctCname")&"님께서 협조문 답변을 남기셨습니다.(No." & iDoc_Idx & ")")
			'Call SendNormalSMS_LINK(fnGetMemberHp(sDoc_RegistId),"",""&session("ssBctCname")&"님께서 협조문 답변을 남기셨습니다.(No." & iDoc_Idx & ")")

			dim docMsg
			docMsg = session("ssBctCname")&"님께서 협조문 답변을 남기셨습니다.(No." & iDoc_Idx & ")" & vbCrLf
			docMsg = docMsg & "----------" & vbCrLf
			docMsg = docMsg & sDoc_Content
			docMsg = chrbyte(Trim(docMsg),2000,"Y")
			Call SendRadioWebHookMessage(fnGetMemberEmail(sDoc_RegistId),"admin","어드민 알림","업무협조 답변",docMsg,"")

			'####### 로그 저장 (insert:1, 협조문 작업자에게 SMS 전송:8) #######
			Call LogInsert(iDoc_Idx,"1","8")
			'####### 로그 저장 #######
		End If
		
	Else
	
		If requestCheckVar(Request("del"),1) = "o" Then
			
			'####### 답변 삭제 #######
			strSql = " UPDATE [db_partner].[dbo].tbl_cooperate_ans SET " & _
					 "		ans_useyn = 'N' " & _
					 "	WHERE " & _
					 "		ans_idx = '" & iAns_Idx & "' "
			dbget.execute strSql
			
			'####### 로그 저장 (delete:1, 협조문 답변 삭제:7) #######
			Call LogInsert(iDoc_Idx,"3","7")
			'####### 로그 저장 #######
			
			
		Else
		
			'####### 답변 저장 #######
			strSql = " UPDATE [db_partner].[dbo].tbl_cooperate_ans SET " & _
					 "		ans_content = '" & sDoc_Content & "' " & _
					 "	WHERE " & _
					 "		ans_idx = '" & iAns_Idx & "' "
			dbget.execute strSql
			
			'####### 로그 저장 (update:2, 협조문 답변 수정:6) #######
			Call LogInsert(iDoc_Idx,"2","6")
			'####### 로그 저장 #######
			
			
		End IF

	End If


	Response.Write "<script type='text/javascript'>alert('처리되었습니다.');location.href='/admin/notice/cooperate/iframe_cooperate_ans.asp?didx="&iDoc_Idx&"&iC="&iCurrentpage&"&registid="&sDoc_RegistId&"';</script>"
	dbget.close()
	Response.End
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
