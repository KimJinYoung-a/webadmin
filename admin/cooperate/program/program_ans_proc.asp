<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 프로그램변경내역
' Hieditor : 강준구 생성
'			 2022.07.11 한용민 수정(isms취약점보안조치, 표준코드로변경)
'###########################################################
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
	Dim iDoc_Idx, iAns_Idx, sDoc_Id, sDoc_Content, sAns_Type, sDoc_SMS, sDoc_RegistId, vIsPop
	Dim vPIdx, vAIdx
	
	vPIdx			= NullFillWith(requestCheckVar(Request("pidx"),10),"")
	vAIdx			= NullFillWith(requestCheckVar(Request("aidx"),10),"")
	iCurrentpage 	= NullFillWith(requestCheckVar(Request("iC"),10),1)
	sDoc_Id			= session("ssBctId")
	sAns_Type		= "1"
	sDoc_Content	= html2db(Request("ans_content"))
	sDoc_SMS		= NullFillWith(Request("sms_send"),"")
	sDoc_RegistId	= NullFillWith(Request("registid"),"")

	If vAIdx = "" Then
		if sDoc_Content <> "" and not(isnull(sDoc_Content)) then
			sDoc_Content = ReplaceBracket(sDoc_Content)
		end If

		'####### 답변 저장 #######
		strSql = " INSERT INTO [db_board].[dbo].tbl_program_change_comment " & _
				 "		(pidx, userid, comment, useyn) " & _
				 "	VALUES " & _
				 "		('" & vPIdx & "', '" & sDoc_Id & "', '" & sDoc_Content & "', 'Y') "
		dbget.execute strSql

	Else
		If requestCheckVar(Request("del"),1) = "o" Then
			
			'####### 답변 삭제 #######
			strSql = " UPDATE [db_board].[dbo].tbl_program_change_comment SET " & _
					 "		useyn = 'N' " & _
					 "	WHERE " & _
					 "		idx = '" & vAIdx & "' "
			dbget.execute strSql

		Else
			if sDoc_Content <> "" and not(isnull(sDoc_Content)) then
				sDoc_Content = ReplaceBracket(sDoc_Content)
			end If

			'####### 답변 저장 #######
			strSql = " UPDATE [db_board].[dbo].tbl_program_change_comment SET " & _
					 "		comment = '" & sDoc_Content & "' " & _
					 "	WHERE " & _
					 "		idx = '" & vAIdx & "' "
			dbget.execute strSql
		End IF

	End If

	Response.Write "<script type='text/javascript'>alert('처리되었습니다.');parent.location.reload();</script>"
	dbget.close()
	Response.End
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
