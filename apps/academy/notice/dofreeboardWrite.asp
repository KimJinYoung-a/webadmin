<%@ codepage="65001" language="VBScript" %>
<% option explicit %>
<%
'###########################################################
' Description :  핑거스 강사 게시판
' History : 2010.03.30 한용민 생성
'###########################################################
%>
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/academy/lib/classes/board/lecturer/lecturer_cls.asp"-->

<%
Dim strSql, vWorkerTemp, vWorkerViewTemp, vFileTemp, i, page
Dim iDoc_Idx, iAns_Idx, sDoc_Id, sDoc_Content, sAns_Type, sDoc_RegistId
	iDoc_Idx		= NullFillWith(requestCheckVar(Request("didx"),10),"")
	iAns_Idx		= NullFillWith(requestCheckVar(Request("aidx"),10),"")
	page 	= NullFillWith(requestCheckVar(Request("page"),10),1)
	sDoc_Id			= requestCheckVar(request.cookies("partner")("userid"),32)
	sAns_Type		= "1"
	sDoc_Content	= html2db(Request("ans_content"))
		
If iAns_Idx = "" Then

	'####### 답변 저장 #######
	strSql = " INSERT INTO [db_academy].[dbo].tbl_lecturer_board_ans " & _
			 "		(doc_idx, id, ans_type, ans_content) " & _
			 "	VALUES " & _
			 "		('" & iDoc_Idx & "', '" & sDoc_Id & "', '" & sAns_Type & "', '" & html2db(sDoc_Content) & "') " & _
			 " UPDATE [db_academy].[dbo].tbl_lecturer_board_document SET doc_ans_ox = 'o' WHERE doc_idx = '" & iDoc_Idx & "' "
	
	'response.write strSql &"<br>"
	'Response.end
	dbacademyget.execute strSql

Else

	If requestCheckVar(Request("del"),1) = "o" Then
		
		'####### 답변 삭제 #######
		strSql = " UPDATE [db_academy].[dbo].tbl_lecturer_board_ans SET " & _
				 "		ans_useyn = 'N' " & _
				 "	WHERE " & _
				 "		ans_idx = '" & iAns_Idx & "' "
		
		'response.write strSql &"<br>"
		dbacademyget.execute strSql

	Else
	
		'####### 답변 저장 #######
		strSql = " UPDATE [db_academy].[dbo].tbl_lecturer_board_ans SET " & _
				 "		ans_content = '" & html2db(sDoc_Content) & "' " & _
				 "	WHERE " & _
				 "		ans_idx = '" & iAns_Idx & "' "

		'response.write strSql &"<br>"
		dbacademyget.execute strSql
		
	End IF

End If
%>
<script>
<!--
	parent.fnFreeBoardRelyEnd();
//-->
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->