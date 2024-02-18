<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  공지사항 답글 프로세스
' History : 2011.03.02 김진영 생성
'			2020.06.30 한용민 수정(남의 글이 수정/삭제가 되어서 수정)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/member_board/boardCls.asp"-->
<%
Dim rMode, rid, sBrd_sn, rcmt_content, cidx
Dim strSql, i
	rid					= session("ssBctId")
	sBrd_sn				= Request("brd_sn")
	rcmt_content		= trim(Request("cmt_content"))
	rMode 				= Request("mode")
	cidx 				= Request("cidx")

if rcmt_content <> "" and not(isnull(rcmt_content)) then
	rcmt_content = ReplaceBracket(rcmt_content)

	'if checkNotValidHTML(rcmt_content) then
	'	response.write "<script type='text/javascript'>"
	'	response.write "	alert('공지사항 답글에는 HTML을 사용하실 수 없습니다.');history.back();"
	'	response.write "</script>"
	'	response.End
	'end if
end If

If cidx <> "" and rMode <> "del" Then
	rMode = "modify"
End If

If rMode = "add" Then
	strSql = ""
	strSql = strSql & " INSERT INTO db_partner.dbo.tbl_cooperate_board_comment " & vbcrlf
	strSql = strSql & " (id, cmt_content, brd_sn) " & vbcrlf
	strSql = strSql & "	VALUES " & vbcrlf
	strSql = strSql & "	('" & rid & "', '" & html2db(rcmt_content) & "', '" & sBrd_sn & "')"
	dbget.execute strSql
	Response.Write "<script>alert('저장 되었습니다.');location.href='iframe_board_reply.asp?brd_sn="&sBrd_sn&"&rid="&rid&"';</script>"
	dbget.close()
	response.end

ElseIf rMode = "modify" Then
	strSql = ""
	strSql = strSql & " Update db_partner.dbo.tbl_cooperate_board_comment set " & vbcrlf
	strSql = strSql & " cmt_content = '" & html2db(rcmt_content) & "' " & vbcrlf
	strSql = strSql & " where cmt_sn = " & cidx & " and brd_sn = " & sBrd_sn

	if not(C_CSPowerUser) and not(C_ADMIN_AUTH) then
		strSql = strSql & "  and id='"& rid &"'" & vbcrlf
	end if

	'response.write strSql & "<Br>"
	dbget.execute strSql
	Response.Write "<script>alert('수정 되었습니다.');location.href='iframe_board_reply.asp?brd_sn="&sBrd_sn&"&rid="&rid&"';</script>"
	dbget.close()
	response.end

ElseIf rMode = "del" Then
	strSql = ""
	strSql = strSql & " delete from db_partner.dbo.tbl_cooperate_board_comment " & vbcrlf
	strSql = strSql & " where cmt_sn = " & cidx & " and brd_sn = " & sBrd_sn

	if not(C_CSPowerUser) and not(C_ADMIN_AUTH) then
		strSql = strSql & "  and id='"& rid &"'" & vbcrlf
	end if

	'response.write strSql & "<Br>"
	dbget.execute strSql
	Response.Write "<script>alert('삭제 되었습니다.');location.href='iframe_board_reply.asp?brd_sn="&sBrd_sn&"&rid="&rid&"';</script>"
	dbget.close()
	response.end
End If
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->