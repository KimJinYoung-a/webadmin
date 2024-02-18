<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �������� ��� ���μ���
' History : 2011.03.02 ������ ����
'			2020.06.30 �ѿ�� ����(���� ���� ����/������ �Ǿ ����)
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
	'	response.write "	alert('�������� ��ۿ��� HTML�� ����Ͻ� �� �����ϴ�.');history.back();"
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
	Response.Write "<script>alert('���� �Ǿ����ϴ�.');location.href='iframe_board_reply.asp?brd_sn="&sBrd_sn&"&rid="&rid&"';</script>"
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
	Response.Write "<script>alert('���� �Ǿ����ϴ�.');location.href='iframe_board_reply.asp?brd_sn="&sBrd_sn&"&rid="&rid&"';</script>"
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
	Response.Write "<script>alert('���� �Ǿ����ϴ�.');location.href='iframe_board_reply.asp?brd_sn="&sBrd_sn&"&rid="&rid&"';</script>"
	dbget.close()
	response.end
End If
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->