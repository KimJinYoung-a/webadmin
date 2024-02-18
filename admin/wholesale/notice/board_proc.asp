<%@  codepage="65001" language="VBScript" %>
<% option explicit %>
<%
'###########################################################
' Description : 텐바이텐 대량구매 사이트 게시판 관리
' Hieditor : 2013.05.09 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/wholesale/notice/boardCls.asp"-->

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">

<% if session("sslgnMethod")<>"S" then %>
	<!-- USB키 처리 시작 (2008.06.23;허진원) -->
	<OBJECT ID='MaGerAuth' WIDTH='' HEIGHT=''	CLASSID='CLSID:781E60AE-A0AD-4A0D-A6A1-C9C060736CFC' codebase='/lib/util/MaGer/MagerAuth.cab#Version=1,0,2,4'></OBJECT>
	<script language="javascript" src="/js/check_USBToken.js"></script>
	<!-- USB키 처리 끝 -->
<% end if %>
</head>
<body bgcolor="#F4F4F4" onload="checkUSBKey()">

<%
dim adminuserid, brd_username, brd_subject, brd_content, brd_hit, brd_regdate, brd_fixed, brd_isusing, brd_type, brd_sn
dim menupos, strSql, i, mode
	brd_sn = request("brd_sn")
	menupos 	= request("menupos")
	brd_subject 		= Request("brd_subject")
	brd_content 		= Request("brd_content")
	brd_fixed 			= Request("brd_fixed")
	brd_isusing		= Request("brd_isusing")
	brd_type			= Request("brd_type")
	mode	 			= Request("mode")

	adminuserid				= session("ssBctId")

dim referer
	referer = request.ServerVariables("HTTP_REFERER")

response.write "mode : " & mode & "<Br>"
	
If mode = "brdreg" Then

	'if checkNotValidHTML(brd_content) or brd_content = "" or checkNotValidHTML(brd_subject) or brd_subject = "" then
	if brd_content = "" or brd_subject = "" then
		response.write "<script language='javascript'>"
		response.write "	alert('내용이 올바르지 않습니다1.');"
		response.write "	location.href='"&referer&"';"
		response.write "</script>"	
		dbget.close()	:	response.End
	end if

	if brd_fixed = "" or brd_isusing = "" or brd_type = "" then
		response.write "<script language='javascript'>"
		response.write "	alert('내용이 올바르지 않습니다2.');"
		response.write "	location.href='"&referer&"';"
		response.write "</script>"	
		dbget.close()	:	response.End
	end if
	
	
	strSql = "if exists(" + vbcrlf
	strSql = strSql & " 	select top 1 brd_sn from db_board.dbo.tbl_board_wholesale where brd_sn='"&brd_sn&"'" + vbcrlf
	strSql = strSql & " )" + vbcrlf
	strSql = strSql & " 	update db_board.dbo.tbl_board_wholesale set" + vbcrlf
	strSql = strSql & " 	brd_type = N'"&brd_type&"'" + vbcrlf
	strSql = strSql & " 	,brd_subject = N'"&html2db(brd_subject)&"'" + vbcrlf
	strSql = strSql & " 	,brd_content = N'"&html2db(brd_content)&"'" + vbcrlf
	strSql = strSql & " 	,brd_fixed = N'"&brd_fixed&"'" + vbcrlf
	strSql = strSql & " 	,brd_isusing = N'"&brd_isusing&"'" + vbcrlf
	strSql = strSql & " 	,brd_lastupdate = getdate()" + vbcrlf
	strSql = strSql & " 	,lastuserid = N'"&adminuserid&"'" + vbcrlf		
	strSql = strSql & " 	where brd_sn= '"&brd_sn&"'" + vbcrlf
	strSql = strSql & " else" + vbcrlf
	strSql = strSql & " 	insert into db_board.dbo.tbl_board_wholesale(" + vbcrlf
	strSql = strSql & " 	brd_gubun, brd_type, brd_subject, brd_content, brd_fixed, brd_isusing" + vbcrlf
	strSql = strSql & " 	, userid, lastuserid) values (" + vbcrlf
	strSql = strSql & "		N'1', N'"&brd_type&"', N'" & html2db(brd_subject) & "', N'" & html2db(brd_content) & "'" + vbcrlf
	strSql = strSql & "		, N'" & brd_fixed & "', N'Y', N'" & adminuserid & "', N'" & adminuserid & "')"
	
	'response.write strSql & "<Br>"
	dbget.execute strSql

	response.write "<script language='javascript'>"
	response.write "	alert('OK');"
	response.write "	location.href='/admin/wholesale/notice/board_list.asp?menupos="&menupos&"';"
	response.write "</script>"
	dbget.close() : response.end

else
	response.write "<script language='javascript'>"
	response.write "	alert('MODE 구분자가 지정되지 않았습니다.');"
	response.write "	location.href='"&referer&"';"
	response.write "</script>"
	dbget.close() : response.end
End If
%>

</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->