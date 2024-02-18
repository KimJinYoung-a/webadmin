<%@  codepage="65001" language="VBScript" %>
<% option explicit %>
<%
'###########################################################
' Description : 아이띵소 제휴 관리
' Hieditor : 2013.05.14 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/ithinkso/notice/boardCls.asp"-->

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
dim menupos, strSql, i, mode, isusing, idx
	menupos 	= request("menupos")
	mode	 			= Request("mode")
	idx	 			= Request("idx")
	isusing	 			= Request("isusing")

dim referer
	referer = request.ServerVariables("HTTP_REFERER")

response.write "mode : " & mode & "<Br>"
	
If mode = "contactreg" Then
	if idx = "" then
		response.write "<script language='javascript'>"
		response.write "	alert('IDX가 없습니다.');"
		response.write "	self.close();"
		response.write "</script>"
		dbget.close() : response.end
	end if
	if isusing = "" then
		response.write "<script language='javascript'>"
		response.write "	alert('사용여부가 지정되지 않았습니다.');"
		response.write "	self.close();"
		response.write "</script>"
		dbget.close() : response.end
	end if

	strSql = "if exists(" + vbcrlf
	strSql = strSql & " 	select top 1 idx from db_board.dbo.tbl_contact_ithinkso where idx='"&idx&"'" + vbcrlf
	strSql = strSql & " )" + vbcrlf
	strSql = strSql & " 	update db_board.dbo.tbl_contact_ithinkso set" + vbcrlf
	strSql = strSql & " 	isusing = N'"&isusing&"'" + vbcrlf
	strSql = strSql & " 	where idx= '"&idx&"'"
	
	'response.write strSql & "<Br>"
	dbget.execute strSql

	response.write "<script language='javascript'>"
	response.write "	alert('OK');"
	response.write "	opener.location.reload();"
	response.write "	self.close();"	
	response.write "</script>"
	dbget.close() : response.end
	
else
	response.write "<script language='javascript'>"
	response.write "	alert('MODE 구분자가 지정되지 않았습니다.');"
	response.write "	self.close();"
	response.write "</script>"
	dbget.close() : response.end
End If		
%>

</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->