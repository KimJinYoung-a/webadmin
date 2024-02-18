<%@  codepage="65001" language="VBScript" %>
<% option explicit %>
<% response.Charset="UTF-8" %>
<% session.codePage = 65001 %>
<%
'###########################################################
' Description : 아이띵소 사이트 관리
' Hieditor : 2013.05.15 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->

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

<!-- #include virtual="/lib/classes/ithinkso/sitemanager/sitemanager_cls_ithinkso_utf8.asp"-->

<%
dim code , codename,imagetype, imagewidth ,isusing ,imageheight,imagecount, mode, orgcode, adminuserid
	code   = request("code")
	codename   = html2Db(request("codename"))
	imagetype  = request("imagetype")
	imagetype   = request("imagetype")
	imagewidth= request("imagewidth")
	isusing   = request("isusing")
	imageheight= request("imageheight")
	imagecount= request("imagecount")
	orgcode= request("orgcode")
	mode= request("mode")
	
adminuserid				= session("ssBctId")
	
response.write "MODE : " & mode & "<Br>"

dim referer
	referer = request.ServerVariables("HTTP_REFERER")
	
dim sqlStr, ItemExists

if mode = "codereg" then
	
	if orgcode <> "" then
	    sqlStr = " update db_contents.dbo.tbl_sitemanager_code_ithinkso" + VbCrlf
	    sqlStr = sqlStr + " set codename=N'" + html2db(codename) + "'" + VbCrlf
	    sqlStr = sqlStr + " ,imagetype=N'" + imagetype + "'" + VbCrlf
	    sqlStr = sqlStr + " ,imagewidth=N'" + imagewidth + "'" + VbCrlf
	    sqlStr = sqlStr + " ,imageheight=N'" + imageheight + "'" + VbCrlf
	    sqlStr = sqlStr + " ,imagecount=N'" + imagecount + "'" + VbCrlf
	    sqlStr = sqlStr + " ,isusing=N'" + isusing + "'" + VbCrlf
	    sqlStr = sqlStr + " ,lastupdateadminid=N'" + adminuserid + "'" + VbCrlf
	    sqlStr = sqlStr + " ,lastdate=getdate()" + VbCrlf	    	    
	    sqlStr = sqlStr + " where codetype=1 and code="&orgcode&""
	    
	    'response.write sqlStr
	    dbget.Execute sqlStr
	else
	    sqlStr = " insert into db_contents.dbo.tbl_sitemanager_code_ithinkso (" + VbCrlf
	    sqlStr = sqlStr + " code, codetype, codename,imagetype,imagewidth,imageheight,isusing,imagecount, regadminid" + VbCrlf
	    sqlStr = sqlStr + " , lastupdateadminid) values ("+ VbCrlf
	    sqlStr = sqlStr + " N'" + code + "'" + VbCrlf
	    sqlStr = sqlStr + " ,N'1'" + VbCrlf	    
	    sqlStr = sqlStr + " ,N'" + html2db(codename) + "'" + VbCrlf
	    sqlStr = sqlStr + " ,N'" + imagetype + "'" + VbCrlf
	    sqlStr = sqlStr + " ,N'" + imagewidth + "'" + VbCrlf
	    sqlStr = sqlStr + " ,N'" + imageheight + "'" + VbCrlf
	    sqlStr = sqlStr + " ,N'" + isusing + "'" + VbCrlf
	    sqlStr = sqlStr + " ,N'" + imagecount + "'" + VbCrlf
	    sqlStr = sqlStr + " ,N'" + adminuserid + "'" + VbCrlf
	    sqlStr = sqlStr + " ,N'" + adminuserid + "'" + VbCrlf	    
	    sqlStr = sqlStr + " )"
	    
	    'response.write sqlStr    
	    dbget.Execute sqlStr
	end if

	session.codePage = 949
	response.write "<script language='javascript'>"
	response.write "	alert('OK');"
	response.write "	location.href='"&referer&"';"
	response.write "</script>"
	dbget.close() : response.end
	
else
	session.codePage = 949
	response.write "<script language='javascript'>"
	response.write "	alert('MODE 구분자가 지정되지 않았습니다.');"
	response.write "	location.href='"&referer&"';"
	response.write "</script>"
	dbget.close() : response.end
end if

session.codePage = 949
%>

</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->
