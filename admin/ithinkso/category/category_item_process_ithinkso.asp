<%@  codepage="65001" language="VBScript" %>
<% option explicit %>
<%
'###########################################################
' Description : 아이띵소 카테고리 관리
' Hieditor : 2013.05.09 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/ithinkso/category/category_cls_ithinkso.asp"-->

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
dim CateSeq0, CateSeq1, CateSeq2, CateSeq3, Depth
dim CateOrder, isusing ,CateSeq, itemidarr ,tmpitemidarr, CateDispSeqarr, tmpCateDispSeqarr
dim mode, menupos, sqlStr, i
	menupos = request("menupos")
	mode = request("mode")
	CateSeq0 = request("CateSeq0")
	CateSeq1 = request("CateSeq1")
	CateSeq2 = request("CateSeq2")
	CateSeq3 = request("CateSeq3")
	Depth = request("Depth")
	itemidarr = request("itemidarr")
	CateDispSeqarr = request("CateDispSeqarr")
	
response.write "MODE : " & mode & "<Br>"

dim referer
	referer = request.ServerVariables("HTTP_REFERER")
		
if mode = "categoryitemreg" then
	if CateSeq0 = "" or CateSeq1 = "" or CateSeq2 = "" then	
		response.write "<script language='javascript'>"
		response.write "	alert('카테고리가 지정되지 않았습니다.');"
		response.write "	location.href='"&referer&"';"
		response.write "</script>"
		dbget.close() : response.end
	end if
	
	if CateSeq3 = "" then CateSeq3 = 0
		
	itemidarr = split(itemidarr,",")
	
	for i = 0 to ubound(itemidarr) -1
		tmpitemidarr = trim(itemidarr(i))
		
		if tmpitemidarr = "" then	
			response.write "<script language='javascript'>"
			response.write "	alert('상품번호가 없습니다.');"
			response.write "	location.href='"&referer&"';"			
			response.write "</script>"
			dbget.close() : response.end
		end if

		sqlStr = "insert into db_item.dbo.tbl_ithinkso_Categoryitem(" + vbcrlf
		sqlStr = sqlStr & " CateTypeSeq, Itemid, CateSeq1, CateSeq2, CateSeq3, IsUsing" + vbcrlf
		sqlStr = sqlStr & " )" + vbcrlf
		sqlStr = sqlStr & " 	select N'"& CateSeq0 &"', N'"& tmpitemidarr &"', N'"& CateSeq1 &"', N'"& CateSeq2 &"', N'"& CateSeq3 &"', N'Y'" + vbcrlf
		sqlStr = sqlStr & " 	from db_item.dbo.tbl_item i" + vbcrlf
		sqlStr = sqlStr & " 	left join db_item.dbo.tbl_ithinkso_Categoryitem ci" + vbcrlf
		sqlStr = sqlStr & " 		on i.itemid = ci.itemid" + vbcrlf
		sqlStr = sqlStr & " 		and ci.IsUsing='Y'" + vbcrlf
		sqlStr = sqlStr & " 		and ci.CateTypeSeq = "&CateSeq0&"" + vbcrlf
		sqlStr = sqlStr & " 		and ci.CateSeq1 = "&CateSeq1&"" + vbcrlf
		sqlStr = sqlStr & " 		and ci.CateSeq2 = "&CateSeq2&"" + vbcrlf
		sqlStr = sqlStr & " 		and ci.CateSeq3 = "&CateSeq3&"" + vbcrlf
		sqlStr = sqlStr & " 	where ci.CateDispSeq is null" + vbcrlf
		sqlStr = sqlStr & " 	and i.itemid in ("&tmpitemidarr&")" + vbcrlf
		
		'response.write sqlStr &"<Br>"   
	    dbget.Execute sqlStr
	next

	response.write "<script language='javascript'>"
	response.write "	alert('OK');"
	response.write "	location.href='"&referer&"';"
	response.write "</script>"
	dbget.close() : response.end

elseif mode = "categoryitemdel" then

	CateDispSeqarr = split(CateDispSeqarr,",")
	
	for i = 0 to ubound(CateDispSeqarr) -1
		tmpCateDispSeqarr = trim(CateDispSeqarr(i))
		
		if tmpCateDispSeqarr = "" then	
			response.write "<script language='javascript'>"
			response.write "	alert('Not Category No.');"
			response.write "	location.href='"&referer&"';"			
			response.write "</script>"
			dbget.close() : response.end
		end if

		sqlStr = "update db_item.dbo.tbl_ithinkso_Categoryitem set" + vbcrlf
		sqlStr = sqlStr & " IsUsing='N'" + vbcrlf
		sqlStr = sqlStr & " ,lastupdate = getdate()" + vbcrlf			
		sqlStr = sqlStr & " where CateDispSeq in ("&tmpCateDispSeqarr&")"
		
		'response.write sqlStr &"<Br>"   
	    dbget.Execute sqlStr
	next

	response.write "<script language='javascript'>"
	response.write "	alert('OK');"
	response.write "	parent.location.href='"&referer&"';"
	response.write "</script>"
	dbget.close() : response.end
	
else
	response.write "<script language='javascript'>"
	response.write "	alert('MODE 구분자가 지정되지 않았습니다.');"
	response.write "	location.href='"&referer&"';"
	response.write "</script>"
	dbget.close() : response.end
end if
%>

</body>
</html>

<!-- #include virtual="/lib/db/dbclose.asp" -->