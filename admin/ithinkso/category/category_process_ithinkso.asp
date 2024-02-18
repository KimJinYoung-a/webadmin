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
dim iCateSeq0, iCateSeq1, iCateSeq2, iCateSeq3, Depth, regDepth
dim CateName, CateOrder, isusing ,CateSeq, tmpCateName, tmpCateSeq, tmpCateOrder, tmpisusing
dim newCateName, newCateOrder, newisusing
dim mode, menupos, sqlStr, i
	menupos = request("menupos")
	mode = request("mode")
	iCateSeq0 = request("iCateSeq0")
	iCateSeq1 = request("iCateSeq1")
	iCateSeq2 = request("iCateSeq2")
	iCateSeq3 = request("iCateSeq3")
	Depth = request("Depth")
	regDepth = request("regDepth")
	CateName = request("CateName")
	CateOrder = request("CateOrder")
	isusing = request("isusing")
	CateSeq = request("tCateSeq")
	newCateName = request("newCateName")
	newCateOrder = request("newCateOrder")
	newisusing = request("newisusing")
	
response.write "MODE : " & mode & "<Br>"

dim referer
	referer = request.ServerVariables("HTTP_REFERER")
		
if mode = "categoryedit" then
	if regDepth = "" then	
		response.write "<script language='javascript'>"
		response.write "	alert('카테고리 구분이 지정되지 않았습니다.');"
		response.write "</script>"
		dbget.close() : response.end
	end if
	
	CateSeq = split(CateSeq,",")
	CateName = split(CateName,",")
	CateOrder = split(CateOrder,",")
	isusing = split(isusing,",")		
	
	for i = 0 to ubound(CateSeq)
		tmpCateSeq = trim(CateSeq(i))
		tmpCateName = trim(CateName(i))
		tmpCateOrder = trim(CateOrder(i))
		tmpisusing = trim(isusing(i))
		
		if tmpCateSeq = "" then	
			response.write "<script language='javascript'>"
			response.write "	alert('카테고리 번호가 지정되지 않았습니다.');"
			response.write "</script>"
			dbget.close() : response.end
		end if
		if tmpCateName = "" then	
			response.write "<script language='javascript'>"
			response.write "	alert('카테고리명이 지정되지 않았습니다.');"
			response.write "</script>"
			dbget.close() : response.end
		end if			
		if tmpCateOrder = "" then	
			response.write "<script language='javascript'>"
			response.write "	alert('정렬순서가 지정되지 않았습니다.');"
			response.write "</script>"
			dbget.close() : response.end
		end if	
		if tmpisusing = "" then	
			response.write "<script language='javascript'>"
			response.write "	alert('사용여부가 지정되지 않았습니다.');"
			response.write "</script>"
			dbget.close() : response.end
		end if
		
		'/카테고리분류
		IF regDepth = 0 THEN		
			sqlStr = "update db_item.dbo.tbl_ithinkso_CategoryType set" + vbcrlf
			sqlStr = sqlStr & " IsUsing = N'"& tmpisusing &"'" + vbcrlf
			sqlStr = sqlStr & " ,CateTypeName = N'"& tmpCateName &"'" + vbcrlf
			sqlStr = sqlStr & " ,CateTypeOrder = N'"& tmpCateOrder &"'" + vbcrlf
			sqlStr = sqlStr & " where CateTypeSeq = N'"& tmpCateSeq &"'"
	
			'response.write sqlStr &"<Br>"   
		    dbget.Execute sqlStr

		'/대카테
		elseIF regDepth = 1 THEN
			sqlStr = "update dbo.tbl_ithinkso_CategoryInfo set" + vbcrlf
			sqlStr = sqlStr & " IsUsing = N'"& tmpisusing &"'" + vbcrlf
			sqlStr = sqlStr & " ,CateName = N'"& tmpCateName &"'" + vbcrlf
			sqlStr = sqlStr & " ,CateOrder = N'"& tmpCateOrder &"'" + vbcrlf
			sqlStr = sqlStr & " ,lastupdate = getdate()" + vbcrlf
			sqlStr = sqlStr & " where CateSeq = N'"& tmpCateSeq &"'"
	
			'response.write sqlStr &"<Br>"   
		    dbget.Execute sqlStr

		'/중카테
		elseIF regDepth = 2 THEN
			sqlStr = "update dbo.tbl_ithinkso_CategoryInfo set" + vbcrlf
			sqlStr = sqlStr & " IsUsing = N'"& tmpisusing &"'" + vbcrlf
			sqlStr = sqlStr & " ,CateName = N'"& tmpCateName &"'" + vbcrlf			
			sqlStr = sqlStr & " ,CateOrder = N'"& tmpCateOrder &"'" + vbcrlf
			sqlStr = sqlStr & " ,lastupdate = getdate()" + vbcrlf			
			sqlStr = sqlStr & " where CateSeq = N'"& tmpCateSeq &"'"
	
			'response.write sqlStr &"<Br>"   
		    dbget.Execute sqlStr

		'/소카테
		elseIF regDepth = 3 THEN
			sqlStr = "update dbo.tbl_ithinkso_CategoryInfo set" + vbcrlf
			sqlStr = sqlStr & " IsUsing = N'"& tmpisusing &"'" + vbcrlf
			sqlStr = sqlStr & " ,CateName = N'"& tmpCateName &"'" + vbcrlf			
			sqlStr = sqlStr & " ,CateOrder = N'"& tmpCateOrder &"'" + vbcrlf
			sqlStr = sqlStr & " ,lastupdate = getdate()" + vbcrlf			
			sqlStr = sqlStr & " where CateSeq = N'"& tmpCateSeq &"'"
	
			'response.write sqlStr &"<Br>"   
		    dbget.Execute sqlStr
		    		    
		else
			response.write "<script language='javascript'>"
			response.write "	alert('카테고리 구분이 잘못되었습니다.');"
			response.write "</script>"
			dbget.close() : response.end		
		end if
	next

	response.write "<script language='javascript'>"
	response.write "	alert('OK');"
	response.write "	parent.location.reload();"
	response.write "</script>"
	dbget.close() : response.end

elseif mode = "categoryreg" then

	if regDepth = "" then	
		response.write "<script language='javascript'>"
		response.write "	alert('카테고리 구분이 지정되지 않았습니다.');"
		response.write "</script>"
		dbget.close() : response.end
	end if
	
	newCateName = trim(newCateName)
	newCateOrder = trim(newCateOrder)
	newisusing = trim(newisusing)
		
	if newCateName = "" then	
		response.write "<script language='javascript'>"
		response.write "	alert('카테고리명이 지정되지 않았습니다.');"
		response.write "</script>"
		dbget.close() : response.end
	end if
	if newisusing = "" then	
		response.write "<script language='javascript'>"
		response.write "	alert('사용여부가 지정되지 않았습니다.');"
		response.write "</script>"
		dbget.close() : response.end
	end if	
	if newCateOrder = "" then	
		response.write "<script language='javascript'>"
		response.write "	alert('정렬순서가 지정되지 않았습니다.');"
		response.write "</script>"
		dbget.close() : response.end
	end if

	'/카테고리분류
	IF regDepth = 0 THEN		
		sqlStr = "insert into db_item.dbo.tbl_ithinkso_CategoryType(" + vbcrlf
		sqlStr = sqlStr & " CateTypeName, CateTypeOrder, IsUsing" + vbcrlf
		sqlStr = sqlStr & " ) values (" + vbcrlf
		sqlStr = sqlStr & " N'"& newCateName &"', N'"& newCateOrder &"', N'Y'" + vbcrlf
		sqlStr = sqlStr & " )"

		'response.write sqlStr &"<Br>"   
	    dbget.Execute sqlStr

	'/대카테
	elseIF regDepth = 1 THEN
		sqlStr = "insert into db_item.dbo.tbl_ithinkso_CategoryInfo(" + vbcrlf
		sqlStr = sqlStr & " CateTypeSeq, CateName, subCateSeq1, subCateSeq2, Depth, CateOrder, IsUsing" + vbcrlf
		sqlStr = sqlStr & " ) values (" + vbcrlf
		sqlStr = sqlStr & " N'"& iCateSeq0 &"', N'"& html2db(newCateName) &"', N'0', N'0', N'"& regDepth &"', N'"& newCateOrder &"', N'Y'" + vbcrlf
		sqlStr = sqlStr & " )"

		'response.write sqlStr &"<Br>"   
	    dbget.Execute sqlStr

	'/중카테
	elseIF regDepth = 2 THEN
		sqlStr = "insert into db_item.dbo.tbl_ithinkso_CategoryInfo(" + vbcrlf
		sqlStr = sqlStr & " CateTypeSeq, CateName, subCateSeq1, subCateSeq2, Depth, CateOrder, IsUsing" + vbcrlf
		sqlStr = sqlStr & " ) values (" + vbcrlf
		sqlStr = sqlStr & " N'"& iCateSeq0 &"', N'"& html2db(newCateName) &"', N'"& iCateSeq1 &"', N'0', N'"& regDepth &"', N'"& newCateOrder &"', N'Y'" + vbcrlf
		sqlStr = sqlStr & " )"

		'response.write sqlStr &"<Br>"   
	    dbget.Execute sqlStr
	    
	'/소카테
	elseIF regDepth = 3 THEN
		sqlStr = "insert into db_item.dbo.tbl_ithinkso_CategoryInfo(" + vbcrlf
		sqlStr = sqlStr & " CateTypeSeq, CateName, subCateSeq1, subCateSeq2, Depth, CateOrder, IsUsing" + vbcrlf
		sqlStr = sqlStr & " ) values (" + vbcrlf
		sqlStr = sqlStr & " N'"& iCateSeq0 &"', N'"& html2db(newCateName) &"', N'"& iCateSeq1 &"', N'"& iCateSeq2 &"', N'"& regDepth &"', N'"& newCateOrder &"', N'Y'" + vbcrlf
		sqlStr = sqlStr & " )"

		'response.write sqlStr &"<Br>"   
	    dbget.Execute sqlStr	    
	else
		response.write "<script language='javascript'>"
		response.write "	alert('카테고리 구분이 잘못되었습니다.');"
		response.write "</script>"
		dbget.close() : response.end		
	end if

	response.write "<script language='javascript'>"
	response.write "	alert('OK');"
	response.write "	parent.location.reload();"
	response.write "</script>"
	dbget.close() : response.end
	
else
	response.write "<script language='javascript'>"
	response.write "	alert('MODE 구분자가 지정되지 않았습니다.');"
	response.write "</script>"
	dbget.close() : response.end
end if
%>

</body>
</html>

<!-- #include virtual="/lib/db/dbclose.asp" -->