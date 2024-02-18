<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<html>
<head>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<%
	Dim vQuery, i, vCateCode, vCount, vItemID, vImgLink, vType, vLink, vWWW
	vCateCode 	= Request("catecode")
	vCount		= Request("cnt")
	vItemID		= Request("itemid")
	vImgLink	= Request("imglink")
	
	vQuery = ""
	For i = 1 To vCount Step 1
		If vCateCode = "106" Then
			vType = "category"
		Else
			If (vCount-3) < i Then
				vType = "brand"
			Else
				vType = "category"
			End If
		End If
		
		vQuery = vQuery & "INSERT INTO [db_item].[dbo].[tbl_display_cate_menu](catecode, type, number, value, useyn) "
		vQuery = vQuery & "VALUES('" & vCateCode & "', '" & vType & "', '" & i & "', '" & Request("cate"&i&"code") & "', 'y') " & vbCrLf
	Next
	
	If vItemID <> "" Then
		vQuery = vQuery & "INSERT INTO [db_item].[dbo].[tbl_display_cate_menu](catecode, type, number, value, useyn) "
		vQuery = vQuery & "VALUES('" & vCateCode & "', 'bookitemid', '" & i & "', '"&vItemID&"', 'y') " & vbCrLf
		
		vQuery = vQuery & "INSERT INTO [db_item].[dbo].[tbl_display_cate_menu](catecode, type, number, value, useyn) "
		vQuery = vQuery & "VALUES('" & vCateCode & "', 'bookimg', '" & i+1 & "', '" & vImgLink & "', 'y') " & vbCrLf
	End IF
	
	If vQuery <> "" Then
		vQuery = "UPDATE [db_item].[dbo].[tbl_display_cate_menu] SET useyn = 'n' WHERE catecode = '" & vCateCode & "' " & vQuery
	End IF
	
	dbget.execute vQuery
	
	Call fnSaveCateLog(session("ssBctId"),"menu","cate=" & vCateCode & ",메뉴변경")
%>

<script>
parent.location.reload();
</script>
</head>
<body>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->