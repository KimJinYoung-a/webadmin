<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<html>
<head>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<%
	Dim vQuery, vMode, vCateCode, vType, vPage, vStartDate, vStartDateB, vIdx, i, sql, vEname, vEsubcopy, vEitemid, vEitemimg, vEicon, vELink, vWorkComment
	vMode = Request("mode")
	vType = Request("type")
	vCateCode = Request("catecode")
	vPage = Request("page")
	vStartDate = Request("startdate")
	vStartDateB = Request("startdate_before")
	vWorkComment = html2db(Request("workcomment"))
	
	If vMode = "delete" Then
		vQuery = " DELETE [db_sitemaster].[dbo].[tbl_display_catemain] WHERE catecode = '" & vCateCode & "' AND page = '" & vPage & "' AND startdate = '" & vStartDateB & "' "
		vQuery = vQuery & " DELETE [db_sitemaster].[dbo].[tbl_display_catemain_detail] WHERE catecode = '" & vCateCode & "' AND page = '" & vPage & "' AND startdate = '" & vStartDateB & "' "
		dbget.execute vQuery
		
		Call fnSaveCateLog(session("ssBctId"),"main","cate="&vCateCode&",startdate="&vStartDate&",type="&vType&",page="&vPage&",페이지삭제")
	Else
		If CStr(vStartDate) <> CStr(vStartDateB) Then
			vQuery = " SELECT count(idx) FROM [db_sitemaster].[dbo].[tbl_display_catemain_detail] WHERE startdate = '" & vStartDate & "' AND catecode = '" & vCateCode & "' AND page = '" & vPage & "' "
			rsget.Open vQuery,dbget
			IF rsget(0) > 0 THEN
				Response.Write "<script>alert('해당 반영일에 이미 저장된 페이지가 있습니다.');history.back();</script>"
				rsget.close
				dbget.close()
				Response.End
			Else
				rsget.close
				vQuery = "UPDATE [db_sitemaster].[dbo].[tbl_display_catemain_detail] "
				vQuery = vQuery & "		SET startdate = '" & vStartDate & "' "
				vQuery = vQuery & " WHERE catecode = '" & vCateCode & "' AND page = '" & vPage & "' AND startdate = '" & vStartDateB & "' "
				
				Call fnSaveCateLog(session("ssBctId"),"main","cate="&vCateCode&",startdate="&vStartDate&",type="&vType&",page="&vPage&","&vStartDateB&"->"&vStartDate&"날짜변경")
			END IF
			
		End If
		
		vQuery = vQuery & " UPDATE [db_sitemaster].[dbo].[tbl_display_catemain] "
		vQuery = vQuery & "		SET startdate = '" & vStartDate & "', workcomment = '" & vWorkComment & "', reguserid = '" & session("ssBctId") & "' "
		vQuery = vQuery & " WHERE catecode = '" & vCateCode & "' AND page = '" & vPage & "' AND startdate = '" & vStartDateB & "' "
		dbget.execute vQuery
		
		Call fnSaveCateLog(session("ssBctId"),"main","cate="&vCateCode&",startdate="&vStartDate&",type="&vType&",page="&vPage&",코멘트 수정")
	End If
%>
<script>
<% If vMode = "delete" Then %>
	parent.location.reload();
<% Else %>
	<% If CStr(vStartDate) <> CStr(vStartDateB) Then %>
	parent.location.reload();
	<% Else %>
	location.href = "templete.asp?catecode=<%=vCateCode%>&page=<%=vPage%>&startdate=<%=vStartDate%>";
	<% End If %>
<% End If %>
</script>
</head>
<body>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->