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
	Dim vQuery, vCateCode, vType, vPage, vStartDate, vIdx, i, sql, vEname, vEsubcopy, vEitemid, vEitemimg, vEicon, vELink, vWorkComment
	vType = Request("type")
	vCateCode = Request("catecode")
	vPage = Request("page")
	vStartDate = Request("startdate")
	vWorkComment = html2db(Request("workcomment"))
	

	vQuery = " SELECT count(idx) FROM [db_sitemaster].[dbo].[tbl_display_catemain_detail] WHERE startdate = '" & vStartDate & "' AND catecode = '" & vCateCode & "' AND page = '" & vPage & "' "
	rsget.Open vQuery,dbget
	IF rsget(0) > 0 THEN
		Response.Write "<script>alert('�ش� �ݿ��Ͽ� �̹� ����� �������� �ֽ��ϴ�.');history.back();</script>"
		rsget.close
		dbget.close()
		Response.End
	Else
		rsget.close
	END IF
	
	
	If vPage <> "1" Then
		'Multi üũ
		vQuery = " SELECT count(idx) FROM [db_sitemaster].[dbo].[tbl_display_catemain_detail] WHERE startdate = '" & vStartDate & "' AND catecode = '" & vCateCode & "' AND type = 'multiimg1' AND page = '1' "
		rsget.Open vQuery,dbget
		IF rsget(0) > 0 THEN
			rsget.close
		Else
			Response.Write "<script>alert('�ش� �ݿ��� 1 �������� multiimg �� �����ϴ�.\n�� �����Ͱ� �־�� 2~5�������� ���� �� �ֽ��ϴ�.');history.back();</script>"
			rsget.close
			dbget.close()
			Response.End
		END IF
		
		
		'Book üũ
		vQuery = " SELECT count(idx) FROM [db_sitemaster].[dbo].[tbl_display_catemain_detail] WHERE startdate = '" & vStartDate & "' AND catecode = '" & vCateCode & "' AND type = 'book' AND page = '1' "
		rsget.Open vQuery,dbget
		IF rsget(0) > 0 THEN
			rsget.close
		Else
			Response.Write "<script>alert('�ش� �ݿ��� 1 �������� book �� �����ϴ�.\n�� �����Ͱ� �־�� 2~5�������� ���� �� �ֽ��ϴ�.');history.back();</script>"
			rsget.close
			dbget.close()
			Response.End
		END IF
		
		
		'Recipe üũ
		vQuery = " SELECT count(idx) FROM [db_sitemaster].[dbo].[tbl_display_catemain_detail] WHERE startdate = '" & vStartDate & "' AND catecode = '" & vCateCode & "' AND type = 'recipe' AND page = '1' "
		rsget.Open vQuery,dbget
		IF rsget(0) > 0 THEN
			rsget.close
		Else
			Response.Write "<script>alert('�ش� �ݿ��� 1 �������� recipe �� �����ϴ�.\n�� �����Ͱ� �־�� 2~5�������� ���� �� �ֽ��ϴ�.');history.back();</script>"
			rsget.close
			dbget.close()
			Response.End
		END IF
	END IF
	
	
	vQuery = "INSERT INTO [db_sitemaster].[dbo].[tbl_display_catemain](catecode, page, startdate, workcomment, reguserid) "
	vQuery = vQuery & " VALUES('" & vCateCode & "','" & vPage & "','" & vStartDate & "','" & vWorkComment & "','" & session("ssBctId") & "')"
	dbget.execute vQuery
	
	
	vQuery = ""
	If vPage = "1" Then
		vQuery = vQuery & " INSERT INTO [db_sitemaster].[dbo].[tbl_display_catemain_detail](catecode, page, type, imgurl, linkurl, startdate, reguserid, regdate) "
		vQuery = vQuery & "	VALUES('" & vCateCode & "','" & vPage & "','multiimg1',null,null,'" & vStartDate & "','" & session("ssBctId") & "',getdate()) "
		
		vQuery = vQuery & " INSERT INTO [db_sitemaster].[dbo].[tbl_display_catemain_detail](catecode, page, type, imgurl, linkurl, startdate, reguserid, regdate) "
		vQuery = vQuery & "	VALUES('" & vCateCode & "','" & vPage & "','multiimg2',null,null,'" & vStartDate & "','" & session("ssBctId") & "',getdate()) "
		
		vQuery = vQuery & " INSERT INTO [db_sitemaster].[dbo].[tbl_display_catemain_detail](catecode, page, type, imgurl, linkurl, startdate, reguserid, regdate) "
		vQuery = vQuery & "	VALUES('" & vCateCode & "','" & vPage & "','multiimg3',null,null,'" & vStartDate & "','" & session("ssBctId") & "',getdate()) "
		dbget.execute vQuery
	Else
		'### 2~5 �������� 1�������� ���� �״�� �Է�. 1������ �̸����⿡�� �Ǽ��� �����ϱ� �Ҷ� 2~5������ ���� �Ȱ����ؼ� UPDATE.
		vQuery = vQuery & " INSERT INTO [db_sitemaster].[dbo].[tbl_display_catemain_detail](catecode, page, type, imgurl, linkurl, startdate, reguserid, regdate) "
		vQuery = vQuery & "	SELECT '" & vCateCode & "','" & vPage & "','multiimg1',null,null,'" & vStartDate & "','" & session("ssBctId") & "',getdate() "
		vQuery = vQuery & "	FROM [db_sitemaster].[dbo].[tbl_display_catemain_detail] WHERE startdate = '" & vStartDate & "' AND catecode = '" & vCateCode & "' AND type = 'multiimg1' AND page = '1' "
		
		vQuery = vQuery & " INSERT INTO [db_sitemaster].[dbo].[tbl_display_catemain_detail](catecode, page, type, imgurl, linkurl, startdate, reguserid, regdate) "
		vQuery = vQuery & "	SELECT '" & vCateCode & "','" & vPage & "','multiimg2',null,null,'" & vStartDate & "','" & session("ssBctId") & "',getdate() "
		vQuery = vQuery & "	FROM [db_sitemaster].[dbo].[tbl_display_catemain_detail] WHERE startdate = '" & vStartDate & "' AND catecode = '" & vCateCode & "' AND type = 'multiimg2' AND page = '1' "
		
		vQuery = vQuery & " INSERT INTO [db_sitemaster].[dbo].[tbl_display_catemain_detail](catecode, page, type, imgurl, linkurl, startdate, reguserid, regdate) "
		vQuery = vQuery & "	SELECT '" & vCateCode & "','" & vPage & "','multiimg3',null,null,'" & vStartDate & "','" & session("ssBctId") & "',getdate() "
		vQuery = vQuery & "	FROM [db_sitemaster].[dbo].[tbl_display_catemain_detail] WHERE startdate = '" & vStartDate & "' AND catecode = '" & vCateCode & "' AND type = 'multiimg3' AND page = '1' "
		dbget.execute vQuery
	End If
	
	
	vQuery = ""
	For i=1 To 12
		vQuery = vQuery & " INSERT INTO [db_sitemaster].[dbo].[tbl_display_catemain_detail](catecode, page, type, code, imgurl, linkurl, startdate, reguserid, regdate) "
		vQuery = vQuery & "	VALUES('" & vCateCode & "','" & vPage & "','item"&i&"',null,null,null,'" & vStartDate & "','" & session("ssBctId") & "',getdate()) "
	Next
	dbget.execute vQuery
	
	
	vQuery = ""
	For i=1 To 4
		vQuery = vQuery & " INSERT INTO [db_sitemaster].[dbo].[tbl_display_catemain_detail](catecode, page, type, code, title, subcopy, imgurl, linkurl, icon, startdate, reguserid, regdate) "
		vQuery = vQuery & "	VALUES('" & vCateCode & "','" & vPage & "','event"&i&"',null,null,null,null,null,null,'" & vStartDate & "','" & session("ssBctId") & "',getdate()) "
	Next
	dbget.execute vQuery
	
	
	If vPage = "1" Then
		vQuery = ""
		vQuery = vQuery & " INSERT INTO [db_sitemaster].[dbo].[tbl_display_catemain_detail](catecode, page, type, imgurl, linkurl, startdate, reguserid, regdate) "
		vQuery = vQuery & "	VALUES('" & vCateCode & "','" & vPage & "','book',null,null,'" & vStartDate & "','" & session("ssBctId") & "',getdate()) "
		dbget.execute vQuery
		
		
		vQuery = ""
		vQuery = vQuery & " INSERT INTO [db_sitemaster].[dbo].[tbl_display_catemain_detail](catecode, page, type, imgurl, linkurl, startdate, reguserid, regdate) "
		vQuery = vQuery & "	VALUES('" & vCateCode & "','" & vPage & "','recipe',null,null,'" & vStartDate & "','" & session("ssBctId") & "',getdate()) "
		dbget.execute vQuery
	Else
		'### 2~5 �������� 1�������� ���� �״�� �Է�. 1������ �̸����⿡�� �Ǽ��� �����ϱ� �Ҷ� 2~5������ ���� �Ȱ����ؼ� UPDATE.
		vQuery = ""
		vQuery = vQuery & " INSERT INTO [db_sitemaster].[dbo].[tbl_display_catemain_detail](catecode, page, type, imgurl, linkurl, startdate, reguserid, regdate) "
		vQuery = vQuery & "	SELECT '" & vCateCode & "','" & vPage & "','book',null,null,'" & vStartDate & "','" & session("ssBctId") & "',getdate() "
		vQuery = vQuery & "	FROM [db_sitemaster].[dbo].[tbl_display_catemain_detail] WHERE startdate = '" & vStartDate & "' AND catecode = '" & vCateCode & "' AND type = 'book' AND page = '1' "
		dbget.execute vQuery
		
		
		vQuery = ""
		vQuery = vQuery & " INSERT INTO [db_sitemaster].[dbo].[tbl_display_catemain_detail](catecode, page, type, imgurl, linkurl, startdate, reguserid, regdate) "
		vQuery = vQuery & "	SELECT '" & vCateCode & "','" & vPage & "','recipe',null,null,'" & vStartDate & "','" & session("ssBctId") & "',getdate() "
		vQuery = vQuery & "	FROM [db_sitemaster].[dbo].[tbl_display_catemain_detail] WHERE startdate = '" & vStartDate & "' AND catecode = '" & vCateCode & "' AND type = 'recipe' AND page = '1' "
		dbget.execute vQuery
	End If
	
	Call fnSaveCateLog(session("ssBctId"),"main","cate="&vCateCode&",startdate="&vStartDate&",type="&vType&",page="&vPage&",�űԻ���")
%>
<script>
//location.href = "templete.asp?catecode=<%=vCateCode%>&page=<%=vPage%>&startdate=<%=vStartDate%>";
parent.location.reload();
</script>
</head>
<body>
</body>
</html>

<!-- #include virtual="/lib/db/dbclose.asp" -->