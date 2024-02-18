<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateMainCls.asp"-->
<%
	Dim vQuery, vAction, vCateCode, i, vIdx, vStartDate, vEndDate, vImgURL, vLinkURL, vTitle, vSubCopy
	vAction = Request("action")
	vIdx = Request("idx")
	vCateCode = Request("catecode")
	vStartDate = Request("startdate")
	vEndDate = Request("enddate")
	vImgURL = Request("imgurl")
	vLinkURL = Trim(Request("linkurl"))
	vLinkURL = Replace(vLinkURL,"www1","www")
	vTitle = html2db(Trim(Request("title")))
	vSubCopy = html2db(Trim(Request("subcopy")))
	
	If vIdx = "" Then
		vQuery = "INSERT INTO [db_sitemaster].[dbo].[tbl_display_catemain_issue](catecode, imgurl, linkurl, title, subcopy, startdate, enddate, reguserid) "
		vQuery = vQuery & "VALUES('" & vCateCode & "','" & vImgURL & "','" & vLinkURL & "','" & vTitle & "','" & vSubCopy & "','" & vStartDate & "','" & vEndDate & "','" & session("ssBctId") & "')"
		dbget.execute vQuery
		
		Call fnSaveCateLog(session("ssBctId"),"issue","cate=" & vCateCode & ",새글입력")
	Else
		If vAction = "delete" Then
			vQuery = "DELETE [db_sitemaster].[dbo].[tbl_display_catemain_issue] WHERE idx = '" & vIdx & "' "
			dbget.execute vQuery
			
			Call fnSaveCateLog(session("ssBctId"),"issue","cate=" & vCateCode & ",idx=" & vIdx & ",삭제")
		Else
			vQuery = "UPDATE [db_sitemaster].[dbo].[tbl_display_catemain_issue] SET "
			vQuery = vQuery & "		imgurl = '" & vImgURL & "', "
			vQuery = vQuery & "		linkurl = '" & vLinkURL & "', "
			vQuery = vQuery & "		title = '" & vTitle & "', "
			vQuery = vQuery & "		subcopy = '" & vSubCopy & "', "
			vQuery = vQuery & "		startdate = '" & vStartDate & "', "
			vQuery = vQuery & "		enddate = '" & vEndDate & "', "
			vQuery = vQuery & "		reguserid = '" & session("ssBctId") & "' "
			vQuery = vQuery & "WHERE idx = '" & vIdx & "' "
			dbget.execute vQuery
			
			Call fnSaveCateLog(session("ssBctId"),"issue","cate=" & vCateCode & ",idx=" & vIdx & ",수정")
		End If
	End If
	'rw vQuery
%>
<Script>
location.href = "issue_update.asp?catecode=<%=vCateCode%>";
</script>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->