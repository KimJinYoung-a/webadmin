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
	Dim vQuery, vItemID, vCateCode, vType, vImgURL, vPage, vStartDate
	vType = Request("type")
	vItemID = Trim(Request("itemid"))
	vCateCode = Request("catecode")
	vPage = Request("page")
	vStartDate = Request("startdate")
	
	If isNumeric(vItemID) = False Then
		Response.Write "<script>alert('없는 상품코드입니다.');window.close();</script>"
		dbget.close()
		Response.End
	End IF
	
	vQuery = "SELECT icon1image FROM [db_item].[dbo].[tbl_item] WHERE itemid = '" & vItemID & "'"
	rsget.Open vQuery, dbget, 1
	
	If not rsget.Eof Then
		vImgURL = "http://webimage.10x10.co.kr/image/icon1/" & GetImageSubFolderByItemid(vItemID) & "/" & rsget("icon1image") & ""
		rsget.close()
	Else
		Response.Write "<script>alert('없는 상품코드입니다.');window.close();</script>"
		rsget.close()
		dbget.close()
		Response.End
	End IF
	
	If vItemID <> "" Then
		vQuery = ""
		vQuery = vQuery & "		UPDATE [db_sitemaster].[dbo].[tbl_display_catemain_detail] SET "
		vQuery = vQuery & " 		code = '" & vItemID & "', "
		vQuery = vQuery & " 		imgurl = '" & vImgURL & "', "
		vQuery = vQuery & " 		reguserid = '" & session("ssBctId") & "', "
		vQuery = vQuery & " 		lastupdate = getdate() "
		vQuery = vQuery & " 	WHERE startdate = '" & vStartDate & "' AND catecode = '" & vCateCode & "' AND type = '" & vType & "' AND page = '" & vPage & "' "
		dbget.execute vQuery
		
		Call fnSaveCateLog(session("ssBctId"),"main","cate="&vCateCode&",startdate="&vStartDate&",type="&vType&",page="&vPage&",수정")
		
	End If
%>
<script>
document.domain = "10x10.co.kr";
opener.$("#<%=vType%>").css("background-image","url(<%=vImgURL%>)");
opener.$("#<%=vType%>worker").html("<br>마지막작업자:<%=session("ssBctCname")%>");
window.close()
</script>
</head>
<body>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->