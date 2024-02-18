<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateMainCls.asp"-->
<html>
<head>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<%
	Dim vQuery, vItemID, vCateCode, vType, vImgURL, vPage, vStartDate, vRecipeID, vRname, vRitemimg, vEicon, vRLink
	vRecipeID = Request("recipeid")
	vType = Request("type")
	vItemID = Request("itemid")
	vCateCode = Request("catecode")
	vPage = Request("page")
	vStartDate = Request("startdate")
	
	
	
	'####### 이후 라인부터 수정해야함.
	
	'####### Recipe 키값에 따라 조회해야하므로 PLAY 작업후 반드시 아래 수정.
	vQuery = "SELECT evt_code FROM [db_event].[dbo].[tbl_event] WHERE evt_code = '" & vRecipeID & "'"
	rsget.Open vQuery, dbget, 1
	
	If not rsget.Eof Then
		rsget.close()
	Else
		Response.Write "<script>alert('없는 Recipe 코드입니다.');window.close();</script>"
		rsget.close()
		dbget.close()
		Response.End
	End IF
	
	If vRecipeID <> "" Then
		
		'####### Recipe 키값에 따라 조회해야하므로 PLAY 작업후 반드시 아래 수정.
		
		vQuery = "select e.evt_name, e.evt_subcopyK, d.etc_itemid, "
		vQuery = vQuery & "case when isNull(d.etc_itemimg,'') = '' then (select icon1image from [db_item].[dbo].[tbl_item] where itemid = d.etc_itemid) else d.etc_itemimg end as etc_itemimg, "
		vQuery = vQuery & "d.issale, d.isgift, d.iscoupon, d.isOnlyTen, d.isoneplusone, d.isfreedelivery, d.isbookingsell, d.iscomment, "
		vQuery = vQuery & "case when d.evt_LinkType = 'I' then d.evt_bannerlink else '/event/eventmain.asp?eventid=" & vEventID & "' end as evt_link "
		vQuery = vQuery & "from db_event.dbo.tbl_event as e inner join db_event.dbo.tbl_event_display as d on e.evt_code = d.evt_code where e.evt_code = '" & vEventID & "'"
		rsget.Open vQuery,dbget,1
		IF not rsget.EOF THEN
			vRname		= db2html(rsget("evt_name"))
			vEsubcopy	= db2html(rsget("evt_subcopyK"))
			vEitemid	= rsget("etc_itemid")
			vRLink		= rsget("evt_link")
		END IF
		rsget.Close
		
		vQuery = ""
		vQuery = vQuery & "		UPDATE [db_sitemaster].[dbo].[tbl_display_catemain_detail] SET "
		vQuery = vQuery & " 		code = '" & vRecipeID & "', "
		vQuery = vQuery & " 		title = '" & vRname & "', "
		vQuery = vQuery & " 		imgurl = '" & vRitemimg & "', "
		vQuery = vQuery & " 		linkurl = '" & vRLink & "', "
		vQuery = vQuery & " 		icon = '" & vEicon & "', "
		vQuery = vQuery & " 		reguserid = '" & session("ssBctId") & "', "
		vQuery = vQuery & " 		lastupdate = getdate() "
		vQuery = vQuery & " 	WHERE startdate = '" & vStartDate & "' AND catecode = '" & vCateCode & "' AND type = '" & vType & "' "
		dbget.execute vQuery
		
		Call fnSaveCateLog(session("ssBctId"),"main","cate="&vCateCode&",startdate="&vStartDate&",type="&vType&",page="&vPage&",수정")

	End If
%>
<script>
document.domain = "10x10.co.kr";
opener.$("#<%=vType%>").css("background-image","url(<%=vRitemimg%>)");
opener.$("#<%=vType%>worker").html("<br>마지막작업자:<%=session("ssBctCname")%>");
window.close()
</script>
</head>
<body>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->