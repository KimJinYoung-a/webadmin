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
	Dim vQuery, vItemID, vCateCode, vType, vImgURL, vPage, vStartDate, vEventID, vEname, vEsubcopy, vEitemid, vEitemimg, vEicon, vELink, vEDHtml
	vEventID = Request("eventid")
	vType = Request("type")
	vItemID = Request("itemid")
	vCateCode = Request("catecode")
	vPage = Request("page")
	vStartDate = Request("startdate")
	
	vQuery = "SELECT evt_code FROM [db_event].[dbo].[tbl_event] WHERE evt_code = '" & vEventID & "'"
	rsget.Open vQuery, dbget, 1
	
	If not rsget.Eof Then
		rsget.close()
	Else
		Response.Write "<script>alert('없는 이벤트코드입니다.');window.close();</script>"
		rsget.close()
		dbget.close()
		Response.End
	End IF
	'20930 20924
	If vEventID <> "" Then
		vQuery = "select e.evt_name, e.evt_subcopyK, d.etc_itemid, "
		vQuery = vQuery & "case when isNull(d.etc_itemimg,'') = '' then (select icon1image from [db_item].[dbo].[tbl_item] where itemid = d.etc_itemid) else d.etc_itemimg end as etc_itemimg, "
		vQuery = vQuery & "d.issale, d.isgift, d.iscoupon, d.isOnlyTen, d.isoneplusone, d.isfreedelivery, d.isbookingsell, d.iscomment, "
		vQuery = vQuery & "case when d.evt_LinkType = 'I' then d.evt_bannerlink else '/event/eventmain.asp?eventid=" & vEventID & "' end as evt_link "
		vQuery = vQuery & "from db_event.dbo.tbl_event as e inner join db_event.dbo.tbl_event_display as d on e.evt_code = d.evt_code where e.evt_code = '" & vEventID & "'"
		rsget.Open vQuery,dbget,1
		IF not rsget.EOF THEN
			vEname		= db2html(rsget("evt_name"))
			vEsubcopy	= db2html(rsget("evt_subcopyK"))
			vEitemid	= rsget("etc_itemid")
			vEitemimg	= rsget("etc_itemimg")
			If Left(vEitemimg,1) = "S" Then
				vEitemimg = "http://webimage.10x10.co.kr/image/icon1/" & GetImageSubFolderByItemid(vEitemid) & "/" & vEitemimg & ""
			End IF
			vEicon		= fnGetEventIcon(rsget("issale"),rsget("isgift"),rsget("iscoupon"),rsget("isOnlyTen"),rsget("isoneplusone"),rsget("isfreedelivery"),rsget("isbookingsell"),rsget("iscomment"))
			vELink		= rsget("evt_link")
		END IF
		rsget.Close
		
		vQuery = ""
		vQuery = vQuery & "		UPDATE [db_sitemaster].[dbo].[tbl_display_catemain_detail] SET "
		vQuery = vQuery & " 		code = '" & vEventID & "', "
		vQuery = vQuery & " 		title = '" & vEname & "', "
		vQuery = vQuery & " 		subcopy = '" & vEsubcopy & "', "
		vQuery = vQuery & " 		imgurl = '" & vEitemimg & "', "
		vQuery = vQuery & " 		linkurl = '" & vELink & "', "
		vQuery = vQuery & " 		icon = '" & vEicon & "', "
		vQuery = vQuery & " 		reguserid = '" & session("ssBctId") & "', "
		vQuery = vQuery & " 		lastupdate = getdate() "
		vQuery = vQuery & " 	WHERE startdate = '" & vStartDate & "' AND catecode = '" & vCateCode & "' AND type = '" & vType & "' AND page = '" & vPage & "' "
		dbget.execute vQuery
		
		Call fnSaveCateLog(session("ssBctId"),"main","cate="&vCateCode&",startdate="&vStartDate&",type="&vType&",page="&vPage&",수정")
		
		vEDHtml = chrbyte(vEname,48,"Y") & "<br>"
		vEDHtml = vEDHtml & replace(chrbyte(vEsubcopy,156,"Y"),vbCrLf,"<br>")
		vEDHtml = Replace(vEDHtml,chr(34),"'")
	End If
%>
<script>
document.domain = "10x10.co.kr";
opener.$("#<%=vType%>").css("background-image","url(<%=vEitemimg%>)");
opener.$("#<%=vType%>worker").html("<br>마지막작업자:<%=session("ssBctCname")%>");
opener.$("#<%=vType%>description").html("<%=vEDHtml%>");
window.close()
</script>
</head>
<body>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->