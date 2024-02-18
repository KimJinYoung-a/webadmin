<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/sitemasterclass/main_enjoyContentsManageCls.asp" -->
<%
Dim Itemid
	Itemid = requestCheckVar(request("Itemid"),10)

	response.write GetItemImageLoad(Itemid)
	response.end
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->