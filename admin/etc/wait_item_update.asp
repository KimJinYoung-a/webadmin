<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/tingcls.asp"-->
<%
dim iting,oneitem
dim itemidlist, tingpointlist, tingpoint_blist, userclasslist
dim limitdivlist,limitealist,selectitemlist

itemidlist = request.Form("itemidlist")
tingpointlist = request.Form("tingpointlist")
tingpoint_blist = request.Form("tingpoint_blist")
userclasslist = request.Form("userclasslist")
limitdivlist = request.Form("limitdivlist")
limitealist = request.Form("limitealist")
selectitemlist = request.Form("selectitemlist")

set iting = new CWaititemUpdate

dim i

	itemidlist = split(itemidlist,"|")
	tingpointlist = split(tingpointlist,"|")
	tingpoint_blist = split(tingpoint_blist,"|")
	userclasslist = split(userclasslist,"|")
	limitdivlist = split(limitdivlist,"|")
	limitealist = split(limitealist,"|")
	selectitemlist = split(selectitemlist,"|")
	
	for i = lBound(itemidlist)+1 to Ubound(itemidlist)
		iting.UpdateItem itemidlist(i),tingpointlist(i),tingpoint_blist(i),userclasslist(i),limitdivlist(i),limitealist(i),selectitemlist(i)
	next

set iting = Nothing
set oneitem = Nothing

dim refer
refer = request.ServerVariables("HTTP_REFERER")
%>

<script language="javascript">
alert('수정되었습니다.');
location.replace('wait_item_list.asp');
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->