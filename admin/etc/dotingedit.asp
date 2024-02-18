<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
dim refer
refer = request.ServerVariables("HTTP_REFERER")

dim id, itemid, userclass, limitdiv
dim tingpoint, tingpoint_b, limitea, limitsell
dim isusing, sellyn
dim mode
dim eventdiv, eventcpcode

id = request("id")
itemid = request("itemid")
userclass = request("userclass")
limitdiv = request("limitdiv")
tingpoint = request("tingpoint")
tingpoint_b = request("tingpoint_b")
limitea = request("limitea")
limitsell = request("limitsell")
isusing = request("isusing")
sellyn = request("sellyn")
mode = request("mode")
eventdiv = request("eventdiv")
eventcpcode = request("eventcpcode")

'response.write "itemid=" + itemid + "<br>"
'response.write "userclass=" + userclass + "<br>"
'response.write "limitdiv=" + limitdiv + "<br>"
'response.write "tingpoint=" + tingpoint + "<br>"
'response.write "tingpoint_b=" + tingpoint_b + "<br>"
'response.write "limitea=" + limitea + "<br>"
'response.write "limitsell=" + limitsell + "<br>"
'response.write "isusing=" + isusing + "<br>"
'response.write "sellyn=" + sellyn + "<br>"

dim sqlStr
dim existsitemid

if (mode="add") then
	''중복체크.
	sqlStr = "select top 1 itemid from [db_ting].[dbo].tbl_new_ting_item"
	sqlStr = sqlStr + " where itemid=" + itemid

	rsget.Open sqlStr,dbget,1
	existsitemid = (rsget.RecordCount>0)
	rsget.close

	if existsitemid then
%>
		<script language='javascript'>
			alert('이미 존재하는 상품입니다.');
			location.replace('<%= refer %>');
		</script>
<%
		dbget.close()	:	response.End
	end if

	sqlStr = "insert into [db_ting].[dbo].tbl_new_ting_item(itemid,"
	sqlStr = sqlStr + " tingpoint, tingpoint_b, limitdiv, limitea,"
	sqlStr = sqlStr + " limitsell, "
	sqlStr = sqlStr + " userclass, eventdiv, eventcpcd, isusing, sellyn)"
	sqlStr = sqlStr + " values("
	sqlStr = sqlStr + " " + itemid + ","
	sqlStr = sqlStr + " " + tingpoint + ","
	sqlStr = sqlStr + " " + tingpoint_b + ","
	sqlStr = sqlStr + " '" + limitdiv + "',"
	sqlStr = sqlStr + " " + limitea + ","
	sqlStr = sqlStr + " " + limitsell + ","
	sqlStr = sqlStr + " '" + userclass + "',"
	sqlStr = sqlStr + " '" + eventdiv + "',"
	sqlStr = sqlStr + " '" + eventcpcode + "',"
	sqlStr = sqlStr + " '" + isusing + "',"
	sqlStr = sqlStr + " '" + sellyn + "'"
	sqlStr = sqlStr + " )"
elseif (mode="edit") then
	sqlStr = " update [db_ting].[dbo].tbl_new_ting_item"
	sqlStr = sqlStr + " set itemid=" + itemid + ","
	sqlStr = sqlStr + " tingpoint=" + tingpoint + ","
	sqlStr = sqlStr + " tingpoint_b=" + tingpoint_b + ","
	sqlStr = sqlStr + " limitdiv='" + limitdiv + "',"
	sqlStr = sqlStr + " limitea=" + limitea + ","
	sqlStr = sqlStr + " limitsell=" + limitsell + ","
	'''sqlStr = sqlStr + " limitselldate='" + limitselldate + "',"
	sqlStr = sqlStr + " userclass='" + userclass + "',"
	sqlStr = sqlStr + " eventdiv='" + eventdiv + "',"
	sqlStr = sqlStr + " eventcpcd='" + eventcpcode + "',"
	sqlStr = sqlStr + " isusing='" + isusing + "',"
	sqlStr = sqlStr + " sellyn='" + sellyn + "'"
	sqlStr = sqlStr + " where id=" + id
end if

'response.write sqlStr
rsget.Open sqlStr,dbget,1


%>
<script language='javascript'>
alert('저장되었습니다.');
location.replace('<%= refer %>');
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->