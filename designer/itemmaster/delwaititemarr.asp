<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
dim mode,itemid

mode = request("mode")
itemid = request("itemid")


''response.write mode + "<br>"
'response.write itemid + "<br>"
'response.write itemdiv + "<br>"
''response.write itemid+ "<br>"
''response.write vtinclude + "<br>"
''response.write buycash + "<br>"
''response.write buyvat + "<br>"
''response.write marginrate + "<br>"
'dbget.close()	:	response.End
dim i
dim sqlStr
dim adminid
 adminid = session("ssBctID")

if mode="del" then
	itemid = split(itemid,"|")

	for i=0 to Ubound(itemid)
		sqlStr = "update [db_temp].[dbo].tbl_wait_item" + VbCrlf
		sqlStr = sqlStr + " set currstate='9'" + VbCrlf
		sqlStr = sqlStr + " where itemid=" + CStr(itemid(i))
		dbget.Execute sqlStr
		
		 sqlStr = " INSERT INTO db_temp.dbo.tbl_wait_item_log (itemid, currstate, adminid)"
		sqlStr = sqlStr +	" VALUES("&itemid(i)&", 9,'"&adminid&"')"
		dbget.Execute sqlStr
	next
else
	'response.write "???"
	'oneitem.FItemID = itemid
	'oneitem.FSellPrice = sellcash
	'oneitem.FSellVat = sellvat
	'oneitem.FBuyPrice = buycash
	'oneitem.FBuyVat = buyvat
	'oneitem.FMarginrate = marginrate
	'oneitem.FVatInclude = vtinclude
	'oneitem.FMarginDiv = "1"

	'obuyprice.UpdateOneItem oneitem
end if
'dbget.close()	:	response.End
dim refer
refer = request.ServerVariables("HTTP_REFERER")
%>

<script language="javascript">
alert('수정되었습니다.');
location.replace('<%= refer %>');
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->