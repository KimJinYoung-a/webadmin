<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<%
dim mode,itemid

mode = RequestCheckvar(request("mode"),16)
itemid = RequestCheckvar(request("itemid"),10)


''response.write mode + "<br>"
'response.write itemid + "<br>"
'response.write itemdiv + "<br>"
''response.write itemid+ "<br>"
''response.write vtinclude + "<br>"
''response.write buycash + "<br>"
''response.write buyvat + "<br>"
''response.write marginrate + "<br>"
'dbACADEMYget.close()	:	response.End
dim i
dim sqlStr
dim adminid
 adminid = session("ssBctID")

if mode="del" then
	itemid = split(itemid,"|")

	for i=0 to Ubound(itemid)
		sqlStr = "update [db_academy].[dbo].tbl_mdpick" + VbCrlf
		sqlStr = sqlStr + " set isusing='N'" + VbCrlf
		sqlStr = sqlStr + " where idx=" + CStr(itemid(i))
		dbACADEMYget.Execute sqlStr
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
'dbACADEMYget.close()	:	response.End
dim refer
refer = request.ServerVariables("HTTP_REFERER")
%>

<script language="javascript">
alert('수정되었습니다.');
location.replace('<%= refer %>');
</script>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->