<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/admin/MenuCls.asp"-->
<%

dim mode
dim menu_id, menu_name, userid, arrMenuId, tmpStr
dim i, j, strSql

mode = request("mode")
menu_id = request("menu_id")
menu_name = request("menu_name")

userid = session("ssBctID")

if (mode = "tmpaddfavorite") then
	response.write "<script>parent.fnAddFavorite('" + CStr(menu_id) + "', '" + CStr(menu_name) + "')</script>"
elseif (mode = "tmpdelfavorite") then
	response.write "<script>parent.fnDelFavorite('" + CStr(menu_id) + "', '" + CStr(menu_name) + "')</script>"
elseif (mode = "realaddfavorite") then
	strSql = " update db_partner.dbo.tbl_partner_menu_favorite "
	strSql = strSql + " set useYN = 'Y' "
	strSql = strSql + " where userid = '" + CStr(userid) + "' and menu_id in (" + CStr(menu_id) + ") "
	''response.write strSql
	dbget.Execute(strSql)

	arrMenuId = Split(menu_id, ",")
	for i = 0 to UBound(arrMenuId)
		tmpStr = Trim(arrMenuId(i))
		if (tmpStr <> "") and (tmpStr <> "-1") then
			strSql = " if not exists( "
			strSql = strSql + " 	select top 1 menu_id "
			strSql = strSql + " 	from db_partner.dbo.tbl_partner_menu_favorite "
			strSql = strSql + " 	where userid = '" + CStr(userid) + "' and menu_id = " + CStr(tmpStr) + " "
			strSql = strSql + " ) "
			strSql = strSql + " begin "
			strSql = strSql + " 	insert into db_partner.dbo.tbl_partner_menu_favorite(userid, menu_id, useYN) "
			strSql = strSql + " 	values('" + CStr(userid) + "', " + CStr(tmpStr) + ", 'Y') "
			strSql = strSql + " end "
			''response.write strSql
			dbget.Execute(strSql)
		end if
	next

	response.write "<script>alert('저장되었습니다.'); opener.location.reload(); opener.focus(); window.close();</script>"
elseif (mode = "realdelfavorite") then
	strSql = " update db_partner.dbo.tbl_partner_menu_favorite "
	strSql = strSql + " set useYN = 'N' "
	strSql = strSql + " where userid = '" + CStr(userid) + "' and menu_id in (" + CStr(menu_id) + ") "
	dbget.Execute(strSql)

	response.write "<script>alert('저장되었습니다.'); opener.location.reload(); opener.focus(); window.close();</script>"
elseif (mode = "addonefavorite") then
	strSql = " if not exists( "
	strSql = strSql + " 	select top 1 menu_id "
	strSql = strSql + " 	from db_partner.dbo.tbl_partner_menu_favorite "
	strSql = strSql + " 	where userid = '" + CStr(userid) + "' and menu_id = " + CStr(menu_id) + " "
	strSql = strSql + " ) "
	strSql = strSql + " begin "
	strSql = strSql + " 	insert into db_partner.dbo.tbl_partner_menu_favorite(userid, menu_id, useYN) "
	strSql = strSql + " 	values('" + CStr(userid) + "', " + CStr(menu_id) + ", 'Y') "
	strSql = strSql + " end "
	strSql = strSql + " else "
	strSql = strSql + " begin "
	strSql = strSql + " 	update db_partner.dbo.tbl_partner_menu_favorite "
	strSql = strSql + " 	set useYN = 'Y' "
	strSql = strSql + " 	where userid = '" + CStr(userid) + "' and menu_id = " + CStr(menu_id) + " "
	strSql = strSql + " end "
	''response.write strSql
	dbget.Execute(strSql)

	response.write "<script>alert('추가되었습니다.'); history.back();</script>"
elseif (mode = "delonefavorite") then
	strSql = " update db_partner.dbo.tbl_partner_menu_favorite "
	strSql = strSql + " set useYN = 'N' "
	strSql = strSql + " where userid = '" + CStr(userid) + "' and menu_id = " + CStr(menu_id) + " "
	dbget.Execute(strSql)

	response.write "<script>alert('제외되었습니다.'); history.back();</script>"
else
	'
end if

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
