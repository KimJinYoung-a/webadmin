<%@ language=vbscript %>
<% option explicit %>
<% Server.ScriptTimeOut = 900 %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<%

dim mode, reguserid
dim sqlStr, i
dim arrItemid, itemidlist, itemid, modifiedtext

reguserid = session("ssBctId")

mode = requestCheckvar(request("mode"),64)
arrItemid = requestCheckvar(request("arrItemid"),64000)
itemid = requestCheckvar(request("itemid"),32)
modifiedtext = requestCheckvar(request("modifiedtext"),64000)

select case mode
	case "ins"
		'// aaa
		itemidlist = "-1"
		arrItemid = Split(arrItemid, vbCrLf)
		for i = 0 to UBound(arrItemid)
			if (Trim(arrItemid(i)) <> "") then
				itemidlist = itemidlist + "," + Trim(arrItemid(i))
			end if
		next

		sqlStr = " update [db_contents].[dbo].[tbl_itemImageText] "
		sqlStr = sqlStr + " set req_yyyymmdd = convert(varchar(10), getdate(), 121), fin_yyyymmdd = NULL, lastuserid = '" & reguserid & "', lastupdate = getdate() "
		sqlStr = sqlStr + " where itemid in (" & itemidlist & ") "
		dbget.Execute sqlStr

		sqlStr = " insert into [db_contents].[dbo].[tbl_itemImageText](itemid, req_yyyymmdd, updatecnt, reguserid, lastuserid, regdate, lastupdate)"
		sqlStr = sqlStr + " select i.itemid, convert(varchar(10), getdate(), 121), 0, '" & reguserid & "', '" & reguserid & "', getdate(), getdate() "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	[db_item].[dbo].tbl_item i "
		sqlStr = sqlStr + " 	left join [db_contents].[dbo].[tbl_itemImageText] t "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		t.itemid = i.itemid "
		sqlStr = sqlStr + " where i.itemid in (" & itemidlist & ") and t.itemid is NULL "
		dbget.Execute sqlStr

		response.write "<script language='javascript'>alert('저장되었습니다.'); opener.focus(); opener.GoPage(1); window.close();</script>"
		dbget.close : response.end
	case "modi"
		sqlStr = " update [db_contents].[dbo].[tbl_itemImageText] "
		sqlStr = sqlStr + " set modifiedtext = '" & html2db(modifiedtext) & "', lastuserid = '" & reguserid & "', lastupdate = getdate() "
		sqlStr = sqlStr + " where itemid = " & itemid
		dbget.Execute sqlStr
	case else
		'// error
		response.write "ERROR"
		dbget.close : response.end
end select

%>
<script language='javascript'>
alert('저장되었습니다.');
opener.location.reload(); opener.focus(); window.close();
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
