<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<%

dim mode, i
mode = Request("mode")

dim keywords, itemid
itemid = Request("itemid")
keywords = Request("keywords")

dim KeyRows, KeyRow, KeyItemid, KeyKeywords, KeyItemidArr

dim sqlStr
select case mode
	case "editone"

		sqlStr = "update db_item.dbo.tbl_item_Contents "
		sqlStr = sqlStr + " set keywords = '" & chrbyte(html2db(Request("keywords")),128,"") & "' "
		sqlStr = sqlStr + "where itemid = " + CStr(itemid) + " "
		dbget.execute(sqlStr)

		sqlStr = "update [db_item].[dbo].tbl_item" + vbCrlf
		sqlStr = sqlStr & " set lastupdate=getdate()"
		sqlStr = sqlStr & " where itemid=" & CStr(itemid) & "" + vbCrlf
		dbget.execute(sqlStr)

    	Response.Write	"<script language=javascript>" &_
    	"	alert('데이터를 저장하였습니다.');" &_
    	"	opener.history.go(0);" &_
    	"	self.close();" &_
    	"</script>"

	case "editmulti"
		KeyRows = Split(keywords, vbCrLf)
		KeyItemidArr = "-1"

		for i = 0 to UBound(KeyRows) - 1
			KeyRow = Trim(KeyRows(i))
			if (KeyRow <> "") then
				KeyRow = Split(KeyRow, vbTab)
				KeyItemid = KeyRow(0)
				KeyKeywords = KeyRow(1)

				sqlStr = "update db_item.dbo.tbl_item_Contents "
				sqlStr = sqlStr + " set keywords = '" & chrbyte(html2db(KeyKeywords),128,"") & "' "
				sqlStr = sqlStr + "where itemid = " + CStr(KeyItemid) + " "

				''response.write sqlStr & "<br>"
				dbget.execute(sqlStr)

				KeyItemidArr = KeyItemidArr & "," & CStr(KeyItemid)
			end if
		next

		'sqlStr = " insert into watcher.dbo.kw_klog "
		'sqlStr = sqlStr + " select ITEMID,NULL,NULL,NULL,NULL,'U',getdate(),NULL,'P','db_item','dbo','TBL_ITEM','db_item','dbo','vw_item_DispCate', NULL "
		'sqlStr = sqlStr + " from db_item.dbo.tbl_item "
		'sqlStr = sqlStr + " where itemid in (" + CStr(KeyItemidArr) + ") "
		'dbget.execute(sqlStr)
		'' 수정 2015/04/17
		sqlStr = " insert watcher.dbo.kc_job_log"
		sqlStr = sqlStr + " (job_id,pk_value,iud_type)"
		sqlStr = sqlStr + " select 'itemDisp',ITEMID,'U'"
        sqlStr = sqlStr + " from db_item.dbo.tbl_item "
        sqlStr = sqlStr + " where itemid in (" + CStr(KeyItemidArr) + ") "
        dbget.execute(sqlStr)
        
        
		Response.Write	"<script language=javascript>" &_
    	"	alert('데이터를 저장하였습니다.');" &_
    	"	opener.history.go(0);" &_
    	"	self.close();" &_
    	"</script>"
	case else
		''
end select

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
