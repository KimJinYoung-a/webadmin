<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
dim itemgubun, itemid, itemoption

itemgubun   = request("itemgubun")
itemid      = request("itemid")
itemoption  = request("itemoption")


response.write itemgubun & "<br>"
response.write itemid & "<br>"
response.write itemoption & "<br>"



''삭제 가능한 내역인지 Check.
dim sqlStr
dim ErrStr
ErrStr = ""


''최근 판매내역 : 판매내역은 취소인경우도.. Code를 바꾸어야함?
sqlStr = "select top 1 * from "
sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d"
sqlStr = sqlStr + " where d.itemid=" + CStr(itemid)
sqlStr = sqlStr + " and d.itemoption='" + itemoption + "'"

rsget.Open sqlStr, dbget, 1
if Not rsget.Eof then
	ErrStr = "삭제하려는 옵션으로 판매된 내역(6개월이내)이 있습니다. 삭제하실 수 없습니다."
end if
rsget.close

''6개월 이전 판매내역 : 판매내역은 취소인경우도.. Code를 바꾸어야함?
if ErrStr="" then
	sqlStr = "select top 1 * from "
	sqlStr = sqlStr + " [db_log].[dbo].tbl_old_order_detail_2003 d"
	sqlStr = sqlStr + " where d.itemid=" + CStr(itemid)
	sqlStr = sqlStr + " and d.itemoption='" + itemoption + "'"

	rsget.Open sqlStr, dbget, 1
	if Not rsget.Eof then
		ErrStr = "삭제하려는 옵션으로 판매된 내역(6개월이전)이 있습니다. 삭제하실 수 없습니다."
	end if
	rsget.close
end if

''입출고내역
if ErrStr="" then
	sqlStr = "select top 1 * from [db_storage].[dbo].tbl_acount_storage_detail d,"
	sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_master m"
	sqlStr = sqlStr + " where m.code=d.mastercode"
	sqlStr = sqlStr + " and m.deldt is NULL"
	sqlStr = sqlStr + " and d.iitemgubun='10'"
	sqlStr = sqlStr + " and d.itemid=" + CStr(itemid)
	sqlStr = sqlStr + " and d.itemoption='" + itemoption + "'"
    sqlStr = sqlStr + " and d.deldt is NULL"
    
	rsget.Open sqlStr, dbget, 1
	if Not rsget.Eof then
		ErrStr = "삭제하려는 옵션으로 입출고 내역이 있습니다. 삭제하실 수 없습니다."
	end if
	rsget.close
end if
	        	

'' 온라인 정산내역
if ErrStr="" then
    sqlStr = "select top 1 * from [db_jungsan].[dbo].tbl_designer_jungsan_detail"
    sqlStr = sqlStr + " where itemid=" & itemid
    sqlStr = sqlStr + " and itemoption='" & itemoption & "'"
    
    rsget.Open sqlStr, dbget, 1
	if Not rsget.Eof then
		ErrStr = "삭제하려는 옵션으로 온라인 정산 내역이 있습니다. 삭제하실 수 없습니다."
	end if
	rsget.close
end if

'' 오프라인 정산내역
if ErrStr="" then
    sqlStr = "select top 1 * from [db_jungsan].[dbo].tbl_off_jungsan_detail"
    sqlStr = sqlStr + " where itemgubun='10'"
    sqlStr = sqlStr + " and itemid=" & itemid
    sqlStr = sqlStr + " and itemoption='" & itemoption & "'"
    
    rsget.Open sqlStr, dbget, 1
	if Not rsget.Eof then
		ErrStr = "삭제하려는 옵션으로 OFF 정산 내역이 있습니다. 삭제하실 수 없습니다."
	end if
	rsget.close
end if



if (ErrStr<>"") then
    response.write "<script>alert('삭제할 수 없습니다.\n\n" & ErrStr & "');</script>"
    response.write "<script>window.close();</script>"
    dbget.close()	:	response.End
end if

'response.write "삭제process"
'dbget.close()	:	response.End


sqlStr = " delete from [db_summary].[dbo].tbl_daily_logisstock_summary" + VbCrlf 
sqlStr = sqlStr + " where itemgubun='" + CStr(itemgubun) + "'" + VbCrlf 
sqlStr = sqlStr + " and itemid=" + CStr(itemid) + "" + VbCrlf 
sqlStr = sqlStr + " and itemoption='" + CStr(itemoption) + "'"

dbget.Execute sqlStr


sqlStr = " delete from [db_summary].[dbo].tbl_erritem_daily_summary"
sqlStr = sqlStr + " where itemgubun='" + CStr(itemgubun) + "'" + VbCrlf 
sqlStr = sqlStr + " and itemid=" + CStr(itemid) + "" + VbCrlf 
sqlStr = sqlStr + " and itemoption='" + CStr(itemoption) + "'"
dbget.Execute sqlStr



sqlStr = " delete from [db_summary].[dbo].tbl_monthly_logisstock_summary"
sqlStr = sqlStr + " where itemgubun='" + CStr(itemgubun) + "'" + VbCrlf 
sqlStr = sqlStr + " and itemid=" + CStr(itemid) + "" + VbCrlf 
sqlStr = sqlStr + " and itemoption='" + CStr(itemoption) + "'"
dbget.Execute sqlStr



sqlStr = " delete from [db_summary].[dbo].tbl_LAST_monthly_logisstock"
sqlStr = sqlStr + " where itemgubun='" + CStr(itemgubun) + "'" + VbCrlf 
sqlStr = sqlStr + " and itemid=" + CStr(itemid) + "" + VbCrlf 
sqlStr = sqlStr + " and itemoption='" + CStr(itemoption) + "'"
dbget.Execute sqlStr


sqlStr = " delete from [db_summary].[dbo].tbl_monthly_accumulated_logisstock_summary"
sqlStr = sqlStr + " where itemgubun='" + CStr(itemgubun) + "'" + VbCrlf 
sqlStr = sqlStr + " and itemid=" + CStr(itemid) + "" + VbCrlf 
sqlStr = sqlStr + " and itemoption='" + CStr(itemoption) + "'"
dbget.Execute sqlStr


sqlStr = " delete from [db_summary].[dbo].tbl_current_logisstock_summary"
sqlStr = sqlStr + " where itemgubun='" + CStr(itemgubun) + "'" + VbCrlf 
sqlStr = sqlStr + " and itemid=" + CStr(itemid) + "" + VbCrlf 
sqlStr = sqlStr + " and itemoption='" + CStr(itemoption) + "'"
dbget.Execute sqlStr


response.write "<script>alert('삭제 되었습니다.');</script>"
response.write "<script>window.close();</script>"
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->