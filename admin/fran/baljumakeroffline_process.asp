<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%

dim mode, masteridxarr
dim baljunum, baljuid, baljucode, baljudate
dim baljuidarr, baljucodearr

mode = request("mode")
masteridxarr = request("masteridxarr")

dim refer
refer = request.ServerVariables("HTTP_REFERER")

dim sqlStr, i, cnt, buf

if (mode = "makebalju") then
    masteridxarr = replace(masteridxarr,"|",",") + "-1"

    '마이너스 주문이 있는지 확인
    sqlStr = " select distinct m.baljucode "
    sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_master m,"
    sqlStr = sqlStr + " [db_storage].[dbo].tbl_ordersheet_detail d "
    sqlStr = sqlStr + " where m.idx = d.masteridx "
    sqlStr = sqlStr + " and m.deldt is null "
    sqlStr = sqlStr + " and d.deldt is null "
    sqlStr = sqlStr + " and m.idx in (" + CStr(masteridxarr) + ") "
    sqlStr = sqlStr + " and d.baljuitemno < 0 "
        
	rsget.Open sqlStr, dbget, 1

	buf = ""
	if Not rsget.Eof then
		do until rsget.eof
			if (buf = "") then
			        buf = rsget("baljucode")
			else
			        buf = buf + "," + rsget("baljucode")
			end if
			rsget.movenext
		loop
	end if
	rsget.close

	if (buf <> "") then
	        response.write "<script>alert('주문중에 마이너스 주문이 있는 주문(" + buf + ")이 있습니다.');</script>"
	        response.write "<script>history.back();</script>"
	        dbget.close()	:	response.End
	end if


    ''상태변경
    sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master "
    sqlStr = sqlStr + " set statecd='1' "
    sqlStr = sqlStr + " where 1 = 1 "
    sqlStr = sqlStr + " and statecd = '0' "
    sqlStr = sqlStr + " and idx in (" + CStr(masteridxarr) + ") "
    dbget.Execute sqlStr
    
    ''실출고수량 0 Reset
    sqlStr = " update [db_storage].[dbo].tbl_ordersheet_detail "
    sqlStr = sqlStr + " set realitemno=0 "
    sqlStr = sqlStr + " where 1 = 1 "
    sqlStr = sqlStr + " and masteridx in (" + CStr(masteridxarr) + ") "
    dbget.Execute sqlStr

    sqlStr = " select max(isnull(baljunum,0)) as maxbaljunum, convert(varchar,getdate(),109) as baljudate "
    sqlStr = sqlStr + " from [db_storage].[dbo].tbl_shopbalju "
	rsget.Open sqlStr, dbget, 1
	if Not rsget.Eof then
		baljunum = rsget("maxbaljunum") + 1
		baljudate = rsget("baljudate")
	end if
	rsget.close

    sqlStr = " select baljuid, baljucode "
    sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_master "
    sqlStr = sqlStr + " where 1 = 1 "
    sqlStr = sqlStr + " and idx in (" + CStr(masteridxarr) + ") "
	rsget.Open sqlStr, dbget, 1
	if Not rsget.Eof then
		baljuidarr = ""
		baljucodearr = ""
		do until rsget.eof
			baljuidarr = baljuidarr + rsget("baljuid") + "|"
			baljucodearr = baljucodearr + rsget("baljucode") + "|"
			rsget.movenext
		loop
	end if
	rsget.close

	baljuidarr = split(baljuidarr,"|")
	baljucodearr = split(baljucodearr,"|")

	cnt = ubound(baljuidarr)
	for i = 0 to cnt
        if (baljuidarr(i) <> "") then
            sqlStr = " insert into [db_storage].[dbo].tbl_shopbalju(baljunum, baljuid, baljucode, baljudate) "
            sqlStr = sqlStr + " values(" + CStr(baljunum) + ", '" + CStr(baljuidarr(i)) + "', '" + CStr(baljucodearr(i)) + "', convert(datetime,'" + CStr(baljudate) + "',109)) "
            rsget.Open sqlStr, dbget, 1
            'response.write sqlStr
        end if
	next
    
    
    ''재고 변경 (오프 접수 -> 오프상품준비) : 오프접수, 상품준비 수량 재계산
    sqlstr = " exec [db_summary].[dbo].sp_ten_RealtimeStock_offjupsuAll"
    dbget.Execute sqlStr
    
    ''기존 Process
    '5.출고된 내역 한정판매 재설정("주문수" 만큼 더해준다.)
    sqlstr = " update [db_item].[dbo].tbl_item "
    sqlstr = sqlstr + " set limitsold=(case when limitno<limitsold + T.itemno then limitno else limitsold + T.itemno end) "
    sqlstr = sqlstr + " from ( "
    sqlstr = sqlstr + "     select sum(d.baljuitemno) as itemno, d.itemid "
    sqlstr = sqlstr + "     from [db_storage].[dbo].tbl_ordersheet_detail d, [db_item].[dbo].tbl_item i "
    sqlstr = sqlstr + "     where d.masteridx in (" + CStr(masteridxarr) + ") "
    sqlstr = sqlstr + "     and d.itemid=i.itemid "
    sqlstr = sqlstr + "     and d.deldt is null "
    sqlstr = sqlstr + "     and d.itemgubun = '10' "
    sqlstr = sqlstr + "     and d.itemoption = '0000' "
    sqlstr = sqlstr + "     and i.limityn='Y' "
    sqlstr = sqlstr + "     and i.mwdiv<>'U'"
    sqlstr = sqlstr + "     group by d.itemid "
    sqlstr = sqlstr + " ) as T "
    sqlstr = sqlstr + " where [db_item].[dbo].tbl_item.itemid=T.itemid "
    rsget.Open sqlStr, dbget, 1

    '5.출고된 내역 한정판매 재설정("주문수" 만큼 더해준다.)
    sqlstr = " update [db_item].[dbo].tbl_item_option "
    sqlstr = sqlstr + " set optlimitsold=(case when optlimitno<optlimitsold+T.itemno then optlimitno else optlimitsold+T.itemno end) "
    sqlstr = sqlstr + " from ( "
    sqlstr = sqlstr + "     select sum(d.baljuitemno) as itemno, d.itemid, d.itemoption "
    sqlstr = sqlstr + "     from [db_storage].[dbo].tbl_ordersheet_detail d, [db_item].[dbo].tbl_item i "
    sqlstr = sqlstr + "     where d.masteridx in (" + CStr(masteridxarr) + ") "
    sqlstr = sqlstr + "     and d.itemid=i.itemid "
    sqlstr = sqlstr + "     and d.deldt is null "
    sqlstr = sqlstr + "     and d.itemgubun = '10' "
    sqlstr = sqlstr + "     and d.itemoption <> '0000' "
    sqlstr = sqlstr + "     and i.limityn='Y' "
    sqlstr = sqlstr + "     and i.mwdiv<>'U'"
    sqlstr = sqlstr + "     group by d.itemid, d.itemoption "
    sqlstr = sqlstr + " ) as T "
    sqlstr = sqlstr + " where [db_item].[dbo].tbl_item_option.itemid=T.itemid and [db_item].[dbo].tbl_item_option.itemoption=T.itemoption "
    rsget.Open sqlStr, dbget, 1
    
    ''
	sqlStr = "update [db_item].[dbo].tbl_item" + VBCrlf
	sqlStr = sqlStr + " set limitno=IsNULL(T.optlimitno,0), limitsold=IsNULL(T.optlimitsold,0)" + VBCrlf
	sqlStr = sqlStr + " from (" + VBCrlf
	sqlStr = sqlStr + " 	select itemid, sum(optlimitno) as optlimitno, sum(optlimitsold) as optlimitsold" + VBCrlf
	sqlStr = sqlStr + " 	from [db_item].[dbo].tbl_item_option" + VBCrlf
	sqlStr = sqlStr + " 	where itemid in (select itemid from [db_storage].[dbo].tbl_ordersheet_detail where masteridx in (" + CStr(masteridxarr) + ") and itemoption <> '0000') " + VBCrlf
	sqlStr = sqlStr + " 	and isusing='Y'" + VBCrlf
	sqlStr = sqlStr + " 	group by itemid " + VBCrlf
	sqlStr = sqlStr + " ) T" + VBCrlf
	sqlStr = sqlStr + " where [db_item].[dbo].tbl_item.itemid= T.itemid " + VBCrlf
	sqlStr = sqlStr + " and [db_item].[dbo].tbl_item.optioncnt>0"

	rsget.Open sqlStr, dbget, 1
end if

%>
<script language="javascript">
alert('저장 되었습니다.');
location.replace('<%= refer %>');
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->
