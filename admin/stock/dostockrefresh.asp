<%@ language=vbscript %>
<% option explicit %>
<%
''불량상품 등록제외한 모든 내용 사용불가 불량및실사내역삭제 - baditem_process.asp : 프로시저로 변경요망
'response.write "관리자 문의 요망!! - 사용불가메뉴." & "<br>"
'response.write "mode=" & request.form("mode")
'dbget.close()	:	response.End
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
dim refer
	refer = request.ServerVariables("HTTP_REFERER")

dim mode, refreshstartdate, itemgubun, itemid, itemoption, realstock, makerid, sqlStr
dim orgrealstock, todayerrno, errstock, yyyymmdd, recent7day, recent3day, LastYYYYMM, nowdate, itemExists, oTimer, assignedNo
	mode	= request.form("mode")
	refreshstartdate = request.form("refreshstartdate")
	itemgubun = request.form("itemgubun")
	itemid = request.form("itemid")
	itemoption = request.form("itemoption")
	realstock = request.form("realstock")
	makerid = request.form("makerid")

sqlStr = "select convert(varchar(10),getdate(),21) as nowdate"

rsget.Open sqlStr,dbget,1
	nowdate = rsget("nowdate")
rsget.Close

recent3day = CStr(DateSerial(Left(nowdate,4),Mid(nowdate,6,2),Mid(nowdate,9,2)-3))
recent7day = CStr(DateSerial(Left(nowdate,4),Mid(nowdate,6,2),Mid(nowdate,9,2)-7))
LastYYYYMM = Left(CStr(DateSerial(Left(nowdate,4),Mid(nowdate,6,2)-1,1-1)),7)

response.write "mode=" & mode

''불량상품 등록
if (mode<>"tmpbaditem2input") then
    response.write "관리자 문의 요망 <br>"
    response.write "mode=" & mode
    response.flush
    dbget.close()	:	response.End
end if

if mode="dellog" then
    sqlStr = " delete from [db_summary].[dbo].tbl_daily_logisstock_summary" + VbCrlf
    sqlStr = sqlStr + " where itemgubun='" + CStr(itemgubun) + "'" + VbCrlf
	sqlStr = sqlStr + " and itemid=" + CStr(itemid) + "" + VbCrlf
	sqlStr = sqlStr + " and itemoption='" + CStr(itemoption) + "'"

	dbget.Execute sqlStr, assignedNo

	response.write "tbl_daily_logisstock_summary.." + CStr(assignedNo)+ " deleted..<br>"
    response.flush

	sqlStr = " delete from [db_summary].[dbo].tbl_monthly_logisstock_summary"
	sqlStr = sqlStr + " where itemgubun='" + CStr(itemgubun) + "'" + VbCrlf
	sqlStr = sqlStr + " and itemid=" + CStr(itemid) + "" + VbCrlf
	sqlStr = sqlStr + " and itemoption='" + CStr(itemoption) + "'"
	dbget.Execute sqlStr, assignedNo

	response.write "tbl_monthly_logisstock_summary.." + CStr(assignedNo)+ " deleted..<br>"
    response.flush

	sqlStr = " delete from [db_summary].[dbo].tbl_current_logisstock_summary"
	sqlStr = sqlStr + " where itemgubun='" + CStr(itemgubun) + "'" + VbCrlf
	sqlStr = sqlStr + " and itemid=" + CStr(itemid) + "" + VbCrlf
	sqlStr = sqlStr + " and itemoption='" + CStr(itemoption) + "'"
	dbget.Execute sqlStr, assignedNo

	response.write "tbl_current_logisstock_summary.." + CStr(assignedNo)+ " deleted..<br>"
    response.flush

	sqlStr = " delete from [db_summary].[dbo].tbl_erritem_daily_summary"
	sqlStr = sqlStr + " where itemgubun='" + CStr(itemgubun) + "'" + VbCrlf
	sqlStr = sqlStr + " and itemid=" + CStr(itemid) + "" + VbCrlf
	sqlStr = sqlStr + " and itemoption='" + CStr(itemoption) + "'"
	dbget.Execute sqlStr, assignedNo

	response.write "tbl_erritem_daily_summary.." + CStr(assignedNo)+ " deleted..<br>"
    response.flush

elseif mode="itemrecentipchulrefresh" then
	if (itemid="") then
		response.write "<script>alert('상품코드를 입력하세요.');</script>"
		response.write "<script>history.back();</script>"
		dbget.close()	:	response.End
	end if

	oTimer = Timer()

	sqlStr = " update [db_summary].[dbo].tbl_daily_logisstock_summary"
	sqlStr = sqlStr + " set sellno=0"
	sqlStr = sqlStr + " ,resellno=0"
	sqlStr = sqlStr + " ,totsellno=0"
	sqlStr = sqlStr + " where yyyymmdd>='" + refreshstartdate + "'"
	sqlStr = sqlStr + " and itemgubun='" + CStr(itemgubun) + "'"
	sqlStr = sqlStr + " and itemid=" + CStr(itemid) + ""
	sqlStr = sqlStr + " and itemoption='" + CStr(itemoption) + "'"

	rsget.Open sqlStr,dbget,1

	''재고 일별 판매입력 - 3일간 온라인 출고건
	sqlStr = " insert into [db_summary].[dbo].tbl_daily_logisstock_summary"
	sqlStr = sqlStr + " (yyyymmdd,itemgubun,itemid,itemoption,sellno,resellno,totsellno)"
	sqlStr = sqlStr + " select T.yyyymmdd,'10',T.itemid,T.itemoption,T.sellno,T.resellno,T.totsellno"
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " ("
	sqlStr = sqlStr + " select convert(varchar(10),beadaldate,21) as yyyymmdd, d.itemid,d.itemoption,"
	sqlStr = sqlStr + " sum(case when d.itemno>0 then d.itemno else 0 end ) as sellno,"
	sqlStr = sqlStr + " sum(case when d.itemno<0 then d.itemno else 0 end ) as resellno,"
	sqlStr = sqlStr + " sum(d.itemno) as totsellno"
	sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m,"
	sqlStr = sqlStr + "  [db_order].[dbo].tbl_order_detail d"
	sqlStr = sqlStr + " where m.orderserial=d.orderserial"
	sqlStr = sqlStr + " and m.regdate>='" + LastYYYYMM + "-01'"
	sqlStr = sqlStr + " and m.beadaldate>='" + refreshstartdate + "'"
	sqlStr = sqlStr + " and m.ipkumdiv>=7"
	sqlStr = sqlStr + " and m.cancelyn='N'"
	sqlStr = sqlStr + " and d.cancelyn<>'Y'"
	sqlStr = sqlStr + " and d.itemid=" + CStr(itemid) + ""
	sqlStr = sqlStr + " and d.itemoption='" + CStr(itemoption) + "'"
	sqlStr = sqlStr + " and d.isupchebeasong<>'Y'"
	sqlStr = sqlStr + " group by convert(varchar(10),beadaldate,21), d.itemid, d.itemoption"
	sqlStr = sqlStr + " ) T"
	sqlStr = sqlStr + " left join [db_summary].[dbo].tbl_daily_logisstock_summary s"
	sqlStr = sqlStr + " on T.yyyymmdd=s.yyyymmdd"
	sqlStr = sqlStr + " and T.itemid=s.itemid"
	sqlStr = sqlStr + " and T.itemoption=s.itemoption"
	sqlStr = sqlStr + " where s.itemid is null"

	rsget.Open sqlStr,dbget,1

	sqlStr = " update [db_summary].[dbo].tbl_daily_logisstock_summary"
	sqlStr = sqlStr + " set sellno=T.sellno"
	sqlStr = sqlStr + " ,resellno=T.resellno"
	sqlStr = sqlStr + " ,totsellno=T.totsellno"
	sqlStr = sqlStr + " ,lastupdate=getdate()"
	sqlStr = sqlStr + " from ("
	sqlStr = sqlStr + " select convert(varchar(10),beadaldate,21) as yyyymmdd, d.itemid,d.itemoption,"
	sqlStr = sqlStr + " sum(case when d.itemno>0 then d.itemno else 0 end ) as sellno,"
	sqlStr = sqlStr + " sum(case when d.itemno<0 then d.itemno else 0 end ) as resellno,"
	sqlStr = sqlStr + " sum(d.itemno) as totsellno"
	sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m,"
	sqlStr = sqlStr + "  [db_order].[dbo].tbl_order_detail d"
	sqlStr = sqlStr + " where  m.orderserial=d.orderserial"
	sqlStr = sqlStr + " and m.regdate>='" + LastYYYYMM + "-01'"
	sqlStr = sqlStr + " and m.beadaldate>='" + refreshstartdate + "'"
	sqlStr = sqlStr + " and m.ipkumdiv>=7"
	sqlStr = sqlStr + " and m.cancelyn='N'"
	sqlStr = sqlStr + " and d.cancelyn<>'Y'"
	sqlStr = sqlStr + " and d.itemid=" + CStr(itemid) + ""
	sqlStr = sqlStr + " and d.itemoption='" + CStr(itemoption) + "'"
	sqlStr = sqlStr + " and d.isupchebeasong<>'Y'"
	sqlStr = sqlStr + " group by convert(varchar(10),beadaldate,21), d.itemid, d.itemoption"
	sqlStr = sqlStr + " ) T"
	sqlStr = sqlStr + " where [db_summary].[dbo].tbl_daily_logisstock_summary.yyyymmdd=T.yyyymmdd"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_daily_logisstock_summary.itemgubun='10'"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_daily_logisstock_summary.itemid=T.itemid"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_daily_logisstock_summary.itemoption=T.itemoption"

	rsget.Open sqlStr,dbget,1

	response.write "<small>재고 일별 판매입력... finish(" + FormatNumber(Timer()-oTimer,2) + " sec)</small><br>"
	response.flush

	oTimer = Timer()
	''재고 일별테이블 업데이트(입출고) - 수정일 이후

	''입출고 있는내역 입력
	sqlStr = " insert into [db_summary].[dbo].tbl_daily_logisstock_summary"
	sqlStr = sqlStr + " (yyyymmdd,itemgubun,itemid,itemoption)"
	sqlStr = sqlStr + " select T.yyyymmdd, T.iitemgubun, T.itemid, T.itemoption"
	sqlStr = sqlStr + " from ("
	sqlStr = sqlStr + " select distinct convert(varchar(10),m.executedt,21) as yyyymmdd ,iitemgubun,itemid,itemoption"
	sqlStr = sqlStr + "   from [db_storage].[dbo].tbl_acount_storage_master m, "
	sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_detail d"
	sqlStr = sqlStr + " where m.executedt>='" + refreshstartdate + "'"
	sqlStr = sqlStr + " and m.code=d.mastercode"
	sqlStr = sqlStr + " and m.deldt is null"
	sqlStr = sqlStr + " and d.deldt is null"
	'sqlStr = sqlStr + " and d.itemno<>0"
	sqlStr = sqlStr + " and d.iitemgubun='" + itemgubun + "'"
	sqlStr = sqlStr + " and d.itemid=" + CStr(itemid)
	sqlStr = sqlStr + " and d.itemoption='" + itemoption + "'"
	sqlStr = sqlStr + " ) T"
	sqlStr = sqlStr + " left join [db_summary].[dbo].tbl_daily_logisstock_summary s"
	sqlStr = sqlStr + " on T.yyyymmdd=s.yyyymmdd"
	sqlStr = sqlStr + " and T.iitemgubun=s.itemgubun"
	sqlStr = sqlStr + " and T.itemid=s.itemid"
	sqlStr = sqlStr + " and T.itemoption=s.itemoption"
	sqlStr = sqlStr + " where s.yyyymmdd is null"

 	rsget.Open sqlStr,dbget,1

 	sqlStr = "update [db_summary].[dbo].tbl_daily_logisstock_summary"
 	sqlStr = sqlStr + " set ipgono=T.ipgono"
 	sqlStr = sqlStr + " ,reipgono=T.reipgono"
 	sqlStr = sqlStr + " ,totipgono=T.ipgono+T.reipgono"
	sqlStr = sqlStr + " ,offchulgono=T.offchulgono"
	sqlStr = sqlStr + " ,offrechulgono=T.offrechulgono"
	sqlStr = sqlStr + " ,etcchulgono=T.etcchulgono"
	sqlStr = sqlStr + " ,etcrechulgono=T.etcrechulgono"
	sqlStr = sqlStr + " ,totchulgono=T.offchulgono+T.offrechulgono+T.etcchulgono+T.etcrechulgono"
	sqlStr = sqlStr + " ,lastupdate=getdate()"
 	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " ( select convert(varchar(10),m.executedt,21) as yyyymmdd"
	sqlStr = sqlStr + " ,iitemgubun,itemid,itemoption, "
	sqlStr = sqlStr + " sum(case when  ipchulflag='I' and itemno>0 then itemno else 0 end) as ipgono,"
	sqlStr = sqlStr + " sum(case when  ipchulflag='I' and itemno<0 then itemno else 0 end) as reipgono,"
	sqlStr = sqlStr + " sum(case when  ipchulflag='S' and itemno<0 then itemno else 0 end) as offchulgono,"
	sqlStr = sqlStr + " sum(case when  ipchulflag='S' and itemno>0 then itemno else 0 end) as offrechulgono,"
	sqlStr = sqlStr + " sum(case when  ipchulflag='E' and itemno<0 then itemno else 0 end) as etcchulgono,"
	sqlStr = sqlStr + " sum(case when  ipchulflag='E' and itemno>0 then itemno else 0 end) as etcrechulgono"
	sqlStr = sqlStr + "   from [db_storage].[dbo].tbl_acount_storage_master m, "
	sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_detail d"
	sqlStr = sqlStr + " where m.executedt>='" + refreshstartdate + "'"
	sqlStr = sqlStr + " and m.code=d.mastercode"
	sqlStr = sqlStr + " and m.deldt is null"
	sqlStr = sqlStr + " and d.iitemgubun='" + itemgubun + "'"
	sqlStr = sqlStr + " and d.itemid=" + CStr(itemid) + ""
	sqlStr = sqlStr + " and d.itemoption='" + itemoption + "'"
	sqlStr = sqlStr + " and d.deldt is null"
	'sqlStr = sqlStr + " and d.itemno<>0"
	sqlStr = sqlStr + " group by convert(varchar(10),m.executedt,21),iitemgubun,itemid,itemoption"
	sqlStr = sqlStr + " ) T"
	sqlStr = sqlStr + " where [db_summary].[dbo].tbl_daily_logisstock_summary.yyyymmdd=T.yyyymmdd"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_daily_logisstock_summary.itemgubun=T.iitemgubun"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_daily_logisstock_summary.itemid=T.itemid"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_daily_logisstock_summary.itemoption=T.itemoption"

	rsget.Open sqlStr,dbget,1

	response.write "<small>재고 일별테이블 업데이트(입출고)... finish(" + FormatNumber(Timer()-oTimer,2) + " sec)</small><br>"
	response.flush

	oTimer = Timer()
	''재고 일별테이블 업데이트(오차)

	sqlStr = " insert into [db_summary].[dbo].tbl_daily_logisstock_summary"
	sqlStr = sqlStr + " (yyyymmdd,itemgubun,itemid,itemoption,"
	sqlStr = sqlStr + " errcsno,errbaditemno,errrealcheckno,erretcno,toterrno)"
	sqlStr = sqlStr + " select s.yyyymmdd,s.itemgubun,s.itemid,s.itemoption,"
	sqlStr = sqlStr + " s.errcsno,s.errbaditemno,s.errrealcheckno,s.erretcno,s.toterrno"
	sqlStr = sqlStr + " from [db_summary].[dbo].tbl_erritem_daily_summary s"
	sqlStr = sqlStr + " left join [db_summary].[dbo].tbl_daily_logisstock_summary l"
	sqlStr = sqlStr + " on s.yyyymmdd=l.yyyymmdd"
	sqlStr = sqlStr + " and s.itemgubun=l.itemgubun"
	sqlStr = sqlStr + " and s.itemid=l.itemid"
	sqlStr = sqlStr + " and s.itemoption=l.itemoption"
	sqlStr = sqlStr + " where s.yyyymmdd>='" + nowdate + "'"
	sqlStr = sqlStr + " and s.itemgubun='" + itemgubun + "'"
	sqlStr = sqlStr + " and s.itemid=" + CStr(itemid) + ""
	sqlStr = sqlStr + " and s.itemoption='" + itemoption + "'"
	sqlStr = sqlStr + " and l.yyyymmdd is null"

	rsget.Open sqlStr,dbget,1

	sqlStr = " update [db_summary].[dbo].tbl_daily_logisstock_summary"
	sqlStr = sqlStr + " set errcsno=s.errcsno"
	sqlStr = sqlStr + " ,errbaditemno=s.errbaditemno"
	sqlStr = sqlStr + " ,errrealcheckno=s.errrealcheckno"
	sqlStr = sqlStr + " ,erretcno=s.erretcno"
	sqlStr = sqlStr + " ,toterrno=s.toterrno"
	sqlStr = sqlStr + " ,lastupdate=getdate()"
	sqlStr = sqlStr + " from  [db_summary].[dbo].tbl_erritem_daily_summary s"
	sqlStr = sqlStr + " where s.yyyymmdd>='" + nowdate + "'"
	sqlStr = sqlStr + " and s.itemgubun='" + itemgubun + "'"
	sqlStr = sqlStr + " and s.itemid=" + CStr(itemid) + ""
	sqlStr = sqlStr + " and s.itemoption='" + itemoption + "'"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_daily_logisstock_summary.yyyymmdd=s.yyyymmdd"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_daily_logisstock_summary.itemgubun=s.itemgubun"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_daily_logisstock_summary.itemid=s.itemid"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_daily_logisstock_summary.itemoption=s.itemoption"

	rsget.Open sqlStr,dbget,1

	response.write "<small>재고 일별테이블 업데이트(오차)... finish(" + FormatNumber(Timer()-oTimer,2) + " sec)</small><br>"
	response.flush

	oTimer = Timer()
	''서머리.
	sqlStr = " update [db_summary].[dbo].tbl_daily_logisstock_summary"
	sqlStr = sqlStr + " set totsysstock=totipgono+totchulgono-totsellno"
	sqlStr = sqlStr + " ,availsysstock=totipgono+totchulgono-totsellno+errcsno+errbaditemno+erretcno"
	sqlStr = sqlStr + " ,realstock=totipgono+totchulgono-totsellno+errcsno+errbaditemno+erretcno+errrealcheckno"
	sqlStr = sqlStr + " ,lastupdate=getdate()"
	sqlStr = sqlStr + " where lastupdate>='" + nowdate + "'"
	sqlStr = sqlStr + " and itemgubun='" + itemgubun + "'"
	sqlStr = sqlStr + " and itemid=" + CStr(itemid) + ""
	sqlStr = sqlStr + " and itemoption='" + itemoption + "'"

	rsget.Open sqlStr,dbget,1

	response.write "<small>재고 일별테이블 업데이트(서머리)... finish(" + FormatNumber(Timer()-oTimer,2) + " sec)</small><br>"
	response.flush

	oTimer = Timer()
	''재고 월별테이블 업데이트

	sqlStr = " delete from [db_summary].[dbo].tbl_monthly_logisstock_summary"
	sqlStr = sqlStr + " where yyyymm>='" + Left(refreshstartdate,7) + "'"
	sqlStr = sqlStr + " and itemgubun='" + itemgubun + "'"
	sqlStr = sqlStr + " and itemid=" + CStr(itemid) + ""
	sqlStr = sqlStr + " and itemoption='" + itemoption + "'"

	rsget.Open sqlStr,dbget,1

	sqlStr = " insert into [db_summary].[dbo].tbl_monthly_logisstock_summary"
	sqlStr = sqlStr + " (yyyymm,itemgubun,itemid,itemoption,"
	sqlStr = sqlStr + " ipgono,reipgono,totipgono,offchulgono,offrechulgono,"
	sqlStr = sqlStr + " etcchulgono,etcrechulgono,totchulgono,"
	sqlStr = sqlStr + " sellno,resellno,totsellno,errcsno,"
	sqlStr = sqlStr + " errbaditemno,errrealcheckno,erretcno,"
	sqlStr = sqlStr + " toterrno,offsellno,totsysstock,availsysstock,realstock)"
	sqlStr = sqlStr + " select"
	sqlStr = sqlStr + " convert(varchar(7),yyyymmdd,21) as yyyymm,itemgubun,itemid,itemoption,"
	sqlStr = sqlStr + " sum(ipgono),sum(reipgono),sum(totipgono),sum(offchulgono),sum(offrechulgono),"
	sqlStr = sqlStr + " sum(etcchulgono),sum(etcrechulgono),sum(totchulgono),"
	sqlStr = sqlStr + " sum(sellno),sum(resellno),sum(totsellno),sum(errcsno),"
	sqlStr = sqlStr + " sum(errbaditemno),sum(errrealcheckno),sum(erretcno),"
	sqlStr = sqlStr + " sum(toterrno),sum(offsellno),sum(totsysstock),sum(availsysstock),sum(realstock)"
	sqlStr = sqlStr + "  from [db_summary].[dbo].tbl_daily_logisstock_summary"
	sqlStr = sqlStr + " where yyyymmdd>='" + refreshstartdate + "'"
	sqlStr = sqlStr + " and itemgubun='" + itemgubun + "'"
	sqlStr = sqlStr + " and itemid=" + CStr(itemid) + ""
	sqlStr = sqlStr + " and itemoption='" + itemoption + "'"
	sqlStr = sqlStr + " group by convert(varchar(7),yyyymmdd,21) ,itemgubun,itemid,itemoption"

	rsget.Open sqlStr,dbget,1

	response.write "<small>재고 월별테이블 업데이트... finish(" + FormatNumber(Timer()-oTimer,2) + " sec)</small><br>"
	response.flush

	oTimer = Timer()
	''현재재고업데이트

	sqlStr = " delete from [db_summary].[dbo].tbl_current_logisstock_summary"
	sqlStr = sqlStr + " where itemgubun='" + itemgubun + "'"
	sqlStr = sqlStr + " and itemid=" + CStr(itemid) + ""
	sqlStr = sqlStr + " and itemoption='" + itemoption + "'"

	rsget.Open sqlStr,dbget,1

	if itemgubun="10" then
		sqlStr = " insert into [db_summary].[dbo].tbl_current_logisstock_summary "
		sqlStr = sqlStr + " (itemgubun,itemid,itemoption,imgsmall)"
		sqlStr = sqlStr + " select top 1 '10'," + CStr(itemid) + ",'" + itemoption + "',T.smallimage"
		sqlStr = sqlStr + " from [db_item].[dbo].tbl_item T"
		sqlStr = sqlStr + " where T.itemid=" + CStr(itemid) + ""

		rsget.Open sqlStr,dbget,1
	else
		sqlStr = " insert into [db_summary].[dbo].tbl_current_logisstock_summary "
		sqlStr = sqlStr + " (itemgubun,itemid,itemoption,imgsmall)"
		sqlStr = sqlStr + " select top 1 '" + itemgubun + "'," + CStr(itemid) + ",'" + itemoption + "',T.offimgsmall"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item T"
		sqlStr = sqlStr + " where and T.itemgubun='" + itemgubun + "'"
		sqlStr = sqlStr + " and T.shopitemid=" + CStr(itemid) + ""
		sqlStr = sqlStr + " and T.itemoption='" + itemoption + "'"

		rsget.Open sqlStr,dbget,1
	end if

	sqlStr = "update [db_summary].[dbo].tbl_current_logisstock_summary"
	sqlStr = sqlStr + " set ipgono=T.ipgono"
	sqlStr = sqlStr + " ,reipgono=T.reipgono"
	sqlStr = sqlStr + " ,totipgono=T.totipgono"
	sqlStr = sqlStr + " ,offchulgono=T.offchulgono"
	sqlStr = sqlStr + " ,offrechulgono=T.offrechulgono"
	sqlStr = sqlStr + " ,etcchulgono=T.etcchulgono"
	sqlStr = sqlStr + " ,etcrechulgono=T.etcrechulgono"
	sqlStr = sqlStr + " ,totchulgono=T.totchulgono"
	sqlStr = sqlStr + " ,sellno=T.sellno"
	sqlStr = sqlStr + " ,resellno=T.resellno"
	sqlStr = sqlStr + " ,totsellno=T.totsellno"
	sqlStr = sqlStr + " ,errcsno=T.errcsno"
	sqlStr = sqlStr + " ,errbaditemno=T.errbaditemno"
	sqlStr = sqlStr + " ,errrealcheckno=T.errrealcheckno"
	sqlStr = sqlStr + " ,erretcno=T.erretcno"
	sqlStr = sqlStr + " ,toterrno=T.toterrno"
	sqlStr = sqlStr + " ,offsellno=T.offsellno"
	sqlStr = sqlStr + " ,totsysstock=T.totsysstock"
	sqlStr = sqlStr + " ,availsysstock=T.availsysstock"
	sqlStr = sqlStr + " ,realstock=T.realstock"
	sqlStr = sqlStr + " from  [db_summary].[dbo].tbl_LAST_monthly_logisstock T"
	sqlStr = sqlStr + " where T.itemgubun='" + itemgubun + "'"
	sqlStr = sqlStr + " and T.itemid=" + CStr(itemid) + ""
	sqlStr = sqlStr + " and T.itemoption='" + itemoption + "'"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_current_logisstock_summary.itemgubun=T.itemgubun"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_current_logisstock_summary.itemid=T.itemid"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_current_logisstock_summary.itemoption=T.itemoption"

	rsget.Open sqlStr,dbget,1

	sqlStr = "update [db_summary].[dbo].tbl_current_logisstock_summary"
	sqlStr = sqlStr + " set ipgono=[db_summary].[dbo].tbl_current_logisstock_summary.ipgono + IsNULL(T.ipgono,0)"
	sqlStr = sqlStr + " ,reipgono=[db_summary].[dbo].tbl_current_logisstock_summary.reipgono + IsNULL(T.reipgono,0)"
	sqlStr = sqlStr + " ,totipgono=[db_summary].[dbo].tbl_current_logisstock_summary.totipgono + IsNULL(T.totipgono,0)"
	sqlStr = sqlStr + " ,offchulgono=[db_summary].[dbo].tbl_current_logisstock_summary.offchulgono + IsNULL(T.offchulgono,0)"
	sqlStr = sqlStr + " ,offrechulgono=[db_summary].[dbo].tbl_current_logisstock_summary.offrechulgono + IsNULL(T.offrechulgono,0)"
	sqlStr = sqlStr + " ,etcchulgono=[db_summary].[dbo].tbl_current_logisstock_summary.etcchulgono + IsNULL(T.etcchulgono,0)"
	sqlStr = sqlStr + " ,etcrechulgono=[db_summary].[dbo].tbl_current_logisstock_summary.etcrechulgono + IsNULL(T.etcrechulgono,0)"
	sqlStr = sqlStr + " ,totchulgono=[db_summary].[dbo].tbl_current_logisstock_summary.totchulgono + IsNULL(T.totchulgono,0)"
	sqlStr = sqlStr + " ,sellno=[db_summary].[dbo].tbl_current_logisstock_summary.sellno + IsNULL(T.sellno,0)"
	sqlStr = sqlStr + " ,resellno=[db_summary].[dbo].tbl_current_logisstock_summary.resellno + IsNULL(T.resellno,0)"
	sqlStr = sqlStr + " ,totsellno=[db_summary].[dbo].tbl_current_logisstock_summary.totsellno + IsNULL(T.totsellno,0)"
	sqlStr = sqlStr + " ,errcsno=[db_summary].[dbo].tbl_current_logisstock_summary.errcsno + IsNULL(T.errcsno,0)"
	sqlStr = sqlStr + " ,errbaditemno=[db_summary].[dbo].tbl_current_logisstock_summary.errbaditemno + IsNULL(T.errbaditemno,0)"
	sqlStr = sqlStr + " ,errrealcheckno=[db_summary].[dbo].tbl_current_logisstock_summary.errrealcheckno + IsNULL(T.errrealcheckno,0)"
	sqlStr = sqlStr + " ,erretcno=[db_summary].[dbo].tbl_current_logisstock_summary.erretcno + IsNULL(T.erretcno,0)"
	sqlStr = sqlStr + " ,toterrno=[db_summary].[dbo].tbl_current_logisstock_summary.toterrno + IsNULL(T.toterrno,0)"
	sqlStr = sqlStr + " ,offsellno=[db_summary].[dbo].tbl_current_logisstock_summary.offsellno + IsNULL(T.offsellno,0)"
	sqlStr = sqlStr + " ,totsysstock=[db_summary].[dbo].tbl_current_logisstock_summary.totsysstock + IsNULL(T.totsysstock,0)"
	sqlStr = sqlStr + " ,availsysstock=[db_summary].[dbo].tbl_current_logisstock_summary.availsysstock + IsNULL(T.availsysstock,0)"
	sqlStr = sqlStr + " ,realstock=[db_summary].[dbo].tbl_current_logisstock_summary.realstock + IsNULL(T.realstock,0)"
	sqlStr = sqlStr + " from  ("
	sqlStr = sqlStr + " select"
	sqlStr = sqlStr + " sum(ipgono) as ipgono,sum(reipgono) as reipgono,sum(totipgono) as totipgono,sum(offchulgono) as offchulgono,sum(offrechulgono) as offrechulgono,"
	sqlStr = sqlStr + " sum(etcchulgono) as etcchulgono,sum(etcrechulgono) as etcrechulgono,sum(totchulgono) as totchulgono,"
	sqlStr = sqlStr + " sum(sellno) as sellno,sum(resellno) as resellno,sum(totsellno) as totsellno,sum(errcsno) as errcsno,"
	sqlStr = sqlStr + " sum(errbaditemno) as errbaditemno,sum(errrealcheckno) as errrealcheckno,sum(erretcno) as erretcno,"
	sqlStr = sqlStr + " sum(toterrno) as toterrno,sum(offsellno) as offsellno, sum(totsysstock) as totsysstock,sum(availsysstock) as availsysstock,sum(realstock) as realstock"
	sqlStr = sqlStr + "  from [db_summary].[dbo].tbl_daily_logisstock_summary"
	sqlStr = sqlStr + " where yyyymmdd>='" + refreshstartdate + "'"
	sqlStr = sqlStr + " and itemgubun='" + itemgubun + "'"
	sqlStr = sqlStr + " and itemid=" + CStr(itemid) + ""
	sqlStr = sqlStr + " and itemoption='" + itemoption + "'"
	sqlStr = sqlStr + " ) T"
	sqlStr = sqlStr + " where [db_summary].[dbo].tbl_current_logisstock_summary.itemgubun='" + itemgubun + "'"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_current_logisstock_summary.itemid=" + CStr(itemid) + ""
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_current_logisstock_summary.itemoption='" + itemoption + "'"

	rsget.Open sqlStr,dbget,1

	response.write "<small>현재재고 업데이트... finish(" + FormatNumber(Timer()-oTimer,2) + " sec)</small><br>"
	response.flush

	oTimer = Timer()

	sqlStr = " update [db_summary].[dbo].tbl_current_logisstock_summary"
	sqlStr = sqlStr + " set ipkumdiv5=IsNULL(T.ipkumdiv5,0)"
	sqlStr = sqlStr + " ,ipkumdiv4=IsNULL(T.ipkumdiv4,0)"
	sqlStr = sqlStr + " ,ipkumdiv2=IsNULL(T.ipkumdiv2,0)"
	sqlStr = sqlStr + " from ("
	sqlStr = sqlStr + " select d.itemid,d.itemoption, "
	sqlStr = sqlStr + " sum(case when m.ipkumdiv='2' then d.itemno*-1 else 0 end ) as ipkumdiv2,"
	sqlStr = sqlStr + " sum(case when m.ipkumdiv='4' then d.itemno*-1 else 0 end ) as ipkumdiv4,"
	sqlStr = sqlStr + " sum(case when m.ipkumdiv in ('5','6') then d.itemno*-1 else 0 end ) as ipkumdiv5"
	sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m,"
	sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d"
	sqlStr = sqlStr + " where m.orderserial=d.orderserial"
	sqlStr = sqlStr + " and m.regdate>='" + refreshstartdate + "'"
	sqlStr = sqlStr + " and m.cancelyn='N'"
	sqlStr = sqlStr + " and d.cancelyn<>'Y'"
	sqlStr = sqlStr + " and d.itemid=" + CStr(itemid) + ""
	sqlStr = sqlStr + " and d.itemoption='" + itemoption + "'"
	sqlStr = sqlStr + " and d.isupchebeasong<>'Y'"
	sqlStr = sqlStr + " group by d.itemid,d.itemoption"
	sqlStr = sqlStr + " ) T"
	sqlStr = sqlStr + " where [db_summary].[dbo].tbl_current_logisstock_summary.itemgubun='10'"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_current_logisstock_summary.itemid=" + CStr(itemid) + ""
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_current_logisstock_summary.itemoption='" + itemoption + "'"

	rsget.Open sqlStr,dbget,1

	response.write "<small>현재 판매 업데이트... finish(" + FormatNumber(Timer()-oTimer,2) + " sec)</small><br>"
	response.flush

	oTimer = Timer()

	sqlStr = " update [db_summary].[dbo].tbl_current_logisstock_summary"
	sqlStr = sqlStr + " set sell7days=IsNULL(T.sell7days,0)"
	sqlStr = sqlStr + " from ("
	sqlStr = sqlStr + " select sum(d.itemno*-1) as sell7days"
	sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m,"
	sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d"
	sqlStr = sqlStr + " where m.orderserial=d.orderserial"
	sqlStr = sqlStr + " and m.regdate>='" + recent7day + "'"
	sqlStr = sqlStr + " and m.jumundiv<>'9'"
	sqlStr = sqlStr + " and m.ipkumdiv>1"
	sqlStr = sqlStr + " and m.cancelyn='N'"
	sqlStr = sqlStr + " and d.itemid=" + CStr(itemid) + ""
	sqlStr = sqlStr + " and d.itemoption='" + itemoption + "'"
	sqlStr = sqlStr + " and d.cancelyn<>'Y'"
	sqlStr = sqlStr + " and d.isupchebeasong<>'Y'"
	sqlStr = sqlStr + " ) as T"
	sqlStr = sqlStr + " where [db_summary].[dbo].tbl_current_logisstock_summary.itemgubun='10'"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_current_logisstock_summary.itemid=" + CStr(itemid) + ""
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_current_logisstock_summary.itemoption='" + itemoption + "'"

	rsget.Open sqlStr,dbget,1

	response.write "<small>온라인 7일 주문수량.. 업데이트... finish(" + FormatNumber(Timer()-oTimer,2) + " sec)</small><br>"
	response.flush

	oTimer = Timer()

	sqlStr = " update [db_summary].[dbo].tbl_current_logisstock_summary"
	sqlStr = sqlStr + " set offchulgo7days=IsNULL(T.chulno7*-1,0)"
	sqlStr = sqlStr + " from ("
	sqlStr = sqlStr + " select "
	sqlStr = sqlStr + " Sum(d.itemno) as chulno7"
	sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shopjumun_master m,"
	sqlStr = sqlStr + " [db_shop].[dbo].tbl_shopjumun_detail d"
	sqlStr = sqlStr + " where m.idx=d.masteridx"
	sqlStr = sqlStr + " and m.shopregdate>='" + recent7day + "'"
	sqlStr = sqlStr + " and m.cancelyn='N'"
	sqlStr = sqlStr + " and d.cancelyn='N'"
	sqlStr = sqlStr + " and d.itemgubun='" + itemgubun + "'"
	sqlStr = sqlStr + " and d.itemid=" + CStr(itemid) + ""
	sqlStr = sqlStr + " and d.itemoption='" + itemoption + "'"
	sqlStr = sqlStr + " and d.itemno>0"
	sqlStr = sqlStr + " ) as T"
	sqlStr = sqlStr + " where [db_summary].[dbo].tbl_current_logisstock_summary.itemgubun='" + itemgubun + "'"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_current_logisstock_summary.itemid=" + CStr(itemid) + ""
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_current_logisstock_summary.itemoption='" + itemoption + "'"

	rsget.Open sqlStr,dbget,1

	response.write "<small>오프 7일 주문수량 업데이트... finish(" + FormatNumber(Timer()-oTimer,2) + " sec)</small><br>"
	response.flush

	oTimer = Timer()

	sqlStr = " update [db_summary].[dbo].tbl_current_logisstock_summary"
	sqlStr = sqlStr + " set offconfirmno=IsNULL(T.offconfirmno,0)"
	sqlStr = sqlStr + " ,offjupno=IsNULL(T.offjupno,0)"
	sqlStr = sqlStr + " from ( select "
	sqlStr = sqlStr + " sum(case  when statecd='0' then baljuitemno*-1 else 0 end ) as offjupno,"
	sqlStr = sqlStr + " sum(case  when statecd<>'0' then baljuitemno*-1 else 0 end ) as offconfirmno"
	sqlStr = sqlStr + "  from "
	sqlStr = sqlStr + " [db_storage].[dbo].tbl_ordersheet_master m,"
	sqlStr = sqlStr + " [db_storage].[dbo].tbl_ordersheet_detail d"
	sqlStr = sqlStr + " where m.idx=d.masteridx"
	sqlStr = sqlStr + " and m.deldt is null"
	sqlStr = sqlStr + " and m.statecd<7"
	sqlStr = sqlStr + " and m.targetid='10x10'"
	sqlStr = sqlStr + " and m.divcode in ("
	sqlStr = sqlStr + " '501','502','503'"
	sqlStr = sqlStr + " )"
	sqlStr = sqlStr + " and d.itemgubun='" + itemgubun + "'"
	sqlStr = sqlStr + " and d.itemid=" + CStr(itemid) + ""
	sqlStr = sqlStr + " and d.itemoption='" + itemoption + "'"
	sqlStr = sqlStr + " ) T"
	sqlStr = sqlStr + " where [db_summary].[dbo].tbl_current_logisstock_summary.itemgubun='10'"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_current_logisstock_summary.itemid=" + CStr(itemid) + ""
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_current_logisstock_summary.itemoption='" + itemoption + "'"

	rsget.Open sqlStr,dbget,1

	response.write "<small>오프 주문 업데이트... finish(" + FormatNumber(Timer()-oTimer,2) + " sec)</small><br>"
	response.flush

	oTimer = Timer()

	sqlStr = " update [db_summary].[dbo].tbl_current_logisstock_summary"
	sqlStr = sqlStr + " set preorderno=IsNULL(T.preorderno,0)"
	sqlStr = sqlStr + " ,preordernofix=IsNULL(T.preordernofix,0)"
	sqlStr = sqlStr + " from ("
	sqlStr = sqlStr + " select sum(baljuitemno) as preorderno, sum(realitemno) as preordernofix  "
	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_master m,"
	sqlStr = sqlStr + " [db_storage].[dbo].tbl_ordersheet_detail d"
	sqlStr = sqlStr + " where m.idx=d.masteridx"
	sqlStr = sqlStr + " and m.deldt is null"
	sqlStr = sqlStr + " and m.ipgodate is null"
	sqlStr = sqlStr + " and datediff(d,m.scheduledate,getdate())<10"
	sqlStr = sqlStr + " and m.baljuid='10x10'"
	sqlStr = sqlStr + " and m.statecd<9"
	sqlStr = sqlStr + " and m.divcode in ("
	sqlStr = sqlStr + " '300','301','302'"
	sqlStr = sqlStr + " )"
	sqlStr = sqlStr + " and d.itemgubun='" + itemgubun + "'"
	sqlStr = sqlStr + " and d.itemid=" + CStr(itemid) + ""
	sqlStr = sqlStr + " and d.itemoption='" + itemoption + "'"
	sqlStr = sqlStr + " and d.deldt is null"
	sqlStr = sqlStr + " ) as T"
	sqlStr = sqlStr + " where [db_summary].[dbo].tbl_current_logisstock_summary.itemgubun='" + itemgubun + "'"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_current_logisstock_summary.itemid=" + CStr(itemid)
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_current_logisstock_summary.itemoption='" + itemoption + "'"

	rsget.Open sqlStr,dbget,1

	response.write "<small>기주문 업데이트... finish(" + FormatNumber(Timer()-oTimer,2) + " sec)</small><br>"
	response.flush

	sqlStr = " update [db_summary].[dbo].tbl_current_logisstock_summary"
	sqlStr = sqlStr + " set itemoptionname=v.codeview"
	sqlStr = sqlStr + " from [db_item].[dbo].vw_all_option v "
	sqlStr = sqlStr + " where [db_summary].[dbo].tbl_current_logisstock_summary.itemgubun='" + itemgubun + "'"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_current_logisstock_summary.itemid=" + CStr(itemid) + ""
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_current_logisstock_summary.itemoption='" + itemoption + "'"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_current_logisstock_summary.itemoption=v.optioncode"

	rsget.Open sqlStr,dbget,1

	oTimer = Timer()

	if itemgubun="10" then
		sqlStr = " update [db_summary].[dbo].tbl_current_logisstock_summary"
		sqlStr = sqlStr + " set maxsellday=(case when datediff(d,T.regdate,getdate()) >7 then 7 else datediff(d,T.regdate,getdate()) end)"
		sqlStr = sqlStr + " from [db_item].[dbo].tbl_item T"
		sqlStr = sqlStr + " where T.itemid=" + CStr(itemid)
		sqlStr = sqlStr + " and [db_summary].[dbo].tbl_current_logisstock_summary.itemid=T.itemid"

		rsget.Open sqlStr,dbget,1

		response.write "<small>판매일 업데이트... finish(" + FormatNumber(Timer()-oTimer,2) + " sec)</small><br>"
		response.flush
	end if

	oTimer = Timer()

	sqlStr = " update [db_summary].[dbo].tbl_current_logisstock_summary"
	sqlStr = sqlStr + " set requireno=convert(int,(sell7days+offchulgo7days)*7/maxsellday)"
	sqlStr = sqlStr + " where itemgubun='" + itemgubun + "'"
	sqlStr = sqlStr + " and itemid=" + CStr(itemid)
	sqlStr = sqlStr + " and itemoption='" + itemoption + "'"
	sqlStr = sqlStr + " and maxsellday<>0"

	rsget.Open sqlStr,dbget,1

	sqlStr = " update [db_summary].[dbo].tbl_current_logisstock_summary"
	sqlStr = sqlStr + " set shortageno=realstock+requireno"
	'sqlStr = sqlStr + " set shortageno=realstock+requireno+offjupno+offconfirmno+ipkumdiv2"
	sqlStr = sqlStr + " where itemgubun='" + itemgubun + "'"
	sqlStr = sqlStr + " and itemid=" + CStr(itemid)
	sqlStr = sqlStr + " and itemoption='" + itemoption + "'"

	rsget.Open sqlStr,dbget,1

	'// 재고기준일수 별도 등록 상품만
	sqlStr = " exec [db_summary].[dbo].[usp_Ten_Refresh_MakeItem_RequireNO] '" + CStr(itemgubun) + "', " + CStr(itemid) + ", '" + CStr(itemoption) + "' "
	rsget.Open sqlStr,dbget,1
	''response.write sqlStr

	response.write "<small>최종 업데이트... finish(" + FormatNumber(Timer()-oTimer,2) + " sec)</small><br>"
	response.flush

	'sqlStr = sqlStr + " ,sell7days,offchulgo7days,ipkumdiv5,ipkumdiv4,ipkumdiv2,"
	'sqlStr = sqlStr + " offconfirmno,offjupno,requireno,shortageno,preorderno,maxsellday,"

elseif mode="ipchulallrefreshbyitemid" then
	if (itemid="") then
		response.write "<script>alert('상품코드를 입력하세요.');</script>"
		response.write "<script>history.back();</script>"
		dbget.close()	:	response.End
	end if

	''daily_logisstock_summary
	oTimer = Timer()

	''입출고 있는내역 입력
	sqlStr = " insert into [db_summary].[dbo].tbl_daily_logisstock_summary"
	sqlStr = sqlStr + " (yyyymmdd,itemgubun,itemid,itemoption)"
	sqlStr = sqlStr + " select T.yyyymmdd, T.iitemgubun, T.itemid, T.itemoption"
	sqlStr = sqlStr + " from ("
	sqlStr = sqlStr + " 	select distinct convert(varchar(10),m.executedt,21) as yyyymmdd ,iitemgubun,itemid,itemoption"
	sqlStr = sqlStr + "   	from [db_storage].[dbo].tbl_acount_storage_master m, "
	sqlStr = sqlStr + " 	[db_storage].[dbo].tbl_acount_storage_detail d"
	sqlStr = sqlStr + " 	where m.code=d.mastercode"
	sqlStr = sqlStr + " 	and m.deldt is null"
	sqlStr = sqlStr + " 	and m.executedt is not null"
	sqlStr = sqlStr + " 	and d.deldt is null"
	sqlStr = sqlStr + " 	and d.iitemgubun='" + itemgubun + "'"
	sqlStr = sqlStr + " 	and d.itemid=" + CStr(itemid)
	sqlStr = sqlStr + " 	and d.itemoption='" + itemoption + "'"
	sqlStr = sqlStr + " ) T"
	sqlStr = sqlStr + " left join [db_summary].[dbo].tbl_daily_logisstock_summary s"
	sqlStr = sqlStr + " on T.yyyymmdd=s.yyyymmdd"
	sqlStr = sqlStr + " and T.iitemgubun=s.itemgubun"
	sqlStr = sqlStr + " and T.itemid=s.itemid"
	sqlStr = sqlStr + " and T.itemoption=s.itemoption"
	sqlStr = sqlStr + " where s.yyyymmdd is null"

 	rsget.Open sqlStr,dbget,1

	''deldt is not counting
	sqlStr = "update [db_summary].[dbo].tbl_daily_logisstock_summary"
 	sqlStr = sqlStr + " set ipgono=0"
 	sqlStr = sqlStr + " ,reipgono=0"
 	sqlStr = sqlStr + " ,totipgono=0"
	sqlStr = sqlStr + " ,offchulgono=0"
	sqlStr = sqlStr + " ,offrechulgono=0"
	sqlStr = sqlStr + " ,etcchulgono=0"
	sqlStr = sqlStr + " ,etcrechulgono=0"
	sqlStr = sqlStr + " ,totchulgono=0"
	sqlStr = sqlStr + " ,lastupdate=getdate()"
	sqlStr = sqlStr + " where itemgubun='" + itemgubun + "'"
	sqlStr = sqlStr + " and itemid=" + CStr(itemid) + ""
	sqlStr = sqlStr + " and itemoption='" + itemoption + "'"

	rsget.Open sqlStr,dbget,1

	sqlStr = "update [db_summary].[dbo].tbl_daily_logisstock_summary"
 	sqlStr = sqlStr + " set ipgono=T.ipgono"
 	sqlStr = sqlStr + " ,reipgono=T.reipgono"
 	sqlStr = sqlStr + " ,totipgono=T.ipgono+T.reipgono"
	sqlStr = sqlStr + " ,offchulgono=T.offchulgono"
	sqlStr = sqlStr + " ,offrechulgono=T.offrechulgono"
	sqlStr = sqlStr + " ,etcchulgono=T.etcchulgono"
	sqlStr = sqlStr + " ,etcrechulgono=T.etcrechulgono"
	sqlStr = sqlStr + " ,totchulgono=T.offchulgono+T.offrechulgono+T.etcchulgono+T.etcrechulgono"
	sqlStr = sqlStr + " ,lastupdate=getdate()"
 	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	( select convert(varchar(10),m.executedt,21) as yyyymmdd"
	sqlStr = sqlStr + " 	,iitemgubun,itemid,itemoption, "
	sqlStr = sqlStr + " 	sum(case when  ipchulflag='I' and itemno>0 then itemno else 0 end) as ipgono,"
	sqlStr = sqlStr + " 	sum(case when  ipchulflag='I' and itemno<0 then itemno else 0 end) as reipgono,"
	sqlStr = sqlStr + " 	sum(case when  ipchulflag='S' and itemno<0 then itemno else 0 end) as offchulgono,"
	sqlStr = sqlStr + " 	sum(case when  ipchulflag='S' and itemno>0 then itemno else 0 end) as offrechulgono,"
	sqlStr = sqlStr + " 	sum(case when  ipchulflag='E' and itemno<0 then itemno else 0 end) as etcchulgono,"
	sqlStr = sqlStr + " 	sum(case when  ipchulflag='E' and itemno>0 then itemno else 0 end) as etcrechulgono"
	sqlStr = sqlStr + "   	from [db_storage].[dbo].tbl_acount_storage_master m, "
	sqlStr = sqlStr + " 	[db_storage].[dbo].tbl_acount_storage_detail d"
	sqlStr = sqlStr + " 	where m.code=d.mastercode"
	sqlStr = sqlStr + " 	and m.deldt is null"
	sqlStr = sqlStr + " 	and d.iitemgubun='" + itemgubun + "'"
	sqlStr = sqlStr + " 	and d.itemid=" + CStr(itemid) + ""
	sqlStr = sqlStr + " 	and d.itemoption='" + itemoption + "'"
	sqlStr = sqlStr + " 	and d.deldt is null"
	sqlStr = sqlStr + " 	group by convert(varchar(10),m.executedt,21),iitemgubun,itemid,itemoption"
	sqlStr = sqlStr + " ) T"
	sqlStr = sqlStr + " where [db_summary].[dbo].tbl_daily_logisstock_summary.yyyymmdd=T.yyyymmdd"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_daily_logisstock_summary.itemgubun=T.iitemgubun"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_daily_logisstock_summary.itemid=T.itemid"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_daily_logisstock_summary.itemoption=T.itemoption"

	rsget.Open sqlStr,dbget,1

	sqlStr = " update [db_summary].[dbo].tbl_daily_logisstock_summary"
	sqlStr = sqlStr + " set totsysstock=totipgono+totchulgono-totsellno"
	sqlStr = sqlStr + " ,availsysstock=totipgono+totchulgono-totsellno+errcsno+errbaditemno+erretcno"
	sqlStr = sqlStr + " ,realstock=totipgono+totchulgono-totsellno+errcsno+errbaditemno+erretcno+errrealcheckno"
	sqlStr = sqlStr + " ,lastupdate=getdate()"
	sqlStr = sqlStr + " where lastupdate>='" + nowdate + "'"
	sqlStr = sqlStr + " and itemgubun='" + itemgubun + "'"
	sqlStr = sqlStr + " and itemid=" + CStr(itemid) + ""
	sqlStr = sqlStr + " and itemoption='" + itemoption + "'"

	rsget.Open sqlStr,dbget,1

	response.write "<small>일별 재고 업데이트... finish(" + FormatNumber(Timer()-oTimer,2) + " sec)</small><br>"
	response.flush

	oTimer = Timer()

	''입출고 있는내역 입력
	sqlStr = " insert into [db_summary].[dbo].tbl_monthly_logisstock_summary"
	sqlStr = sqlStr + " (yyyymm,itemgubun,itemid,itemoption)"
	sqlStr = sqlStr + " select T.yyyymm, T.iitemgubun, T.itemid, T.itemoption"
	sqlStr = sqlStr + " from ("
	sqlStr = sqlStr + " 	select distinct convert(varchar(7),m.executedt,21) as yyyymm ,iitemgubun,itemid,itemoption"
	sqlStr = sqlStr + "   	from [db_storage].[dbo].tbl_acount_storage_master m, "
	sqlStr = sqlStr + " 	[db_storage].[dbo].tbl_acount_storage_detail d"
	sqlStr = sqlStr + " 	where m.code=d.mastercode"
	sqlStr = sqlStr + " 	and m.deldt is null"
	sqlStr = sqlStr + " 	and m.executedt is not null"
	sqlStr = sqlStr + " 	and d.deldt is null"
	sqlStr = sqlStr + " 	and d.iitemgubun='" + itemgubun + "'"
	sqlStr = sqlStr + " 	and d.itemid=" + CStr(itemid)
	sqlStr = sqlStr + " 	and d.itemoption='" + itemoption + "'"
	sqlStr = sqlStr + " ) T"
	sqlStr = sqlStr + " left join [db_summary].[dbo].tbl_monthly_logisstock_summary s"
	sqlStr = sqlStr + " on T.yyyymm=s.yyyymm"
	sqlStr = sqlStr + " and T.iitemgubun=s.itemgubun"
	sqlStr = sqlStr + " and T.itemid=s.itemid"
	sqlStr = sqlStr + " and T.itemoption=s.itemoption"
	sqlStr = sqlStr + " where s.yyyymm is null"

 	rsget.Open sqlStr,dbget,1

	''deldt is not counting
	sqlStr = "update [db_summary].[dbo].tbl_monthly_logisstock_summary"
 	sqlStr = sqlStr + " set ipgono=0"
 	sqlStr = sqlStr + " ,reipgono=0"
 	sqlStr = sqlStr + " ,totipgono=0"
	sqlStr = sqlStr + " ,offchulgono=0"
	sqlStr = sqlStr + " ,offrechulgono=0"
	sqlStr = sqlStr + " ,etcchulgono=0"
	sqlStr = sqlStr + " ,etcrechulgono=0"
	sqlStr = sqlStr + " ,totchulgono=0"
	sqlStr = sqlStr + " ,lastupdate=getdate()"
	sqlStr = sqlStr + " where itemgubun='" + itemgubun + "'"
	sqlStr = sqlStr + " and itemid=" + CStr(itemid) + ""
	sqlStr = sqlStr + " and itemoption='" + itemoption + "'"

	rsget.Open sqlStr,dbget,1

 	sqlStr = "update [db_summary].[dbo].tbl_monthly_logisstock_summary"
 	sqlStr = sqlStr + " set ipgono=T.ipgono"
 	sqlStr = sqlStr + " ,reipgono=T.reipgono"
 	sqlStr = sqlStr + " ,totipgono=T.ipgono+T.reipgono"
	sqlStr = sqlStr + " ,offchulgono=T.offchulgono"
	sqlStr = sqlStr + " ,offrechulgono=T.offrechulgono"
	sqlStr = sqlStr + " ,etcchulgono=T.etcchulgono"
	sqlStr = sqlStr + " ,etcrechulgono=T.etcrechulgono"
	sqlStr = sqlStr + " ,totchulgono=T.offchulgono+T.offrechulgono+T.etcchulgono+T.etcrechulgono"
	sqlStr = sqlStr + " ,lastupdate=getdate()"
 	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	( select convert(varchar(7),m.executedt,21) as yyyymm"
	sqlStr = sqlStr + " 	,iitemgubun,itemid,itemoption, "
	sqlStr = sqlStr + " 	sum(case when  ipchulflag='I' and itemno>0 then itemno else 0 end) as ipgono,"
	sqlStr = sqlStr + " 	sum(case when  ipchulflag='I' and itemno<0 then itemno else 0 end) as reipgono,"
	sqlStr = sqlStr + " 	sum(case when  ipchulflag='S' and itemno<0 then itemno else 0 end) as offchulgono,"
	sqlStr = sqlStr + " 	sum(case when  ipchulflag='S' and itemno>0 then itemno else 0 end) as offrechulgono,"
	sqlStr = sqlStr + " 	sum(case when  ipchulflag='E' and itemno<0 then itemno else 0 end) as etcchulgono,"
	sqlStr = sqlStr + " 	sum(case when  ipchulflag='E' and itemno>0 then itemno else 0 end) as etcrechulgono"
	sqlStr = sqlStr + "   	from [db_storage].[dbo].tbl_acount_storage_master m, "
	sqlStr = sqlStr + " 	[db_storage].[dbo].tbl_acount_storage_detail d"
	sqlStr = sqlStr + " 	where m.code=d.mastercode"
	sqlStr = sqlStr + " 	and m.deldt is null"
	sqlStr = sqlStr + " 	and d.iitemgubun='" + itemgubun + "'"
	sqlStr = sqlStr + " 	and d.itemid=" + CStr(itemid) + ""
	sqlStr = sqlStr + " 	and d.itemoption='" + itemoption + "'"
	sqlStr = sqlStr + " 	and d.deldt is null"
	sqlStr = sqlStr + " 	group by convert(varchar(7),m.executedt,21),iitemgubun,itemid,itemoption"
	sqlStr = sqlStr + " ) T"
	sqlStr = sqlStr + " where [db_summary].[dbo].tbl_monthly_logisstock_summary.yyyymm=T.yyyymm"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_monthly_logisstock_summary.itemgubun=T.iitemgubun"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_monthly_logisstock_summary.itemid=T.itemid"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_monthly_logisstock_summary.itemoption=T.itemoption"

	rsget.Open sqlStr,dbget,1

	response.write "<small>재고 월별 입출고입력... finish(" + FormatNumber(Timer()-oTimer,2) + " sec)</small><br>"
	response.flush

	oTimer = Timer()

	sqlStr = " update [db_summary].[dbo].tbl_monthly_logisstock_summary"
	sqlStr = sqlStr + " set totsysstock=totipgono+totchulgono-totsellno"
	sqlStr = sqlStr + " ,availsysstock=totipgono+totchulgono-totsellno+errcsno+errbaditemno+erretcno"
	sqlStr = sqlStr + " ,realstock=totipgono+totchulgono-totsellno+errcsno+errbaditemno+erretcno+errrealcheckno"
	sqlStr = sqlStr + " ,lastupdate=getdate()"
	sqlStr = sqlStr + " where lastupdate>='" + nowdate + "'"
	sqlStr = sqlStr + " and itemgubun='" + itemgubun + "'"
	sqlStr = sqlStr + " and itemid=" + CStr(itemid) + ""
	sqlStr = sqlStr + " and itemoption='" + itemoption + "'"

	rsget.Open sqlStr,dbget,1

	sqlStr = " update [db_summary].[dbo].tbl_LAST_monthly_logisstock"
	sqlStr = sqlStr + " set lastyyyymm='" + LastYYYYMM + "'"
	sqlStr = sqlStr + " ,ipgono=IsNULL(T.ipgono,0)"
	sqlStr = sqlStr + " ,reipgono=IsNULL(T.reipgono,0)"
	sqlStr = sqlStr + " ,totipgono=IsNULL(T.totipgono,0)"
	sqlStr = sqlStr + " ,offchulgono=IsNULL(T.offchulgono,0)"
	sqlStr = sqlStr + " ,offrechulgono=IsNULL(T.offrechulgono,0)"
	sqlStr = sqlStr + " ,etcchulgono=IsNULL(T.etcchulgono,0)"
	sqlStr = sqlStr + " ,etcrechulgono=IsNULL(T.etcrechulgono,0)"
	sqlStr = sqlStr + " ,totchulgono=IsNULL(T.totchulgono,0)"
	sqlStr = sqlStr + " ,totsysstock=IsNULL(T.totsysstock,0)"
	sqlStr = sqlStr + " ,availsysstock=IsNULL(T.availsysstock,0)"
	sqlStr = sqlStr + " ,realstock=IsNULL(T.realstock,0)"
	sqlStr = sqlStr + " ,lastupdate=getdate()"
	sqlStr = sqlStr + " from ("
	sqlStr = sqlStr + " select sum(ipgono) as ipgono,"
	sqlStr = sqlStr + " sum(reipgono) as reipgono, sum(totipgono) as totipgono, sum(offchulgono) as offchulgono,"
	sqlStr = sqlStr + " sum(offrechulgono) as offrechulgono, sum(etcchulgono) as etcchulgono,"
	sqlStr = sqlStr + " sum(etcrechulgono) as etcrechulgono, sum(totchulgono) as totchulgono,"
	sqlStr = sqlStr + " sum(totsysstock) as totsysstock,"
	sqlStr = sqlStr + " sum(availsysstock) as availsysstock, sum(realstock) as realstock"
	sqlStr = sqlStr + " from [db_summary].[dbo].tbl_monthly_logisstock_summary"
	sqlStr = sqlStr + " where yyyymm<'" + Left(refreshstartdate,7) + "'"
	sqlStr = sqlStr + " and itemgubun='" + itemgubun + "'"
	sqlStr = sqlStr + " and itemid=" + CStr(itemid) + ""
	sqlStr = sqlStr + " and itemoption='" + itemoption + "'"
	sqlStr = sqlStr + " ) T"
	sqlStr = sqlStr + " where [db_summary].[dbo].tbl_LAST_monthly_logisstock.itemgubun='" + itemgubun + "'"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_LAST_monthly_logisstock.itemid=" + CStr(itemid) + ""
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_LAST_monthly_logisstock.itemoption='" + itemoption + "'"

	rsget.Open sqlStr,dbget,1

	response.write "<small>2달전 최종 재고 업데이트... finish(" + FormatNumber(Timer()-oTimer,2) + " sec)</small><br>"
	response.flush

	oTimer = Timer()

	sqlStr = "insert into [db_summary].[dbo].tbl_current_logisstock_summary"
	sqlStr = sqlStr + " (itemgubun,itemid,itemoption)"
	sqlStr = sqlStr + " select T.itemgubun,T.itemid,T.itemoption"
	sqlStr = sqlStr + " from [db_summary].[dbo].tbl_LAST_monthly_logisstock T"
	sqlStr = sqlStr + " left join [db_summary].[dbo].tbl_current_logisstock_summary s"
	sqlStr = sqlStr + " on T.itemgubun=s.itemgubun"
	sqlStr = sqlStr + " and T.itemid=s.itemid"
	sqlStr = sqlStr + " and T.itemoption=s.itemoption"
	sqlStr = sqlStr + " where T.itemgubun='" + itemgubun + "'"
	sqlStr = sqlStr + " and T.itemid=" + CStr(itemid) + ""
	sqlStr = sqlStr + " and T.itemoption='" + itemoption + "'"
	sqlStr = sqlStr + " and s.itemgubun is NULL"

	rsget.Open sqlStr,dbget,1

	sqlStr = "update [db_summary].[dbo].tbl_current_logisstock_summary"
	sqlStr = sqlStr + " set ipgono=0"
	sqlStr = sqlStr + " ,reipgono=0"
	sqlStr = sqlStr + " ,totipgono=0"
	sqlStr = sqlStr + " ,offchulgono=0"
	sqlStr = sqlStr + " ,offrechulgono=0"
	sqlStr = sqlStr + " ,etcchulgono=0"
	sqlStr = sqlStr + " ,etcrechulgono=0"
	sqlStr = sqlStr + " ,totchulgono=0"
	sqlStr = sqlStr + " ,totsysstock=0"
	sqlStr = sqlStr + " ,availsysstock=0"
	sqlStr = sqlStr + " ,realstock=0"
	sqlStr = sqlStr + " where itemgubun='" + itemgubun + "'"
	sqlStr = sqlStr + " and itemid=" + CStr(itemid) + ""
	sqlStr = sqlStr + " and itemoption='" + itemoption + "'"
'response.write sqlStr
	rsget.Open sqlStr,dbget,1

	sqlStr = "update [db_summary].[dbo].tbl_current_logisstock_summary"
	sqlStr = sqlStr + " set ipgono=IsNULL(T.ipgono,0)"
	sqlStr = sqlStr + " ,reipgono=IsNULL(T.reipgono,0)"
	sqlStr = sqlStr + " ,totipgono=IsNULL(T.totipgono,0)"
	sqlStr = sqlStr + " ,offchulgono=IsNULL(T.offchulgono,0)"
	sqlStr = sqlStr + " ,offrechulgono=IsNULL(T.offrechulgono,0)"
	sqlStr = sqlStr + " ,etcchulgono=IsNULL(T.etcchulgono,0)"
	sqlStr = sqlStr + " ,etcrechulgono=IsNULL(T.etcrechulgono,0)"
	sqlStr = sqlStr + " ,totchulgono=IsNULL(T.totchulgono,0)"
	sqlStr = sqlStr + " ,totsysstock=IsNULL(T.totsysstock,0)"
	sqlStr = sqlStr + " ,availsysstock=IsNULL(T.availsysstock,0)"
	sqlStr = sqlStr + " ,realstock=IsNULL(T.realstock,0)"
	sqlStr = sqlStr + " from [db_summary].[dbo].tbl_LAST_monthly_logisstock T"
	sqlStr = sqlStr + " where T.itemgubun='" + itemgubun + "'"
	sqlStr = sqlStr + " and T.itemid=" + CStr(itemid) + ""
	sqlStr = sqlStr + " and T.itemoption='" + itemoption + "'"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_current_logisstock_summary.itemgubun=T.itemgubun"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_current_logisstock_summary.itemid=T.itemid"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_current_logisstock_summary.itemoption=T.itemoption"

	rsget.Open sqlStr,dbget,1

	sqlStr = "insert into [db_summary].[dbo].tbl_current_logisstock_summary"
	sqlStr = sqlStr + " (itemgubun,itemid,itemoption)"
	sqlStr = sqlStr + " select T.itemgubun,T.itemid,T.itemoption from"
	sqlStr = sqlStr + " ("
	sqlStr = sqlStr + " 	select distinct itemgubun,itemid,itemoption"
	sqlStr = sqlStr + "  	from [db_summary].[dbo].tbl_daily_logisstock_summary"
	sqlStr = sqlStr + " 	where yyyymmdd>='" + refreshstartdate + "'"
	sqlStr = sqlStr + " 	and itemgubun='" + itemgubun + "'"
	sqlStr = sqlStr + " 	and itemid=" + CStr(itemid) + ""
	sqlStr = sqlStr + " 	and itemoption='" + itemoption + "'"
	sqlStr = sqlStr + " ) T"
	sqlStr = sqlStr + " left join [db_summary].[dbo].tbl_current_logisstock_summary s"
	sqlStr = sqlStr + " on T.itemgubun=s.itemgubun"
	sqlStr = sqlStr + " and T.itemid=s.itemid"
	sqlStr = sqlStr + " and T.itemoption=s.itemoption"
	sqlStr = sqlStr + " where s.itemgubun is NULL"

	rsget.Open sqlStr,dbget,1

	sqlStr = "update [db_summary].[dbo].tbl_current_logisstock_summary"
	sqlStr = sqlStr + " set ipgono=[db_summary].[dbo].tbl_current_logisstock_summary.ipgono + IsNULL(T.ipgono,0)"
	sqlStr = sqlStr + " ,reipgono=[db_summary].[dbo].tbl_current_logisstock_summary.reipgono + IsNULL(T.reipgono,0)"
	sqlStr = sqlStr + " ,totipgono=[db_summary].[dbo].tbl_current_logisstock_summary.totipgono + IsNULL(T.totipgono,0)"
	sqlStr = sqlStr + " ,offchulgono=[db_summary].[dbo].tbl_current_logisstock_summary.offchulgono + IsNULL(T.offchulgono,0)"
	sqlStr = sqlStr + " ,offrechulgono=[db_summary].[dbo].tbl_current_logisstock_summary.offrechulgono + IsNULL(T.offrechulgono,0)"
	sqlStr = sqlStr + " ,etcchulgono=[db_summary].[dbo].tbl_current_logisstock_summary.etcchulgono + IsNULL(T.etcchulgono,0)"
	sqlStr = sqlStr + " ,etcrechulgono=[db_summary].[dbo].tbl_current_logisstock_summary.etcrechulgono + IsNULL(T.etcrechulgono,0)"
	sqlStr = sqlStr + " ,totchulgono=[db_summary].[dbo].tbl_current_logisstock_summary.totchulgono + IsNULL(T.totchulgono,0)"
	sqlStr = sqlStr + " ,totsysstock=[db_summary].[dbo].tbl_current_logisstock_summary.totsysstock + IsNULL(T.totsysstock,0)"
	sqlStr = sqlStr + " ,availsysstock=[db_summary].[dbo].tbl_current_logisstock_summary.availsysstock + IsNULL(T.availsysstock,0)"
	sqlStr = sqlStr + " ,realstock=[db_summary].[dbo].tbl_current_logisstock_summary.realstock + IsNULL(T.realstock,0)"
	sqlStr = sqlStr + " from  ("
	sqlStr = sqlStr + " select"
	sqlStr = sqlStr + " sum(ipgono) as ipgono,sum(reipgono) as reipgono,sum(totipgono) as totipgono,sum(offchulgono) as offchulgono,sum(offrechulgono) as offrechulgono,"
	sqlStr = sqlStr + " sum(etcchulgono) as etcchulgono,sum(etcrechulgono) as etcrechulgono,sum(totchulgono) as totchulgono,"
	sqlStr = sqlStr + " sum(sellno) as sellno,sum(resellno) as resellno,sum(totsellno) as totsellno,sum(errcsno) as errcsno,"
	sqlStr = sqlStr + " sum(errbaditemno) as errbaditemno,sum(errrealcheckno) as errrealcheckno,sum(erretcno) as erretcno,"
	sqlStr = sqlStr + " sum(toterrno) as toterrno,sum(offsellno) as offsellno, sum(totsysstock) as totsysstock,sum(availsysstock) as availsysstock,sum(realstock) as realstock"
	sqlStr = sqlStr + "  from [db_summary].[dbo].tbl_daily_logisstock_summary"
	sqlStr = sqlStr + " where yyyymmdd>='" + refreshstartdate + "'"
	sqlStr = sqlStr + " and itemgubun='" + itemgubun + "'"
	sqlStr = sqlStr + " and itemid=" + CStr(itemid) + ""
	sqlStr = sqlStr + " and itemoption='" + itemoption + "'"
	sqlStr = sqlStr + " ) T"
	sqlStr = sqlStr + " where [db_summary].[dbo].tbl_current_logisstock_summary.itemgubun='" + itemgubun + "'"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_current_logisstock_summary.itemid=" + CStr(itemid) + ""
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_current_logisstock_summary.itemoption='" + itemoption + "'"

	rsget.Open sqlStr,dbget,1

	response.write "<small>현재 재고 업데이트... finish(" + FormatNumber(Timer()-oTimer,2) + " sec)</small><br>"
	response.flush

	sqlStr = " update [db_summary].[dbo].tbl_current_logisstock_summary"
	sqlStr = sqlStr + " set shortageno=realstock+requireno"
	sqlStr = sqlStr + " where itemgubun='" + itemgubun + "'"
	sqlStr = sqlStr + " and itemid=" + CStr(itemid)
	sqlStr = sqlStr + " and itemoption='" + itemoption + "'"

	rsget.Open sqlStr,dbget,1

elseif mode="ipchuldellbyitemid" then
	if (itemid="") then
		response.write "<script>alert('상품코드를 입력하세요.');</script>"
		response.write "<script>history.back();</script>"
		dbget.close()	:	response.End
	end if

	sqlStr = " delete from [db_summary].[dbo].tbl_daily_logisstock_summary"
	sqlStr = sqlStr + " where itemgubun='" + itemgubun + "'"
	sqlStr = sqlStr + " and itemid=" + CStr(itemid) + ""
	sqlStr = sqlStr + " and itemoption='" + itemoption + "'"

	rsget.Open sqlStr,dbget,1

	sqlStr = " delete from [db_summary].[dbo].tbl_monthly_logisstock_summary"
	sqlStr = sqlStr + " where itemgubun='" + itemgubun + "'"
	sqlStr = sqlStr + " and itemid=" + CStr(itemid) + ""
	sqlStr = sqlStr + " and itemoption='" + itemoption + "'"

	rsget.Open sqlStr,dbget,1

	sqlStr = " delete [db_summary].[dbo].tbl_LAST_monthly_logisstock"
	sqlStr = sqlStr + " where itemgubun='" + itemgubun + "'"
	sqlStr = sqlStr + " and itemid=" + CStr(itemid) + ""
	sqlStr = sqlStr + " and itemoption='" + itemoption + "'"

	rsget.Open sqlStr,dbget,1

	sqlStr = "delete [db_summary].[dbo].tbl_current_logisstock_summary"
	sqlStr = sqlStr + " where itemgubun='" + itemgubun + "'"
	sqlStr = sqlStr + " and itemid=" + CStr(itemid) + ""
	sqlStr = sqlStr + " and itemoption='" + itemoption + "'"

	rsget.Open sqlStr,dbget,1

elseif mode="ipchulallrefreshbybrand" then
	if (makerid="") then
		response.write "<script>alert('브랜드를 입력하세요.');</script>"
		response.write "<script>history.back();</script>"
		dbget.close()	:	response.End
	end if

	''daily_logisstock_summary
	oTimer = Timer()

	''deldt is not counting
	sqlStr = "update [db_summary].[dbo].tbl_daily_logisstock_summary"
 	sqlStr = sqlStr + " set ipgono=0"
 	sqlStr = sqlStr + " ,reipgono=0"
 	sqlStr = sqlStr + " ,totipgono=0"
	sqlStr = sqlStr + " ,offchulgono=0"
	sqlStr = sqlStr + " ,offrechulgono=0"
	sqlStr = sqlStr + " ,etcchulgono=0"
	sqlStr = sqlStr + " ,etcrechulgono=0"
	sqlStr = sqlStr + " ,totchulgono=0"
	sqlStr = sqlStr + " ,lastupdate=getdate()"
	sqlStr = sqlStr + " from ("
	sqlStr = sqlStr + " 	select s.itemid,s.itemoption "
	sqlStr = sqlStr + " 	from [db_summary].[dbo].tbl_daily_logisstock_summary s"
	sqlStr = sqlStr + " 	,[db_item].[dbo].tbl_item i "
	sqlStr = sqlStr + " 	where s.itemgubun='10' "
	sqlStr = sqlStr + " 	and s.itemid=i.itemid "
	sqlStr = sqlStr + " 	and i.makerid='" + makerid + "'"
	sqlStr = sqlStr + " ) T"
	sqlStr = sqlStr + " where [db_summary].[dbo].tbl_daily_logisstock_summary.itemgubun='10'"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_daily_logisstock_summary.itemid=T.itemid"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_daily_logisstock_summary.itemoption=T.itemoption"

	rsget.Open sqlStr,dbget,1

	''입출고 있는내역 입력 - ToDo 현재 브랜드 기준으로 변경..
	sqlStr = " insert into [db_summary].[dbo].tbl_daily_logisstock_summary"
	sqlStr = sqlStr + " (yyyymmdd,itemgubun,itemid,itemoption)"
	sqlStr = sqlStr + " select T.yyyymmdd, T.iitemgubun, T.itemid, T.itemoption"
	sqlStr = sqlStr + " from ("
	sqlStr = sqlStr + " 	select distinct convert(varchar(10),m.executedt,21) as yyyymmdd ,d.iitemgubun,d.itemid,d.itemoption"
	sqlStr = sqlStr + "   	from [db_storage].[dbo].tbl_acount_storage_master m, "
	sqlStr = sqlStr + " 	[db_storage].[dbo].tbl_acount_storage_detail d,"
	sqlStr = sqlStr + " 	[db_item].[dbo].tbl_item i"
	sqlStr = sqlStr + " 	where m.code=d.mastercode"
	sqlStr = sqlStr + " 	and m.deldt is null"
	sqlStr = sqlStr + " 	and d.itemid=i.itemid"
	sqlStr = sqlStr + " 	and d.iitemgubun='10'"
	sqlStr = sqlStr + " 	and i.makerid='" + makerid + "'"
	sqlStr = sqlStr + " 	and d.deldt is null"
	sqlStr = sqlStr + " ) T"
	sqlStr = sqlStr + " left join [db_summary].[dbo].tbl_daily_logisstock_summary s"
	sqlStr = sqlStr + " on T.yyyymmdd=s.yyyymmdd"
	sqlStr = sqlStr + " and T.iitemgubun=s.itemgubun"
	sqlStr = sqlStr + " and T.itemid=s.itemid"
	sqlStr = sqlStr + " and T.itemoption=s.itemoption"
	sqlStr = sqlStr + " where s.yyyymmdd is null"

 	rsget.Open sqlStr,dbget,1

	sqlStr = "update [db_summary].[dbo].tbl_daily_logisstock_summary"
 	sqlStr = sqlStr + " set ipgono=T.ipgono"
 	sqlStr = sqlStr + " ,reipgono=T.reipgono"
 	sqlStr = sqlStr + " ,totipgono=T.ipgono+T.reipgono"
	sqlStr = sqlStr + " ,offchulgono=T.offchulgono"
	sqlStr = sqlStr + " ,offrechulgono=T.offrechulgono"
	sqlStr = sqlStr + " ,etcchulgono=T.etcchulgono"
	sqlStr = sqlStr + " ,etcrechulgono=T.etcrechulgono"
	sqlStr = sqlStr + " ,totchulgono=T.offchulgono+T.offrechulgono+T.etcchulgono+T.etcrechulgono"
	sqlStr = sqlStr + " ,lastupdate=getdate()"
 	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	( select convert(varchar(10),m.executedt,21) as yyyymmdd"
	sqlStr = sqlStr + " 	,d.iitemgubun,d.itemid,d.itemoption, "
	sqlStr = sqlStr + " 	sum(case when  ipchulflag='I' and itemno>0 then itemno else 0 end) as ipgono,"
	sqlStr = sqlStr + " 	sum(case when  ipchulflag='I' and itemno<0 then itemno else 0 end) as reipgono,"
	sqlStr = sqlStr + " 	sum(case when  ipchulflag='S' and itemno<0 then itemno else 0 end) as offchulgono,"
	sqlStr = sqlStr + " 	sum(case when  ipchulflag='S' and itemno>0 then itemno else 0 end) as offrechulgono,"
	sqlStr = sqlStr + " 	sum(case when  ipchulflag='E' and itemno<0 then itemno else 0 end) as etcchulgono,"
	sqlStr = sqlStr + " 	sum(case when  ipchulflag='E' and itemno>0 then itemno else 0 end) as etcrechulgono"
	sqlStr = sqlStr + "   	from [db_storage].[dbo].tbl_acount_storage_master m, "
	sqlStr = sqlStr + " 	[db_storage].[dbo].tbl_acount_storage_detail d,"
	sqlStr = sqlStr + " 	[db_item].[dbo].tbl_item i"
	sqlStr = sqlStr + " 	where m.code=d.mastercode"
	sqlStr = sqlStr + " 	and m.deldt is null"
	sqlStr = sqlStr + " 	and d.itemid=i.itemid"
	sqlStr = sqlStr + " 	and d.iitemgubun='10'"
	sqlStr = sqlStr + " 	and i.makerid='" + makerid + "'"
	sqlStr = sqlStr + " 	and d.deldt is null"
	sqlStr = sqlStr + " 	group by convert(varchar(10),m.executedt,21),d.iitemgubun,d.itemid,d.itemoption"
	sqlStr = sqlStr + " ) T"
	sqlStr = sqlStr + " where [db_summary].[dbo].tbl_daily_logisstock_summary.yyyymmdd=T.yyyymmdd"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_daily_logisstock_summary.itemgubun=T.iitemgubun"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_daily_logisstock_summary.itemid=T.itemid"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_daily_logisstock_summary.itemoption=T.itemoption"

	rsget.Open sqlStr,dbget,1

	sqlStr = " update [db_summary].[dbo].tbl_daily_logisstock_summary"
	sqlStr = sqlStr + " set totsysstock=totipgono+totchulgono-totsellno"
	sqlStr = sqlStr + " ,availsysstock=totipgono+totchulgono-totsellno+errcsno+errbaditemno+erretcno"
	sqlStr = sqlStr + " ,realstock=totipgono+totchulgono-totsellno+errcsno+errbaditemno+erretcno+errrealcheckno"
	sqlStr = sqlStr + " ,lastupdate=getdate()"
	sqlStr = sqlStr + " from ("
	sqlStr = sqlStr + " 	select s.itemid,s.itemoption "
	sqlStr = sqlStr + " 	from [db_summary].[dbo].tbl_daily_logisstock_summary s"
	sqlStr = sqlStr + " 	,[db_item].[dbo].tbl_item i "
	sqlStr = sqlStr + " 	where s.itemgubun='10' "
	sqlStr = sqlStr + " 	and s.itemid=i.itemid "
	sqlStr = sqlStr + " 	and i.makerid='" + makerid + "'"
	sqlStr = sqlStr + " ) T"
	sqlStr = sqlStr + " where lastupdate>='" + nowdate + "'"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_daily_logisstock_summary.itemgubun='10'"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_daily_logisstock_summary.itemid=T.itemid"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_daily_logisstock_summary.itemoption=T.itemoption"

	rsget.Open sqlStr,dbget,1

	response.write "<small>일별 재고 업데이트... finish(" + FormatNumber(Timer()-oTimer,2) + " sec)</small><br>"
	response.flush

	oTimer = Timer()
	''재고 일별테이블 업데이트(입출고)

	sqlStr = "update [db_summary].[dbo].tbl_monthly_logisstock_summary"
	sqlStr = sqlStr + " set ipgono=0"
 	sqlStr = sqlStr + " ,reipgono=0"
 	sqlStr = sqlStr + " ,totipgono=0"
	sqlStr = sqlStr + " ,offchulgono=0"
	sqlStr = sqlStr + " ,offrechulgono=0"
	sqlStr = sqlStr + " ,etcchulgono=0"
	sqlStr = sqlStr + " ,etcrechulgono=0"
	sqlStr = sqlStr + " ,totchulgono=0"
	sqlStr = sqlStr + " ,lastupdate=getdate()"
	sqlStr = sqlStr + " from ("
	sqlStr = sqlStr + " 	select s.itemid,s.itemoption "
	sqlStr = sqlStr + " 	from [db_summary].[dbo].tbl_monthly_logisstock_summary s"
	sqlStr = sqlStr + " 	,[db_item].[dbo].tbl_item i "
	sqlStr = sqlStr + " 	where s.itemgubun='10' "
	sqlStr = sqlStr + " 	and s.itemid=i.itemid "
	sqlStr = sqlStr + " 	and i.makerid='" + makerid + "'"
	sqlStr = sqlStr + " ) T"
	sqlStr = sqlStr + " where [db_summary].[dbo].tbl_monthly_logisstock_summary.itemgubun='10'"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_monthly_logisstock_summary.itemid=T.itemid"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_monthly_logisstock_summary.itemoption=T.itemoption"

	rsget.Open sqlStr,dbget,1

	''입출고 있는내역 입력
	sqlStr = " insert into [db_summary].[dbo].tbl_monthly_logisstock_summary"
	sqlStr = sqlStr + " (yyyymm,itemgubun,itemid,itemoption)"
	sqlStr = sqlStr + " select T.yyyymm, T.iitemgubun, T.itemid, T.itemoption"
	sqlStr = sqlStr + " from ("
	sqlStr = sqlStr + " 	select distinct convert(varchar(7),m.executedt,21) as yyyymm ,d.iitemgubun,d.itemid,d.itemoption"
	sqlStr = sqlStr + "   	from [db_storage].[dbo].tbl_acount_storage_master m, "
	sqlStr = sqlStr + " 	[db_storage].[dbo].tbl_acount_storage_detail d,"
	sqlStr = sqlStr + " 	[db_item].[dbo].tbl_item i "
	sqlStr = sqlStr + " 	where m.code=d.mastercode"
	sqlStr = sqlStr + " 	and m.deldt is null"
	sqlStr = sqlStr + " 	and d.itemid=i.itemid"
	sqlStr = sqlStr + " 	and d.iitemgubun='10'"
	sqlStr = sqlStr + " 	and i.makerid='" + makerid + "'"
	sqlStr = sqlStr + " 	and d.deldt is null"
	sqlStr = sqlStr + " ) T"
	sqlStr = sqlStr + " left join [db_summary].[dbo].tbl_monthly_logisstock_summary s"
	sqlStr = sqlStr + " on T.yyyymm=s.yyyymm"
	sqlStr = sqlStr + " and T.iitemgubun=s.itemgubun"
	sqlStr = sqlStr + " and T.itemid=s.itemid"
	sqlStr = sqlStr + " and T.itemoption=s.itemoption"
	sqlStr = sqlStr + " where s.yyyymm is null"

 	rsget.Open sqlStr,dbget,1

 	sqlStr = "update [db_summary].[dbo].tbl_monthly_logisstock_summary"
 	sqlStr = sqlStr + " set ipgono=T.ipgono"
 	sqlStr = sqlStr + " ,reipgono=T.reipgono"
 	sqlStr = sqlStr + " ,totipgono=T.ipgono+T.reipgono"
	sqlStr = sqlStr + " ,offchulgono=T.offchulgono"
	sqlStr = sqlStr + " ,offrechulgono=T.offrechulgono"
	sqlStr = sqlStr + " ,etcchulgono=T.etcchulgono"
	sqlStr = sqlStr + " ,etcrechulgono=T.etcrechulgono"
	sqlStr = sqlStr + " ,totchulgono=T.offchulgono+T.offrechulgono+T.etcchulgono+T.etcrechulgono"
	sqlStr = sqlStr + " ,lastupdate=getdate()"
 	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " ( 	select convert(varchar(7),m.executedt,21) as yyyymm"
	sqlStr = sqlStr + " 	,d.iitemgubun,d.itemid,d.itemoption, "
	sqlStr = sqlStr + " 	sum(case when  ipchulflag='I' and itemno>0 then itemno else 0 end) as ipgono,"
	sqlStr = sqlStr + " 	sum(case when  ipchulflag='I' and itemno<0 then itemno else 0 end) as reipgono,"
	sqlStr = sqlStr + " 	sum(case when  ipchulflag='S' and itemno<0 then itemno else 0 end) as offchulgono,"
	sqlStr = sqlStr + " 	sum(case when  ipchulflag='S' and itemno>0 then itemno else 0 end) as offrechulgono,"
	sqlStr = sqlStr + " 	sum(case when  ipchulflag='E' and itemno<0 then itemno else 0 end) as etcchulgono,"
	sqlStr = sqlStr + " 	sum(case when  ipchulflag='E' and itemno>0 then itemno else 0 end) as etcrechulgono"
	sqlStr = sqlStr + "   	from [db_storage].[dbo].tbl_acount_storage_master m, "
	sqlStr = sqlStr + " 	[db_storage].[dbo].tbl_acount_storage_detail d,"
	sqlStr = sqlStr + " 	[db_item].[dbo].tbl_item i "
	sqlStr = sqlStr + " 	where m.code=d.mastercode"
	sqlStr = sqlStr + " 	and m.deldt is null"
	sqlStr = sqlStr + " 	and d.itemid=i.itemid"
	sqlStr = sqlStr + " 	and d.iitemgubun='10'"
	sqlStr = sqlStr + " 	and i.makerid='" + makerid + "'"
	sqlStr = sqlStr + " 	and d.deldt is null"
	sqlStr = sqlStr + " 	group by convert(varchar(7),m.executedt,21),d.iitemgubun,d.itemid,d.itemoption"
	sqlStr = sqlStr + " ) T"
	sqlStr = sqlStr + " where [db_summary].[dbo].tbl_monthly_logisstock_summary.yyyymm=T.yyyymm"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_monthly_logisstock_summary.itemgubun=T.iitemgubun"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_monthly_logisstock_summary.itemid=T.itemid"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_monthly_logisstock_summary.itemoption=T.itemoption"

	rsget.Open sqlStr,dbget,1

	response.write "<small>재고 월별 입출고입력... finish(" + FormatNumber(Timer()-oTimer,2) + " sec)</small><br>"
	response.flush

	oTimer = Timer()

	sqlStr = " update [db_summary].[dbo].tbl_monthly_logisstock_summary"
	sqlStr = sqlStr + " set totsysstock=totipgono+totchulgono-totsellno"
	sqlStr = sqlStr + " ,availsysstock=totipgono+totchulgono-totsellno+errcsno+errbaditemno+erretcno"
	sqlStr = sqlStr + " ,realstock=totipgono+totchulgono-totsellno+errcsno+errbaditemno+erretcno+errrealcheckno"
	sqlStr = sqlStr + " ,lastupdate=getdate()"
	sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i"
	sqlStr = sqlStr + " where [db_summary].[dbo].tbl_monthly_logisstock_summary.lastupdate>='" + nowdate + "'"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_monthly_logisstock_summary.itemgubun='10'"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_monthly_logisstock_summary.itemid=i.itemid"
	sqlStr = sqlStr + " and i.makerid='" + makerid + "'"

	rsget.Open sqlStr,dbget,1

	sqlStr = " update [db_summary].[dbo].tbl_LAST_monthly_logisstock"
	sqlStr = sqlStr + " set lastyyyymm='" + LastYYYYMM + "'"
	sqlStr = sqlStr + " ,ipgono=IsNULL(T.ipgono,0)"
	sqlStr = sqlStr + " ,reipgono=IsNULL(T.reipgono,0)"
	sqlStr = sqlStr + " ,totipgono=IsNULL(T.totipgono,0)"
	sqlStr = sqlStr + " ,offchulgono=IsNULL(T.offchulgono,0)"
	sqlStr = sqlStr + " ,offrechulgono=IsNULL(T.offrechulgono,0)"
	sqlStr = sqlStr + " ,etcchulgono=IsNULL(T.etcchulgono,0)"
	sqlStr = sqlStr + " ,etcrechulgono=IsNULL(T.etcrechulgono,0)"
	sqlStr = sqlStr + " ,totchulgono=IsNULL(T.totchulgono,0)"
	sqlStr = sqlStr + " ,totsysstock=IsNULL(T.totsysstock,0)"
	sqlStr = sqlStr + " ,availsysstock=IsNULL(T.availsysstock,0)"
	sqlStr = sqlStr + " ,realstock=IsNULL(T.realstock,0)"
	sqlStr = sqlStr + " ,lastupdate=getdate()"
	sqlStr = sqlStr + " from ("
	sqlStr = sqlStr + " select s.itemgubun,s.itemid,s.itemoption, sum(s.ipgono) as ipgono,"
	sqlStr = sqlStr + " sum(s.reipgono) as reipgono, sum(s.totipgono) as totipgono, sum(s.offchulgono) as offchulgono,"
	sqlStr = sqlStr + " sum(s.offrechulgono) as offrechulgono, sum(s.etcchulgono) as etcchulgono,"
	sqlStr = sqlStr + " sum(s.etcrechulgono) as etcrechulgono, sum(s.totchulgono) as totchulgono,"
	sqlStr = sqlStr + " sum(s.totsysstock) as totsysstock,"
	sqlStr = sqlStr + " sum(s.availsysstock) as availsysstock, sum(s.realstock) as realstock"
	sqlStr = sqlStr + " from [db_summary].[dbo].tbl_monthly_logisstock_summary s, [db_item].[dbo].tbl_item i"
	sqlStr = sqlStr + " where s.yyyymm<'" + Left(refreshstartdate,7) + "'"
	sqlStr = sqlStr + " and s.itemgubun='10'"
	sqlStr = sqlStr + " and s.itemid=i.itemid"
	sqlStr = sqlStr + " and i.makerid='" + makerid + "'"
	sqlStr = sqlStr + " group by s.itemgubun,s.itemid,s.itemoption"
	sqlStr = sqlStr + " ) T"

	sqlStr = sqlStr + " where [db_summary].[dbo].tbl_LAST_monthly_logisstock.itemgubun=T.itemgubun"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_LAST_monthly_logisstock.itemid=T.itemid"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_LAST_monthly_logisstock.itemoption=T.itemoption"

	rsget.Open sqlStr,dbget,1

	response.write "<small>2달전 최종 재고 업데이트... finish(" + FormatNumber(Timer()-oTimer,2) + " sec)</small><br>"
	response.flush

	oTimer = Timer()

	sqlStr = "insert into [db_summary].[dbo].tbl_current_logisstock_summary"
	sqlStr = sqlStr + " (itemgubun,itemid,itemoption)"
	sqlStr = sqlStr + " select T.itemgubun,T.itemid,T.itemoption"
	sqlStr = sqlStr + " from ("
	sqlStr = sqlStr + " 	select L.itemgubun, L.itemid, L.itemoption "
	sqlStr = sqlStr + " 	from [db_summary].[dbo].tbl_LAST_monthly_logisstock L, [db_item].[dbo].tbl_item i"
	sqlStr = sqlStr + " 	where L.itemgubun='10' "
	sqlStr = sqlStr + " 	and L.itemid=i.itemid "
	sqlStr = sqlStr + " 	and i.makerid='" + makerid + "'"
	sqlStr = sqlStr + " ) T"
	sqlStr = sqlStr + " left join [db_summary].[dbo].tbl_current_logisstock_summary s"
	sqlStr = sqlStr + " on T.itemgubun=s.itemgubun"
	sqlStr = sqlStr + " and T.itemid=s.itemid"
	sqlStr = sqlStr + " and T.itemoption=s.itemoption"
	sqlStr = sqlStr + " where s.itemgubun is NULL"

	rsget.Open sqlStr,dbget,1

	sqlStr = "update [db_summary].[dbo].tbl_current_logisstock_summary"
	sqlStr = sqlStr + " set ipgono=0"
	sqlStr = sqlStr + " ,reipgono=0"
	sqlStr = sqlStr + " ,totipgono=0"
	sqlStr = sqlStr + " ,offchulgono=0"
	sqlStr = sqlStr + " ,offrechulgono=0"
	sqlStr = sqlStr + " ,etcchulgono=0"
	sqlStr = sqlStr + " ,etcrechulgono=0"
	sqlStr = sqlStr + " ,totchulgono=0"
	sqlStr = sqlStr + " ,totsysstock=0"
	sqlStr = sqlStr + " ,availsysstock=0"
	sqlStr = sqlStr + " ,realstock=0"
	sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i"
	sqlStr = sqlStr + " where [db_summary].[dbo].tbl_current_logisstock_summary.itemgubun='10'"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_current_logisstock_summary.itemid=i.itemid"
	sqlStr = sqlStr + " and i.makerid='" + makerid + "'"

	rsget.Open sqlStr,dbget,1

	sqlStr = "update [db_summary].[dbo].tbl_current_logisstock_summary"
	sqlStr = sqlStr + " set ipgono=T.ipgono"
	sqlStr = sqlStr + " ,reipgono=T.reipgono"
	sqlStr = sqlStr + " ,totipgono=T.totipgono"
	sqlStr = sqlStr + " ,offchulgono=T.offchulgono"
	sqlStr = sqlStr + " ,offrechulgono=T.offrechulgono"
	sqlStr = sqlStr + " ,etcchulgono=T.etcchulgono"
	sqlStr = sqlStr + " ,etcrechulgono=T.etcrechulgono"
	sqlStr = sqlStr + " ,totchulgono=T.totchulgono"
	sqlStr = sqlStr + " ,totsysstock=T.totsysstock"
	sqlStr = sqlStr + " ,availsysstock=T.availsysstock"
	sqlStr = sqlStr + " ,realstock=T.realstock"
	sqlStr = sqlStr + " from ("
	sqlStr = sqlStr + " 	select L.itemgubun, L.itemid, L.itemoption, "
	sqlStr = sqlStr + " 	L.ipgono, L.reipgono, L.totipgono, L.offchulgono, "
	sqlStr = sqlStr + " 	L.offrechulgono, L.etcchulgono, "
	sqlStr = sqlStr + " 	L.etcrechulgono, L.totchulgono, "
	sqlStr = sqlStr + " 	L.totsysstock, L.availsysstock, L.realstock "
	sqlStr = sqlStr + " 	from [db_summary].[dbo].tbl_LAST_monthly_logisstock L, [db_item].[dbo].tbl_item i"
	sqlStr = sqlStr + " 	where L.itemgubun='10' "
	sqlStr = sqlStr + " 	and L.itemid=i.itemid "
	sqlStr = sqlStr + " 	and i.makerid='" + makerid + "'"
	sqlStr = sqlStr + " ) T"
	sqlStr = sqlStr + " where [db_summary].[dbo].tbl_current_logisstock_summary.itemgubun=T.itemgubun"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_current_logisstock_summary.itemid=T.itemid"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_current_logisstock_summary.itemoption=T.itemoption"

	rsget.Open sqlStr,dbget,1

	sqlStr = "insert into [db_summary].[dbo].tbl_current_logisstock_summary"
	sqlStr = sqlStr + " (itemgubun,itemid,itemoption)"
	sqlStr = sqlStr + " select T.itemgubun,T.itemid,T.itemoption from"
	sqlStr = sqlStr + " ("
	sqlStr = sqlStr + " 	select distinct s.itemgubun,s.itemid,s.itemoption"
	sqlStr = sqlStr + " 	from [db_summary].[dbo].tbl_daily_logisstock_summary s, [db_item].[dbo].tbl_item i"
	sqlStr = sqlStr + " 	where yyyymmdd>='" + refreshstartdate + "'"
	sqlStr = sqlStr + " 	and s.itemgubun='10'"
	sqlStr = sqlStr + " 	and s.itemid=i.itemid"
	sqlStr = sqlStr + " 	and i.makerid='" + makerid + "'"
	sqlStr = sqlStr + " ) T"
	sqlStr = sqlStr + " left join [db_summary].[dbo].tbl_current_logisstock_summary s"
	sqlStr = sqlStr + " on T.itemgubun=s.itemgubun"
	sqlStr = sqlStr + " and T.itemid=s.itemid"
	sqlStr = sqlStr + " and T.itemoption=s.itemoption"
	sqlStr = sqlStr + " where s.itemgubun is NULL"

	rsget.Open sqlStr,dbget,1

	sqlStr = "update [db_summary].[dbo].tbl_current_logisstock_summary"
	sqlStr = sqlStr + " set ipgono=[db_summary].[dbo].tbl_current_logisstock_summary.ipgono + IsNULL(T.ipgono,0)"
	sqlStr = sqlStr + " ,reipgono=[db_summary].[dbo].tbl_current_logisstock_summary.reipgono + IsNULL(T.reipgono,0)"
	sqlStr = sqlStr + " ,totipgono=[db_summary].[dbo].tbl_current_logisstock_summary.totipgono + IsNULL(T.totipgono,0)"
	sqlStr = sqlStr + " ,offchulgono=[db_summary].[dbo].tbl_current_logisstock_summary.offchulgono + IsNULL(T.offchulgono,0)"
	sqlStr = sqlStr + " ,offrechulgono=[db_summary].[dbo].tbl_current_logisstock_summary.offrechulgono + IsNULL(T.offrechulgono,0)"
	sqlStr = sqlStr + " ,etcchulgono=[db_summary].[dbo].tbl_current_logisstock_summary.etcchulgono + IsNULL(T.etcchulgono,0)"
	sqlStr = sqlStr + " ,etcrechulgono=[db_summary].[dbo].tbl_current_logisstock_summary.etcrechulgono + IsNULL(T.etcrechulgono,0)"
	sqlStr = sqlStr + " ,totchulgono=[db_summary].[dbo].tbl_current_logisstock_summary.totchulgono + IsNULL(T.totchulgono,0)"
	sqlStr = sqlStr + " ,totsysstock=[db_summary].[dbo].tbl_current_logisstock_summary.totsysstock + IsNULL(T.totsysstock,0)"
	sqlStr = sqlStr + " ,availsysstock=[db_summary].[dbo].tbl_current_logisstock_summary.availsysstock + IsNULL(T.availsysstock,0)"
	sqlStr = sqlStr + " ,realstock=[db_summary].[dbo].tbl_current_logisstock_summary.realstock + IsNULL(T.realstock,0)"
	sqlStr = sqlStr + " from  ("
	sqlStr = sqlStr + " 	select s.itemgubun,s.itemid,s.itemoption,"
	sqlStr = sqlStr + " 	sum(ipgono) as ipgono,sum(reipgono) as reipgono,sum(totipgono) as totipgono,sum(offchulgono) as offchulgono,sum(offrechulgono) as offrechulgono,"
	sqlStr = sqlStr + " 	sum(etcchulgono) as etcchulgono,sum(etcrechulgono) as etcrechulgono,sum(totchulgono) as totchulgono,"
	sqlStr = sqlStr + " 	sum(totsysstock) as totsysstock,sum(availsysstock) as availsysstock,sum(realstock) as realstock"
	sqlStr = sqlStr + " 	from [db_summary].[dbo].tbl_daily_logisstock_summary s, [db_item].[dbo].tbl_item i"
	sqlStr = sqlStr + " 	where s.yyyymmdd>='" + refreshstartdate + "'"
	sqlStr = sqlStr + " 	and s.itemgubun='10'"
	sqlStr = sqlStr + " 	and s.itemid=i.itemid"
	sqlStr = sqlStr + " 	and i.makerid='" + makerid + "'"
	sqlStr = sqlStr + " 	group by s.itemgubun,s.itemid,s.itemoption"
	sqlStr = sqlStr + " ) T"
	sqlStr = sqlStr + " where [db_summary].[dbo].tbl_current_logisstock_summary.itemgubun=T.itemgubun"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_current_logisstock_summary.itemid=T.itemid"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_current_logisstock_summary.itemoption=T.itemoption"

	rsget.Open sqlStr,dbget,1

	response.write "<small>현재 재고 업데이트... finish(" + FormatNumber(Timer()-oTimer,2) + " sec)</small><br>"
	response.flush

	sqlStr = " update [db_summary].[dbo].tbl_current_logisstock_summary"
	sqlStr = sqlStr + " set shortageno=realstock+requireno"
	sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i"
	sqlStr = sqlStr + " where [db_summary].[dbo].tbl_current_logisstock_summary.itemgubun='10'"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_current_logisstock_summary.itemid=i.itemid"
	sqlStr = sqlStr + " and i.makerid='" + makerid + "'"

	rsget.Open sqlStr,dbget,1

elseif mode="itemtodayipchulrefresh" then
	'response.write "<script>alert('수정중..');</script>"
	'response.write "<script>history.back();</script>"
	dbget.close()	:	response.End

	''금일 필드여부확인
	sqlStr = " select count(*) as cnt from [db_summary].[dbo].tbl_deliver_itemsell_daily_summary"
	sqlStr = sqlStr + " where itemid=" + CStr(itemid)
	sqlStr = sqlStr + " and itemoption='" + itemoption + "'"

	rsget.Open sqlStr,dbget,1
		itemExists = rsget("cnt")>0
	rsget.close

	oTimer = Timer()

	if (itemExists) then
		sqlStr = " update [db_summary].[dbo].tbl_deliver_itemsell_daily_summary"
		sqlStr = sqlStr + " set sellno=IsNULL(T.sellno,0)"
		sqlStr = sqlStr + " ,reno=IsNULL(T.reno,0)"
		sqlStr = sqlStr + " ,totsellno=IsNULL(T.totsellno,0)"
		sqlStr = sqlStr + " from ("
		sqlStr = sqlStr + " select "
		sqlStr = sqlStr + " sum(case when d.itemno>0 then d.itemno else 0 end ) as sellno,"
		sqlStr = sqlStr + " sum(case when d.itemno<0 then d.itemno else 0 end ) as reno,"
		sqlStr = sqlStr + " sum(d.itemno) as totsellno"
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m,"
		sqlStr = sqlStr + "  [db_order].[dbo].tbl_order_detail d"
		sqlStr = sqlStr + " where m.orderserial=d.orderserial"
		sqlStr = sqlStr + " and m.ipkumdiv>=7"
		sqlStr = sqlStr + " and m.beadaldate>='" + nowdate + "'"
		sqlStr = sqlStr + " and m.cancelyn='N'"
		sqlStr = sqlStr + " and d.cancelyn<>'Y'"
		sqlStr = sqlStr + " and d.itemid=" + CStr(itemid) + ""
		sqlStr = sqlStr + " and d.itemoption='" + CStr(itemoption) + "'"
		sqlStr = sqlStr + " and d.isupchebeasong<>'Y'"
		sqlStr = sqlStr + " ) T"
		sqlStr = sqlStr + " where [db_summary].[dbo].tbl_deliver_itemsell_daily_summary.yyyymmdd='" + nowdate + "'"
		sqlStr = sqlStr + " and [db_summary].[dbo].tbl_deliver_itemsell_daily_summary.itemid=" + CStr(itemid) + ""
		sqlStr = sqlStr + " and [db_summary].[dbo].tbl_deliver_itemsell_daily_summary.itemoption='" + CStr(itemoption) + "'"

		rsget.Open sqlStr,dbget,1

	else
		sqlStr = " insert into [db_summary].[dbo].tbl_deliver_itemsell_daily_summary"
		sqlStr = sqlStr + " (yyyymmdd,itemid,itemoption,sellno,reno,totsellno)"
		sqlStr = sqlStr + " select '" + nowdate + "'," + CStr(itmid) + ",'" + itemoption + "',"
		sqlStr = sqlStr + " IsNULL(sum(case when d.itemno>0 then d.itemno else 0 end ),0) as sellno,"
		sqlStr = sqlStr + " IsNULL(sum(case when d.itemno<0 then d.itemno else 0 end ),0) as reno,"
		sqlStr = sqlStr + " IsNULL(sum(d.itemno),0) as totsellno"
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m,"
		sqlStr = sqlStr + "  [db_order].[dbo].tbl_order_detail d"
		sqlStr = sqlStr + " where m.orderserial=d.orderserial"
		sqlStr = sqlStr + " and m.ipkumdiv>=7"
		sqlStr = sqlStr + " and m.beadaldate>='" + nowdate + "'"
		sqlStr = sqlStr + " and m.cancelyn='N'"
		sqlStr = sqlStr + " and d.cancelyn<>'Y'"
		sqlStr = sqlStr + " and d.itemid=" + CStr(itemid) + ""
		sqlStr = sqlStr + " and d.itemoption='" + CStr(itemoption) + "'"
		sqlStr = sqlStr + " and d.isupchebeasong<>'Y'"

		rsget.Open sqlStr,dbget,1
	end if

	response.write "<small>금일 출고 완료 업데이트... finish (" + FormatNumber(Timer()-oTimer,2) + " sec)</small><br>"
	response.flush

	''금일 출고완료입력
	''금일 입출고입력
	oTimer = Timer()

	sqlStr = " update [db_summary].[dbo].tbl_current_logisstock_summary"
	sqlStr = sqlStr + " set ipkumdiv5=IsNULL(T.ipkumdiv5,0)"
	sqlStr = sqlStr + " ,ipkumdiv4=IsNULL(T.ipkumdiv4,0)"
	sqlStr = sqlStr + " ,ipkumdiv2=IsNULL(T.ipkumdiv2,0)"
	sqlStr = sqlStr + " from ("
	sqlStr = sqlStr + " select "
	sqlStr = sqlStr + " sum(case when m.ipkumdiv='2' then d.itemno*-1 else 0 end ) as ipkumdiv2,"
	sqlStr = sqlStr + " sum(case when m.ipkumdiv='4' then d.itemno*-1 else 0 end ) as ipkumdiv4,"
	sqlStr = sqlStr + " sum(case when m.ipkumdiv in ('5','6') then d.itemno*-1 else 0 end ) as ipkumdiv5"
	sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m,"
	sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d"
	sqlStr = sqlStr + " where m.orderserial=d.orderserial"
	sqlStr = sqlStr + " and m.ipkumdiv>1"
	sqlStr = sqlStr + " and m.ipkumdiv<7"
	sqlStr = sqlStr + " and m.regdate>='" + recent7day + "'"
	sqlStr = sqlStr + " and m.cancelyn='N'"
	sqlStr = sqlStr + " and d.cancelyn<>'Y'"
	sqlStr = sqlStr + " and d.itemid=" + CStr(itemid) + ""
	sqlStr = sqlStr + " and d.itemoption='" + itemoption + "'"
	sqlStr = sqlStr + " ) T"
	sqlStr = sqlStr + " where [db_summary].[dbo].tbl_current_logisstock_summary.itemgubun='10'"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_current_logisstock_summary.itemid=" + CStr(itemid) + ""
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_current_logisstock_summary.itemoption='" + itemoption + "'"

	rsget.Open sqlStr,dbget,1

	response.write "<small>현재 판매 업데이트... finish(" + FormatNumber(Timer()-oTimer,2) + " sec)</small><br>"
	response.flush

	oTimer = Timer()

	sqlStr = " update [db_summary].[dbo].tbl_current_logisstock_summary"
	sqlStr = sqlStr + " set sell7days=IsNULL(T.sell7days,0)"
	sqlStr = sqlStr + " from ("
	sqlStr = sqlStr + " select sum(d.itemno*-1) as sell7days"
	sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m,"
	sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d"
	sqlStr = sqlStr + " where m.orderserial=d.orderserial"
	sqlStr = sqlStr + " and m.regdate>='" + recent7day + "'"
	sqlStr = sqlStr + " and m.jumundiv<>'9'"
	sqlStr = sqlStr + " and m.ipkumdiv>1"
	sqlStr = sqlStr + " and m.cancelyn='N'"
	sqlStr = sqlStr + " and d.itemid=" + CStr(itemid) + ""
	sqlStr = sqlStr + " and d.itemoption='" + itemoption + "'"
	sqlStr = sqlStr + " and d.cancelyn<>'Y'"
	sqlStr = sqlStr + " ) as T"
	sqlStr = sqlStr + " where [db_summary].[dbo].tbl_current_logisstock_summary.itemgubun='10'"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_current_logisstock_summary.itemid=" + CStr(itemid) + ""
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_current_logisstock_summary.itemoption='" + itemoption + "'"

	rsget.Open sqlStr,dbget,1

	response.write "<small>온라인 7일 주문수량.. 업데이트... finish(" + FormatNumber(Timer()-oTimer,2) + " sec)</small><br>"
	response.flush

	oTimer = Timer()

	sqlStr = " update [db_summary].[dbo].tbl_current_logisstock_summary"
	sqlStr = sqlStr + " set offchulgo7days=IsNULL(T.chulno7,0)"
	sqlStr = sqlStr + " from ("
	sqlStr = sqlStr + " select "
	sqlStr = sqlStr + " Sum(d.itemno) as chulno7"
	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_master m,"
	sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_detail d"
	sqlStr = sqlStr + " where m.code=d.mastercode"
	sqlStr = sqlStr + " and m.ipchulflag='S'"
	sqlStr = sqlStr + " and m.scheduledt>='" + recent7day + "'"
	sqlStr = sqlStr + " and m.deldt is NULL"
	sqlStr = sqlStr + " and d.deldt is NULL"
	sqlStr = sqlStr + " and d.iitemgubun='" + itemgubun + "'"
	sqlStr = sqlStr + " and d.itemid=" + CStr(itemid) + ""
	sqlStr = sqlStr + " and d.itemoption='" + itemoption + "'"
	sqlStr = sqlStr + " ) as T"
	sqlStr = sqlStr + " where [db_summary].[dbo].tbl_current_logisstock_summary.itemgubun='" + itemgubun + "'"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_current_logisstock_summary.itemid=" + CStr(itemid) + ""
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_current_logisstock_summary.itemoption='" + itemoption + "'"

	rsget.Open sqlStr,dbget,1

	response.write "<small>오프 7일 주문수량 업데이트... finish(" + FormatNumber(Timer()-oTimer,2) + " sec)</small><br>"
	response.flush

	oTimer = Timer()

	sqlStr = " update [db_summary].[dbo].tbl_current_logisstock_summary"
	sqlStr = sqlStr + " set offconfirmno=IsNULL(T.offconfirmno,0)"
	sqlStr = sqlStr + " ,offjupno=IsNULL(T.offjupno,0)"
	sqlStr = sqlStr + " from ( select "
	sqlStr = sqlStr + " sum(case  when statecd='0' then baljuitemno*-1 else 0 end ) as offjupno,"
	sqlStr = sqlStr + " sum(case  when statecd<>'0' then baljuitemno*-1 else 0 end ) as offconfirmno"
	sqlStr = sqlStr + "  from "
	sqlStr = sqlStr + " [db_storage].[dbo].tbl_ordersheet_master m,"
	sqlStr = sqlStr + " [db_storage].[dbo].tbl_ordersheet_detail d"
	sqlStr = sqlStr + " where m.idx=d.masteridx"
	sqlStr = sqlStr + " and m.deldt is null"
	sqlStr = sqlStr + " and m.statecd<7"
	sqlStr = sqlStr + " and m.targetid='10x10'"
	sqlStr = sqlStr + " and m.divcode in ("
	sqlStr = sqlStr + " '501','502','503'"
	sqlStr = sqlStr + " )"
	sqlStr = sqlStr + " and d.itemgubun='" + itemgubun + "'"
	sqlStr = sqlStr + " and d.itemid=" + CStr(itemid) + ""
	sqlStr = sqlStr + " and d.itemoption='" + itemoption + "'"
	sqlStr = sqlStr + " ) T"
	sqlStr = sqlStr + " where [db_summary].[dbo].tbl_current_logisstock_summary.itemgubun='10'"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_current_logisstock_summary.itemid=" + CStr(itemid) + ""
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_current_logisstock_summary.itemoption='" + itemoption + "'"

	rsget.Open sqlStr,dbget,1

	response.write "<small>오프 주문 업데이트... finish(" + FormatNumber(Timer()-oTimer,2) + " sec)</small><br>"
	response.flush

	oTimer = Timer()

	sqlStr = " update [db_summary].[dbo].tbl_current_logisstock_summary"
	sqlStr = sqlStr + " set preorderno=IsNULL(T.preorderno,0)"
	sqlStr = sqlStr + " , preordernofix=IsNULL(T.preordernofix,0)"
	sqlStr = sqlStr + " from ("
	sqlStr = sqlStr + " select sum(baljuitemno) as preorderno, sum(realitemno) as preordernofix  "
	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_master m,"
	sqlStr = sqlStr + " [db_storage].[dbo].tbl_ordersheet_detail d"
	sqlStr = sqlStr + " where m.idx=d.masteridx"
	sqlStr = sqlStr + " and m.deldt is null"
	sqlStr = sqlStr + " and m.ipgodate is null"
	sqlStr = sqlStr + " and datediff(d,m.scheduledate,getdate())<10"
	sqlStr = sqlStr + " and m.baljuid='10x10'"
	sqlStr = sqlStr + " and m.statecd<9"
	sqlStr = sqlStr + " and m.divcode in ("
	sqlStr = sqlStr + " '300','301','302'"
	sqlStr = sqlStr + " )"
	sqlStr = sqlStr + " and d.itemgubun='" + itemgubun + "'"
	sqlStr = sqlStr + " and d.itemid=" + CStr(itemid) + ""
	sqlStr = sqlStr + " and d.itemoption='" + itemoption + "'"
	sqlStr = sqlStr + " and d.deldt is null"
	sqlStr = sqlStr + " ) as T"
	sqlStr = sqlStr + " where [db_summary].[dbo].tbl_current_logisstock_summary.itemgubun='" + itemgubun + "'"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_current_logisstock_summary.itemid=" + CStr(itemid)
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_current_logisstock_summary.itemoption='" + itemoption + "'"

	rsget.Open sqlStr,dbget,1

	response.write "<small>기주문 업데이트... finish(" + FormatNumber(Timer()-oTimer,2) + " sec)</small><br>"
	response.flush

	oTimer = Timer()

	if itemgubun="10" then
		sqlStr = " update [db_summary].[dbo].tbl_current_logisstock_summary"
		sqlStr = sqlStr + " set maxsellday=(case when datediff(d,T.regdate,getdate()) >7 then 7 else datediff(d,T.regdate,getdate()) end)"
		sqlStr = sqlStr + " from [db_item].[dbo].tbl_item T"
		sqlStr = sqlStr + " where T.itemid=" + CStr(itemid)
		sqlStr = sqlStr + " and [db_summary].[dbo].tbl_current_logisstock_summary.itemid=T.itemid"

		rsget.Open sqlStr,dbget,1

		response.write "<small>판매일 업데이트... finish(" + FormatNumber(Timer()-oTimer,2) + " sec)</small><br>"
		response.flush
	end if

	oTimer = Timer()

	sqlStr = " update [db_summary].[dbo].tbl_current_logisstock_summary"
	sqlStr = sqlStr + " set requireno=convert(int,(sell7days+offchulgo7days)*7/maxsellday)"
	sqlStr = sqlStr + " where itemgubun='" + itemgubun + "'"
	sqlStr = sqlStr + " and itemid=" + CStr(itemid)
	sqlStr = sqlStr + " and itemoption='" + itemoption + "'"
	sqlStr = sqlStr + " and maxsellday<>0"

	rsget.Open sqlStr,dbget,1

	'// 재고기준일수 별도 등록 상품만
	sqlStr = " exec [db_summary].[dbo].[usp_Ten_Refresh_MakeItem_RequireNO] '" + CStr(itemgubun) + "', " + CStr(itemid) + ", '" + CStr(itemoption) + "' "
	rsget.Open sqlStr,dbget,1
	''response.write sqlStr

	sqlStr = " update [db_summary].[dbo].tbl_current_logisstock_summary"
	sqlStr = sqlStr + " set shortageno=realstock+requireno"
	sqlStr = sqlStr + " where itemgubun='" + itemgubun + "'"
	sqlStr = sqlStr + " and itemid=" + CStr(itemid)
	sqlStr = sqlStr + " and itemoption='" + itemoption + "'"

	rsget.Open sqlStr,dbget,1

	response.write "<small>최종 업데이트... finish(" + FormatNumber(Timer()-oTimer,2) + " sec)</small><br>"
	response.flush

elseif mode="errcheckupdate" then
	''오차 입력일
	sqlStr = "select convert(varchar(10),getdate(),21) as yyyymmdd"
	rsget.Open sqlStr,dbget,1
		yyyymmdd = rsget("yyyymmdd")
	rsget.Close

	''현 실사 재고
	sqlStr = "select top 1 * from [db_summary].[dbo].tbl_current_logisstock_summary"
	sqlStr = sqlStr + " where itemgubun='" + itemgubun + "'"
	sqlStr = sqlStr + " and itemid=" + CStr(itemid)
	sqlStr = sqlStr + " and itemoption='" + itemoption + "'"

	rsget.Open sqlStr,dbget,1
		orgrealstock = rsget("realstock") + rsget("ipkumdiv5") + rsget("offconfirmno")
	rsget.Close

	''금일 입력 오차
	sqlStr = "select top 1 errrealcheckno from [db_summary].[dbo].tbl_erritem_daily_summary"
	sqlStr = sqlStr + " where itemgubun='" + itemgubun + "'"
	sqlStr = sqlStr + " and itemid=" + CStr(itemid)
	sqlStr = sqlStr + " and itemoption='" + itemoption + "'"
	sqlStr = sqlStr + " and yyyymmdd='" + yyyymmdd + "'"

	rsget.Open sqlStr,dbget,1
	if Not rsget.Eof then
		itemExists = true
		todayerrno = rsget("errrealcheckno")
	else
		itemExists = false
		todayerrno = 0
	end if
	rsget.Close

	''현재까지 입력된 입력 오차 + 금일 입력 오차
	errstock = realstock - orgrealstock + todayerrno

	if (itemExists) then
		sqlStr = " update [db_summary].[dbo].tbl_erritem_daily_summary"
		sqlStr = sqlStr + " set errrealcheckno=" + CStr(errstock)
		sqlStr = sqlStr + " ,modiuser='" + session("ssBctId") + "'"
		sqlStr = sqlStr + " ,lastupdate=getdate()"
		sqlStr = sqlStr + " where itemgubun='" + itemgubun + "'"
		sqlStr = sqlStr + " and itemid=" + CStr(itemid)
		sqlStr = sqlStr + " and itemoption='" + itemoption + "'"
		sqlStr = sqlStr + " and yyyymmdd='" + yyyymmdd + "'"

		rsget.Open sqlStr,dbget,1
	else
		sqlStr = " insert into [db_summary].[dbo].tbl_erritem_daily_summary"
		sqlStr = sqlStr + " (yyyymmdd,itemgubun,itemid,itemoption,errrealcheckno,reguser)"
		sqlStr = sqlStr + " values("
		sqlStr = sqlStr + " '" + yyyymmdd + "','" + itemgubun + "'," + CStr(itemid) + ",'" + itemoption + "'"
		sqlStr = sqlStr + " ," + CStr(errstock) + ""
		sqlStr = sqlStr + " ,'" + session("ssBctId") + "'"
		sqlStr = sqlStr + " )"

		rsget.Open sqlStr,dbget,1
	end if

	sqlStr = "update [db_summary].[dbo].tbl_erritem_daily_summary"
	sqlStr = sqlStr + " set toterrno=errcsno+errbaditemno+erretcno+errrealcheckno"
	sqlStr = sqlStr + " ,lastupdate=getdate()"
	sqlStr = sqlStr + " ,modiuser='" + session("ssBctId") + "'"
	sqlStr = sqlStr + " where itemgubun='" + itemgubun + "'"
	sqlStr = sqlStr + " and itemid=" + CStr(itemid)
	sqlStr = sqlStr + " and itemoption='" + itemoption + "'"
	sqlStr = sqlStr + " and yyyymmdd='" + yyyymmdd + "'"

	rsget.Open sqlStr,dbget,1

	''일별 재고로그에 추가
	sqlStr = "select top 1 * from [db_summary].[dbo].tbl_daily_logisstock_summary"
	sqlStr = sqlStr + " where itemgubun='" + itemgubun + "'"
	sqlStr = sqlStr + " and itemid=" + CStr(itemid)
	sqlStr = sqlStr + " and itemoption='" + itemoption + "'"
	sqlStr = sqlStr + " and yyyymmdd='" + yyyymmdd + "'"

	rsget.Open sqlStr,dbget,1
		itemExists = Not rsget.Eof
	rsget.Close

	if (itemExists) then
		sqlStr = " update [db_summary].[dbo].tbl_daily_logisstock_summary"
		sqlStr = sqlStr + " set errrealcheckno=" + CStr(errstock)
		sqlStr = sqlStr + " ,lastupdate=getdate()"
		sqlStr = sqlStr + " where yyyymmdd='" + yyyymmdd + "'"
		sqlStr = sqlStr + " and itemgubun='" + itemgubun + "'"
		sqlStr = sqlStr + " and itemid=" + CStr(itemid)
		sqlStr = sqlStr + " and itemoption='" + itemoption + "'"

		rsget.Open sqlStr,dbget,1
	else
		sqlStr = " insert into [db_summary].[dbo].tbl_daily_logisstock_summary"
		sqlStr = sqlStr + " (yyyymmdd,itemgubun,itemid,itemoption,errrealcheckno)"
		sqlStr = sqlStr + " values("
		sqlStr = sqlStr + " '" + yyyymmdd + "','" + itemgubun + "'," + CStr(itemid) + ",'" + itemoption + "'"
		sqlStr = sqlStr + " ," + CStr(errstock) + ""
		sqlStr = sqlStr + " )"

		rsget.Open sqlStr,dbget,1
	end if

	''서머리.
	sqlStr = " update [db_summary].[dbo].tbl_daily_logisstock_summary"
	sqlStr = sqlStr + " set toterrno=errcsno+errbaditemno+erretcno+errrealcheckno"
	sqlStr = sqlStr + " ,totsysstock=totipgono+totchulgono-totsellno"
	sqlStr = sqlStr + " ,availsysstock=totipgono+totchulgono-totsellno+errcsno+errbaditemno+erretcno"
	sqlStr = sqlStr + " ,realstock=totipgono+totchulgono-totsellno+errcsno+errbaditemno+erretcno+errrealcheckno"
	sqlStr = sqlStr + " ,lastupdate=getdate()"
	sqlStr = sqlStr + " where yyyymmdd='" + yyyymmdd + "'"
	sqlStr = sqlStr + " and itemgubun='" + itemgubun + "'"
	sqlStr = sqlStr + " and itemid=" + CStr(itemid) + ""
	sqlStr = sqlStr + " and itemoption='" + itemoption + "'"

	rsget.Open sqlStr,dbget,1

	dim toterrrealcheckno
	sqlStr = "select sum(errrealcheckno) as toterrrealcheckno from [db_summary].[dbo].tbl_erritem_daily_summary"
	sqlStr = sqlStr + " where itemgubun='" + itemgubun + "'"
	sqlStr = sqlStr + " and itemid=" + CStr(itemid)
	sqlStr = sqlStr + " and itemoption='" + itemoption + "'"
	rsget.Open sqlStr,dbget,1
		toterrrealcheckno = rsget("toterrrealcheckno")
	rsget.close

	if IsNULL(toterrrealcheckno) then toterrrealcheckno=0

	''현재고테이블 업데이트
	sqlStr = "update [db_summary].[dbo].tbl_current_logisstock_summary"
	sqlStr = sqlStr + " set errrealcheckno=" + CStr(toterrrealcheckno)
	sqlStr = sqlStr + " ,toterrno=errcsno+errbaditemno+erretcno+" + CStr(toterrrealcheckno)
	sqlStr = sqlStr + " ,totsysstock=totipgono+totchulgono-totsellno"
	sqlStr = sqlStr + " ,availsysstock=totipgono+totchulgono-totsellno+errcsno+errbaditemno+erretcno"
	sqlStr = sqlStr + " ,realstock=totipgono+totchulgono-totsellno+errcsno+errbaditemno+erretcno+" + CStr(toterrrealcheckno)
	'sqlStr = sqlStr + " ,lastupdate=getdate()"
	sqlStr = sqlStr + " where itemgubun='" + itemgubun + "'"
	sqlStr = sqlStr + " and itemid=" + CStr(itemid) + ""
	sqlStr = sqlStr + " and itemoption='" + itemoption + "'"

	rsget.Open sqlStr,dbget,1

elseif mode="tmpbaditem2input" then
	''오차 입력일
	sqlStr = "select convert(varchar(10),getdate(),21) as yyyymmdd"
	rsget.Open sqlStr,dbget,1
		yyyymmdd = rsget("yyyymmdd")
	rsget.Close

	''기입력된 오차에 추가함 (테이블에 상품이 존재할 경우)
	sqlStr = " update [db_summary].[dbo].tbl_erritem_daily_summary"
	sqlStr = sqlStr + " set errbaditemno=errbaditemno + IsNULL(T.itemno,0)*-1"
	sqlStr = sqlStr + " from ( "
	sqlStr = sqlStr + " select b.itemgubun, b.itemid, b.itemoption, b.itemno"
	sqlStr = sqlStr + " from [db_summary].[dbo].tbl_temp_baditem b, [db_summary].[dbo].tbl_erritem_daily_summary s"
	sqlStr = sqlStr + " where s.yyyymmdd='" + yyyymmdd + "'"
	sqlStr = sqlStr + " and b.itemgubun=s.itemgubun"
	sqlStr = sqlStr + " and b.itemid=s.itemid"
	sqlStr = sqlStr + " and b.itemoption=s.itemoption"
	sqlStr = sqlStr + " ) T"
	sqlStr = sqlStr + " where [db_summary].[dbo].tbl_erritem_daily_summary.yyyymmdd='" + yyyymmdd + "'"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_erritem_daily_summary.itemgubun=T.itemgubun"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_erritem_daily_summary.itemid=T.itemid"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_erritem_daily_summary.itemoption=T.itemoption"

	rsget.Open sqlStr,dbget,1

	''기입력된 오차에 추가함 (테이블에 상품이 없을 경우)
	sqlStr = " insert into [db_summary].[dbo].tbl_erritem_daily_summary"
	sqlStr = sqlStr + " (yyyymmdd,itemgubun,itemid,itemoption,errbaditemno,reguser)"
	sqlStr = sqlStr + " select "
	sqlStr = sqlStr + " '" + yyyymmdd + "'"
	sqlStr = sqlStr + " ,T.itemgubun,T.itemid,T.itemoption,T.itemno*-1,'" + session("ssBctId") + "'"
	sqlStr = sqlStr + " from ("
	sqlStr = sqlStr + " select b.* "
	sqlStr = sqlStr + " from [db_summary].[dbo].tbl_temp_baditem b "
	sqlStr = sqlStr + " left join [db_summary].[dbo].tbl_erritem_daily_summary s on s.yyyymmdd='" + yyyymmdd + "'"
	sqlStr = sqlStr + " and b.itemgubun=s.itemgubun"
	sqlStr = sqlStr + " and b.itemid=s.itemid"
	sqlStr = sqlStr + " and b.itemoption=s.itemoption"
	sqlStr = sqlStr + " where s.itemid is null"
	sqlStr = sqlStr + " ) T"

	rsget.Open sqlStr,dbget,1

	sqlStr = "update [db_summary].[dbo].tbl_erritem_daily_summary"
	sqlStr = sqlStr + " set toterrno=errbaditemno+erretcno+errrealcheckno" ''errcsno는 오차테이블에서 일반 입출고로 개념변경  errcsno+
	sqlStr = sqlStr + " ,lastupdate=getdate()"
	sqlStr = sqlStr + " ,modiuser='" + session("ssBctId") + "'"
	sqlStr = sqlStr + " from [db_summary].[dbo].tbl_temp_baditem b "
	sqlStr = sqlStr + " where [db_summary].[dbo].tbl_erritem_daily_summary.yyyymmdd='" + yyyymmdd + "'"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_erritem_daily_summary.itemgubun=b.itemgubun"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_erritem_daily_summary.itemid=b.itemid"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_erritem_daily_summary.itemoption=b.itemoption"

	rsget.Open sqlStr,dbget,1

	''일별 재고로그에 추가
	sqlStr = " update [db_summary].[dbo].tbl_daily_logisstock_summary"
	sqlStr = sqlStr + " set errbaditemno=errbaditemno + b.itemno*-1"
	sqlStr = sqlStr + " from [db_summary].[dbo].tbl_temp_baditem b "
	sqlStr = sqlStr + " where [db_summary].[dbo].tbl_daily_logisstock_summary.yyyymmdd='" + yyyymmdd + "'"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_daily_logisstock_summary.itemgubun=b.itemgubun"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_daily_logisstock_summary.itemid=b.itemid"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_daily_logisstock_summary.itemoption=b.itemoption"

	rsget.Open sqlStr,dbget,1

	sqlStr = " insert into [db_summary].[dbo].tbl_daily_logisstock_summary"
	sqlStr = sqlStr + " (yyyymmdd,itemgubun,itemid,itemoption,errbaditemno)"
	sqlStr = sqlStr + " select "
	sqlStr = sqlStr + " '" + yyyymmdd + "'"
	sqlStr = sqlStr + " ,T.itemgubun,T.itemid,T.itemoption,T.itemno*-1"
	sqlStr = sqlStr + " from ("
	sqlStr = sqlStr + " select b.* "
	sqlStr = sqlStr + " from [db_summary].[dbo].tbl_temp_baditem b "
	sqlStr = sqlStr + " left join [db_summary].[dbo].tbl_daily_logisstock_summary s on s.yyyymmdd='" + yyyymmdd + "'"
	sqlStr = sqlStr + " and b.itemgubun=s.itemgubun"
	sqlStr = sqlStr + " and b.itemid=s.itemid"
	sqlStr = sqlStr + " and b.itemoption=s.itemoption"
	sqlStr = sqlStr + " where s.itemid is null"
	sqlStr = sqlStr + " ) T"

	rsget.Open sqlStr,dbget,1

	''서머리.
	sqlStr = " update [db_summary].[dbo].tbl_daily_logisstock_summary"
	sqlStr = sqlStr + " set toterrno=errbaditemno+erretcno+errrealcheckno"  ''errcsno+
	sqlStr = sqlStr + " ,totsysstock=totipgono+totchulgono-totsellno+errcsno"
	sqlStr = sqlStr + " ,availsysstock=totipgono+totchulgono-totsellno+errcsno+errbaditemno+erretcno"
	sqlStr = sqlStr + " ,realstock=totipgono+totchulgono-totsellno+errcsno+errbaditemno+erretcno+errrealcheckno"
	sqlStr = sqlStr + " ,lastupdate=getdate()"
	sqlStr = sqlStr + " from [db_summary].[dbo].tbl_temp_baditem b "
	sqlStr = sqlStr + " where [db_summary].[dbo].tbl_daily_logisstock_summary.yyyymmdd='" + yyyymmdd + "'"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_daily_logisstock_summary.itemgubun=b.itemgubun"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_daily_logisstock_summary.itemid=b.itemid"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_daily_logisstock_summary.itemoption=b.itemoption"

	rsget.Open sqlStr,dbget,1

	''현재고테이블 업데이트
	sqlStr = "update [db_summary].[dbo].tbl_current_logisstock_summary"
	sqlStr = sqlStr + " set errbaditemno=errbaditemno+IsNULL(T.itemno*-1,0)"
	sqlStr = sqlStr + " ,toterrno=toterrno+IsNULL(T.itemno*-1,0)"  ''errcsno+
	sqlStr = sqlStr + " ,availsysstock=availsysstock+IsNULL(T.itemno*-1,0)"
	sqlStr = sqlStr + " ,realstock=realstock+IsNULL(T.itemno*-1,0)"
	sqlStr = sqlStr + " ,shortageno = shortageno+IsNULL(T.itemno*-1,0)"
	sqlStr = sqlStr + " ,lastupdate=getdate()"
	sqlStr = sqlStr + " from [db_summary].[dbo].tbl_temp_baditem T "
	sqlStr = sqlStr + " where [db_summary].[dbo].tbl_current_logisstock_summary.itemgubun=T.itemgubun"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_current_logisstock_summary.itemid=T.itemid"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_current_logisstock_summary.itemoption=T.itemoption"

	rsget.Open sqlStr,dbget,1

	'/----------------------------- 한정 처리 	'/2016.07.18 한용민 생성
	'/상품 한정 내리고
	sqlStr = " update i set " & vbCrLf
	sqlStr = sqlStr & " i.limitsold=(case when 0>i.limitsold + T.itemno then 0 else i.limitsold + T.itemno end) " & vbCrLf
	sqlStr = sqlStr & " from db_item.dbo.tbl_item i " & vbCrLf
	sqlStr = sqlStr & " join ( " & vbCrLf
	sqlStr = sqlStr & " 	select " & vbCrLf
	sqlStr = sqlStr & " 	itemid, sum(IsNull(itemno,0)) as itemno " & vbCrLf
	sqlStr = sqlStr & " 	from [db_summary].[dbo].tbl_temp_baditem bi " & vbCrLf
	sqlStr = sqlStr & " 	where bi.itemgubun='10' " & vbCrLf
	sqlStr = sqlStr & " 	group by itemid " & vbCrLf
	sqlStr = sqlStr & " ) T " & vbCrLf
	sqlStr = sqlStr & " 	on i.itemid = T.itemid " & vbCrLf

	'response.write sqlStr & "<Br>"
	dbget.Execute sqlStr

	'/옵션별 한정 내리고
	sqlStr = "update o set" & vbcrlf
	sqlStr = sqlStr & " o.optlimitsold=(case when 0>o.optlimitsold + bi.itemno then 0 else o.optlimitsold + bi.itemno end)" & vbcrlf
	sqlStr = sqlStr & " from [db_summary].[dbo].tbl_temp_baditem bi" & vbcrlf
	sqlStr = sqlStr & " join [db_item].[dbo].tbl_item_option o" & vbcrlf
	sqlStr = sqlStr & " 	on bi.itemid=o.itemid" & vbcrlf
	sqlStr = sqlStr & " 	and bi.itemoption = o.itemoption" & vbcrlf
	sqlStr = sqlStr & " where bi.itemgubun='10'" & vbcrlf
	sqlStr = sqlStr & " and bi.itemoption<>'0000'" & vbcrlf

	'response.write sqlStr & "<Br>"
	dbget.Execute sqlStr

	'/상품 품절처리 판매중인 상품중 한정품절 된 상품 일시 품절로 변경
	sqlStr = "update i" & vbcrlf
	sqlStr = sqlStr & " set i.sellyn='S' , i.lastupdate=getdate()" & vbcrlf
	sqlStr = sqlStr & " from (" & vbcrlf
	sqlStr = sqlStr & " 	select bi.itemid" & vbcrlf
	sqlStr = sqlStr & " 	from [db_summary].[dbo].tbl_temp_baditem bi" & vbcrlf
	sqlStr = sqlStr & " 	group by bi.itemid" & vbcrlf
	sqlStr = sqlStr & " ) as t" & vbcrlf
	sqlStr = sqlStr & " join [db_item].[dbo].tbl_item i" & vbcrlf
	sqlStr = sqlStr & " 	on t.itemid = i.itemid" & vbcrlf
	sqlStr = sqlStr & " 	and i.sellyn='Y'" & vbcrlf
	sqlStr = sqlStr & " 	and i.limityn='Y'" & vbcrlf
	sqlStr = sqlStr & " 	and (i.limitno-i.limitSold<1)" & vbcrlf

	'response.write sqlStr & "<Br>"
	dbget.Execute sqlStr

	'/일시 품절이나 한정수량>0 경우 판매로 변경
	sqlStr = "update i" & vbcrlf
	sqlStr = sqlStr & " set i.sellyn='Y' , i.lastupdate=getdate()" & vbcrlf
	sqlStr = sqlStr & " from (" & vbcrlf
	sqlStr = sqlStr & " 	select bi.itemid" & vbcrlf
	sqlStr = sqlStr & " 	from [db_summary].[dbo].tbl_temp_baditem bi" & vbcrlf
	sqlStr = sqlStr & " 	group by bi.itemid" & vbcrlf
	sqlStr = sqlStr & " ) as t" & vbcrlf
	sqlStr = sqlStr & " join [db_item].[dbo].tbl_item i" & vbcrlf
	sqlStr = sqlStr & " 	on t.itemid = i.itemid" & vbcrlf
	sqlStr = sqlStr & " 	and i.sellyn='S'" & vbcrlf
	sqlStr = sqlStr & " 	and i.limityn='Y'" & vbcrlf
	sqlStr = sqlStr & " 	and (i.limitno-i.limitSold>0)" & vbcrlf

	'response.write sqlStr & "<Br>"
	dbget.Execute sqlStr
	'/----------------------------- 한정 처리 ------------------------------------------------

	''임시 불량상품 삭제
	sqlStr = "delete from [db_summary].[dbo].tbl_temp_baditem"

	rsget.Open sqlStr,dbget,1

elseif mode="itemsellrefresholdall" then
	''특정상품 옵션포함전체.
	if (itemid="") then
		response.write "<script>alert('상품코드를 입력하세요.');</script>"
		response.write "<script>history.back();</script>"
		dbget.close()	:	response.End
	end if

	oTimer = Timer()

	sqlStr = " update [db_summary].[dbo].tbl_monthly_logisstock_summary"
	sqlStr = sqlStr + " set sellno=0"
	sqlStr = sqlStr + " ,resellno=0"
	sqlStr = sqlStr + " ,totsellno=0"
	sqlStr = sqlStr + " where yyyymm<'" + Left(refreshstartdate,7) + "'"
	sqlStr = sqlStr + " and itemgubun='" + CStr(itemgubun) + "'"
	sqlStr = sqlStr + " and itemid=" + CStr(itemid) + ""

	rsget.Open sqlStr,dbget,1

	''재고 일별 판매입력 - 기존내역 없는것들.
	sqlStr = " insert into [db_summary].[dbo].tbl_monthly_logisstock_summary"
	sqlStr = sqlStr + " (yyyymm,itemgubun,itemid,itemoption,sellno,resellno,totsellno)"
	sqlStr = sqlStr + " select T.yyyymm,'10',T.itemid,T.itemoption,T.sellno,T.resellno,T.totsellno"
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " ("
	sqlStr = sqlStr + " select convert(varchar(7),beadaldate,21) as yyyymm, d.itemid,d.itemoption,"
	sqlStr = sqlStr + " sum(case when d.itemno>0 then d.itemno else 0 end ) as sellno,"
	sqlStr = sqlStr + " sum(case when d.itemno<0 then d.itemno else 0 end ) as resellno,"
	sqlStr = sqlStr + " sum(d.itemno) as totsellno"
	sqlStr = sqlStr + " from [db_log].[dbo].tbl_old_order_master_2003 m,"
	sqlStr = sqlStr + "  [db_log].[dbo].tbl_old_order_detail_2003 d"
	sqlStr = sqlStr + " where m.orderserial=d.orderserial"
	sqlStr = sqlStr + " and m.ipkumdiv>=7"
	sqlStr = sqlStr + " and m.cancelyn='N'"
	sqlStr = sqlStr + " and d.cancelyn<>'Y'"
	sqlStr = sqlStr + " and d.itemid=" + CStr(itemid) + ""
	sqlStr = sqlStr + " and ((d.isupchebeasong is null ) or (d.isupchebeasong<>'Y'))"
	sqlStr = sqlStr + " group by convert(varchar(7),beadaldate,21), d.itemid, d.itemoption"
	sqlStr = sqlStr + " ) T"
	sqlStr = sqlStr + " left join [db_summary].[dbo].tbl_monthly_logisstock_summary s"
	sqlStr = sqlStr + " on T.yyyymm=s.yyyymm"
	sqlStr = sqlStr + " and T.itemid=s.itemid"
	sqlStr = sqlStr + " and T.itemoption=s.itemoption"
	sqlStr = sqlStr + " where s.itemid is null"

	rsget.Open sqlStr,dbget,1

	sqlStr = " insert into [db_summary].[dbo].tbl_monthly_logisstock_summary"
	sqlStr = sqlStr + " (yyyymm,itemgubun,itemid,itemoption,sellno,resellno,totsellno)"
	sqlStr = sqlStr + " select T.yyyymm,'10',T.itemid,T.itemoption,T.sellno,T.resellno,T.totsellno"
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " ("
	sqlStr = sqlStr + " select convert(varchar(7),beadaldate,21) as yyyymm, d.itemid,d.itemoption,"
	sqlStr = sqlStr + " sum(case when d.itemno>0 then d.itemno else 0 end ) as sellno,"
	sqlStr = sqlStr + " sum(case when d.itemno<0 then d.itemno else 0 end ) as resellno,"
	sqlStr = sqlStr + " sum(d.itemno) as totsellno"
	sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m,"
	sqlStr = sqlStr + "  [db_order].[dbo].tbl_order_detail d"
	sqlStr = sqlStr + " where m.orderserial=d.orderserial"
	sqlStr = sqlStr + " and m.ipkumdiv>=7"
	sqlStr = sqlStr + " and m.cancelyn='N'"
	sqlStr = sqlStr + " and d.cancelyn<>'Y'"
	sqlStr = sqlStr + " and d.itemid=" + CStr(itemid) + ""
	sqlStr = sqlStr + " and ((d.isupchebeasong is null ) or (d.isupchebeasong<>'Y'))"
	sqlStr = sqlStr + " group by convert(varchar(7),beadaldate,21), d.itemid, d.itemoption"
	sqlStr = sqlStr + " ) T"
	sqlStr = sqlStr + " left join [db_summary].[dbo].tbl_monthly_logisstock_summary s"
	sqlStr = sqlStr + " on T.yyyymm=s.yyyymm"
	sqlStr = sqlStr + " and T.itemid=s.itemid"
	sqlStr = sqlStr + " and T.itemoption=s.itemoption"
	sqlStr = sqlStr + " where s.itemid is null"

	rsget.Open sqlStr,dbget,1

	sqlStr = " update [db_summary].[dbo].tbl_monthly_logisstock_summary"
	sqlStr = sqlStr + " set sellno=T.sellno"
	sqlStr = sqlStr + " ,resellno=T.resellno"
	sqlStr = sqlStr + " ,totsellno=T.totsellno"
	sqlStr = sqlStr + " ,lastupdate=getdate()"
	sqlStr = sqlStr + " from ("
	sqlStr = sqlStr + " select convert(varchar(7),beadaldate,21) as yyyymm, d.itemid,d.itemoption,"
	sqlStr = sqlStr + " sum(case when d.itemno>0 then d.itemno else 0 end ) as sellno,"
	sqlStr = sqlStr + " sum(case when d.itemno<0 then d.itemno else 0 end ) as resellno,"
	sqlStr = sqlStr + " sum(d.itemno) as totsellno"
	sqlStr = sqlStr + " from [db_log].[dbo].tbl_old_order_master_2003 m,"
	sqlStr = sqlStr + "  [db_log].[dbo].tbl_old_order_detail_2003 d"
	sqlStr = sqlStr + " where  m.orderserial=d.orderserial"
	sqlStr = sqlStr + " and m.ipkumdiv>=7"
	sqlStr = sqlStr + " and m.cancelyn='N'"
	sqlStr = sqlStr + " and d.cancelyn<>'Y'"
	sqlStr = sqlStr + " and d.itemid=" + CStr(itemid) + ""
	sqlStr = sqlStr + " and ((d.isupchebeasong is null ) or (d.isupchebeasong<>'Y'))"
	sqlStr = sqlStr + " group by convert(varchar(7),beadaldate,21), d.itemid, d.itemoption"
	sqlStr = sqlStr + " ) T"
	sqlStr = sqlStr + " where [db_summary].[dbo].tbl_monthly_logisstock_summary.yyyymm=T.yyyymm"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_monthly_logisstock_summary.itemgubun='10'"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_monthly_logisstock_summary.itemid=T.itemid"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_monthly_logisstock_summary.itemoption=T.itemoption"

	rsget.Open sqlStr,dbget,1

	sqlStr = " update [db_summary].[dbo].tbl_monthly_logisstock_summary"
	sqlStr = sqlStr + " set sellno=[db_summary].[dbo].tbl_monthly_logisstock_summary.sellno + T.sellno"
	sqlStr = sqlStr + " ,resellno=[db_summary].[dbo].tbl_monthly_logisstock_summary.resellno + T.resellno"
	sqlStr = sqlStr + " ,totsellno=[db_summary].[dbo].tbl_monthly_logisstock_summary.totsellno + T.totsellno"
	sqlStr = sqlStr + " ,lastupdate=getdate()"
	sqlStr = sqlStr + " from ("
	sqlStr = sqlStr + " select convert(varchar(7),beadaldate,21) as yyyymm, d.itemid,d.itemoption,"
	sqlStr = sqlStr + " sum(case when d.itemno>0 then d.itemno else 0 end ) as sellno,"
	sqlStr = sqlStr + " sum(case when d.itemno<0 then d.itemno else 0 end ) as resellno,"
	sqlStr = sqlStr + " sum(d.itemno) as totsellno"
	sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m,"
	sqlStr = sqlStr + "  [db_order].[dbo].tbl_order_detail d"
	sqlStr = sqlStr + " where  m.orderserial=d.orderserial"
	sqlStr = sqlStr + " and m.ipkumdiv>=7"
	sqlStr = sqlStr + " and m.cancelyn='N'"
	sqlStr = sqlStr + " and d.cancelyn<>'Y'"
	sqlStr = sqlStr + " and d.itemid=" + CStr(itemid) + ""
	sqlStr = sqlStr + " and ((d.isupchebeasong is null ) or (d.isupchebeasong<>'Y'))"
	sqlStr = sqlStr + " group by convert(varchar(7),beadaldate,21), d.itemid, d.itemoption"
	sqlStr = sqlStr + " ) T"
	sqlStr = sqlStr + " where [db_summary].[dbo].tbl_monthly_logisstock_summary.yyyymm=T.yyyymm"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_monthly_logisstock_summary.itemgubun='10'"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_monthly_logisstock_summary.itemid=T.itemid"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_monthly_logisstock_summary.itemoption=T.itemoption"

	rsget.Open sqlStr,dbget,1

	sqlStr = " update [db_summary].[dbo].tbl_monthly_logisstock_summary"
	sqlStr = sqlStr + " set totsysstock=totipgono+totchulgono-totsellno"
	sqlStr = sqlStr + " ,availsysstock=totipgono+totchulgono-totsellno+errcsno+errbaditemno+erretcno"
	sqlStr = sqlStr + " ,realstock=totipgono+totchulgono-totsellno+errcsno+errbaditemno+erretcno+errrealcheckno"
	sqlStr = sqlStr + " ,lastupdate=getdate()"
	sqlStr = sqlStr + " where itemgubun='10'"
	sqlStr = sqlStr + " and itemid=" + CStr(itemid) + ""

	rsget.Open sqlStr,dbget,1

	sqlStr = " insert into [db_summary].[dbo].tbl_LAST_monthly_logisstock"
	sqlStr = sqlStr + " (lastyyyymm,itemgubun,itemid,itemoption)"
	sqlStr = sqlStr + " select distinct '" + LastYYYYMM + "', '10', T.itemid, T.itemoption"
	sqlStr = sqlStr + " from [db_summary].[dbo].tbl_monthly_logisstock_summary T"
	sqlStr = sqlStr + " left join [db_summary].[dbo].tbl_LAST_monthly_logisstock s"
	sqlStr = sqlStr + " on T.itemgubun='10'"
	sqlStr = sqlStr + " and T.itemid=" + CStr(itemid)
	sqlStr = sqlStr + " and T.itemid=s.itemid"
	sqlStr = sqlStr + " and T.itemoption=s.itemoption"
	sqlStr = sqlStr + " where T.itemgubun='10'"
	sqlStr = sqlStr + " and T.itemid=" + CStr(itemid)
	sqlStr = sqlStr + " and s.itemid is null"

	rsget.Open sqlStr,dbget,1

	sqlStr = " update [db_summary].[dbo].tbl_LAST_monthly_logisstock"
	sqlStr = sqlStr + " set lastyyyymm='" + LastYYYYMM + "'"
	sqlStr = sqlStr + " ,sellno=IsNULL(T.sellno,0)"
	sqlStr = sqlStr + " ,resellno=IsNULL(T.resellno,0)"
	sqlStr = sqlStr + " ,totsellno=IsNULL(T.totsellno,0)"
	sqlStr = sqlStr + " ,totsysstock=IsNULL(T.totsysstock,0)"
	sqlStr = sqlStr + " ,availsysstock=IsNULL(T.availsysstock,0)"
	sqlStr = sqlStr + " ,realstock=IsNULL(T.realstock,0)"
	sqlStr = sqlStr + " ,lastupdate=getdate()"
	sqlStr = sqlStr + " from ("
	sqlStr = sqlStr + " select itemid, itemoption, sum(sellno) as sellno,"
	sqlStr = sqlStr + " sum(resellno) as resellno, sum(totsellno) as totsellno,"
	sqlStr = sqlStr + " sum(totsysstock) as totsysstock,"
	sqlStr = sqlStr + " sum(availsysstock) as availsysstock, sum(realstock) as realstock"
	sqlStr = sqlStr + " from [db_summary].[dbo].tbl_monthly_logisstock_summary"
	sqlStr = sqlStr + " where yyyymm<'" + Left(refreshstartdate,7) + "'"
	sqlStr = sqlStr + " and itemgubun='10'"
	sqlStr = sqlStr + " and itemid=" + CStr(itemid) + ""
	sqlStr = sqlStr + " group by itemid,itemoption"
	sqlStr = sqlStr + " ) T"

	sqlStr = sqlStr + " where [db_summary].[dbo].tbl_LAST_monthly_logisstock.itemgubun='10'"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_LAST_monthly_logisstock.itemid=" + CStr(itemid) + ""
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_LAST_monthly_logisstock.itemid=T.itemid"

	rsget.Open sqlStr,dbget,1

	sqlStr = "update [db_summary].[dbo].tbl_current_logisstock_summary"
	sqlStr = sqlStr + " set ipgono=T.ipgono"
	sqlStr = sqlStr + " ,reipgono=T.reipgono"
	sqlStr = sqlStr + " ,totipgono=T.totipgono"
	sqlStr = sqlStr + " ,offchulgono=T.offchulgono"
	sqlStr = sqlStr + " ,offrechulgono=T.offrechulgono"
	sqlStr = sqlStr + " ,etcchulgono=T.etcchulgono"
	sqlStr = sqlStr + " ,etcrechulgono=T.etcrechulgono"
	sqlStr = sqlStr + " ,totchulgono=T.totchulgono"
	sqlStr = sqlStr + " ,sellno=T.sellno"
	sqlStr = sqlStr + " ,resellno=T.resellno"
	sqlStr = sqlStr + " ,totsellno=T.totsellno"
	sqlStr = sqlStr + " ,errcsno=T.errcsno"
	sqlStr = sqlStr + " ,errbaditemno=T.errbaditemno"
	sqlStr = sqlStr + " ,errrealcheckno=T.errrealcheckno"
	sqlStr = sqlStr + " ,erretcno=T.erretcno"
	sqlStr = sqlStr + " ,toterrno=T.toterrno"
	sqlStr = sqlStr + " ,offsellno=T.offsellno"
	sqlStr = sqlStr + " ,totsysstock=T.totsysstock"
	sqlStr = sqlStr + " ,availsysstock=T.availsysstock"
	sqlStr = sqlStr + " ,realstock=T.realstock"
	sqlStr = sqlStr + " from  [db_summary].[dbo].tbl_LAST_monthly_logisstock T"
	sqlStr = sqlStr + " where T.itemgubun='10'"
	sqlStr = sqlStr + " and T.itemid=" + CStr(itemid) + ""
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_current_logisstock_summary.itemgubun='10'"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_current_logisstock_summary.itemid=T.itemid"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_current_logisstock_summary.itemoption=T.itemoption"

	rsget.Open sqlStr,dbget,1

	sqlStr = "update [db_summary].[dbo].tbl_current_logisstock_summary"
	sqlStr = sqlStr + " set ipgono=[db_summary].[dbo].tbl_current_logisstock_summary.ipgono + IsNULL(T.ipgono,0)"
	sqlStr = sqlStr + " ,reipgono=[db_summary].[dbo].tbl_current_logisstock_summary.reipgono + IsNULL(T.reipgono,0)"
	sqlStr = sqlStr + " ,totipgono=[db_summary].[dbo].tbl_current_logisstock_summary.totipgono + IsNULL(T.totipgono,0)"
	sqlStr = sqlStr + " ,offchulgono=[db_summary].[dbo].tbl_current_logisstock_summary.offchulgono + IsNULL(T.offchulgono,0)"
	sqlStr = sqlStr + " ,offrechulgono=[db_summary].[dbo].tbl_current_logisstock_summary.offrechulgono + IsNULL(T.offrechulgono,0)"
	sqlStr = sqlStr + " ,etcchulgono=[db_summary].[dbo].tbl_current_logisstock_summary.etcchulgono + IsNULL(T.etcchulgono,0)"
	sqlStr = sqlStr + " ,etcrechulgono=[db_summary].[dbo].tbl_current_logisstock_summary.etcrechulgono + IsNULL(T.etcrechulgono,0)"
	sqlStr = sqlStr + " ,totchulgono=[db_summary].[dbo].tbl_current_logisstock_summary.totchulgono + IsNULL(T.totchulgono,0)"
	sqlStr = sqlStr + " ,sellno=[db_summary].[dbo].tbl_current_logisstock_summary.sellno + IsNULL(T.sellno,0)"
	sqlStr = sqlStr + " ,resellno=[db_summary].[dbo].tbl_current_logisstock_summary.resellno + IsNULL(T.resellno,0)"
	sqlStr = sqlStr + " ,totsellno=[db_summary].[dbo].tbl_current_logisstock_summary.totsellno + IsNULL(T.totsellno,0)"
	sqlStr = sqlStr + " ,errcsno=[db_summary].[dbo].tbl_current_logisstock_summary.errcsno + IsNULL(T.errcsno,0)"
	sqlStr = sqlStr + " ,errbaditemno=[db_summary].[dbo].tbl_current_logisstock_summary.errbaditemno + IsNULL(T.errbaditemno,0)"
	sqlStr = sqlStr + " ,errrealcheckno=[db_summary].[dbo].tbl_current_logisstock_summary.errrealcheckno + IsNULL(T.errrealcheckno,0)"
	sqlStr = sqlStr + " ,erretcno=[db_summary].[dbo].tbl_current_logisstock_summary.erretcno + IsNULL(T.erretcno,0)"
	sqlStr = sqlStr + " ,toterrno=[db_summary].[dbo].tbl_current_logisstock_summary.toterrno + IsNULL(T.toterrno,0)"
	sqlStr = sqlStr + " ,offsellno=[db_summary].[dbo].tbl_current_logisstock_summary.offsellno + IsNULL(T.offsellno,0)"
	sqlStr = sqlStr + " ,totsysstock=[db_summary].[dbo].tbl_current_logisstock_summary.totsysstock + IsNULL(T.totsysstock,0)"
	sqlStr = sqlStr + " ,availsysstock=[db_summary].[dbo].tbl_current_logisstock_summary.availsysstock + IsNULL(T.availsysstock,0)"
	sqlStr = sqlStr + " ,realstock=[db_summary].[dbo].tbl_current_logisstock_summary.realstock + IsNULL(T.realstock,0)"
	sqlStr = sqlStr + " from  ("
	sqlStr = sqlStr + " select itemid,itemoption,"
	sqlStr = sqlStr + " sum(ipgono) as ipgono,sum(reipgono) as reipgono,sum(totipgono) as totipgono,sum(offchulgono) as offchulgono,sum(offrechulgono) as offrechulgono,"
	sqlStr = sqlStr + " sum(etcchulgono) as etcchulgono,sum(etcrechulgono) as etcrechulgono,sum(totchulgono) as totchulgono,"
	sqlStr = sqlStr + " sum(sellno) as sellno,sum(resellno) as resellno,sum(totsellno) as totsellno,sum(errcsno) as errcsno,"
	sqlStr = sqlStr + " sum(errbaditemno) as errbaditemno,sum(errrealcheckno) as errrealcheckno,sum(erretcno) as erretcno,"
	sqlStr = sqlStr + " sum(toterrno) as toterrno,sum(offsellno) as offsellno, sum(totsysstock) as totsysstock,sum(availsysstock) as availsysstock,sum(realstock) as realstock"
	sqlStr = sqlStr + "  from [db_summary].[dbo].tbl_daily_logisstock_summary"
	sqlStr = sqlStr + " where yyyymmdd>='" + refreshstartdate + "'"
	sqlStr = sqlStr + " and itemgubun='10'"
	sqlStr = sqlStr + " and itemid=" + CStr(itemid) + ""
	sqlStr = sqlStr + " group by itemid,itemoption"
	sqlStr = sqlStr + " ) T"
	sqlStr = sqlStr + " where [db_summary].[dbo].tbl_current_logisstock_summary.itemgubun='10'"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_current_logisstock_summary.itemid=T.itemid"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_current_logisstock_summary.itemoption=T.itemoption"

	rsget.Open sqlStr,dbget,1

	response.write "<small>재고 월별 판매입력... finish(" + FormatNumber(Timer()-oTimer,2) + " sec)</small><br>"
	response.flush
end if
%>

<script type="text/javascript">

<% if mode="tmpbaditem2input" then %>
	alert('->. [한정 처리] 불량상품 한정처리 완료\n\n저장 되었습니다.');
	opener.location.reload();
	window.close();
<% elseif mode="dellog" then %>
	alert('저장 되었습니다.');
	opener.location.reload();
	window.close();
<% else %>
	alert('저장 되었습니다.');
	location.replace('<%= refer %>');
<% end if %>

</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->