<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
dim refer
refer = request.ServerVariables("HTTP_REFERER")

dim mode,refreshstartdate,itemgubun,itemid,itemoption
dim realstock

mode	= request.form("mode")
refreshstartdate = request.form("refreshstartdate")
itemgubun = request.form("itemgubun")
itemid = request.form("itemid")
itemoption = request.form("itemoption")
realstock = request.form("realstock")

dim sqlStr

dim orgrealstock, todayerrno
dim errstock
dim itemExists
dim yyyymmdd

response.write "Starting Job..<br>"
response.flush

if mode="itemrecentipchulrefresh" then
	''기존 판매 일별/월별 데이타 삭제
	sqlStr = "delete from [db_summary].[dbo].tbl_deliver_itemsell_daily_summary"
	sqlStr = sqlStr + " where yyyymmdd>='" + refreshstartdate + "'"
	sqlStr = sqlStr + " and itemid=" + CStr(itemid) + ""
	sqlStr = sqlStr + " and itemoption='" + itemoption + "'"

	rsget.Open sqlStr,dbget,1

	sqlStr = "delete from [db_summary].[dbo].tbl_deliver_itemsell_monthly_summary"
	sqlStr = sqlStr + " where yyyymm>='" + Left(refreshstartdate,7) + "'"
	sqlStr = sqlStr + " and itemid=" + CStr(itemid) + ""
	sqlStr = sqlStr + " and itemoption='" + itemoption + "'"

	rsget.Open sqlStr,dbget,1

response.write "기존 판매 일별/월별 데이타 삭제... finish<br>"
response.flush

	''기존 입출고 일별/월별 데이타 삭제
	sqlStr = "delete from [db_summary].[dbo].tbl_ipchul_daily_summary"
	sqlStr = sqlStr + " where yyyymmdd>='" + refreshstartdate + "'"
	sqlStr = sqlStr + " and itemgubun='" + itemgubun + "'"
	sqlStr = sqlStr + " and itemid=" + CStr(itemid) + ""
	sqlStr = sqlStr + " and itemoption='" + itemoption + "'"

	rsget.Open sqlStr,dbget,1

	sqlStr = "delete from [db_summary].[dbo].tbl_ipchul_monthly_summary"
	sqlStr = sqlStr + " where yyyymm>='" + Left(refreshstartdate,7) + "'"
	sqlStr = sqlStr + " and itemgubun='" + itemgubun + "'"
	sqlStr = sqlStr + " and itemid=" + CStr(itemid) + ""
	sqlStr = sqlStr + " and itemoption='" + itemoption + "'"

	rsget.Open sqlStr,dbget,1

response.write "기존 입출고 일별/월별 데이타 삭제... finish<br>"
response.flush

	''일별판매로그 입력
	sqlStr = " insert into [db_summary].[dbo].tbl_deliver_itemsell_daily_summary"
	sqlStr = sqlStr + " (yyyymmdd,itemid,itemoption,sellno,reno,totsellno)"
	sqlStr = sqlStr + " select "
	sqlStr = sqlStr + " convert(varchar(10),beadaldate,21) as yyyymmdd, d.itemid,d.itemoption,"
	sqlStr = sqlStr + " sum(case when d.itemno>0 then d.itemno else 0 end ) as sellno,"
	sqlStr = sqlStr + " sum(case when d.itemno<0 then d.itemno else 0 end ) as reno,"
	sqlStr = sqlStr + " sum(d.itemno) as totsellno"
	sqlStr = sqlStr + " from [db_order].[10x10].tbl_order_master m,"
	sqlStr = sqlStr + "  [db_order].[10x10].tbl_order_detail d"
	sqlStr = sqlStr + " where m.beadaldate>='" + refreshstartdate + "'"
	sqlStr = sqlStr + " and m.orderserial=d.orderserial"
	sqlStr = sqlStr + " and m.cancelyn='N'"
	sqlStr = sqlStr + " and d.cancelyn<>'Y'"
	sqlStr = sqlStr + " and m.ipkumdiv>4"
	sqlStr = sqlStr + " and d.itemid=" + CStr(itemid) + ""
	sqlStr = sqlStr + " and d.itemoption='" + itemoption + "'"
	sqlStr = sqlStr + " and d.isupchebeasong<>'Y'"
	sqlStr = sqlStr + " group by convert(varchar(10),beadaldate,21) , d.itemid, d.itemoption"

	rsget.Open sqlStr,dbget,1

response.write "일별판매로그 입력... finish<br>"
response.flush

	''월별판매로그 입력
	sqlStr = " insert into [db_summary].[dbo].tbl_deliver_itemsell_monthly_summary "
	sqlStr = sqlStr + " (yyyymm,itemid,itemoption,sellno,reno,totsellno)"
	sqlStr = sqlStr + " select Left(yyyymmdd,7) as yyyymm, itemid,itemoption,"
	sqlStr = sqlStr + " sum(sellno) as sellno, sum(reno) as reno, sum(totsellno) as totsellno"
	sqlStr = sqlStr + " from [db_summary].[dbo].tbl_deliver_itemsell_daily_summary"
	sqlStr = sqlStr + " where yyyymmdd>='" + refreshstartdate + "'"
	sqlStr = sqlStr + " and itemid=" + CStr(itemid) + ""
	sqlStr = sqlStr + " and itemoption='" + itemoption + "'"
	sqlStr = sqlStr + " group by Left(yyyymmdd,7), itemid,itemoption"

	rsget.Open sqlStr,dbget,1


response.write "월별판매로그 입력... finish<br>"
response.flush

	sqlStr = " update [db_storage].[10x10].tbl_acount_storage_master"
	sqlStr = sqlStr + " set ipchulflag='I'"
	sqlStr = sqlStr + " where Left(code,2)='ST'"
	sqlStr = sqlStr + " and ipchulflag is null"

	rsget.Open sqlStr,dbget,1


	sqlStr = " update [db_storage].[10x10].tbl_acount_storage_master"
	sqlStr = sqlStr + " set ipchulflag='E'"
	sqlStr = sqlStr + " where Left(code,2)='SO'"
	sqlStr = sqlStr + " and Left(socid,10)<>'streetshop'"
	sqlStr = sqlStr + " and ipchulflag is null"

	rsget.Open sqlStr,dbget,1


	sqlStr = " update [db_storage].[10x10].tbl_acount_storage_master"
	sqlStr = sqlStr + " set ipchulflag='S'"
	sqlStr = sqlStr + " where Left(code,2)='SO'"
	sqlStr = sqlStr + " and Left(socid,10)='streetshop'"
	sqlStr = sqlStr + " and ipchulflag is null"

	rsget.Open sqlStr,dbget,1


	''입출고 일별데이타 입력
	sqlStr = "insert into [db_summary].[dbo].tbl_ipchul_daily_summary"
	sqlStr = sqlStr + " (yyyymmdd,itemgubun,itemid,itemoption,ipno,reipno,totipno,"
	sqlStr = sqlStr + " shopchulno,reshopchulno,totshopchulno,etcchulno,reetcchulno,totetcchulno)"
	sqlStr = sqlStr + " select T.yyyymmdd,T.iitemgubun,T.itemid,T.itemoption"
	sqlStr = sqlStr + " ,T.ipno,T.reipno,(T.ipno+T.reipno)"
	sqlStr = sqlStr + " ,T.shopchulno,T.reshopchulno,(T.shopchulno+T.reshopchulno)"
	sqlStr = sqlStr + " ,T.etcchulno,T.reetcchulno,(T.etcchulno+T.reetcchulno)"
	sqlStr = sqlStr + " from ("
	sqlStr = sqlStr + " select convert(varchar(10),m.executedt,21) as yyyymmdd"
	sqlStr = sqlStr + " ,iitemgubun,itemid,itemoption, "
	sqlStr = sqlStr + " sum(case when  ipchulflag='I' and itemno>0 then itemno else 0 end) as ipno,"
	sqlStr = sqlStr + " sum(case when  ipchulflag='I' and itemno<0 then itemno else 0 end) as reipno,"
	sqlStr = sqlStr + " sum(case when  ipchulflag='S' and itemno<0 then itemno else 0 end) as shopchulno,"
	sqlStr = sqlStr + " sum(case when  ipchulflag='S' and itemno>0 then itemno else 0 end) as reshopchulno,"
	sqlStr = sqlStr + " sum(case when  ipchulflag='E' and itemno<0 then itemno else 0 end) as etcchulno,"
	sqlStr = sqlStr + " sum(case when  ipchulflag='E' and itemno>0 then itemno else 0 end) as reetcchulno"
	sqlStr = sqlStr + "   from [db_storage].[10x10].tbl_acount_storage_master m, "
	sqlStr = sqlStr + " [db_storage].[10x10].tbl_acount_storage_detail d"
	sqlStr = sqlStr + " where m.executedt>='" + refreshstartdate + "'"
	sqlStr = sqlStr + " and m.code=d.mastercode"
	sqlStr = sqlStr + " and m.deldt is null"
	sqlStr = sqlStr + " and d.iitemgubun='" + itemgubun + "'"
	sqlStr = sqlStr + " and d.itemid=" + CStr(itemid) + ""
	sqlStr = sqlStr + " and d.itemoption='" + itemoption + "'"
	sqlStr = sqlStr + " and d.deldt is null"
	sqlStr = sqlStr + " and d.itemno<>0"
	sqlStr = sqlStr + " group by convert(varchar(10),m.executedt,21),iitemgubun,itemid,itemoption"
	sqlStr = sqlStr + " ) T"

	rsget.Open sqlStr,dbget,1

response.write "입출고 일별데이타 입력... finish<br>"
response.flush

	''입출고 월별데이타 입력

	sqlStr = " insert into [db_summary].[dbo].tbl_ipchul_monthly_summary"
	sqlStr = sqlStr + " (yyyymm,itemgubun,itemid,itemoption,ipno,reipno,totipno,"
	sqlStr = sqlStr + " shopchulno,reshopchulno,totshopchulno,etcchulno,reetcchulno,totetcchulno)"
	sqlStr = sqlStr + " select convert(varchar(7),yyyymmdd,21) as yyyymm,itemgubun,itemid,itemoption,"
	sqlStr = sqlStr + " sum(ipno),sum(reipno),sum(totipno),"
	sqlStr = sqlStr + " sum(shopchulno),sum(reshopchulno),sum(totshopchulno),"
	sqlStr = sqlStr + " sum(etcchulno),sum(reetcchulno),sum(totetcchulno)"
	sqlStr = sqlStr + "  from [db_summary].[dbo].tbl_ipchul_daily_summary"
	sqlStr = sqlStr + " where yyyymmdd>='" + refreshstartdate + "'"
	sqlStr = sqlStr + " and itemgubun='" + itemgubun + "'"
	sqlStr = sqlStr + " and itemid=" + CStr(itemid) + ""
	sqlStr = sqlStr + " and itemoption='" + itemoption + "'"
	sqlStr = sqlStr + " group by convert(varchar(7),yyyymmdd,21),itemgubun,itemid,itemoption "

	rsget.Open sqlStr,dbget,1

response.write "입출고 월별데이타 입력... finish<br>"
response.flush

	''재고 일별/월별테이블 삭제
	sqlStr = " delete from [db_summary].[dbo].tbl_daily_logisstock_summary"
	sqlStr = sqlStr + " where yyyymmdd>='" + refreshstartdate + "'"
	sqlStr = sqlStr + " and itemgubun='" + itemgubun + "'"
	sqlStr = sqlStr + " and itemid=" + CStr(itemid) + ""
	sqlStr = sqlStr + " and itemoption='" + itemoption + "'"

	rsget.Open sqlStr,dbget,1


	sqlStr = " delete from [db_summary].[dbo].tbl_monthly_logisstock_summary"
	sqlStr = sqlStr + " where yyyymm>='" + Left(refreshstartdate,7) + "'"
	sqlStr = sqlStr + " and itemgubun='" + itemgubun + "'"
	sqlStr = sqlStr + " and itemid=" + CStr(itemid) + ""
	sqlStr = sqlStr + " and itemoption='" + itemoption + "'"

	rsget.Open sqlStr,dbget,1

response.write "재고 일별/월별테이블 삭제... finish<br>"
response.flush


	''재고 일별 판매입력
	sqlStr = " insert into [db_summary].[dbo].tbl_daily_logisstock_summary"
	sqlStr = sqlStr + " (yyyymmdd,itemgubun,itemid,itemoption,sellno,resellno,totsellno)"
	sqlStr = sqlStr + " select yyyymmdd,'10',itemid,itemoption,sellno,reno,totsellno"
	sqlStr = sqlStr + " from [db_summary].[dbo].tbl_deliver_itemsell_daily_summary "
	sqlStr = sqlStr + " where yyyymmdd>='" + refreshstartdate + "'"
	sqlStr = sqlStr + " and itemid=" + CStr(itemid) + ""
	sqlStr = sqlStr + " and itemoption='" + itemoption + "'"

	rsget.Open sqlStr,dbget,1

response.write "재고 일별 판매입력... finish<br>"
response.flush

	''재고 일별테이블 업데이트(입출고)

	sqlStr = " insert into [db_summary].[dbo].tbl_daily_logisstock_summary"
	sqlStr = sqlStr + " (yyyymmdd,itemgubun,itemid,itemoption,ipgono,reipgono,totipgono,"
	sqlStr = sqlStr + " offchulgono,offrechulgono,etcchulgono,etcrechulgono,totchulgono)"
	sqlStr = sqlStr + " select s.yyyymmdd,s.itemgubun,s.itemid,s.itemoption,s.ipno,s.reipno,s.totipno,"
	sqlStr = sqlStr + " s.shopchulno,s.reshopchulno,s.etcchulno,s.reetcchulno,(s.shopchulno+s.reshopchulno+s.etcchulno+s.reetcchulno)"
	sqlStr = sqlStr + " from [db_summary].[dbo].tbl_ipchul_daily_summary s"
	sqlStr = sqlStr + " left join [db_summary].[dbo].tbl_daily_logisstock_summary l"
	sqlStr = sqlStr + " on s.yyyymmdd=l.yyyymmdd"
	sqlStr = sqlStr + " and s.itemgubun=l.itemgubun"
	sqlStr = sqlStr + " and s.itemid=l.itemid"
	sqlStr = sqlStr + " and s.itemoption=l.itemoption"
	sqlStr = sqlStr + " where s.yyyymmdd>='" + refreshstartdate + "'"
	sqlStr = sqlStr + " and s.itemgubun='" + itemgubun + "'"
	sqlStr = sqlStr + " and s.itemid=" + CStr(itemid) + ""
	sqlStr = sqlStr + " and s.itemoption='" + itemoption + "'"
	sqlStr = sqlStr + " and l.yyyymmdd is null"

	rsget.Open sqlStr,dbget,1


	sqlStr = " update [db_summary].[dbo].tbl_daily_logisstock_summary"
	sqlStr = sqlStr + " set ipgono=s.ipno"
	sqlStr = sqlStr + " ,reipgono=s.reipno"
	sqlStr = sqlStr + " ,totipgono=s.totipno"
	sqlStr = sqlStr + " ,offchulgono=s.shopchulno"
	sqlStr = sqlStr + " ,offrechulgono=s.reshopchulno"
	sqlStr = sqlStr + " ,etcchulgono=s.etcchulno"
	sqlStr = sqlStr + " ,etcrechulgono=s.reetcchulno"
	sqlStr = sqlStr + " ,totchulgono=s.shopchulno+s.reshopchulno+s.etcchulno+s.reetcchulno"
	sqlStr = sqlStr + " from  [db_summary].[dbo].tbl_ipchul_daily_summary s"
	sqlStr = sqlStr + " where s.yyyymmdd>='" + refreshstartdate + "'"
	sqlStr = sqlStr + " and s.itemgubun='" + itemgubun + "'"
	sqlStr = sqlStr + " and s.itemid=" + CStr(itemid) + ""
	sqlStr = sqlStr + " and s.itemoption='" + itemoption + "'"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_daily_logisstock_summary.yyyymmdd=s.yyyymmdd"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_daily_logisstock_summary.itemgubun=s.itemgubun"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_daily_logisstock_summary.itemid=s.itemid"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_daily_logisstock_summary.itemoption=s.itemoption"

	rsget.Open sqlStr,dbget,1

response.write "재고 일별테이블 업데이트(입출고)... finish<br>"
response.flush

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
	sqlStr = sqlStr + " where s.yyyymmdd>='" + refreshstartdate + "'"
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
	sqlStr = sqlStr + " from  [db_summary].[dbo].tbl_erritem_daily_summary s"
	sqlStr = sqlStr + " where s.yyyymmdd>='" + refreshstartdate + "'"
	sqlStr = sqlStr + " and s.itemgubun='" + itemgubun + "'"
	sqlStr = sqlStr + " and s.itemid=" + CStr(itemid) + ""
	sqlStr = sqlStr + " and s.itemoption='" + itemoption + "'"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_daily_logisstock_summary.yyyymmdd=s.yyyymmdd"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_daily_logisstock_summary.itemgubun=s.itemgubun"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_daily_logisstock_summary.itemid=s.itemid"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_daily_logisstock_summary.itemoption=s.itemoption"

	rsget.Open sqlStr,dbget,1

response.write "재고 일별테이블 업데이트(오차)... finish<br>"
response.flush

	sqlStr = " update [db_summary].[dbo].tbl_daily_logisstock_summary"
	sqlStr = sqlStr + " set offsellno=IsNULL(T.offsellno,0)"
	sqlStr = sqlStr + " from ("
	sqlStr = sqlStr + " select convert(varchar(10),m.shopregdate,21) as yyyymmdd,sum(itemno) as offsellno"
	sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shopjumun_master m,"
	sqlStr = sqlStr + "  [db_shop].[dbo].tbl_shopjumun_detail d"
	sqlStr = sqlStr + " where m.idx=d.masteridx"
	sqlStr = sqlStr + " and m.shopregdate>='" + refreshstartdate + "'"
	sqlStr = sqlStr + " and m.cancelyn='N'"
	sqlStr = sqlStr + " and d.cancelyn='N'"
	sqlStr = sqlStr + " and d.itemgubun='" + itemgubun + "'"
	sqlStr = sqlStr + " and d.itemid=" + CStr(itemid)
	sqlStr = sqlStr + " and d.itemoption='" + CStr(itemoption) + "'"
	sqlStr = sqlStr + " group by convert(varchar(10),m.shopregdate,21)"
	sqlStr = sqlStr + " ) T"
	sqlStr = sqlStr + " where [db_summary].[dbo].tbl_daily_logisstock_summary.yyyymmdd=T.yyyymmdd"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_daily_logisstock_summary.itemgubun='" + itemgubun + "'"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_daily_logisstock_summary.itemid=" + CStr(itemid)
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_daily_logisstock_summary.itemoption='" + itemoption + "'"

	rsget.Open sqlStr,dbget,1


	sqlStr = " insert into [db_summary].[dbo].tbl_daily_logisstock_summary"
	sqlStr = sqlStr + " (yyyymmdd, itemgubun, itemid, itemoption,offsellno)"
	sqlStr = sqlStr + " select T.yyyymmdd, '" + itemgubun + "'," + itemid + ",'" + itemoption + "', IsNULL(T.offsellno,0)"
	sqlStr = sqlStr + " from ("
	sqlStr = sqlStr + " select convert(varchar(10),m.shopregdate,21) as yyyymmdd, sum(itemno) as offsellno"
	sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shopjumun_master m,"
	sqlStr = sqlStr + "  [db_shop].[dbo].tbl_shopjumun_detail d"
	sqlStr = sqlStr + " where m.idx=d.masteridx"
	sqlStr = sqlStr + " and m.shopregdate>='" + refreshstartdate + "'"
	sqlStr = sqlStr + " and m.cancelyn='N'"
	sqlStr = sqlStr + " and d.cancelyn='N'"
	sqlStr = sqlStr + " and d.itemgubun='" + itemgubun + "'"
	sqlStr = sqlStr + " and d.itemid=" + CStr(itemid)
	sqlStr = sqlStr + " and d.itemoption='" + CStr(itemoption) + "'"
	sqlStr = sqlStr + " group by convert(varchar(10),m.shopregdate,21)"
	sqlStr = sqlStr + " ) T"
	sqlStr = sqlStr + " left join [db_summary].[dbo].tbl_daily_logisstock_summary s "
	sqlStr = sqlStr + " on T.yyyymmdd=s.yyyymmdd"
	sqlStr = sqlStr + " and s.itemgubun='" + itemgubun + "'"
	sqlStr = sqlStr + " and s.itemid=" + CStr(itemid)
	sqlStr = sqlStr + " and s.itemoption='" + CStr(itemoption) + "'"
	sqlStr = sqlStr + " where s.itemgubun is null"

	rsget.Open sqlStr,dbget,1

response.write "오프샾 판매 업데이트... finish<br>"
response.flush

	''서머리.
	sqlStr = " update [db_summary].[dbo].tbl_daily_logisstock_summary"
	sqlStr = sqlStr + " set totsysstock=totipgono+totchulgono-totsellno"
	sqlStr = sqlStr + " ,availsysstock=totipgono+totchulgono-totsellno+errcsno+errbaditemno+erretcno"
	sqlStr = sqlStr + " ,realstock=totipgono+totchulgono-totsellno+errcsno+errbaditemno+erretcno+errrealcheckno"
	sqlStr = sqlStr + " ,lastupdate=getdate()"
	sqlStr = sqlStr + " where yyyymmdd>='" + refreshstartdate + "'"
	sqlStr = sqlStr + " and itemgubun='" + itemgubun + "'"
	sqlStr = sqlStr + " and itemid=" + CStr(itemid) + ""
	sqlStr = sqlStr + " and itemoption='" + itemoption + "'"

	rsget.Open sqlStr,dbget,1

response.write "재고 일별테이블 업데이트(서머리)... finish<br>"
response.flush

	''재고 월별테이블 업데이트
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

response.write "재고 월별테이블 업데이트... finish<br>"
response.flush


	''현재재고업데이트

	sqlStr = " delete from [db_summary].[dbo].tbl_current_logisstock_summary"
	sqlStr = sqlStr + " where itemgubun='" + itemgubun + "'"
	sqlStr = sqlStr + " and itemid=" + CStr(itemid) + ""
	sqlStr = sqlStr + " and itemoption='" + itemoption + "'"

	rsget.Open sqlStr,dbget,1



	'sqlStr = " insert into [db_summary].[dbo].tbl_current_logisstock_summary"
	'sqlStr = sqlStr + " (itemgubun,itemid,itemoption,ipgono,reipgono,totipgono,offchulgono,offrechulgono,"
	'sqlStr = sqlStr + " etcchulgono,etcrechulgono,totchulgono,sellno,resellno,"
	'sqlStr = sqlStr + " totsellno,errcsno,errbaditemno,errrealcheckno,erretcno,"
	'sqlStr = sqlStr + " toterrno,totsysstock,availsysstock,realstock)"
	'sqlStr = sqlStr + " select"
	'sqlStr = sqlStr + " itemgubun,itemid,itemoption,"
	'sqlStr = sqlStr + " sum(ipgono) as ipgono,sum(reipgono) as reipgono,sum(totipgono) as totipgono,sum(offchulgono) as offchulgono,sum(offrechulgono) as offrechulgono,"
	'sqlStr = sqlStr + " sum(etcchulgono) as etcchulgono,sum(etcrechulgono) as etcrechulgono,sum(totchulgono) as totchulgono,"
	'sqlStr = sqlStr + " sum(sellno) as sellno,sum(resellno) as resellno,sum(totsellno) as totsellno,sum(errcsno) as errcsno,"
	'sqlStr = sqlStr + " sum(errbaditemno) as errbaditemno,sum(errrealcheckno) as errrealcheckno,sum(erretcno) as erretcno,"
	'sqlStr = sqlStr + " sum(toterrno) as toterrno,sum(totsysstock) as totsysstock,sum(availsysstock) as availsysstock,sum(realstock) as realstock"
	'sqlStr = sqlStr + "  from [db_summary].[dbo].tbl_daily_logisstock_summary"
	'sqlStr = sqlStr + " where itemgubun='" + itemgubun + "'"
	'sqlStr = sqlStr + " and itemid=" + CStr(itemid) + ""
	'sqlStr = sqlStr + " and itemoption='" + itemoption + "'"
	'sqlStr = sqlStr + " group by itemgubun,itemid,itemoption"

	'rsget.Open sqlStr,dbget,1


	if itemgubun="10" then
		sqlStr = " insert into [db_summary].[dbo].tbl_current_logisstock_summary "
		sqlStr = sqlStr + " (itemgubun,itemid,itemoption,imgsmall)"
		sqlStr = sqlStr + " select top 1 '10'," + CStr(itemid) + ",'" + itemoption + "',T.smallimage"
		sqlStr = sqlStr + " from [db_item].[10x10].tbl_item T"
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
	'sqlStr = "update [db_summary].[dbo].tbl_current_logisstock_summary"
	'sqlStr = sqlStr + " set ipgono=IsNULL(T.ipgono,0)"
	'sqlStr = sqlStr + " ,reipgono=IsNULL(T.reipgono,0)"
	'sqlStr = sqlStr + " ,totipgono=IsNULL(T.totipgono,0)"
	'sqlStr = sqlStr + " ,offchulgono=IsNULL(T.offchulgono,0)"
	'sqlStr = sqlStr + " ,offrechulgono=IsNULL(T.offrechulgono,0)"
	'sqlStr = sqlStr + " ,etcchulgono=IsNULL(T.etcchulgono,0)"
	'sqlStr = sqlStr + " ,etcrechulgono=IsNULL(T.etcrechulgono,0)"
	'sqlStr = sqlStr + " ,totchulgono=IsNULL(T.totchulgono,0)"
	'sqlStr = sqlStr + " ,sellno=IsNULL(T.sellno,0)"
	'sqlStr = sqlStr + " ,resellno=IsNULL(T.resellno,0)"
	'sqlStr = sqlStr + " ,totsellno=IsNULL(T.totsellno,0)"
	'sqlStr = sqlStr + " ,errcsno=IsNULL(T.errcsno,0)"
	'sqlStr = sqlStr + " ,errbaditemno=IsNULL(T.errbaditemno,0)"
	'sqlStr = sqlStr + " ,errrealcheckno=IsNULL(T.errrealcheckno,0)"
	'sqlStr = sqlStr + " ,erretcno=IsNULL(T.erretcno,0)"
	'sqlStr = sqlStr + " ,toterrno=IsNULL(T.toterrno,0)"
	'sqlStr = sqlStr + " ,offsellno=IsNULL(T.offsellno,0)"
	'sqlStr = sqlStr + " ,totsysstock=IsNULL(T.totsysstock,0)"
	'sqlStr = sqlStr + " ,availsysstock=IsNULL(T.availsysstock,0)"
	'sqlStr = sqlStr + " ,realstock=IsNULL(T.realstock,0)"
	'sqlStr = sqlStr + " from ("
	'sqlStr = sqlStr + " select"
	'sqlStr = sqlStr + " sum(ipgono) as ipgono,sum(reipgono) as reipgono,sum(totipgono) as totipgono,sum(offchulgono) as offchulgono,sum(offrechulgono) as offrechulgono,"
	'sqlStr = sqlStr + " sum(etcchulgono) as etcchulgono,sum(etcrechulgono) as etcrechulgono,sum(totchulgono) as totchulgono,"
	'sqlStr = sqlStr + " sum(sellno) as sellno,sum(resellno) as resellno,sum(totsellno) as totsellno,sum(errcsno) as errcsno,"
	'sqlStr = sqlStr + " sum(errbaditemno) as errbaditemno,sum(errrealcheckno) as errrealcheckno,sum(erretcno) as erretcno,"
	'sqlStr = sqlStr + " sum(toterrno) as toterrno,sum(offsellno) as offsellno, sum(totsysstock) as totsysstock,sum(availsysstock) as availsysstock,sum(realstock) as realstock"
	'sqlStr = sqlStr + "  from [db_summary].[dbo].tbl_monthly_logisstock_summary"
	'sqlStr = sqlStr + " where yyyymm<'" + Left(refreshstartdate,7) + "'"
	'sqlStr = sqlStr + " and itemgubun='" + itemgubun + "'"
	'sqlStr = sqlStr + " and itemid=" + CStr(itemid) + ""
	'sqlStr = sqlStr + " and itemoption='" + itemoption + "'"
	'sqlStr = sqlStr + " ) T"

	'sqlStr = sqlStr + " where [db_summary].[dbo].tbl_current_logisstock_summary.itemgubun='" + itemgubun + "'"
	'sqlStr = sqlStr + " and [db_summary].[dbo].tbl_current_logisstock_summary.itemid=" + CStr(itemid) + ""
	'sqlStr = sqlStr + " and [db_summary].[dbo].tbl_current_logisstock_summary.itemoption='" + itemoption + "'"

	'rsget.Open sqlStr,dbget,1



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




response.write "현재재고 업데이트... finish<br>"
response.flush


	sqlStr = " update [db_summary].[dbo].tbl_current_logisstock_summary"
	sqlStr = sqlStr + " set ipkumdiv5=IsNULL(T.ipkumdiv5,0)"
	sqlStr = sqlStr + " ,ipkumdiv4=IsNULL(T.ipkumdiv4,0)"
	sqlStr = sqlStr + " ,ipkumdiv2=IsNULL(T.ipkumdiv2,0)"
	sqlStr = sqlStr + " from ("
	sqlStr = sqlStr + " select d.itemid,d.itemoption, "
	sqlStr = sqlStr + " sum(case when m.ipkumdiv='2' then d.itemno*-1 else 0 end ) as ipkumdiv2,"
	sqlStr = sqlStr + " sum(case when m.ipkumdiv='4' then d.itemno*-1 else 0 end ) as ipkumdiv4,"
	sqlStr = sqlStr + " sum(case when m.ipkumdiv='5' then d.itemno*-1 else 0 end ) as ipkumdiv5"
	sqlStr = sqlStr + " from [db_order].[10x10].tbl_order_master m,"
	sqlStr = sqlStr + " [db_order].[10x10].tbl_order_detail d"
	sqlStr = sqlStr + " where m.orderserial=d.orderserial"
	sqlStr = sqlStr + " and m.regdate>='" + refreshstartdate + "'"
	sqlStr = sqlStr + " and m.cancelyn='N'"
	sqlStr = sqlStr + " and d.cancelyn<>'Y'"
	sqlStr = sqlStr + " and d.itemid=" + CStr(itemid) + ""
	sqlStr = sqlStr + " and d.itemoption='" + itemoption + "'"
	sqlStr = sqlStr + " group by d.itemid,d.itemoption"
	sqlStr = sqlStr + " ) T"
	sqlStr = sqlStr + " where [db_summary].[dbo].tbl_current_logisstock_summary.itemgubun='10'"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_current_logisstock_summary.itemid=" + CStr(itemid) + ""
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_current_logisstock_summary.itemoption='" + itemoption + "'"

	rsget.Open sqlStr,dbget,1


	response.write "현재 판매 업데이트... finish<br>"
	response.flush


	sqlStr = " update [db_summary].[dbo].tbl_current_logisstock_summary"
	sqlStr = sqlStr + " set sell7days=IsNULL(T.sell7days,0)"
	sqlStr = sqlStr + " from ("
	sqlStr = sqlStr + " select sum(d.itemno*-1) as sell7days"
	sqlStr = sqlStr + " from [db_order].[10x10].tbl_order_master m,"
	sqlStr = sqlStr + " [db_order].[10x10].tbl_order_detail d"
	sqlStr = sqlStr + " where m.orderserial=d.orderserial"
	sqlStr = sqlStr + " and datediff(d,m.regdate,getdate())<8"
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

	response.write "온라인 7일 주문수량.. 업데이트... finish<br>"
	response.flush


	sqlStr = " update [db_summary].[dbo].tbl_current_logisstock_summary"
	sqlStr = sqlStr + " set offchulgo7days=IsNULL(T.chulno7,0)"
	sqlStr = sqlStr + " from ("
	sqlStr = sqlStr + " select "
	sqlStr = sqlStr + " Sum(d.itemno) as chulno7"
	sqlStr = sqlStr + " from [db_storage].[10x10].tbl_acount_storage_master m,"
	sqlStr = sqlStr + " [db_storage].[10x10].tbl_acount_storage_detail d"
	sqlStr = sqlStr + " where m.code=d.mastercode"
	sqlStr = sqlStr + " and m.ipchulflag='S'"
	'sqlStr = sqlStr + " and datediff(d,m.executedt,getdate())<8"
	sqlStr = sqlStr + " and datediff(d,m.scheduledt,getdate())<8"
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

	response.write "오프 7일 주문수량 업데이트... finish<br>"
	response.flush



	sqlStr = " update [db_summary].[dbo].tbl_current_logisstock_summary"
	sqlStr = sqlStr + " set offconfirmno=IsNULL(T.offconfirmno,0)"
	sqlStr = sqlStr + " ,offjupno=IsNULL(T.offjupno,0)"
	sqlStr = sqlStr + " from ( select "
	sqlStr = sqlStr + " sum(case  when statecd='0' then realitemno*-1 else 0 end ) as offjupno,"
	sqlStr = sqlStr + " sum(case  when statecd<>'0' then realitemno*-1 else 0 end ) as offconfirmno"
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

	response.write "오프 주문 업데이트... finish<br>"
	response.flush


	sqlStr = " update [db_summary].[dbo].tbl_current_logisstock_summary"
	sqlStr = sqlStr + " set preorderno=IsNULL(T.preorderno,0)"
	sqlStr = sqlStr + " from ("
	sqlStr = sqlStr + " select sum(realitemno) as preorderno  "
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

	response.write "기주문 업데이트... finish<br>"
	response.flush


	if itemgubun="10" then
		sqlStr = " update [db_summary].[dbo].tbl_current_logisstock_summary"
		sqlStr = sqlStr + " set maxsellday=(case when datediff(d,T.regdate,getdate()) >7 then 7 else datediff(d,T.regdate,getdate()) end)"
		sqlStr = sqlStr + " from [db_item].[10x10].tbl_item T"
		sqlStr = sqlStr + " where T.itemid=" + CStr(itemid)
		sqlStr = sqlStr + " and [db_summary].[dbo].tbl_current_logisstock_summary.itemid=T.itemid"

		rsget.Open sqlStr,dbget,1

		response.write "판매일 업데이트... finish<br>"
		response.flush
	end if


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



	'sqlStr = sqlStr + " ,sell7days,offchulgo7days,ipkumdiv5,ipkumdiv4,ipkumdiv2,"
	'sqlStr = sqlStr + " offconfirmno,offjupno,requireno,shortageno,preorderno,maxsellday,"

elseif mode="itemtodayipchulrefresh" then
	''금일 필드여부확인

	''금일 출고완료입력

	''금일 입출고입력


	sqlStr = " update [db_summary].[dbo].tbl_current_logisstock_summary"
	sqlStr = sqlStr + " set ipkumdiv5=IsNULL(T.ipkumdiv5,0)"
	sqlStr = sqlStr + " ,ipkumdiv4=IsNULL(T.ipkumdiv4,0)"
	sqlStr = sqlStr + " ,ipkumdiv2=IsNULL(T.ipkumdiv2,0)"
	sqlStr = sqlStr + " from ("
	sqlStr = sqlStr + " select "
	sqlStr = sqlStr + " sum(case when m.ipkumdiv='2' then d.itemno*-1 else 0 end ) as ipkumdiv2,"
	sqlStr = sqlStr + " sum(case when m.ipkumdiv='4' then d.itemno*-1 else 0 end ) as ipkumdiv4,"
	sqlStr = sqlStr + " sum(case when m.ipkumdiv='5' then d.itemno*-1 else 0 end ) as ipkumdiv5"
	sqlStr = sqlStr + " from [db_order].[10x10].tbl_order_master m,"
	sqlStr = sqlStr + " [db_order].[10x10].tbl_order_detail d"
	sqlStr = sqlStr + " where m.orderserial=d.orderserial"
	sqlStr = sqlStr + " and m.regdate>='" + refreshstartdate + "'"
	sqlStr = sqlStr + " and m.cancelyn='N'"
	sqlStr = sqlStr + " and d.cancelyn<>'Y'"
	sqlStr = sqlStr + " and d.itemid=" + CStr(itemid) + ""
	sqlStr = sqlStr + " and d.itemoption='" + itemoption + "'"
	sqlStr = sqlStr + " ) T"
	sqlStr = sqlStr + " where [db_summary].[dbo].tbl_current_logisstock_summary.itemgubun='10'"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_current_logisstock_summary.itemid=" + CStr(itemid) + ""
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_current_logisstock_summary.itemoption='" + itemoption + "'"

	rsget.Open sqlStr,dbget,1


	response.write "현재 판매 업데이트... finish<br>"
	response.flush


	sqlStr = " update [db_summary].[dbo].tbl_current_logisstock_summary"
	sqlStr = sqlStr + " set sell7days=IsNULL(T.sell7days,0)"
	sqlStr = sqlStr + " from ("
	sqlStr = sqlStr + " select sum(d.itemno*-1) as sell7days"
	sqlStr = sqlStr + " from [db_order].[10x10].tbl_order_master m,"
	sqlStr = sqlStr + " [db_order].[10x10].tbl_order_detail d"
	sqlStr = sqlStr + " where m.orderserial=d.orderserial"
	sqlStr = sqlStr + " and datediff(d,m.regdate,getdate())<8"
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

	response.write "온라인 7일 주문수량.. 업데이트... finish<br>"
	response.flush


	sqlStr = " update [db_summary].[dbo].tbl_current_logisstock_summary"
	sqlStr = sqlStr + " set offchulgo7days=IsNULL(T.chulno7,0)"
	sqlStr = sqlStr + " from ("
	sqlStr = sqlStr + " select "
	sqlStr = sqlStr + " Sum(d.itemno) as chulno7"
	sqlStr = sqlStr + " from [db_storage].[10x10].tbl_acount_storage_master m,"
	sqlStr = sqlStr + " [db_storage].[10x10].tbl_acount_storage_detail d"
	sqlStr = sqlStr + " where m.code=d.mastercode"
	sqlStr = sqlStr + " and m.ipchulflag='S'"
	'sqlStr = sqlStr + " and datediff(d,m.executedt,getdate())<8"
	sqlStr = sqlStr + " and datediff(d,m.scheduledt,getdate())<8"
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

	response.write "오프 7일 주문수량 업데이트... finish<br>"
	response.flush



	sqlStr = " update [db_summary].[dbo].tbl_current_logisstock_summary"
	sqlStr = sqlStr + " set offconfirmno=IsNULL(T.offconfirmno,0)"
	sqlStr = sqlStr + " ,offjupno=IsNULL(T.offjupno,0)"
	sqlStr = sqlStr + " from ( select "
	sqlStr = sqlStr + " sum(case  when statecd='0' then realitemno*-1 else 0 end ) as offjupno,"
	sqlStr = sqlStr + " sum(case  when statecd<>'0' then realitemno*-1 else 0 end ) as offconfirmno"
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

	response.write "오프 주문 업데이트... finish<br>"
	response.flush


	sqlStr = " update [db_summary].[dbo].tbl_current_logisstock_summary"
	sqlStr = sqlStr + " set preorderno=IsNULL(T.preorderno,0)"
	sqlStr = sqlStr + " from ("
	sqlStr = sqlStr + " select sum(realitemno) as preorderno  "
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

	response.write "기주문 업데이트... finish<br>"
	response.flush


	if itemgubun="10" then
		sqlStr = " update [db_summary].[dbo].tbl_current_logisstock_summary"
		sqlStr = sqlStr + " set maxsellday=(case when datediff(d,T.regdate,getdate()) >7 then 7 else datediff(d,T.regdate,getdate()) end)"
		sqlStr = sqlStr + " from [db_item].[10x10].tbl_item T"
		sqlStr = sqlStr + " where T.itemid=" + CStr(itemid)
		sqlStr = sqlStr + " and [db_summary].[dbo].tbl_current_logisstock_summary.itemid=T.itemid"

		rsget.Open sqlStr,dbget,1

		response.write "판매일 업데이트... finish<br>"
		response.flush
	end if


	sqlStr = " update [db_summary].[dbo].tbl_current_logisstock_summary"
	sqlStr = sqlStr + " set requireno=convert(int,(sell7days+offchulgo7days)*7/maxsellday)"
	sqlStr = sqlStr + " where itemgubun='" + itemgubun + "'"
	sqlStr = sqlStr + " and itemid=" + CStr(itemid)
	sqlStr = sqlStr + " and itemoption='" + itemoption + "'"
	sqlStr = sqlStr + " and maxsellday<>0"

	rsget.Open sqlStr,dbget,1


	sqlStr = " update [db_summary].[dbo].tbl_current_logisstock_summary"
	sqlStr = sqlStr + " set shortageno=realstock+requireno"
	sqlStr = sqlStr + " where itemgubun='" + itemgubun + "'"
	sqlStr = sqlStr + " and itemid=" + CStr(itemid)
	sqlStr = sqlStr + " and itemoption='" + itemoption + "'"

	rsget.Open sqlStr,dbget,1
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
		orgrealstock = rsget("realstock")
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

	''현재까지 입력된 입력 오차 + 금일오차
	errstock = realstock-orgrealstock + todayerrno

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


	''재고 월별테이블 삭제
	'sqlStr = " delete from [db_summary].[dbo].tbl_monthly_logisstock_summary"
	'sqlStr = sqlStr + " where yyyymm>='" + Left(yyyymmdd,7) + "'"
	'sqlStr = sqlStr + " and itemgubun='" + itemgubun + "'"
	'sqlStr = sqlStr + " and itemid=" + CStr(itemid) + ""
	'sqlStr = sqlStr + " and itemoption='" + itemoption + "'"

	'rsget.Open sqlStr,dbget,1

	''재고 월별테이블 업데이트
	'sqlStr = " insert into [db_summary].[dbo].tbl_monthly_logisstock_summary"
	'sqlStr = sqlStr + " (yyyymm,itemgubun,itemid,itemoption,"
	'sqlStr = sqlStr + " ipgono,reipgono,totipgono,offchulgono,offrechulgono,"
	'sqlStr = sqlStr + " etcchulgono,etcrechulgono,totchulgono,"
	'sqlStr = sqlStr + " sellno,resellno,totsellno,errcsno,"
	'sqlStr = sqlStr + " errbaditemno,errrealcheckno,erretcno,"
	'sqlStr = sqlStr + " toterrno,totsysstock,availsysstock,realstock)"
	'sqlStr = sqlStr + " select"
	'sqlStr = sqlStr + " convert(varchar(7),yyyymmdd,21) as yyyymm,itemgubun,itemid,itemoption,"
	'sqlStr = sqlStr + " sum(ipgono),sum(reipgono),sum(totipgono),sum(offchulgono),sum(offrechulgono),"
	'sqlStr = sqlStr + " sum(etcchulgono),sum(etcrechulgono),sum(totchulgono),"
	'sqlStr = sqlStr + " sum(sellno),sum(resellno),sum(totsellno),sum(errcsno),"
	'sqlStr = sqlStr + " sum(errbaditemno),sum(errrealcheckno),sum(erretcno),"
	'sqlStr = sqlStr + " sum(toterrno),sum(totsysstock),sum(availsysstock),sum(realstock)"
	'sqlStr = sqlStr + "  from [db_summary].[dbo].tbl_daily_logisstock_summary"
	'sqlStr = sqlStr + " where yyyymmdd>='" + yyyymmdd + "-01" + "'"
	'sqlStr = sqlStr + " and itemgubun='" + itemgubun + "'"
	'sqlStr = sqlStr + " and itemid=" + CStr(itemid) + ""
	'sqlStr = sqlStr + " and itemoption='" + itemoption + "'"
	'sqlStr = sqlStr + " group by convert(varchar(7),yyyymmdd,21) ,itemgubun,itemid,itemoption"

	'rsget.Open sqlStr,dbget,1
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
	sqlStr = sqlStr + " set toterrno=errcsno+errbaditemno+erretcno+errrealcheckno"
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
	sqlStr = sqlStr + " set toterrno=errcsno+errbaditemno+erretcno+errrealcheckno"
	sqlStr = sqlStr + " ,totsysstock=totipgono+totchulgono-totsellno"
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
	sqlStr = sqlStr + " set errbaditemno=IsNULL(T.errbaditemno,0)"
	sqlStr = sqlStr + " ,toterrno=errcsno+errrealcheckno+erretcno+IsNULL(T.errbaditemno,0)"
	sqlStr = sqlStr + " ,totsysstock=totipgono+totchulgono-totsellno"
	sqlStr = sqlStr + " ,availsysstock=totipgono+totchulgono-totsellno+errcsno+erretcno+IsNULL(T.errbaditemno,0)"
	sqlStr = sqlStr + " ,realstock=totipgono+totchulgono-totsellno+errcsno+errrealcheckno+erretcno+IsNULL(T.errbaditemno,0)"
	'sqlStr = sqlStr + " ,lastupdate=getdate()"
	sqlStr = sqlStr + " from ("
	sqlStr = sqlStr + " select s.itemgubun, s.itemid, s.itemoption, sum(s.errbaditemno) as errbaditemno"
	sqlStr = sqlStr + " from [db_summary].[dbo].tbl_erritem_daily_summary s,"
	sqlStr = sqlStr + " [db_summary].[dbo].tbl_temp_baditem b "
	sqlStr = sqlStr + " where s.itemgubun=b.itemgubun"
	sqlStr = sqlStr + " and s.itemid=b.itemid"
	sqlStr = sqlStr + " and s.itemoption=b.itemoption"
	sqlStr = sqlStr + " group by s.itemgubun, s.itemid, s.itemoption"
	sqlStr = sqlStr + " ) T"
	sqlStr = sqlStr + " where [db_summary].[dbo].tbl_current_logisstock_summary.itemgubun=T.itemgubun"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_current_logisstock_summary.itemid=T.itemid"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_current_logisstock_summary.itemoption=T.itemoption"

	rsget.Open sqlStr,dbget,1


	''임시 불량상품 삭제
	sqlStr = "delete from [db_summary].[dbo].tbl_temp_baditem"

	rsget.Open sqlStr,dbget,1
end if
%>


<script language="javascript">
<% if mode="tmpbaditem2input" then %>
alert('저장 되었습니다.');
opener.location.reload();
window.close();
<% else %>
alert('저장 되었습니다.');
location.replace('<%= refer %>');
<% end if %>
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->