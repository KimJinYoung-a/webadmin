<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<%

dim refer
refer = request.ServerVariables("HTTP_REFERER")

dim mode, targetid, targetname, divcode, defaultmargine, itemgubunarr, itemidarr, itemoptionarr, itemnoarr
dim designer, sellcash, suplycash, buycash, mwdiv, itemname, itemoptionname
dim itemgubun, itemid, itemoption, itemno
dim pmwdiv
dim sqlStr, i

mode = requestCheckVar(request.form("mode"),32)
targetid = requestCheckVar(request.form("brandid"),32)

itemgubunarr = request.form("itemgubunarr")
itemidarr = request.form("itemidarr")
itemoptionarr = request.form("itemoptionarr")
itemnoarr = request.form("itemnoarr")
pmwdiv    = requestCheckVar(request.form("pmwdiv"),1)

itemgubunarr = split(itemgubunarr, "|")
itemidarr = split(itemidarr, "|")
itemoptionarr = split(itemoptionarr, "|")
itemnoarr = split(itemnoarr, "|")

dim iid, ipgocode

if (mode = "ipgoarr") then
	'======================================================================
	'반품입고등록
	'1.온라인 입고 마스타

	'업체명 검색 - 수정요망 매입위탁반품 .
	sqlStr = " select top 1 socname_kor,maeipdiv,defaultmargine from [db_user].[dbo].tbl_user_c"
	sqlStr = sqlStr + " where userid='" + targetid + "'"
	rsget.Open sqlStr, dbget, 1
	if Not rsget.Eof then
		targetname = rsget("socname_kor")
		divcode = rsget("maeipdiv")
		defaultmargine = rsget("defaultmargine")

        	if divcode="M" then
        		divcode = "001"
        	elseif divcode="W" then
        		divcode = "002"
        	end if
	end if
	rsget.close
    
    ''매입구분.
    if (pmwdiv="M") then 
        divcode = "001"
    elseif (pmwdiv="W") then
        divcode = "002"
    end if
    
	sqlStr = " select * from [db_storage].[dbo].tbl_acount_storage_master where 1=0"
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew
	rsget("code") = ""
	rsget("socid") = targetid
	rsget("socname") = targetname
	rsget("chargeid") = session("ssBctid")
	rsget("chargename") = session("ssBctCname")
	rsget("divcode") = divcode ''001-매입, 002-위탁
	rsget("vatcode") = "008"   ''부가세.(이것만 받는다.)
	rsget("comment") = "불량상품반품처리"
	rsget("scheduledt") = Left(now(), 10)
	rsget("executedt") = Left(now(), 10)
	rsget("ipchulflag") = "I"

	rsget.update
	iid = rsget("id")
	rsget.close

	ipgocode = "ST" + Format00(6,Right(CStr(iid),6))

	sqlStr = " update [db_storage].[dbo].tbl_acount_storage_master " + VBCrlf
	sqlStr = sqlStr + " set code='" + ipgocode + "'" + VBCrlf
	sqlStr = sqlStr + " where id=" + CStr(iid)
	rsget.Open sqlStr,dbget,1

	'''2.온라인 입고 디테일 입력
	for i=0 to UBound(itemgubunarr) - 1
		if (trim(itemgubunarr(i)) <> "") then
			itemgubun = trim(itemgubunarr(i))
			itemid = trim(itemidarr(i))
			itemoption = trim(itemoptionarr(i))
			itemno = CStr(-1 * trim(itemnoarr(i)))
			designer = targetid

			sellcash = "0"
			suplycash = "0"
			buycash = "0"
			mwdiv = "0"
			itemname = ""
			itemoptionname = ""

			sqlStr = " insert into [db_storage].[dbo].tbl_acount_storage_detail " + VBCrlf
			sqlStr = sqlStr + " (mastercode,itemid,itemoption,sellcash,suplycash, " + VBCrlf
			sqlStr = sqlStr + " itemno,indt,updt,buycash,mwgubun,iitemgubun,iitemname,iitemoptionname,imakerid) " + VBCrlf
			sqlStr = sqlStr + " values('" + ipgocode + "'," + requestCheckVar(itemid,10) + ", '" + requestCheckVar(itemoption,4) + "', " + sellcash + ", " + suplycash + ", " + itemno + ", getdate(), getdate(), " + buycash + ", '" + mwdiv + "', '" + requestCheckVar(itemgubun,2) + "', '" + itemname + "', '" + itemoptionname + "', '" + designer + "') " + VBCrlf
			rsget.Open sqlStr,dbget,1
		end if
	next

    ''상품정보
	sqlStr = " update [db_storage].[dbo].tbl_acount_storage_detail" + vbCrlf
	sqlStr = sqlStr + " set iitemname=[db_item].[dbo].tbl_item.itemname"
	sqlStr = sqlStr + " , sellcash=[db_item].[dbo].tbl_item.sellcash"
	sqlStr = sqlStr + " , suplycash=[db_item].[dbo].tbl_item.buycash"
	sqlStr = sqlStr + " , buycash=[db_item].[dbo].tbl_item.buycash"
	sqlStr = sqlStr + " , mwgubun=[db_item].[dbo].tbl_item.mwdiv"
	sqlStr = sqlStr + " from [db_item].[dbo].tbl_item "
	sqlStr = sqlStr + " where mastercode='" + CStr(ipgocode) + "'"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.iitemgubun='10'"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.itemid=[db_item].[dbo].tbl_item.itemid"
	rsget.Open sqlStr, dbget, 1
    
    ''옵션명
	sqlStr = " update [db_storage].[dbo].tbl_acount_storage_detail" + vbCrlf
	sqlStr = sqlStr + " set iitemoptionname=IsNULL([db_item].[dbo].vw_all_option.codeview,'')"
	sqlStr = sqlStr + " from [db_item].[dbo].vw_all_option "
	sqlStr = sqlStr + " where mastercode='" + CStr(ipgocode) + "'"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.itemoption=[db_item].[dbo].vw_all_option.optioncode"
	rsget.Open sqlStr, dbget, 1


	'''2.온라인 입고 마스타 업데이트
	sqlStr = " update [db_storage].[dbo].tbl_acount_storage_master " + VBCrlf
	sqlStr = sqlStr + " set totalsellcash=T.totsell" + VBCrlf
	sqlStr = sqlStr + " ,totalsuplycash=T.totsupp" + VBCrlf
	sqlStr = sqlStr + " ,totalbuycash=T.totbuy" + VBCrlf
	sqlStr = sqlStr + " ,indt=getdate()" + VBCrlf
	sqlStr = sqlStr + " ,updt=getdate()" + VBCrlf
	sqlStr = sqlStr + " from (select sum(sellcash*itemno) as totsell, " + vbCrlf
	sqlStr = sqlStr + " sum(suplycash*itemno) as totsupp, " + vbCrlf
	sqlStr = sqlStr + " sum(buycash*itemno) as totbuy " + vbCrlf
	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_detail" + vbCrlf
	sqlStr = sqlStr + " where mastercode='"  + CStr(ipgocode) + "'" + vbCrlf
	sqlStr = sqlStr + " and deldt is null" + vbCrlf
	sqlStr = sqlStr + " ) as T"
	sqlStr = sqlStr + " where id=" + CStr(iid)
	rsget.Open sqlStr,dbget,1


	'======================================================================
	'불량상품서머리정보 업데이트 -> 프로시져로 수정요망..
     dim yyyymmdd

	''오차 입력일
	sqlStr = "select convert(varchar(10),getdate(),21) as yyyymmdd "
	rsget.Open sqlStr,dbget,1
		yyyymmdd = rsget("yyyymmdd")
	rsget.Close

	''기입력된 오차에 추가함 (테이블에 상품이 존재할 경우)
	sqlStr = " update [db_summary].[dbo].tbl_erritem_daily_summary"
	sqlStr = sqlStr + " set errbaditemno=errbaditemno + IsNULL(T.itemno,0)*-1"
	sqlStr = sqlStr + " from ( "
	sqlStr = sqlStr + " select b.iitemgubun as itemgubun, b.itemid, b.itemoption, b.itemno "
	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_detail b, [db_summary].[dbo].tbl_erritem_daily_summary s"
	sqlStr = sqlStr + " where s.yyyymmdd='" + yyyymmdd + "'"
	sqlStr = sqlStr + " and b.iitemgubun=s.itemgubun"
	sqlStr = sqlStr + " and b.itemid=s.itemid"
	sqlStr = sqlStr + " and b.itemoption=s.itemoption"
	sqlStr = sqlStr + " and b.mastercode='" + CStr(ipgocode) + "' "
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
	sqlStr = sqlStr + " select b.iitemgubun as itemgubun, b.itemid, b.itemoption, b.itemno "
	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_detail b "
	sqlStr = sqlStr + " left join [db_summary].[dbo].tbl_erritem_daily_summary s on s.yyyymmdd='" + yyyymmdd + "'"
	sqlStr = sqlStr + " and b.iitemgubun=s.itemgubun"
	sqlStr = sqlStr + " and b.itemid=s.itemid"
	sqlStr = sqlStr + " and b.itemoption=s.itemoption"
	sqlStr = sqlStr + " where s.itemid is null and b.mastercode='" + CStr(ipgocode) + "' "
	sqlStr = sqlStr + " ) T"
	rsget.Open sqlStr,dbget,1

	sqlStr = "update [db_summary].[dbo].tbl_erritem_daily_summary"
	sqlStr = sqlStr + " set toterrno=errcsno+errbaditemno+erretcno+errrealcheckno"
	sqlStr = sqlStr + " ,lastupdate=getdate()"
	sqlStr = sqlStr + " ,modiuser='" + session("ssBctId") + "'"
	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_detail b "
	sqlStr = sqlStr + " where [db_summary].[dbo].tbl_erritem_daily_summary.yyyymmdd='" + yyyymmdd + "'"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_erritem_daily_summary.itemgubun=b.iitemgubun"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_erritem_daily_summary.itemid=b.itemid"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_erritem_daily_summary.itemoption=b.itemoption"
	sqlStr = sqlStr + " and b.mastercode='" + CStr(ipgocode) + "' "
	rsget.Open sqlStr,dbget,1

	''일별 재고로그에 추가
	sqlStr = " update [db_summary].[dbo].tbl_daily_logisstock_summary "
	sqlStr = sqlStr + " set errbaditemno=errbaditemno + b.itemno*-1"
	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_detail b "
	sqlStr = sqlStr + " where [db_summary].[dbo].tbl_daily_logisstock_summary.yyyymmdd='" + yyyymmdd + "'"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_daily_logisstock_summary.itemgubun=b.iitemgubun"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_daily_logisstock_summary.itemid=b.itemid"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_daily_logisstock_summary.itemoption=b.itemoption"
	sqlStr = sqlStr + " and b.mastercode='" + CStr(ipgocode) + "' "
	rsget.Open sqlStr,dbget,1

	sqlStr = " insert into [db_summary].[dbo].tbl_daily_logisstock_summary"
	sqlStr = sqlStr + " (yyyymmdd,itemgubun,itemid,itemoption,errbaditemno)"
	sqlStr = sqlStr + " select "
	sqlStr = sqlStr + " '" + yyyymmdd + "'"
	sqlStr = sqlStr + " ,T.itemgubun,T.itemid,T.itemoption,T.itemno*-1"
	sqlStr = sqlStr + " from ("
	sqlStr = sqlStr + " select b.iitemgubun as itemgubun, b.itemid, b.itemoption, b.itemno "
	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_detail b "
	sqlStr = sqlStr + " left join [db_summary].[dbo].tbl_daily_logisstock_summary s on s.yyyymmdd='" + yyyymmdd + "'"
	sqlStr = sqlStr + " and b.iitemgubun=s.itemgubun"
	sqlStr = sqlStr + " and b.itemid=s.itemid"
	sqlStr = sqlStr + " and b.itemoption=s.itemoption"
	sqlStr = sqlStr + " where s.itemid is null and b.mastercode='" + CStr(ipgocode) + "' "
	sqlStr = sqlStr + " ) T"
	rsget.Open sqlStr,dbget,1

    
	''서머리.
	sqlStr = " update [db_summary].[dbo].tbl_daily_logisstock_summary"
	sqlStr = sqlStr + " set toterrno=errcsno+errbaditemno+erretcno+errrealcheckno"
	sqlStr = sqlStr + " ,totsysstock=totipgono+totchulgono-totsellno"
	sqlStr = sqlStr + " ,availsysstock=totipgono+totchulgono-totsellno+errcsno+errbaditemno+erretcno"
	sqlStr = sqlStr + " ,realstock=totipgono+totchulgono-totsellno+errcsno+errbaditemno+erretcno+errrealcheckno"
	sqlStr = sqlStr + " ,lastupdate=getdate()"
	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_detail b "
	sqlStr = sqlStr + " where [db_summary].[dbo].tbl_daily_logisstock_summary.yyyymmdd='" + yyyymmdd + "'"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_daily_logisstock_summary.itemgubun=b.iitemgubun"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_daily_logisstock_summary.itemid=b.itemid"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_daily_logisstock_summary.itemoption=b.itemoption"
	sqlStr = sqlStr + " and b.mastercode='" + CStr(ipgocode) + "' "

	rsget.Open sqlStr,dbget,1


	''현재고테이블 업데이트 : 시스템 재고만 변화?
	sqlStr = "update [db_summary].[dbo].tbl_current_logisstock_summary"
	sqlStr = sqlStr + " set errbaditemno=IsNULL(T.errbaditemno,0)"
	sqlStr = sqlStr + " ,toterrno=errrealcheckno+erretcno+IsNULL(T.errbaditemno,0)"
	sqlStr = sqlStr + " ,totsysstock=totipgono+totchulgono-totsellno+errcsno"
	sqlStr = sqlStr + " ,availsysstock=totipgono+totchulgono-totsellno+errcsno+erretcno+IsNULL(T.errbaditemno,0)"
	sqlStr = sqlStr + " ,realstock=totipgono+totchulgono-totsellno+errcsno+errrealcheckno+erretcno+IsNULL(T.errbaditemno,0)"
	sqlStr = sqlStr + " ,shortageno=totipgono+totchulgono-totsellno+errcsno+errrealcheckno+erretcno+requireno+ipkumdiv5+ipkumdiv4+ipkumdiv2+offconfirmno+offjupno+IsNULL(T.errbaditemno,0)"
	sqlStr = sqlStr + " ,lastupdate=getdate()"
	sqlStr = sqlStr + " from ("
	sqlStr = sqlStr + "     select s.itemgubun, s.itemid, s.itemoption, sum(s.errbaditemno) as errbaditemno"
	sqlStr = sqlStr + "     from [db_summary].[dbo].tbl_erritem_daily_summary s,"
	sqlStr = sqlStr + "     [db_storage].[dbo].tbl_acount_storage_detail b "
	sqlStr = sqlStr + "     where s.itemgubun=b.iitemgubun"
	sqlStr = sqlStr + "     and s.itemid=b.itemid"
	sqlStr = sqlStr + "     and s.itemoption=b.itemoption"
	sqlStr = sqlStr + "     and b.mastercode='" + CStr(ipgocode) + "' "
	sqlStr = sqlStr + "     group by s.itemgubun, s.itemid, s.itemoption"
	sqlStr = sqlStr + " ) T"
	sqlStr = sqlStr + " where [db_summary].[dbo].tbl_current_logisstock_summary.itemgubun=T.itemgubun"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_current_logisstock_summary.itemid=T.itemid"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_current_logisstock_summary.itemoption=T.itemoption"
	rsget.Open sqlStr,dbget,1
	
	''입출고 업데이트
	sqlStr = "exec db_summary.dbo.ten_recentIpChul_Update '" & ipgocode & "','','',0,'',''"
	dbget.Execute sqlStr
end if

%>

<script type='text/javascript'>
	alert('저장 되었습니다.');
	location.replace('<%= refer %>');
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->