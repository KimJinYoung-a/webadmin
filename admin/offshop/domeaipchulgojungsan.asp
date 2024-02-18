<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프라인 가맹점 정산관리
' History : 2009.04.07 서동석 생성
'			2010.05.13 한용민 수정
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
dim mode, check, idx, topidx ,shopid, makerid, yyyy, mm ,workidx , adminuserid ,adminusername
dim sqlStr, i, iid, cnt ,MaybeYYYYMM,Maydiffkey, shopname, maybetitle, shopdiv
dim ckidx, suplycasharr,itemnoarr ,orgsellcasharr,sellcasharr,buycasharr ,b2ccharge
	mode    = requestCheckVar(request("mode"),32)
	check   = request("check")
	shopid  = requestCheckVar(request("shopid"),32)
	idx     = requestCheckVar(request("idx"),10)
	topidx  = requestCheckVar(request("topidx"),10)
	makerid = requestCheckVar(request("makerid"),32)
	yyyy    = requestCheckVar(request("yyyy1"),4)
	mm      = requestCheckVar(request("mm1"),2)
	workidx      = requestCheckVar(request("workidx"),10)
	b2ccharge      = requestCheckVar(request("b2ccharge"),20)

adminuserid = session("ssBctId")
adminusername = session("ssBctCname")

if (workidx = "") then
	workidx = "NULL"
end if

dim refer
	refer = request.ServerVariables("HTTP_REFERER")

'response.write mode & "!!!<Br>"

if mode="chulgo" then

	'' insert master
	sqlStr = " select * from [db_shop].[dbo].tbl_fran_meachuljungsan_master where 1=0"
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew

	rsget("shopid") = shopid
	rsget("title")  = shopid & " 작성중"
	rsget("totalsum")   = 0
	rsget("divcode")    = "MC"
	rsget("etcstr")     = ""
	rsget("reguserid")  = session("ssBctId")
	rsget("regusername") = session("ssBctCname")

	if (workidx <> "NULL") then
		rsget("workidx") = workidx
	end if

	rsget.update
	iid = rsget("idx")
	rsget.close

	if (workidx <> "NULL") then
		sqlStr = " update db_storage.dbo.tbl_cartoonbox_master "
		sqlStr = sqlStr + " set "
		sqlStr = sqlStr + " 	jungsanidx = " + CStr(iid) + " "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + " 	and idx = " + CStr(workidx) + " "
		'response.write "aaaaaaaaaaaa" & sqlStr
		rsget.Open sqlStr, dbget, 1
	end if

	'' insert sub master
	sqlStr = " insert into [db_shop].[dbo].tbl_fran_meachuljungsan_submaster "
	sqlStr = sqlStr + " (masteridx,linkidx,shopid,code01,code02,execdate, "
	sqlStr = sqlStr + " totalcount,totalsellcash,totalbuycash,totalsuplycash)"
	sqlStr = sqlStr + " select " + CStr(iid) + ",m.id, m.socid, m.code, s.baljucode, m.executedt,"
	sqlStr = sqlStr + " 0, m.totalsellcash*-1, m.totalbuycash*-1, m.totalsuplycash*-1"
	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_master m"
	sqlStr = sqlStr + " left join [db_storage].[dbo].tbl_ordersheet_master s on m.code=s.alinkcode"
	sqlStr = sqlStr + " where m.id in (" + check + ")"

	rsget.Open sqlStr, dbget, 1

	'' insert sub detail
	sqlStr = " insert into [db_shop].[dbo].tbl_fran_meachuljungsan_subdetail "
	sqlStr = sqlStr + " (masteridx, topmasteridx, linkbaljucode, linkmastercode, linkdetailidx,"
	sqlStr = sqlStr + " itemgubun, itemid, itemoption, itemname, itemoptionname,"
	sqlStr = sqlStr + " makerid, itemno, sellcash, suplycash, buycash)"
	sqlStr = sqlStr + " select  m.idx," + CStr(iid) + ",'',m.code01,"
	sqlStr = sqlStr + " d.id, d.iitemgubun, d.itemid, d.itemoption, d.iitemname, d.iitemoptionname,"
	sqlStr = sqlStr + " d.imakerid, d.itemno*-1,d.sellcash,d.suplycash,d.buycash"
	sqlStr = sqlStr + " from [db_shop].[dbo].tbl_fran_meachuljungsan_submaster m,"
	sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_detail d"
	sqlStr = sqlStr + " where m.masteridx=" + CStr(iid)
	sqlStr = sqlStr + " and m.code01=d.mastercode"

	rsget.Open sqlStr, dbget, 1

    MaybeYYYYMM = fnGetMayYYYYMM(iid)
    shopname    = fnGetShopName(shopid, shopdiv)
    Maydiffkey  = fnGetMayDiffKey(shopid, MaybeYYYYMM)

    IF (MaybeYYYYMM<>"") and (shopname<>"") then
        maybetitle = shopname + " " + MaybeYYYYMM + " " + CStr(Maydiffkey) + "차 출고분"
    End IF

	sqlStr = " update [db_shop].[dbo].tbl_fran_meachuljungsan_master"
	sqlStr = sqlStr + " set totalsum=T.totalsum"
	sqlStr = sqlStr + " ,totalsellcash=T.totalsellcash"
	sqlStr = sqlStr + " ,totalbuycash=T.totalbuycash"
	sqlStr = sqlStr + " ,totalsuplycash=T.totalsuplycash"

	IF (MaybeYYYYMM<>"") then
	    sqlStr = sqlStr + " ,yyyymm='"& MaybeYYYYMM &"'"
	end if
	IF (MaybeTitle<>"") then
	    sqlStr = sqlStr + " ,title='"&MaybeTitle&"'"
	end if

	sqlStr = sqlStr + " ,diffkey="&Maydiffkey&""
	sqlStr = sqlStr + " ,shopdiv='"&shopdiv&"'"
	sqlStr = sqlStr + " from (select sum(totalsuplycash) as totalsum, sum(totalsellcash) as totalsellcash, "
	sqlStr = sqlStr + "			sum(totalbuycash) as totalbuycash, sum(totalsuplycash) as totalsuplycash "
	sqlStr = sqlStr + " 		from [db_shop].[dbo].tbl_fran_meachuljungsan_submaster"
	sqlStr = sqlStr + " 		where masteridx=" + CStr(iid)
	sqlStr = sqlStr + " 	) as T "
	sqlStr = sqlStr + " where idx=" + CStr(iid)

	rsget.Open sqlStr, dbget, 1

elseif mode="witsksell" then
    '' insert master
	sqlStr = " select * from [db_shop].[dbo].tbl_fran_meachuljungsan_master where 1=0"
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew

	rsget("shopid") = shopid
	rsget("title") = shopid & " 작성중"
	rsget("totalsum") = 0
	rsget("divcode") = "WS"
	rsget("etcstr") = ""
	rsget("reguserid") = session("ssBctId")
	rsget("regusername") = session("ssBctCname")

	rsget.update
	iid = rsget("idx")
	rsget.close

	'' insert sub master
	sqlStr = " insert into [db_shop].[dbo].tbl_fran_meachuljungsan_submaster "
	sqlStr = sqlStr + " (masteridx,linkidx,shopid,code01,code02,execdate, "
	sqlStr = sqlStr + " totalcount,totalsellcash,totalbuycash,totalsuplycash)"
	sqlStr = sqlStr + " select " + CStr(iid) + ", m.idx, d.shopid, convert(varchar(7),m.yyyymm),"
	sqlStr = sqlStr + " m.makerid, convert(varchar(7),m.yyyymm) + '-01', sum(d.itemno) as totitemcnt,"
	sqlStr = sqlStr + " sum(d.realsellprice*d.itemno) as totsum, sum(d.suplyprice*d.itemno) as realjungsansum, 0"
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + "     [db_jungsan].[dbo].tbl_off_jungsan_master m, "
    sqlStr = sqlStr + "     [db_jungsan].[dbo].tbl_off_jungsan_detail d "
    sqlStr = sqlStr + "     where m.idx=d.masteridx "
    sqlStr = sqlStr + "     and m.idx in  (" + check + ")"
    sqlStr = sqlStr + "     and d.gubuncd in ('B012','B013') "
    sqlStr = sqlStr + "     and d.shopid='" + shopid + "'"
    sqlStr = sqlStr + " group by m.idx, d.shopid, convert(varchar(7),m.yyyymm), m.makerid, convert(varchar(7),m.yyyymm) + '-01'"

	rsget.Open sqlStr, dbget, 1

	'' insert sub detail : sellprice, orgsellprice
	sqlStr = " insert into [db_shop].[dbo].tbl_fran_meachuljungsan_subdetail "
	sqlStr = sqlStr + " (masteridx, topmasteridx, linkbaljucode, linkmastercode, linkdetailidx,"
	sqlStr = sqlStr + " itemgubun, itemid, itemoption, itemname, itemoptionname,"
	sqlStr = sqlStr + " makerid, itemno, sellcash, suplycash, buycash, orgsellcash)"
	sqlStr = sqlStr + " select m.idx," + CStr(iid) + ",d.gubuncd, d.orderno, d.detailidx,"
	sqlStr = sqlStr + " d.itemgubun,d.itemid,d.itemoption,d.itemname,d.itemoptionname,"
	sqlStr = sqlStr + " d.makerid,d.itemno,d.realsellprice,0,d.suplyprice,d.sellprice"
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " [db_shop].[dbo].tbl_fran_meachuljungsan_submaster m,"
	sqlStr = sqlStr + " [db_jungsan].[dbo].tbl_off_jungsan_detail d"
	sqlStr = sqlStr + " where m.linkidx=d.masteridx"
	sqlStr = sqlStr + " and m.masteridx=" + CStr(iid)
	sqlStr = sqlStr + " and d.gubuncd in ('B012','B013') "
	sqlStr = sqlStr + " and d.shopid='" + shopid + "'"

	rsget.Open sqlStr, dbget, 1

	'' update Detail shopsuplyprice
	'' 현재 OFF 상품 가격 기준 정산.
	''
	sqlStr = " update [db_shop].[dbo].tbl_fran_meachuljungsan_subdetail"
	sqlStr = sqlStr + " set suplycash=IsNULL(T.shopbuyprice,0)"
	sqlStr = sqlStr + " from ("
	sqlStr = sqlStr + " 	select distinct d.idx, IsNULL(s.defaultsuplymargin,35) as defaultsuplymargin"
	sqlStr = sqlStr + " ,	( case "
	sqlStr = sqlStr + "			when (i.shopbuyprice=0) and (j.discountprice=0) then convert(int,j.sellprice - j.sellprice*IsNULL(s.defaultsuplymargin,35)/100)"
	sqlStr = sqlStr + "  		when (i.shopbuyprice=0) and (j.discountprice<>0) then convert(int,j.discountprice - j.discountprice*IsNULL(s.defaultsuplymargin,35)/100)"
	sqlStr = sqlStr + "    		else i.shopbuyprice "
	sqlStr = sqlStr + "    		end ) as shopbuyprice "
	sqlStr = sqlStr + " 	from [db_shop].[dbo].tbl_fran_meachuljungsan_subdetail d"
	sqlStr = sqlStr + " 		left join [db_shop].[dbo].tbl_shop_designer s "
	sqlStr = sqlStr + " 			on d.makerid=s.makerid and s.shopid='" + shopid + "'"
	sqlStr = sqlStr + " 		left join [db_shop].[dbo].tbl_shopjumun_detail j"
	sqlStr = sqlStr + " 			on d.linkmastercode=j.orderno"
	sqlStr = sqlStr + " 			and d.itemgubun=j.itemgubun"
	sqlStr = sqlStr + " 			and d.itemid=j.itemid"
	sqlStr = sqlStr + " 			and d.itemoption=j.itemoption"
	sqlStr = sqlStr + " 		left join [db_shop].[dbo].tbl_shop_item i"
	sqlStr = sqlStr + " 			on d.itemgubun=i.itemgubun"
	sqlStr = sqlStr + " 			and d.itemid=i.shopitemid"
	sqlStr = sqlStr + " 			and d.itemoption=i.itemoption"
	sqlStr = sqlStr + " 	where d.topmasteridx=" + CStr(iid)
	sqlStr = sqlStr + " ) as T"
	sqlStr = sqlStr + " where [db_shop].[dbo].tbl_fran_meachuljungsan_subdetail.idx=T.idx"

	rsget.Open sqlStr, dbget, 1

	'' update Sub master
	sqlStr = " update [db_shop].[dbo].tbl_fran_meachuljungsan_submaster"
	sqlStr = sqlStr + " set totalsellcash=T.totalsellcash"
	sqlStr = sqlStr + " ,totalbuycash=T.totalbuycash"
	sqlStr = sqlStr + " ,totalsuplycash=T.totalsuplycash"
	sqlStr = sqlStr + " ,totalorgsellcash=T.totalorgsellcash"
	sqlStr = sqlStr + " from (select masteridx, sum(sellcash*itemno) as totalsellcash, "
	sqlStr = sqlStr + "			sum(buycash*itemno) as totalbuycash, sum(suplycash*itemno) as totalsuplycash,"
	sqlStr = sqlStr + "			sum(orgsellcash*itemno) as totalorgsellcash "
	sqlStr = sqlStr + " 		from [db_shop].[dbo].tbl_fran_meachuljungsan_subdetail"
	sqlStr = sqlStr + " 		where topmasteridx=" + CStr(iid)
	sqlStr = sqlStr + "		    group by masteridx"
	sqlStr = sqlStr + " ) as T "
	sqlStr = sqlStr + " where [db_shop].[dbo].tbl_fran_meachuljungsan_submaster.idx=T.masteridx"
	sqlStr = sqlStr + " and [db_shop].[dbo].tbl_fran_meachuljungsan_submaster.masteridx=" + CStr(iid)

	rsget.Open sqlStr, dbget, 1

    MaybeYYYYMM = fnGetMayYYYYMM(iid)
    shopname    = fnGetShopName(shopid, shopdiv)
    Maydiffkey  = fnGetMayDiffKey(shopid, MaybeYYYYMM)

    IF (MaybeYYYYMM<>"") and (shopname<>"") then
        maybetitle = shopname + " " + MaybeYYYYMM + " 업체위탁 상품대"
    End IF

	'' update master
	sqlStr = " update [db_shop].[dbo].tbl_fran_meachuljungsan_master"
	sqlStr = sqlStr + " set totalsum=T.totalsum"
	sqlStr = sqlStr + " ,totalsellcash=T.totalsellcash"
	sqlStr = sqlStr + " ,totalbuycash=T.totalbuycash"
	sqlStr = sqlStr + " ,totalsuplycash=T.totalsuplycash"
	sqlStr = sqlStr + " ,totalorgsellcash=T.totalorgsellcash"

	IF (MaybeYYYYMM<>"") then
	    sqlStr = sqlStr + " ,yyyymm='"& MaybeYYYYMM &"'"
	end if
	IF (MaybeTitle<>"") then
	    sqlStr = sqlStr + " ,title='"&MaybeTitle&"'"
	end if

	sqlStr = sqlStr + " ,diffkey="&Maydiffkey&""
	sqlStr = sqlStr + " ,shopdiv='"&shopdiv&"'"
	sqlStr = sqlStr + " from (select sum(totalsuplycash) as totalsum, sum(totalsellcash) as totalsellcash, "
	sqlStr = sqlStr + "			sum(totalbuycash) as totalbuycash, sum(totalsuplycash) as totalsuplycash, "
	sqlStr = sqlStr + "			sum(totalorgsellcash) as totalorgsellcash "
	sqlStr = sqlStr + " 		from [db_shop].[dbo].tbl_fran_meachuljungsan_submaster"
	sqlStr = sqlStr + " 		where masteridx=" + CStr(iid)
	sqlStr = sqlStr + " 	) as T "
	sqlStr = sqlStr + " where idx=" + CStr(iid)

	rsget.Open sqlStr, dbget, 1

'//b2c 매출 작성
elseif mode="b2cmaechul" then

	if shopid = "" or check = "" or b2ccharge = "" then
	    response.write "<script type='text/javascript'>"
		response.write "	alert('구분값이 없습니다');"
		response.write "	window.close();"
		response.write "</script>"
		response.end	:	dbget.close()
	end if

	dbget.beginTrans

	'//마스터 등록
	sqlStr = " select * from [db_shop].[dbo].tbl_fran_meachuljungsan_master where 1=0"
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew

	rsget("shopid") = shopid
	rsget("title") = shopid & " 작성중"
	rsget("totalsum") = 0
	rsget("divcode") = "TC"
	rsget("etcstr") = ""
	rsget("reguserid") = adminuserid
	rsget("regusername") = adminusername

	rsget.update
		iid = rsget("idx")
	rsget.close

	'//서브 마스터 등록
	sqlStr = " insert into [db_shop].[dbo].tbl_fran_meachuljungsan_submaster" + vbcrlf
	sqlStr = sqlStr & " (masteridx, linkidx, shopid, code01, code02" + vbcrlf
	sqlStr = sqlStr & " , execdate, totalcount" + vbcrlf
	sqlStr = sqlStr & " , totalsellcash" + vbcrlf
	sqlStr = sqlStr & " , totalbuycash" + vbcrlf
	sqlStr = sqlStr & " , totalsuplycash)" + vbcrlf
	sqlStr = sqlStr & " 	select" + vbcrlf
	sqlStr = sqlStr & " 	"& CStr(iid) & ", 0, m.shopid, convert(varchar(10),m.shopregdate,121),''" + vbcrlf
	sqlStr = sqlStr & " 	, convert(varchar(10),m.shopregdate,121), sum(d.itemno) as totitemcnt" + vbcrlf
	sqlStr = sqlStr & " 	,isnull(sum(d.realsellprice*d.itemno),0) as totsum" + vbcrlf
	sqlStr = sqlStr & " 	,isnull(sum(d.suplyprice*d.itemno),0)" + vbcrlf		'//매입가,매장공급가 꺼꾸로되어있슴.헷갈리지말것
	sqlStr = sqlStr & " 	, isnull(sum(d.realsellprice*d.itemno),0)-(isnull(sum(d.realsellprice*d.itemno),0)*"&b2ccharge&"/100)" + vbcrlf		'/isnull(sum(d.shopbuyprice*d.itemno),0)
	sqlStr = sqlStr & "		from [db_shop].[dbo].tbl_shopjumun_master m " + vbcrlf
    sqlStr = sqlStr & "		join [db_shop].[dbo].tbl_shopjumun_detail d " + vbcrlf
	sqlStr = sqlStr & "			on m.idx=d.masteridx" + vbcrlf
    sqlStr = sqlStr & "			and m.cancelyn='N' and d.cancelyn='N'" + vbcrlf
    sqlStr = sqlStr & " 	left join [db_shop].[dbo].tbl_fran_meachuljungsan_submaster sm" + vbcrlf
    sqlStr = sqlStr & " 		on m.shopid=sm.shopid" + vbcrlf
    sqlStr = sqlStr & " 		and convert(varchar(10),m.shopregdate,121)=sm.code01" + vbcrlf
    sqlStr = sqlStr & "     where convert(varchar(10),m.shopregdate,121) in (" & check & ")" + vbcrlf
    sqlStr = sqlStr & " 	and m.shopid='" & shopid & "'" + vbcrlf
    sqlStr = sqlStr & " 	and sm.idx is null" + vbcrlf		'//중복체크
    sqlStr = sqlStr & " 	group by" + vbcrlf
    sqlStr = sqlStr & " 		m.shopid, convert(varchar(10),m.shopregdate,121)"

	'response.write sqlStr & "<Br>"
	dbget.execute sqlStr

	'//서브 디테일 등록(주문마스터)
	sqlStr = " insert into [db_shop].[dbo].tbl_fran_meachuljungsan_subdetail" + vbcrlf
	sqlStr = sqlStr & " (masteridx, topmasteridx, linkbaljucode, linkmastercode, linkdetailidx" + vbcrlf
	sqlStr = sqlStr & " , itemno, sellcash" + vbcrlf
	sqlStr = sqlStr & " , suplycash" + vbcrlf
	sqlStr = sqlStr & " , buycash, orgsellcash)" + vbcrlf
	sqlStr = sqlStr & " 	select" + vbcrlf
	sqlStr = sqlStr & " 	sm.idx, " + CStr(iid) + ",'SB2C', m.orderno, m.idx" + vbcrlf
	sqlStr = sqlStr & " 	,1, isnull(sum(d.realsellprice*d.itemno),0)" + vbcrlf
	sqlStr = sqlStr & " 	, isnull(sum(d.realsellprice*d.itemno),0)-(isnull(sum(d.realsellprice*d.itemno),0)*"&b2ccharge&"/100)" + vbcrlf		'/isnull(sum(d.shopbuyprice*d.itemno),0)
	sqlStr = sqlStr & " 	,isnull(sum(d.suplyprice*d.itemno),0), sum(d.sellprice*d.itemno)" + vbcrlf
	sqlStr = sqlStr & "		from [db_shop].[dbo].tbl_shopjumun_master m " + vbcrlf
    sqlStr = sqlStr & "		join [db_shop].[dbo].tbl_shopjumun_detail d " + vbcrlf
	sqlStr = sqlStr & "			on m.idx=d.masteridx" + vbcrlf
    sqlStr = sqlStr & "			and m.cancelyn='N' and d.cancelyn='N'" + vbcrlf
    sqlStr = sqlStr & " 	join [db_shop].[dbo].tbl_fran_meachuljungsan_submaster sm" + vbcrlf
    sqlStr = sqlStr & " 		on m.shopid=sm.shopid" + vbcrlf
    sqlStr = sqlStr & " 		and convert(varchar(10),m.shopregdate,121)=sm.code01" + vbcrlf
    sqlStr = sqlStr & "		left join [db_shop].[dbo].tbl_fran_meachuljungsan_subdetail sd" + vbcrlf
    sqlStr = sqlStr & "			on m.idx = sd.linkdetailidx" + vbcrlf
    sqlStr = sqlStr & "			and linkbaljucode = 'SB2C'" + vbcrlf
	sqlStr = sqlStr & " 	where sm.masteridx=" + CStr(iid)
	sqlStr = sqlStr & " 	and sd.linkdetailidx is null" + vbcrlf		'//중복체크
    sqlStr = sqlStr & " 	group by" + vbcrlf
    sqlStr = sqlStr & " 		sm.idx, m.orderno, m.idx"

	'//서브 디테일 등록(주문디테일 : 양이 많아서 뺌)
	'sqlStr = " insert into [db_shop].[dbo].tbl_fran_meachuljungsan_subdetail" + vbcrlf
	'sqlStr = sqlStr & " (masteridx, topmasteridx, linkbaljucode, linkmastercode, linkdetailidx" + vbcrlf
	'sqlStr = sqlStr & " ,itemgubun, itemid, itemoption, itemname, itemoptionname" + vbcrlf
	'sqlStr = sqlStr & " ,makerid, itemno, sellcash, suplycash, buycash, orgsellcash)" + vbcrlf
	'sqlStr = sqlStr & " 	select" + vbcrlf
	'sqlStr = sqlStr & " 	sm.idx, " + CStr(iid) + ",'SB2C', d.orderno, d.idx" + vbcrlf
	'sqlStr = sqlStr & " 	,d.itemgubun, d.itemid, d.itemoption, d.itemname, d.itemoptionname" + vbcrlf
	'sqlStr = sqlStr & " 	,d.makerid, d.itemno, d.realsellprice, d.shopbuyprice, d.suplyprice, d.sellprice" + vbcrlf
	'sqlStr = sqlStr & "		from [db_shop].[dbo].tbl_shopjumun_master m " + vbcrlf
    'sqlStr = sqlStr & "		join [db_shop].[dbo].tbl_shopjumun_detail d " + vbcrlf
	'sqlStr = sqlStr & "			on m.idx=d.masteridx" + vbcrlf
    'sqlStr = sqlStr & "			and m.cancelyn='N' and d.cancelyn='N'" + vbcrlf
    'sqlStr = sqlStr & " 	join [db_shop].[dbo].tbl_fran_meachuljungsan_submaster sm" + vbcrlf
    'sqlStr = sqlStr & " 		on m.shopid=sm.shopid" + vbcrlf
    'sqlStr = sqlStr & " 		and convert(varchar(10),m.shopregdate,121)=sm.code01" + vbcrlf
	'sqlStr = sqlStr & " 	where sm.masteridx=" + CStr(iid)

	'response.write sqlStr & "<Br>"
	dbget.execute sqlStr

	'//서브 마스터 업데이트
	sqlStr = " update [db_shop].[dbo].tbl_fran_meachuljungsan_submaster" + vbcrlf
	sqlStr = sqlStr + " set totalsellcash=T.totalsellcash" + vbcrlf
	sqlStr = sqlStr + " ,totalbuycash=T.totalbuycash" + vbcrlf
	sqlStr = sqlStr + " ,totalsuplycash=T.totalsuplycash" + vbcrlf
	sqlStr = sqlStr + " ,totalorgsellcash=T.totalorgsellcash" + vbcrlf
	sqlStr = sqlStr + " from (" + vbcrlf
	sqlStr = sqlStr + " 	select" + vbcrlf
	sqlStr = sqlStr + " 	masteridx, sum(sellcash*itemno) as totalsellcash" + vbcrlf
	sqlStr = sqlStr + "		,sum(buycash*itemno) as totalbuycash, sum(suplycash*itemno) as totalsuplycash" + vbcrlf
	sqlStr = sqlStr + "		,sum(orgsellcash*itemno) as totalorgsellcash " + vbcrlf
	sqlStr = sqlStr + " 	from [db_shop].[dbo].tbl_fran_meachuljungsan_subdetail" + vbcrlf
	sqlStr = sqlStr + " 	where topmasteridx=" + CStr(iid) + vbcrlf
	sqlStr = sqlStr + "		group by masteridx" + vbcrlf
	sqlStr = sqlStr + " ) as T " + vbcrlf
	sqlStr = sqlStr + " where [db_shop].[dbo].tbl_fran_meachuljungsan_submaster.idx=T.masteridx" + vbcrlf
	sqlStr = sqlStr + " and [db_shop].[dbo].tbl_fran_meachuljungsan_submaster.masteridx=" + CStr(iid)

	'response.write sqlStr & "<Br>"
	dbget.execute sqlStr

    MaybeYYYYMM = fnGetMayYYYYMM(iid)		'//정산년월
    shopname    = fnGetShopName(shopid, shopdiv)
    Maydiffkey  = fnGetMayDiffKey(shopid, MaybeYYYYMM)		'//차수

    IF (MaybeYYYYMM<>"") and (shopname<>"") then
        maybetitle = shopname + " " + MaybeYYYYMM + " B2C매출"
    End IF

	'//마스터 업데이트
	sqlStr = " update [db_shop].[dbo].tbl_fran_meachuljungsan_master" + vbcrlf
	sqlStr = sqlStr + " set totalsum=T.totalsum" + vbcrlf
	sqlStr = sqlStr + " ,totalsellcash=T.totalsellcash" + vbcrlf
	sqlStr = sqlStr + " ,totalbuycash=T.totalbuycash" + vbcrlf
	sqlStr = sqlStr + " ,totalsuplycash=T.totalsuplycash" + vbcrlf
	sqlStr = sqlStr + " ,totalorgsellcash=T.totalorgsellcash" + vbcrlf

	IF (MaybeYYYYMM<>"") then
	    sqlStr = sqlStr + " ,yyyymm='"& MaybeYYYYMM &"'" + vbcrlf
	end if
	IF (MaybeTitle<>"") then
	    sqlStr = sqlStr + " ,title='"&MaybeTitle&"'" + vbcrlf
	end if

	sqlStr = sqlStr + " ,diffkey="&Maydiffkey&"" + vbcrlf
	sqlStr = sqlStr + " ,shopdiv='"&shopdiv&"'" + vbcrlf
	sqlStr = sqlStr + " from (" + vbcrlf
	sqlStr = sqlStr + " 	select" + vbcrlf
	sqlStr = sqlStr + " 	sum(totalsuplycash) as totalsum, sum(totalsellcash) as totalsellcash" + vbcrlf
	sqlStr = sqlStr + "		,sum(totalbuycash) as totalbuycash, sum(totalsuplycash) as totalsuplycash" + vbcrlf
	sqlStr = sqlStr + "		,sum(totalorgsellcash) as totalorgsellcash " + vbcrlf
	sqlStr = sqlStr + " 	from [db_shop].[dbo].tbl_fran_meachuljungsan_submaster" + vbcrlf
	sqlStr = sqlStr + " 	where masteridx=" + CStr(iid) + vbcrlf
	sqlStr = sqlStr + " ) as T " + vbcrlf
	sqlStr = sqlStr + " where idx=" + CStr(iid)

	'response.write sqlStr & "<Br>"
	dbget.execute sqlStr

	If Err.Number = 0 Then
	    dbget.CommitTrans

	    response.write "<script type='text/javascript'>"
		response.write "	alert('저장 되었습니다.');"
		response.write "	opener.popMasterEdit('"& iid &"');"
		response.write "	opener.location.reload();"
		response.write "	window.close();"
		response.write "</script>"
		response.end	:	dbget.close()
	Else
	    dbget.RollBackTrans

	    response.write "<script type='text/javascript'>"
		response.write "	alert('잘못된 내역 입니다("&Err.Description&")');"
		response.write "	window.close();"
		response.write "</script>"
		response.end	:	dbget.close()

	End If

elseif mode="witsksell_old" then

'''이전  OFF 정산 테이블
	'' insert master
	sqlStr = " select * from [db_shop].[dbo].tbl_fran_meachuljungsan_master where 1=0"
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew

	rsget("shopid") = shopid
	rsget("title") = shopid & " 작성중"
	rsget("totalsum") = 0
	rsget("divcode") = "WS"
	rsget("etcstr") = ""
	rsget("reguserid") = session("ssBctId")
	rsget("regusername") = session("ssBctCname")

	rsget.update
	iid = rsget("idx")
	rsget.close

	'' insert sub master
	sqlStr = " insert into [db_shop].[dbo].tbl_fran_meachuljungsan_submaster "
	sqlStr = sqlStr + " (masteridx,linkidx,shopid,code01,code02,execdate, "
	sqlStr = sqlStr + " totalcount,totalsellcash,totalbuycash,totalsuplycash)"
	sqlStr = sqlStr + " select " + CStr(iid) + ", m.idx, m.shopid, convert(varchar(7),m.yyyymm),"
	sqlStr = sqlStr + " m.jungsanid, convert(varchar(7),m.yyyymm) + '-01', m.totitemcnt,"
	sqlStr = sqlStr + " m.totsum, m.realjungsansum, 0"
	sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_jungsanmaster m"
	sqlStr = sqlStr + " where idx in  (" + check + ")"

	rsget.Open sqlStr, dbget, 1

	'' insert sub detail : sellprice, orgsellprice
	sqlStr = " insert into [db_shop].[dbo].tbl_fran_meachuljungsan_subdetail "
	sqlStr = sqlStr + " (masteridx, topmasteridx, linkbaljucode, linkmastercode, linkdetailidx,"
	sqlStr = sqlStr + " itemgubun, itemid, itemoption, itemname, itemoptionname,"
	sqlStr = sqlStr + " makerid, itemno, sellcash, suplycash, buycash, orgsellcash)"
	sqlStr = sqlStr + " select m.idx," + CStr(iid) + ",d.jungsangubun, d.orderno, d.idx,"
	sqlStr = sqlStr + " d.itemgubun,d.itemid,d.itemoption,d.itemname,d.itemoptionname,"
	sqlStr = sqlStr + " d.makerid,d.itemno,d.realsellprice,0,d.suplyprice,d.sellprice"
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " [db_shop].[dbo].tbl_fran_meachuljungsan_submaster m,"
	sqlStr = sqlStr + " [db_shop].[dbo].tbl_shop_jungsandetail d"
	sqlStr = sqlStr + " where m.linkidx=d.masteridx"
	sqlStr = sqlStr + " and m.masteridx=" + CStr(iid)

	rsget.Open sqlStr, dbget, 1

	'' update Detail shopsuplyprice
	'' 현재 OFF 상품 가격 기준 정산.
	''
	sqlStr = " update [db_shop].[dbo].tbl_fran_meachuljungsan_subdetail"
	sqlStr = sqlStr + " set suplycash=IsNULL(T.shopbuyprice,0)"
	sqlStr = sqlStr + " from ("
	sqlStr = sqlStr + " 	select distinct d.idx, IsNULL(s.defaultsuplymargin,35) as defaultsuplymargin"
	sqlStr = sqlStr + " ,	( case "
	sqlStr = sqlStr + "			when (i.shopbuyprice=0) and (j.discountprice=0) then convert(int,j.sellprice - j.sellprice*IsNULL(s.defaultsuplymargin,35)/100)"
	sqlStr = sqlStr + "  		when (i.shopbuyprice=0) and (j.discountprice<>0) then convert(int,j.discountprice - j.discountprice*IsNULL(s.defaultsuplymargin,35)/100)"
	sqlStr = sqlStr + "    		else i.shopbuyprice "
	sqlStr = sqlStr + "    		end ) as shopbuyprice "
	sqlStr = sqlStr + " 	from [db_shop].[dbo].tbl_fran_meachuljungsan_subdetail d"
	sqlStr = sqlStr + " 		left join [db_shop].[dbo].tbl_shop_designer s "
	sqlStr = sqlStr + " 			on d.makerid=s.makerid and s.shopid='" + shopid + "'"
	sqlStr = sqlStr + " 		left join [db_shop].[dbo].tbl_shopjumun_detail j"
	sqlStr = sqlStr + " 			on d.linkmastercode=j.orderno"
	sqlStr = sqlStr + " 			and d.itemgubun=j.itemgubun"
	sqlStr = sqlStr + " 			and d.itemid=j.itemid"
	sqlStr = sqlStr + " 			and d.itemoption=j.itemoption"
	sqlStr = sqlStr + " 		left join [db_shop].[dbo].tbl_shop_item i"
	sqlStr = sqlStr + " 			on d.itemgubun=i.itemgubun"
	sqlStr = sqlStr + " 			and d.itemid=i.shopitemid"
	sqlStr = sqlStr + " 			and d.itemoption=i.itemoption"
	sqlStr = sqlStr + " 	where d.topmasteridx=" + CStr(iid)
	sqlStr = sqlStr + " ) as T"
	sqlStr = sqlStr + " where [db_shop].[dbo].tbl_fran_meachuljungsan_subdetail.idx=T.idx"

	rsget.Open sqlStr, dbget, 1

	'' update Sub master
	sqlStr = " update [db_shop].[dbo].tbl_fran_meachuljungsan_submaster"
	sqlStr = sqlStr + " set totalsellcash=T.totalsellcash"
	sqlStr = sqlStr + " ,totalbuycash=T.totalbuycash"
	sqlStr = sqlStr + " ,totalsuplycash=T.totalsuplycash"
	sqlStr = sqlStr + " ,totalorgsellcash=T.totalorgsellcash"
	sqlStr = sqlStr + " from (select masteridx, sum(sellcash*itemno) as totalsellcash, "
	sqlStr = sqlStr + "			sum(buycash*itemno) as totalbuycash, sum(suplycash*itemno) as totalsuplycash,"
	sqlStr = sqlStr + "			sum(orgsellcash*itemno) as totalorgsellcash "
	sqlStr = sqlStr + " 		from [db_shop].[dbo].tbl_fran_meachuljungsan_subdetail"
	sqlStr = sqlStr + " 		where topmasteridx=" + CStr(iid)
	sqlStr = sqlStr + "		    group by masteridx"
	sqlStr = sqlStr + " ) as T "
	sqlStr = sqlStr + " where [db_shop].[dbo].tbl_fran_meachuljungsan_submaster.idx=T.masteridx"
	sqlStr = sqlStr + " and [db_shop].[dbo].tbl_fran_meachuljungsan_submaster.masteridx=" + CStr(iid)

	rsget.Open sqlStr, dbget, 1

	'' update master
	sqlStr = " update [db_shop].[dbo].tbl_fran_meachuljungsan_master"
	sqlStr = sqlStr + " set totalsum=T.totalsum"
	sqlStr = sqlStr + " ,totalsellcash=T.totalsellcash"
	sqlStr = sqlStr + " ,totalbuycash=T.totalbuycash"
	sqlStr = sqlStr + " ,totalsuplycash=T.totalsuplycash"
	sqlStr = sqlStr + " ,totalorgsellcash=T.totalorgsellcash"
	sqlStr = sqlStr + " from (select sum(totalsuplycash) as totalsum, sum(totalsellcash) as totalsellcash, "
	sqlStr = sqlStr + "			sum(totalbuycash) as totalbuycash, sum(totalsuplycash) as totalsuplycash, "
	sqlStr = sqlStr + "			sum(totalorgsellcash) as totalorgsellcash "
	sqlStr = sqlStr + " 		from [db_shop].[dbo].tbl_fran_meachuljungsan_submaster"
	sqlStr = sqlStr + " 		where masteridx=" + CStr(iid)
	sqlStr = sqlStr + " 	) as T "
	sqlStr = sqlStr + " where idx=" + CStr(iid)

	rsget.Open sqlStr, dbget, 1

elseif mode="addmaster" then
	if request("etcstr") <> "" then
		if checkNotValidHTML(request("etcstr")) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
		response.write "</script>"
		dbget.close()	:	response.End
		end if
	end if

	sqlStr = " select * from [db_shop].[dbo].tbl_fran_meachuljungsan_master where 1=0"
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew

	rsget("shopid") = requestCheckVar(request("shopid"),32)
	rsget("title") = requestCheckVar(html2db(request("title")),64)
	rsget("totalsum") = requestCheckVar(request("totalsum"),20)
	rsget("totalsellcash") = requestCheckVar(request("totalsuplycash"),20)
	rsget("totalbuycash") = requestCheckVar(request("totalbuycash"),20)
	rsget("totalsuplycash") = requestCheckVar(request("totalsuplycash"),20)
	rsget("divcode") = requestCheckVar(request("divcode"),3)
	rsget("etcstr") = html2db(request("etcstr"))
	rsget("reguserid") = session("ssBctId")
	rsget("regusername") = session("ssBctCname")

	if request("taxdate")<>"" then
		rsget("taxdate") = requestCheckVar(request("taxdate"),30)
	end if

	if request("ipkumdate")<>"" then
		rsget("ipkumdate") = requestCheckVar(request("ipkumdate"),30)
	end if

	rsget.update
	iid = rsget("idx")
	rsget.close

elseif mode="modimaster" then
	if request("etcstr") <> "" then
		if checkNotValidHTML(request("etcstr")) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
		response.write "</script>"
		dbget.close()	:	response.End
		end if
	end if

	sqlStr = " update [db_shop].[dbo].tbl_fran_meachuljungsan_master" + VbCrlf
	sqlStr = sqlStr + " set title='" + requestCheckVar(html2db(request("title")),64) + "'" + VbCrlf
	sqlStr = sqlStr + " ,totalsum=" + requestCheckVar(request("totalsum"),20) + ""
	sqlStr = sqlStr + " ,statecd='" + requestCheckVar(request("statecd"),1) + "'"

	if request("taxdate")<>"" then
		sqlStr = sqlStr + " ,taxdate='" + requestCheckVar(request("taxdate"),30) + "'"
	else
		sqlStr = sqlStr + " ,taxdate=NULL"
	end if

	if request("ipkumdate")<>"" then
		sqlStr = sqlStr + " ,ipkumdate='" + requestCheckVar(request("ipkumdate"),30) + "'"
	else
		sqlStr = sqlStr + " ,ipkumdate=NULL"
	end if

	sqlStr = sqlStr + " ,etcstr='" + html2db(request("etcstr")) + "'"
	sqlStr = sqlStr + " ,finishuserid='" + session("ssBctId") + "'"
	sqlStr = sqlStr + " ,finishusername='" + session("ssBctCname") + "'"
	sqlStr = sqlStr + " where idx=" + CStr(idx)  + VbCrlf

	rsget.Open sqlStr, dbget, 1

elseif mode="delmaster" then
	''현재상태 0 수정중인 경우만 삭제 가능
	cnt=0
	sqlStr = " select count(idx) as cnt from [db_shop].[dbo].tbl_fran_meachuljungsan_master"
	sqlStr = sqlStr + " where idx=" + CStr(idx)  + VbCrlf

	rsget.Open sqlStr, dbget, 1
		cnt = rsget("cnt")
	rsget.Close

	if cnt<1 then
		response.write "<script type='text/javascript'>alert('수정중 상태에서만 삭제 가능합니다.');</script>"
		response.write "<script type='text/javascript'>window.close();</script>"
		dbget.close()	:	response.End
	end if

	sqlStr = " delete from [db_shop].[dbo].tbl_fran_meachuljungsan_subdetail" + VbCrlf
	sqlStr = sqlStr + " where topmasteridx=" + CStr(idx)

	rsget.Open sqlStr, dbget, 1

	sqlStr = " delete from [db_shop].[dbo].tbl_fran_meachuljungsan_submaster" + VbCrlf
	sqlStr = sqlStr + " where masteridx=" + CStr(idx)  + VbCrlf

	rsget.Open sqlStr, dbget, 1

	sqlStr = " delete from [db_shop].[dbo].tbl_fran_meachuljungsan_master" + VbCrlf
	sqlStr = sqlStr + " where idx=" + CStr(idx)  + VbCrlf

	rsget.Open sqlStr, dbget, 1

elseif mode="modidetail" then

	ckidx = request.form("ckidx") + ","
	itemnoarr= request.form("itemnoarr")
	suplycasharr = request.form("suplycasharr")
	orgsellcasharr = request.form("orgsellcasharr")
	sellcasharr = request.form("sellcasharr")
	buycasharr = request.form("buycasharr")

	ckidx = split(ckidx,",")
	suplycasharr = split(suplycasharr,",")
	orgsellcasharr = split(orgsellcasharr,",")
	sellcasharr = split(sellcasharr,",")
	buycasharr = split(buycasharr,",")
	itemnoarr = split(itemnoarr,",")

	for i=0 to Ubound(ckidx)
		if trim(ckidx(i))<>"" then
			sqlStr = " update [db_shop].[dbo].tbl_fran_meachuljungsan_subdetail" + VbCrlf
			sqlStr = sqlStr + " set orgsellcash=" + CStr(requestCheckVar(orgsellcasharr(i),20))  + VbCrlf
			sqlStr = sqlStr + " , sellcash=" + CStr(requestCheckVar(sellcasharr(i),20))  + VbCrlf
			sqlStr = sqlStr + " , buycash=" + CStr(requestCheckVar(buycasharr(i),20))  + VbCrlf
			sqlStr = sqlStr + " , suplycash=" + CStr(requestCheckVar(suplycasharr(i),20))  + VbCrlf
			sqlStr = sqlStr + " ,itemno=" + CStr(requestCheckVar(itemnoarr(i),10))  + VbCrlf
			sqlStr = sqlStr + " where idx=" + requestCheckVar(trim(ckidx(i)),10)

			rsget.Open sqlStr, dbget, 1
		end if
	next

	'' update Sub master
	sqlStr = " update [db_shop].[dbo].tbl_fran_meachuljungsan_submaster"
	sqlStr = sqlStr + " set totalsellcash=IsNULL(T.totalsellcash,0)"
	sqlStr = sqlStr + " ,totalbuycash=IsNULL(T.totalbuycash,0)"
	sqlStr = sqlStr + " ,totalsuplycash=IsNULL(T.totalsuplycash,0)"
	sqlStr = sqlStr + " ,totalorgsellcash=IsNULL(T.totalorgsellcash,0)"
	sqlStr = sqlStr + " from (select masteridx, sum(sellcash*itemno) as totalsellcash, "
	sqlStr = sqlStr + "			sum(buycash*itemno) as totalbuycash, sum(suplycash*itemno) as totalsuplycash,"
	sqlStr = sqlStr + "			sum(orgsellcash*itemno) as totalorgsellcash "
	sqlStr = sqlStr + " 		from [db_shop].[dbo].tbl_fran_meachuljungsan_subdetail"
	sqlStr = sqlStr + " 		where topmasteridx=" + CStr(topidx)
	sqlStr = sqlStr + "		    group by masteridx"
	sqlStr = sqlStr + " ) as T "
	sqlStr = sqlStr + " where [db_shop].[dbo].tbl_fran_meachuljungsan_submaster.idx=T.masteridx"
	sqlStr = sqlStr + " and [db_shop].[dbo].tbl_fran_meachuljungsan_submaster.masteridx=" + CStr(topidx)

	rsget.Open sqlStr, dbget, 1

	'' update master
	sqlStr = " update [db_shop].[dbo].tbl_fran_meachuljungsan_master"
	sqlStr = sqlStr + " set totalsum=IsNULL(T.totalsum,0)"
	sqlStr = sqlStr + " ,totalsellcash=IsNULL(T.totalsellcash,0)"
	sqlStr = sqlStr + " ,totalbuycash=IsNULL(T.totalbuycash,0)"
	sqlStr = sqlStr + " ,totalsuplycash=IsNULL(T.totalsuplycash,0)"
	sqlStr = sqlStr + " ,totalorgsellcash=IsNULL(T.totalorgsellcash,0)"
	sqlStr = sqlStr + " from (select sum(totalsuplycash) as totalsum, sum(totalsellcash) as totalsellcash, "
	sqlStr = sqlStr + "			sum(totalbuycash) as totalbuycash, sum(totalsuplycash) as totalsuplycash, "
	sqlStr = sqlStr + "			sum(totalorgsellcash) as totalorgsellcash "
	sqlStr = sqlStr + " 		from [db_shop].[dbo].tbl_fran_meachuljungsan_submaster"
	sqlStr = sqlStr + " 		where masteridx=" + CStr(topidx)
	sqlStr = sqlStr + " 	) as T "
	sqlStr = sqlStr + " where idx=" + CStr(topidx)

	rsget.Open sqlStr, dbget, 1

elseif mode="deldetail" then

	ckidx = trim(request.form("ckidx") + ",")

	if Right(ckidx,1)="," then
		ckidx = Left(ckidx,Len(ckidx)-1)
	end if

	sqlStr = " delete from [db_shop].[dbo].tbl_fran_meachuljungsan_subdetail" + VbCrlf
	sqlStr = sqlStr + " where idx in (" + trim(ckidx) + ")"
	''response.write sqlStr
	''dbget.close()	:	response.End

	rsget.Open sqlStr, dbget, 1

	'' update Sub master
	sqlStr = " update [db_shop].[dbo].tbl_fran_meachuljungsan_submaster"
	sqlStr = sqlStr + " set totalsellcash=T.totalsellcash"
	sqlStr = sqlStr + " ,totalbuycash=T.totalbuycash"
	sqlStr = sqlStr + " ,totalsuplycash=T.totalsuplycash"
	sqlStr = sqlStr + " ,totalorgsellcash=T.totalorgsellcash"
	sqlStr = sqlStr + " from (select masteridx, sum(sellcash*itemno) as totalsellcash, "
	sqlStr = sqlStr + "			sum(buycash*itemno) as totalbuycash, sum(suplycash*itemno) as totalsuplycash,"
	sqlStr = sqlStr + "			sum(orgsellcash*itemno) as totalorgsellcash "
	sqlStr = sqlStr + " 		from [db_shop].[dbo].tbl_fran_meachuljungsan_subdetail"
	sqlStr = sqlStr + " 		where topmasteridx=" + CStr(topidx)
	sqlStr = sqlStr + "		    group by masteridx"
	sqlStr = sqlStr + " ) as T "
	sqlStr = sqlStr + " where [db_shop].[dbo].tbl_fran_meachuljungsan_submaster.idx=T.masteridx"
	sqlStr = sqlStr + " and [db_shop].[dbo].tbl_fran_meachuljungsan_submaster.masteridx=" + CStr(topidx)

	rsget.Open sqlStr, dbget, 1

	'' update master
	sqlStr = " update [db_shop].[dbo].tbl_fran_meachuljungsan_master"
	sqlStr = sqlStr + " set totalsum=T.totalsum"
	sqlStr = sqlStr + " ,totalsellcash=T.totalsellcash"
	sqlStr = sqlStr + " ,totalbuycash=T.totalbuycash"
	sqlStr = sqlStr + " ,totalsuplycash=T.totalsuplycash"
	sqlStr = sqlStr + " ,totalorgsellcash=T.totalorgsellcash"
	sqlStr = sqlStr + " from (select sum(totalsuplycash) as totalsum, sum(totalsellcash) as totalsellcash, "
	sqlStr = sqlStr + "			sum(totalbuycash) as totalbuycash, sum(totalsuplycash) as totalsuplycash, "
	sqlStr = sqlStr + "			sum(totalorgsellcash) as totalorgsellcash "
	sqlStr = sqlStr + " 		from [db_shop].[dbo].tbl_fran_meachuljungsan_submaster"
	sqlStr = sqlStr + " 		where masteridx=" + CStr(topidx)
	sqlStr = sqlStr + " 	) as T "
	sqlStr = sqlStr + " where idx=" + CStr(topidx)

	rsget.Open sqlStr, dbget, 1

elseif mode="etcsubadd" then

	sqlStr = " select * from [db_shop].[dbo].tbl_fran_meachuljungsan_submaster where 1=0"
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew

	rsget("masteridx") = topidx
	rsget("linkidx") = 0
	rsget("shopid") = shopid
	rsget("code01") = yyyy + "-" + mm
	rsget("code02") = makerid
	rsget("execdate") = yyyy + "-" + mm + "-01"
	rsget("totalcount") = 0
	rsget("totalsellcash") = 0
	rsget("totalbuycash") = 0
	rsget("totalsuplycash") = 0
	rsget("totalorgsellcash") = 0
	rsget.update
		iid = rsget("idx")
	rsget.close

elseif mode="etcsubdetailadd" then

	sqlStr = " select * from [db_shop].[dbo].tbl_fran_meachuljungsan_subdetail where 1=0"
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew

	rsget("masteridx") = idx
	rsget("topmasteridx") = topidx
	rsget("linkbaljucode") = requestCheckVar(request("linkbaljucode"),16)
	rsget("linkmastercode") = "0"
	rsget("linkdetailidx") = 0
	rsget("itemgubun") = requestCheckVar(request("itemgubun"),2)
	rsget("itemid") = requestCheckVar(request("itemid"),10)
	rsget("itemoption") = requestCheckVar(request("itemoption"),4)
	rsget("itemname") = requestCheckVar(html2Db(request("itemname")),124)
	rsget("itemoptionname") = requestCheckVar(html2Db(request("itemoptionname")),96)
	rsget("makerid") = requestCheckVar(request("makerid"),32)
	rsget("itemno") = requestCheckVar(request("itemno"),10)
	rsget("sellcash") = requestCheckVar(request("sellcash"),20)
	rsget("suplycash") = requestCheckVar(request("suplycash"),20)
	rsget("buycash") = requestCheckVar(request("buycash"),20)
	rsget("orgsellcash") = requestCheckVar(request("orgsellcash"),20)

	rsget.update
	iid = rsget("idx")
	rsget.close

	'' update Sub master
	sqlStr = " update [db_shop].[dbo].tbl_fran_meachuljungsan_submaster"
	sqlStr = sqlStr + " set totalsellcash=T.totalsellcash"
	sqlStr = sqlStr + " ,totalbuycash=T.totalbuycash"
	sqlStr = sqlStr + " ,totalsuplycash=T.totalsuplycash"
	sqlStr = sqlStr + " ,totalorgsellcash=T.totalorgsellcash"
	sqlStr = sqlStr + " from (select masteridx, sum(sellcash*itemno) as totalsellcash, "
	sqlStr = sqlStr + "			sum(buycash*itemno) as totalbuycash, sum(suplycash*itemno) as totalsuplycash,"
	sqlStr = sqlStr + "			sum(orgsellcash*itemno) as totalorgsellcash "
	sqlStr = sqlStr + " 		from [db_shop].[dbo].tbl_fran_meachuljungsan_subdetail"
	sqlStr = sqlStr + " 		where topmasteridx=" + CStr(topidx)
	sqlStr = sqlStr + "		    group by masteridx"
	sqlStr = sqlStr + " ) as T "
	sqlStr = sqlStr + " where [db_shop].[dbo].tbl_fran_meachuljungsan_submaster.idx=T.masteridx"
	sqlStr = sqlStr + " and [db_shop].[dbo].tbl_fran_meachuljungsan_submaster.masteridx=" + CStr(topidx)

	rsget.Open sqlStr, dbget, 1

	'' update master
	sqlStr = " update [db_shop].[dbo].tbl_fran_meachuljungsan_master"
	sqlStr = sqlStr + " set totalsum=T.totalsum"
	sqlStr = sqlStr + " ,totalsellcash=T.totalsellcash"
	sqlStr = sqlStr + " ,totalbuycash=T.totalbuycash"
	sqlStr = sqlStr + " ,totalsuplycash=T.totalsuplycash"
	sqlStr = sqlStr + " ,totalorgsellcash=T.totalorgsellcash"
	sqlStr = sqlStr + " from (select sum(totalsuplycash) as totalsum, sum(totalsellcash) as totalsellcash, "
	sqlStr = sqlStr + "			sum(totalbuycash) as totalbuycash, sum(totalsuplycash) as totalsuplycash, "
	sqlStr = sqlStr + "			sum(totalorgsellcash) as totalorgsellcash "
	sqlStr = sqlStr + " 		from [db_shop].[dbo].tbl_fran_meachuljungsan_submaster"
	sqlStr = sqlStr + " 		where masteridx=" + CStr(topidx)
	sqlStr = sqlStr + " 	) as T "
	sqlStr = sqlStr + " where idx=" + CStr(topidx)

	rsget.Open sqlStr, dbget, 1

end if

function fnGetMayYYYYMM(iid)
    ''정산년월 계산.
    sqlStr = " select top 1 convert(varchar(7),F.execdate,21) as MaybeYYYYMM, count(*) cnt"
    sqlStr = sqlStr + " from [db_shop].[dbo].tbl_fran_meachuljungsan_submaster F"
    sqlStr = sqlStr + " where masteridx=" + CStr(iid)
    sqlStr = sqlStr + " group by convert(varchar(7),F.execdate,21)"
    sqlStr = sqlStr + " order by cnt desc"

	'response.write sqlStr &"<Br>"
    rsget.Open sqlStr, dbget, 1
    If Not rsget.EOF then
        fnGetMayYYYYMM = rsget("MaybeYYYYMM")
    end if
    rsget.close
end function

function fnGetShopName(shopid, byREF shopdiv)
    sqlStr = " select shopname, shopdiv "
    sqlStr = sqlStr + " from db_shop.dbo.tbl_shop_user "
    sqlStr = sqlStr + " where userid='"&shopid&"'"

	'response.write sqlStr &"<Br>"
    rsget.Open sqlStr, dbget, 1
    If Not rsget.EOF then
        fnGetShopName = rsget("shopname")
        shopdiv       = rsget("shopdiv")
    end if
    rsget.close

    if (shopdiv="2") then shopdiv="1"
    if (shopdiv="4") then shopdiv="3"
    if (shopdiv="6") then shopdiv="5"
    if (shopdiv="8") then shopdiv="7"

    ''iTs
    if (shopid="streetshop874") then shopdiv="1"
    if (shopid="streetshop884") then shopdiv="1"
    ''29cm
    if (shopid="streetshop878") then shopdiv="1"
    if (shopid="cafe003") then shopdiv="1"
end function

function fnGetMayDiffKey(shopid, MaybeYYYYMM)
    fnGetMayDiffKey = 1

    sqlStr = " select count(*)+1 as Maydiffkey from [db_shop].[dbo].tbl_fran_meachuljungsan_master where shopid='"&shopid&"' and yyyymm='"&MaybeYYYYMM&"'"

	'response.write sqlStr &"<Br>"
    rsget.Open sqlStr, dbget, 1
    If Not rsget.EOF then
        fnGetMayDiffKey = rsget("Maydiffkey")
    end if
    rsget.close
end function
%>

<script type='text/javascript'>

	<% if (mode="chulgo") or (mode="modimaster") or (mode="witsksell")  then %>
		alert('저장 되었습니다.');
		opener.popMasterEdit('<%= iid %>');
		opener.location.reload();
		window.close();
	<% elseif mode="delmaster" then %>
		alert('삭제 되었습니다.');
		opener.location.reload();
		window.close();
	<% elseif (mode="etcsubadd") or (mode="etcsubdetailadd") then %>
		alert('저장 되었습니다.');
		opener.location.reload();
		window.close();
	<% elseif mode="modidetail" or mode="deldetail" then %>
		alert('수정 되었습니다.');
		location.replace('<%= refer %>');
	<% else %>
		alert('저장 되었습니다.');
		opener.popMasterEdit('<%= iid %>');
		opener.location.reload();
		window.close();
	<% end if %>

</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->