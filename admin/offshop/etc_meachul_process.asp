<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  기타매출관리
' History : 2009.04.07 서동석 생성
'			2010.05.13 한용민 수정
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_taxsheetcls.asp"-->
<%
'// ===========================================================================
dim mode, check, idx, topidx ,shopid, makerid, yyyy, mm ,dd ,workidx , adminuserid ,adminusername
dim sqlStr, i, iid, cnt ,MaybeYYYYMM,Maydiffkey, shopname, maybetitle, shopdiv
dim ckidx, suplycasharr,itemnoarr ,orgsellcasharr,sellcasharr,buycasharr ,b2bcharge
dim bizsection_cd, selltype, divcode, papertype
dim oTax
dim diffkey
dim neoTaxNo, eserotaxkey, paperissuetype
dim AssignedRow
dim ctype

	mode    = requestCheckVar(request("mode"),32)
	check   = request("check")
	shopid  = requestCheckVar(request("shopid"),32)
	idx     = requestCheckVar(request("idx"),10)
	topidx  = requestCheckVar(request("topidx"),10)
	makerid = requestCheckVar(request("makerid"),32)
	yyyy    = requestCheckVar(request("yyyy1"),4)
	mm      = requestCheckVar(request("mm1"),2)
	dd      = requestCheckVar(request("dd1"),2)
	workidx      	= requestCheckVar(request("workidx"),10)
	b2bcharge      	= requestCheckVar(request("b2bcharge"),20)
	bizsection_cd  	= requestCheckVar(request("bizsection_cd"),10)
	selltype  		= requestCheckVar(request("selltype"),10)
	divcode  		= requestCheckVar(request("divcode"),2)
	papertype  		= requestCheckVar(request("papertype"),3)

	diffkey  		= requestCheckVar(request("diffkey"),10)
	shopdiv  		= requestCheckVar(request("shopdiv"),2)

	neoTaxNo  		= requestCheckVar(request("neoTaxNo"),32)
	eserotaxkey  	= requestCheckVar(request("eserotaxkey"),32)
	paperissuetype  = requestCheckVar(request("paperissuetype"),1)
    ctype           = requestCheckVar(request("ctype"),10)

adminuserid = session("ssBctId")
adminusername = session("ssBctCname")

if (workidx = "") then
	workidx = "NULL"
end if

dim refer
	refer = request.ServerVariables("HTTP_REFERER")

'response.write mode & "!!!<Br>"


'// ===========================================================================
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

		sqlStr = " update j "
		sqlStr = sqlStr + " set j.invoiceidx = c.invoiceidx "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	[db_shop].[dbo].tbl_fran_meachuljungsan_master j "
		sqlStr = sqlStr + " 	join db_storage.dbo.tbl_cartoonbox_master c "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		1 = 1 "
		sqlStr = sqlStr + " 		and j.idx =  " & iid & " "
		sqlStr = sqlStr + " 		and c.idx =  " & workidx & " "
		sqlStr = sqlStr + " 		and c.invoiceidx is not NULL "
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
	sqlStr = sqlStr + " d.id, d.iitemgubun, d.itemid, d.itemoption, convert(varchar(64),d.iitemname), d.iitemoptionname,"
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
	sqlStr = sqlStr + " ,bizsection_cd='"&bizsection_cd&"'" + vbcrlf
	sqlStr = sqlStr + " ,papertype='"&papertype&"'" + vbcrlf
	sqlStr = sqlStr + " ,paperissuetype='"&paperissuetype&"'" + vbcrlf
	sqlStr = sqlStr + " ,selltype='"&selltype&"'" + vbcrlf
	sqlStr = sqlStr + " from (select sum(totalsuplycash) as totalsum, sum(totalsellcash) as totalsellcash, "
	sqlStr = sqlStr + "			sum(totalbuycash) as totalbuycash, sum(totalsuplycash) as totalsuplycash "
	sqlStr = sqlStr + " 		from [db_shop].[dbo].tbl_fran_meachuljungsan_submaster"
	sqlStr = sqlStr + " 		where masteridx=" + CStr(iid)
	sqlStr = sqlStr + " 	) as T "
	sqlStr = sqlStr + " where idx=" + CStr(iid)

	rsget.Open sqlStr, dbget, 1

	'// wholesale1075 인 경우, 부가세 10% 가산, skyer9, 2017-04-11
	If (shopid = "wholesale1075") Then
		sqlStr = " update [db_shop].[dbo].tbl_fran_meachuljungsan_master "
		sqlStr = sqlStr + " set totalsum = round(totalsum*1.1,0) "
		sqlStr = sqlStr + " where idx = " + CStr(iid)
		rsget.Open sqlStr, dbget, 1
	End If

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
	sqlStr = sqlStr + " d.itemgubun,d.itemid,d.itemoption,convert(varchar(64),d.itemname),d.itemoptionname," '' convert(varchar(64) 추가
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
	sqlStr = sqlStr + " ,bizsection_cd='"&bizsection_cd&"'" + vbcrlf
	sqlStr = sqlStr + " ,papertype='"&papertype&"'" + vbcrlf
	sqlStr = sqlStr + " ,paperissuetype='"&paperissuetype&"'" + vbcrlf
	sqlStr = sqlStr + " ,selltype='"&selltype&"'" + vbcrlf
	sqlStr = sqlStr + " from (select sum(totalsuplycash) as totalsum, sum(totalsellcash) as totalsellcash, "
	sqlStr = sqlStr + "			sum(totalbuycash) as totalbuycash, sum(totalsuplycash) as totalsuplycash, "
	sqlStr = sqlStr + "			sum(totalorgsellcash) as totalorgsellcash "
	sqlStr = sqlStr + " 		from [db_shop].[dbo].tbl_fran_meachuljungsan_submaster"
	sqlStr = sqlStr + " 		where masteridx=" + CStr(iid)
	sqlStr = sqlStr + " 	) as T "
	sqlStr = sqlStr + " where idx=" + CStr(iid)

	rsget.Open sqlStr, dbget, 1

elseif mode="offipjumshop" then

	if (Is3PLShopid(shopid) = True) then
		bizsection_cd = "3pl"
	end if

	if shopid = "" or check = "" or b2bcharge = "" or bizsection_cd = "" or selltype = "" then
	    response.write "<script type='text/javascript'>"
		response.write "	alert('구분값이 없습니다');"
		response.write "	window.close();"
		response.write "</script>"
		response.end	:	dbget.close()
	end if

	'' response.write "TEST..." & check
	'' response.end

	'//마스터 등록
	sqlStr = " select * from [db_shop].[dbo].tbl_fran_meachuljungsan_master where 1=0"
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew

	rsget("shopid") = shopid
	rsget("title") = shopid & " 작성중"
	rsget("totalsum") = 0
	rsget("divcode") = "AA"
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
	sqlStr = sqlStr & " 	, isnull(sum(d.realsellprice*d.itemno),0)-(isnull(sum(d.realsellprice*d.itemno),0)*"&b2bcharge&"/100)" + vbcrlf
	sqlStr = sqlStr & "		from [db_shop].[dbo].tbl_shopjumun_master m " + vbcrlf
    sqlStr = sqlStr & "		join [db_shop].[dbo].tbl_shopjumun_detail d " + vbcrlf
	sqlStr = sqlStr & "			on m.idx=d.masteridx" + vbcrlf
    sqlStr = sqlStr & "			and m.cancelyn='N' and d.cancelyn='N'" + vbcrlf
    sqlStr = sqlStr & "		left join [db_shop].[dbo].tbl_fran_meachuljungsan_subdetail sd "
    sqlStr = sqlStr & "		on d.idx = sd.linkdetailidx "
    sqlStr = sqlStr & "     where convert(varchar(10),m.shopregdate,121) in (" & check & ")" + vbcrlf
    sqlStr = sqlStr & " 	and m.shopid='" & shopid & "'" + vbcrlf
	sqlStr = sqlStr & " 	and sd.linkdetailidx is null" + vbcrlf		'//중복체크
    sqlStr = sqlStr & " 	group by" + vbcrlf
    sqlStr = sqlStr & " 		m.shopid, convert(varchar(10),m.shopregdate,121)"
    ''response.write sqlStr & "<Br>"

	dbget.execute sqlStr

	'//서브 디테일 등록(주문마스터)
	sqlStr = " insert into [db_shop].[dbo].tbl_fran_meachuljungsan_subdetail" + vbcrlf
	sqlStr = sqlStr & " (masteridx, topmasteridx, linkbaljucode, linkmastercode, linkdetailidx" + vbcrlf
	sqlStr = sqlStr & " , itemgubun, itemid, itemoption, itemname, itemoptionname, makerid " + vbcrlf
	sqlStr = sqlStr & " , itemno, sellcash" + vbcrlf
	sqlStr = sqlStr & " , suplycash" + vbcrlf
	sqlStr = sqlStr & " , buycash, orgsellcash)" + vbcrlf
	sqlStr = sqlStr & " 	select" + vbcrlf
	sqlStr = sqlStr & " 	sm.idx, " + CStr(iid) + ",'SB2C', m.orderno, d.idx" + vbcrlf
	sqlStr = sqlStr & "		, d.itemgubun, d.itemid, d.itemoption, convert(varchar(64),d.itemname), d.itemoptionname, d.makerid " + vbcrlf  '// convert(varchar(64), 추가 2016/01/04
	sqlStr = sqlStr & " 	,d.itemno, isnull(d.realsellprice,0)" + vbcrlf
	sqlStr = sqlStr & " 	, isnull(d.realsellprice,0)-(isnull(d.realsellprice,0)*"&b2bcharge&"/100)" + vbcrlf		'// 디테일에서는 소수점을 버리지 않는다.
	sqlStr = sqlStr & " 	,isnull(d.suplyprice,0), d.sellprice" + vbcrlf
	sqlStr = sqlStr & "		from [db_shop].[dbo].tbl_shopjumun_master m " + vbcrlf
    sqlStr = sqlStr & "		join [db_shop].[dbo].tbl_shopjumun_detail d " + vbcrlf
	sqlStr = sqlStr & "			on m.idx=d.masteridx" + vbcrlf
    sqlStr = sqlStr & "			and m.cancelyn='N' and d.cancelyn='N'" + vbcrlf
    sqlStr = sqlStr & " 	join [db_shop].[dbo].tbl_fran_meachuljungsan_submaster sm" + vbcrlf
    sqlStr = sqlStr & " 		on m.shopid=sm.shopid" + vbcrlf
    sqlStr = sqlStr & " 		and convert(varchar(10),m.shopregdate,121)=sm.code01" + vbcrlf
	sqlStr = sqlStr & " 		and sm.masteridx = " + CStr(iid) + " " + vbcrlf
    sqlStr = sqlStr & "		left join [db_shop].[dbo].tbl_fran_meachuljungsan_subdetail sd" + vbcrlf
    sqlStr = sqlStr & "			on d.idx = sd.linkdetailidx" + vbcrlf
    sqlStr = sqlStr & "			and linkbaljucode = 'SB2C'" + vbcrlf
	sqlStr = sqlStr & " 	where sm.masteridx=" + CStr(iid)
	sqlStr = sqlStr & " 	and sd.linkdetailidx is null" + vbcrlf		'//중복체크

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
        maybetitle = shopname + " " + MaybeYYYYMM + " 정산"
    End IF

	'//마스터 업데이트
	sqlStr = " update [db_shop].[dbo].tbl_fran_meachuljungsan_master" + vbcrlf
	sqlStr = sqlStr + " set totalsum=T.totalsum " + vbcrlf
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

	if (bizsection_cd <> "3pl") then
		sqlStr = sqlStr + " ,bizsection_cd='"&bizsection_cd&"'" + vbcrlf
	end if

	sqlStr = sqlStr + " ,papertype='"&papertype&"'" + vbcrlf
	sqlStr = sqlStr + " ,paperissuetype='"&paperissuetype&"'" + vbcrlf
	sqlStr = sqlStr + " ,selltype='"&selltype&"'" + vbcrlf
	sqlStr = sqlStr + " from (" + vbcrlf
	sqlStr = sqlStr + " 	select" + vbcrlf
	sqlStr = sqlStr + " 	floor(sum(totalsuplycash)) as totalsum, sum(totalsellcash) as totalsellcash" + vbcrlf			'// 마스터는 소수점 버린다.
	sqlStr = sqlStr + "		,sum(totalbuycash) as totalbuycash, floor(sum(totalsuplycash)) as totalsuplycash" + vbcrlf
	sqlStr = sqlStr + "		,sum(totalorgsellcash) as totalorgsellcash " + vbcrlf
	sqlStr = sqlStr + " 	from [db_shop].[dbo].tbl_fran_meachuljungsan_submaster" + vbcrlf
	sqlStr = sqlStr + " 	where masteridx=" + CStr(iid) + vbcrlf
	sqlStr = sqlStr + " ) as T " + vbcrlf
	sqlStr = sqlStr + " where idx=" + CStr(iid)

	'response.write sqlStr & "<Br>"
	'response.end
	dbget.execute sqlStr

elseif mode="onipjumshop" then

	if (Is3PLShopid(shopid) = True) then
		bizsection_cd = "3pl"
	end if

	if shopid = "" or check = "" or b2bcharge = "" or bizsection_cd = "" or selltype = "" then
	    response.write "<script type='text/javascript'>"
		response.write "	alert('구분값이 없습니다');"
		response.write "	window.close();"
		response.write "</script>"
		response.end	:	dbget.close()
	end if

	'//마스터 등록
	sqlStr = " select * from [db_shop].[dbo].tbl_fran_meachuljungsan_master where 1=0"
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew

	rsget("shopid") = shopid
	rsget("title") = shopid & " 작성중"
	rsget("totalsum") = 0
	rsget("divcode") = "BB"
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
	sqlStr = sqlStr & " 	select " + vbcrlf
	sqlStr = sqlStr & " 		"& CStr(iid) & " " + vbcrlf
	sqlStr = sqlStr & " 		, 0 " + vbcrlf
	sqlStr = sqlStr & " 		, m.sitename " + vbcrlf
	sqlStr = sqlStr & " 		, convert(varchar(10),d.beasongdate,121) as yyyymmdd " + vbcrlf
	sqlStr = sqlStr & " 		, '"&makerid&"' " + vbcrlf
	sqlStr = sqlStr & " 		, convert(varchar(10),d.beasongdate,121) as yyyymmdd " + vbcrlf
	sqlStr = sqlStr & " 		, sum(d.itemno) as totitemcnt " + vbcrlf
	sqlStr = sqlStr & " 		, sum(isnull(d.reducedPrice,0)*(d.itemno)) as totsum " + vbcrlf
	sqlStr = sqlStr & " 		, sum(isnull(d.buycash,0)*(d.itemno)) as buyprice " + vbcrlf
	sqlStr = sqlStr & " 		, sum(isnull(d.reducedPrice,0)*d.itemno-(CASE WHEN d.itemid=0 then 0 else isnull(d.reducedPrice,0)*d.itemno*"&b2bcharge&"/100 end )) " + vbcrlf
	sqlStr = sqlStr & " 	from " + vbcrlf
	sqlStr = sqlStr & " 		db_order.dbo.tbl_order_master m " + vbcrlf
	sqlStr = sqlStr & " 		join db_order.dbo.tbl_order_detail d " + vbcrlf
	sqlStr = sqlStr & " 		on " + vbcrlf
	sqlStr = sqlStr & " 			m.orderserial = d.orderserial " + vbcrlf
	sqlStr = sqlStr & " 		left join [db_shop].[dbo].tbl_fran_meachuljungsan_submaster sm " + vbcrlf
	sqlStr = sqlStr & " 		on " + vbcrlf
	sqlStr = sqlStr & " 			1 = 1 " + vbcrlf
	sqlStr = sqlStr & " 			and m.sitename = sm.shopid " + vbcrlf
	sqlStr = sqlStr & " 			and convert(varchar(10),d.beasongdate,121) = sm.code01 " + vbcrlf
	sqlStr = sqlStr & " 			and sm.code02 = '"&makerid&"' " + vbcrlf
	sqlStr = sqlStr & " 	where " + vbcrlf
	sqlStr = sqlStr & " 		1 = 1 " + vbcrlf
	if (makerid<>"") then
	    sqlStr = sqlStr & " 		and d.makerid='"&makerid&"'" + vbcrlf
	end if
	if (ctype<>"P") then
    	sqlStr = sqlStr & " 		and d.itemid <> 0 " + vbcrlf			'// 배송비 제외
    end if
	sqlStr = sqlStr & " 		and convert(varchar(10),d.beasongdate,121) in (" & check & ") " + vbcrlf
	sqlStr = sqlStr & " 		and d.currstate >= '7' " + vbcrlf
	sqlStr = sqlStr & " 		and m.sitename = '" & shopid & "' " + vbcrlf
	sqlStr = sqlStr & " 		and m.cancelyn='N' " + vbcrlf
	sqlStr = sqlStr & " 		and d.cancelyn<>'Y' " + vbcrlf
	sqlStr = sqlStr & " 		and sm.idx is null " + vbcrlf			'//중복체크
	sqlStr = sqlStr & " 	group by " + vbcrlf
	sqlStr = sqlStr & " 		convert(varchar(10),d.beasongdate,121) " + vbcrlf
	sqlStr = sqlStr & " 		, m.sitename " + vbcrlf
    ''response.write sqlStr & "<Br>"

	dbget.execute sqlStr

	'//서브 디테일 등록(주문마스터)
	sqlStr = " insert into [db_shop].[dbo].tbl_fran_meachuljungsan_subdetail" + vbcrlf
	sqlStr = sqlStr & " (masteridx, topmasteridx, linkbaljucode, linkmastercode, linkdetailidx" + vbcrlf
	sqlStr = sqlStr & " , itemgubun, itemid, itemoption, itemname, itemoptionname, makerid " + vbcrlf
	sqlStr = sqlStr & " , itemno, sellcash" + vbcrlf
	sqlStr = sqlStr & " , suplycash" + vbcrlf
	sqlStr = sqlStr & " , buycash, orgsellcash)" + vbcrlf

	sqlStr = sqlStr & " 	select " + vbcrlf
	sqlStr = sqlStr & " 		sm.idx " + vbcrlf
	sqlStr = sqlStr & " 		, " + CStr(iid) + " " + vbcrlf
	sqlStr = sqlStr & " 		, 'OB2C' " + vbcrlf
	sqlStr = sqlStr & " 		, m.orderserial " + vbcrlf
	sqlStr = sqlStr & " 		, d.idx " + vbcrlf
	sqlStr = sqlStr & "			, '10', d.itemid, d.itemoption, d.itemname, d.itemoptionname, d.makerid " + vbcrlf
	sqlStr = sqlStr & " 		, (d.itemno) " + vbcrlf
	sqlStr = sqlStr & " 		, isnull(d.reducedPrice,0) " + vbcrlf
	sqlStr = sqlStr & " 		, (isnull(d.reducedPrice,0) - (CASE WHEN d.itemid=0 then 0 else (isnull(d.reducedPrice,0)*"&b2bcharge&"/100) end )) " + vbcrlf			'// 디테일에서는 소수점을 버리지 않는다.
	sqlStr = sqlStr & " 		, isnull(d.buycash,0) " + vbcrlf
	sqlStr = sqlStr & " 		, isnull(d.orgitemcost,0) " + vbcrlf
	sqlStr = sqlStr & " 	from " + vbcrlf
	sqlStr = sqlStr & " 		db_order.dbo.tbl_order_master m " + vbcrlf
	sqlStr = sqlStr & " 		join db_order.dbo.tbl_order_detail d " + vbcrlf
	sqlStr = sqlStr & " 		on " + vbcrlf
	sqlStr = sqlStr & " 			m.orderserial = d.orderserial " + vbcrlf
	sqlStr = sqlStr & " 		left join [db_shop].[dbo].tbl_fran_meachuljungsan_submaster sm " + vbcrlf
	sqlStr = sqlStr & " 		on " + vbcrlf
	sqlStr = sqlStr & " 			1 = 1 " + vbcrlf
	sqlStr = sqlStr & " 			and m.sitename = sm.shopid " + vbcrlf
	sqlStr = sqlStr & " 			and convert(varchar(10),d.beasongdate,121) = sm.code01 " + vbcrlf
	sqlStr = sqlStr & " 		left join [db_shop].[dbo].tbl_fran_meachuljungsan_subdetail sd " + vbcrlf
	sqlStr = sqlStr & " 		on " + vbcrlf
	sqlStr = sqlStr & " 			1 = 1 " + vbcrlf
	sqlStr = sqlStr & " 			and d.idx = sd.linkdetailidx " + vbcrlf
	sqlStr = sqlStr & " 			and sd.linkbaljucode = 'OB2C' " + vbcrlf
	sqlStr = sqlStr & " 	where " + vbcrlf
	sqlStr = sqlStr & " 		1 = 1 " + vbcrlf
	sqlStr = sqlStr & " 		and sm.masteridx=" + CStr(iid) + " " + vbcrlf
	if (makerid<>"") then
	    sqlStr = sqlStr & " 		and d.makerid='"&makerid&"'" + vbcrlf
	end if
	if (ctype<>"P") then
    	sqlStr = sqlStr & " 		and d.itemid <> 0 " + vbcrlf			'// 배송비 제외
    end if
	sqlStr = sqlStr & " 		and convert(varchar(10),d.beasongdate,121) in (" & check & ") " + vbcrlf
	sqlStr = sqlStr & " 		and d.currstate >= '7' " + vbcrlf
	sqlStr = sqlStr & " 		and m.sitename = '" & shopid & "' " + vbcrlf
	sqlStr = sqlStr & " 		and m.cancelyn='N' " + vbcrlf
	sqlStr = sqlStr & " 		and d.cancelyn<>'Y' " + vbcrlf
	sqlStr = sqlStr & " 		and sd.linkdetailidx is null " + vbcrlf
	''response.write sqlStr & "<Br>"

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
    shopname    = fnGetOnIpjumShopName(shopid, shopdiv)
    Maydiffkey  = fnGetMayDiffKey(shopid, MaybeYYYYMM)		'//차수

    IF (MaybeYYYYMM<>"") and (shopname<>"") then
        maybetitle = shopname + " " + MaybeYYYYMM + " 상품대 정산"
        if (makerid<>"") then maybetitle=maybetitle+" ("&makerid&")"
    End IF

	'//마스터 업데이트
	sqlStr = " update [db_shop].[dbo].tbl_fran_meachuljungsan_master" + vbcrlf
	sqlStr = sqlStr + " set totalsum=T.totalsum " + vbcrlf
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

	if (bizsection_cd <> "3pl") then
		sqlStr = sqlStr + " ,bizsection_cd='"&bizsection_cd&"'" + vbcrlf
	end if

	sqlStr = sqlStr + " ,papertype='"&papertype&"'" + vbcrlf
	sqlStr = sqlStr + " ,paperissuetype='"&paperissuetype&"'" + vbcrlf
	sqlStr = sqlStr + " ,selltype='"&selltype&"'" + vbcrlf
	sqlStr = sqlStr + " from (" + vbcrlf
	sqlStr = sqlStr + " 	select" + vbcrlf
	sqlStr = sqlStr + " 	floor(sum(totalsuplycash)) as totalsum, sum(totalsellcash) as totalsellcash" + vbcrlf				'// 마스터는 소수점 버린다.
	sqlStr = sqlStr + "		,sum(totalbuycash) as totalbuycash, floor(sum(totalsuplycash)) as totalsuplycash" + vbcrlf
	sqlStr = sqlStr + "		,sum(totalorgsellcash) as totalorgsellcash " + vbcrlf
	sqlStr = sqlStr + " 	from [db_shop].[dbo].tbl_fran_meachuljungsan_submaster" + vbcrlf
	sqlStr = sqlStr + " 	where masteridx=" + CStr(iid) + vbcrlf
	sqlStr = sqlStr + " ) as T " + vbcrlf
	sqlStr = sqlStr + " where idx=" + CStr(iid)

	'response.write sqlStr & "<Br>"
	'response.end
	dbget.execute sqlStr

elseif mode="onipjumshopbeasongpay" then

	if (Is3PLShopid(shopid) = True) then
		bizsection_cd = "3pl"
	end if

	if shopid = "" or check = "" or b2bcharge = "" or bizsection_cd = "" or selltype = "" then
	    response.write "<script type='text/javascript'>"
		response.write "	alert('구분값이 없습니다');"
		response.write "	window.close();"
		response.write "</script>"
		response.end	:	dbget.close()
	end if

	'//마스터 등록
	sqlStr = " select * from [db_shop].[dbo].tbl_fran_meachuljungsan_master where 1=0"
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew

	rsget("shopid") = shopid
	rsget("title") = shopid & " 작성중"
	rsget("totalsum") = 0
	rsget("divcode") = "CC"
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
	sqlStr = sqlStr & " 	select " + vbcrlf
	sqlStr = sqlStr & " 		"& CStr(iid) & " " + vbcrlf
	sqlStr = sqlStr & " 		, 0 " + vbcrlf
	sqlStr = sqlStr & " 		, m.sitename " + vbcrlf
	sqlStr = sqlStr & " 		, convert(varchar(10),d.beasongdate,121) as yyyymmdd " + vbcrlf
	sqlStr = sqlStr & " 		, 'beasongpay' " + vbcrlf
	sqlStr = sqlStr & " 		, convert(varchar(10),d.beasongdate,121) as yyyymmdd " + vbcrlf
	sqlStr = sqlStr & " 		, sum(case when d.itemid = 0 then d.itemno else 0 end) as totitemcnt " + vbcrlf
	sqlStr = sqlStr & " 		, sum(isnull(d.reducedPrice,0)*(case when d.itemid = 0 then d.itemno else 0 end)) as totsum " + vbcrlf
	sqlStr = sqlStr & " 		, sum(isnull(d.buycash,0)*(case when d.itemid = 0 then d.itemno else 0 end)) as buyprice " + vbcrlf
	sqlStr = sqlStr & " 		, sum(isnull(d.reducedPrice,0)*(case when d.itemid = 0 then d.itemno else 0 end))-floor(sum(isnull(d.reducedPrice,0)*(case when d.itemid = 0 then d.itemno else 0 end)*"&b2bcharge&"/100)) " + vbcrlf
	sqlStr = sqlStr & " 	from " + vbcrlf
	sqlStr = sqlStr & " 		db_order.dbo.tbl_order_master m " + vbcrlf
	sqlStr = sqlStr & " 		join db_order.dbo.tbl_order_detail d " + vbcrlf
	sqlStr = sqlStr & " 		on " + vbcrlf
	sqlStr = sqlStr & " 			m.orderserial = d.orderserial " + vbcrlf
	sqlStr = sqlStr & " 		left join [db_shop].[dbo].tbl_fran_meachuljungsan_submaster sm " + vbcrlf
	sqlStr = sqlStr & " 		on " + vbcrlf
	sqlStr = sqlStr & " 			1 = 1 " + vbcrlf
	sqlStr = sqlStr & " 			and m.sitename = sm.shopid " + vbcrlf
	sqlStr = sqlStr & " 			and convert(varchar(10),d.beasongdate,121) = sm.code01 " + vbcrlf
	sqlStr = sqlStr & " 			and sm.code02 = 'beasongpay' " + vbcrlf
	sqlStr = sqlStr & " 	where " + vbcrlf
	sqlStr = sqlStr & " 		1 = 1 " + vbcrlf
	sqlStr = sqlStr & " 		and d.itemid = 0 " + vbcrlf			'// 배송비만
	sqlStr = sqlStr & " 		and convert(varchar(10),d.beasongdate,121) in (" & check & ") " + vbcrlf
	sqlStr = sqlStr & " 		and d.currstate >= '7' " + vbcrlf
	sqlStr = sqlStr & " 		and m.sitename = '" & shopid & "' " + vbcrlf
	sqlStr = sqlStr & " 		and m.cancelyn='N' " + vbcrlf
	sqlStr = sqlStr & " 		and d.cancelyn<>'Y' " + vbcrlf
	sqlStr = sqlStr & " 		and sm.idx is null " + vbcrlf			'//중복체크
	sqlStr = sqlStr & " 	group by " + vbcrlf
	sqlStr = sqlStr & " 		convert(varchar(10),d.beasongdate,121) " + vbcrlf
	sqlStr = sqlStr & " 		, m.sitename " + vbcrlf
    ''response.write sqlStr & "<Br>"

	dbget.execute sqlStr

	'//서브 디테일 등록(주문마스터)
	sqlStr = " insert into [db_shop].[dbo].tbl_fran_meachuljungsan_subdetail" + vbcrlf
	sqlStr = sqlStr & " (masteridx, topmasteridx, linkbaljucode, linkmastercode, linkdetailidx" + vbcrlf
	sqlStr = sqlStr & " , itemgubun, itemid, itemoption, itemname, itemoptionname, makerid " + vbcrlf
	sqlStr = sqlStr & " , itemno, sellcash" + vbcrlf
	sqlStr = sqlStr & " , suplycash" + vbcrlf
	sqlStr = sqlStr & " , buycash, orgsellcash)" + vbcrlf

	sqlStr = sqlStr & " 	select " + vbcrlf
	sqlStr = sqlStr & " 		sm.idx " + vbcrlf
	sqlStr = sqlStr & " 		, " + CStr(iid) + " " + vbcrlf
	sqlStr = sqlStr & " 		, 'OBEA' " + vbcrlf
	sqlStr = sqlStr & " 		, m.orderserial " + vbcrlf
	sqlStr = sqlStr & " 		, d.idx " + vbcrlf
	sqlStr = sqlStr & "			, '10', d.itemid, d.itemoption, d.itemname, d.itemoptionname, d.makerid " + vbcrlf
	sqlStr = sqlStr & " 		, (case when d.itemid = 0 then d.itemno else 0 end) " + vbcrlf
	sqlStr = sqlStr & " 		, isnull(d.reducedPrice,0) " + vbcrlf
	sqlStr = sqlStr & " 		, (isnull(d.reducedPrice,0) - (isnull(d.reducedPrice,0)*"&b2bcharge&"/100)) " + vbcrlf			'// 디테일에서는 소수점을 버리지 않는다.
	sqlStr = sqlStr & " 		, isnull(d.buycash,0) " + vbcrlf
	sqlStr = sqlStr & " 		, isnull(d.orgitemcost,0) " + vbcrlf
	sqlStr = sqlStr & " 	from " + vbcrlf
	sqlStr = sqlStr & " 		db_order.dbo.tbl_order_master m " + vbcrlf
	sqlStr = sqlStr & " 		join db_order.dbo.tbl_order_detail d " + vbcrlf
	sqlStr = sqlStr & " 		on " + vbcrlf
	sqlStr = sqlStr & " 			m.orderserial = d.orderserial " + vbcrlf
	sqlStr = sqlStr & " 		left join [db_shop].[dbo].tbl_fran_meachuljungsan_submaster sm " + vbcrlf
	sqlStr = sqlStr & " 		on " + vbcrlf
	sqlStr = sqlStr & " 			1 = 1 " + vbcrlf
	sqlStr = sqlStr & " 			and m.sitename = sm.shopid " + vbcrlf
	sqlStr = sqlStr & " 			and convert(varchar(10),d.beasongdate,121) = sm.code01 " + vbcrlf
	sqlStr = sqlStr & " 		left join [db_shop].[dbo].tbl_fran_meachuljungsan_subdetail sd " + vbcrlf
	sqlStr = sqlStr & " 		on " + vbcrlf
	sqlStr = sqlStr & " 			1 = 1 " + vbcrlf
	sqlStr = sqlStr & " 			and d.idx = sd.linkdetailidx " + vbcrlf
	sqlStr = sqlStr & " 			and sd.linkbaljucode = 'OBEA' " + vbcrlf
	sqlStr = sqlStr & " 	where " + vbcrlf
	sqlStr = sqlStr & " 		1 = 1 " + vbcrlf
	sqlStr = sqlStr & " 		and sm.masteridx=" + CStr(iid) + " " + vbcrlf
	sqlStr = sqlStr & " 		and d.itemid = 0 " + vbcrlf			'// 배송비만
	sqlStr = sqlStr & " 		and convert(varchar(10),d.beasongdate,121) in (" & check & ") " + vbcrlf
	sqlStr = sqlStr & " 		and d.currstate >= '7' " + vbcrlf
	sqlStr = sqlStr & " 		and m.sitename = '" & shopid & "' " + vbcrlf
	sqlStr = sqlStr & " 		and m.cancelyn='N' " + vbcrlf
	sqlStr = sqlStr & " 		and d.cancelyn<>'Y' " + vbcrlf
	sqlStr = sqlStr & " 		and sd.linkdetailidx is null " + vbcrlf
	''response.write sqlStr & "<Br>"

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
    shopname    = fnGetOnIpjumShopName(shopid, shopdiv)
    Maydiffkey  = fnGetMayDiffKey(shopid, MaybeYYYYMM)		'//차수

    IF (MaybeYYYYMM<>"") and (shopname<>"") then
        maybetitle = shopname + " " + MaybeYYYYMM + " 배송비 정산"
    End IF

	'//마스터 업데이트
	sqlStr = " update [db_shop].[dbo].tbl_fran_meachuljungsan_master" + vbcrlf
	sqlStr = sqlStr + " set totalsum=T.totalsum " + vbcrlf
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

	if (bizsection_cd <> "3pl") then
		sqlStr = sqlStr + " ,bizsection_cd='"&bizsection_cd&"'" + vbcrlf
	end if

	sqlStr = sqlStr + " ,papertype='"&papertype&"'" + vbcrlf
	sqlStr = sqlStr + " ,paperissuetype='"&paperissuetype&"'" + vbcrlf
	sqlStr = sqlStr + " ,selltype='"&selltype&"'" + vbcrlf
	sqlStr = sqlStr + " from (" + vbcrlf
	sqlStr = sqlStr + " 	select" + vbcrlf
	sqlStr = sqlStr + " 	floor(sum(totalsuplycash)) as totalsum, sum(totalsellcash) as totalsellcash" + vbcrlf				'// 마스터는 소수점 버린다.
	sqlStr = sqlStr + "		,sum(totalbuycash) as totalbuycash, floor(sum(totalsuplycash)) as totalsuplycash" + vbcrlf
	sqlStr = sqlStr + "		,sum(totalorgsellcash) as totalorgsellcash " + vbcrlf
	sqlStr = sqlStr + " 	from [db_shop].[dbo].tbl_fran_meachuljungsan_submaster" + vbcrlf
	sqlStr = sqlStr + " 	where masteridx=" + CStr(iid) + vbcrlf
	sqlStr = sqlStr + " ) as T " + vbcrlf
	sqlStr = sqlStr + " where idx=" + CStr(iid)

	'response.write sqlStr & "<Br>"
	'response.end
	dbget.execute sqlStr

elseif mode="addmaster" then

	if (IsEtcShopid(shopid) <> True) then
	    response.write "<script type='text/javascript'>"
		response.write "	alert('기타매장, 출고처(기타) 만 선택가능합니다.');"
		response.write "</script>"

		response.Write "기타매장, 출고처(기타) 만 선택가능합니다."

		dbget.Close()
		response.end
	end if

	sqlStr = " select * from [db_shop].[dbo].tbl_fran_meachuljungsan_master where 1=0"
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew

	rsget("shopid") = request("shopid")
	rsget("title") = html2db(request("title"))
	rsget("totalsum") =  request("totalsuplycash") ''request("totalsum") =>발행금액 == 공급액
	rsget("totalsellcash") = request("totalsuplycash")
	rsget("totalbuycash") = request("totalbuycash")
	rsget("totalsuplycash") = request("totalsuplycash")

	rsget("divcode") = request("divcode")
	rsget("etcstr") = html2db(request("etcstr"))
	rsget("reguserid") = session("ssBctId")
	rsget("regusername") = session("ssBctCname")

    rsget("shopdiv") = shopdiv
    rsget("diffKey") = diffKey

    rsget("bizsection_cd") = bizsection_cd
    rsget("selltype") = selltype
    rsget("papertype") = papertype
    rsget("paperissuetype") = paperissuetype

    rsget("shopdiv") = shopdiv
    rsget("yyyymm") = yyyy + "-" + mm

	rsget.update
	iid = rsget("idx")
	rsget.close

elseif mode="modimaster" then

	sqlStr = " update [db_shop].[dbo].tbl_fran_meachuljungsan_master" + VbCrlf
	sqlStr = sqlStr + " set title='" + html2db(request("title")) + "'" + VbCrlf
	sqlStr = sqlStr + " ,totalsum=" + request("totalsum") + ""
	sqlStr = sqlStr + " ,totalbuycash=" + request("totalbuycash") + ""
	''sqlStr = sqlStr + " ,statecd='" + request("statecd") + "'"
	sqlStr = sqlStr + " ,shopdiv='" + request("shopdiv") + "'"
	sqlStr = sqlStr + " ,selltype='" + request("selltype") + "'"
	sqlStr = sqlStr + " ,bizsection_cd='" + request("bizsection_cd") + "'"
	sqlStr = sqlStr + " ,papertype='" + request("papertype") + "'"
	sqlStr = sqlStr + " ,paperissuetype='" + request("paperissuetype") + "'"

	if request("taxdate")<>"" then
		sqlStr = sqlStr + " ,taxdate='" + request("taxdate") + "'"
	else
		sqlStr = sqlStr + " ,taxdate=NULL"
	end if

	if request("ipkumdate")<>"" then
		sqlStr = sqlStr + " ,ipkumdate='" + request("ipkumdate") + "'"
	else
		sqlStr = sqlStr + " ,ipkumdate=NULL"
	end if

	if C_ADMIN_AUTH or C_MngPart or C_PSMngPart then
		sqlStr = sqlStr & " ,yyyymm='" & yyyy & "-" & mm & "'" & vbcrlf
	end if

	sqlStr = sqlStr + " ,etcstr='" + html2db(request("etcstr")) + "'"
	sqlStr = sqlStr + " ,finishuserid='" + session("ssBctId") + "'"
	sqlStr = sqlStr + " ,finishusername='" + session("ssBctCname") + "'"
	sqlStr = sqlStr + " where idx=" + CStr(idx)  + VbCrlf

	'rw sqlStr & "<br>"
	dbget.execute sqlStr

	''if (papertype = "110") and (eserotaxkey <> "") then
	if (paperissuetype = "2") and (eserotaxkey <> "") then
		'// 세금계산서(역발행)

		sqlStr = " update m "
		sqlStr = sqlStr + " 	set m.issuestatecd = 9, m.taxdate = e.appdate, m.taxregdate = e.regdate, eserotaxkey = e.taxkey "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	[db_shop].[dbo].tbl_fran_meachuljungsan_master m "
		sqlStr = sqlStr + " 	join db_Partner.dbo.tbl_Esero_Tax e "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		1 = 1 "
		sqlStr = sqlStr + " 		and m.idx = " + CStr(idx) + " "
		sqlStr = sqlStr + " 		and e.taxkey = '" + CStr(eserotaxkey) + "' "
		sqlStr = sqlStr + " where isNULL(m.eserotaxkey,'')<>isNULL(e.taxkey,'')"
		dbget.Execute sqlStr,AssignedRow

        ''2012/10/02 추가
        ''If AssignedRow>0 then

        sqlStr = " select count(*) as CNT from db_partner.dbo.tbl_esero_TaxMatch"
        sqlStr = sqlStr + " where taxkey='"&eserotaxkey&"'"
        sqlStr = sqlStr + " and matchType<>0"
        rsget.Open sqlStr, dbget, 1
    		cnt = rsget("cnt")
    	rsget.Close

	    if (cnt<1) then
            sqlStr = " exec db_partner.[dbo].[sp_Ten_Esero_Tax_MatchOne_etcSell] "&CStr(idx)&",'"&eserotaxkey&"'"
            dbget.Execute sqlStr
        end if

	end if

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

elseif mode="changeState" then

	If request("statecd") = "0" Then
		'======================================================================
		'이전 방식(tbl_tax_history_master 를 이용하는 경우)
		sqlStr = " SELECT Count(*) FROM [db_jungsan].[dbo].tbl_tax_history_master " + VbCrlf
		sqlStr = sqlStr + " WHERE jungsanGubun = 'OFFSHOP' "
		sqlStr = sqlStr + " AND jungsanID = '" + request("idx") + "'"
		sqlStr = sqlStr + " AND deleteYN = 'N' "

		rsget.Open sqlStr, dbget, 1
		If Not rsget.EOF Then
			If rsget(0) > 0 Then
				response.write "<script type='text/javascript'>" & vbCrLf
				response.write "alert('세금계산서발행로그가 존재합니다.\n\n수정하시려면 로그를 먼저 삭제하십시오.')" & vbCrLf
				response.write "history.back();" & vbCrLf
				response.write "</script>" & vbCrLf
				rsget.close
				dbget.close
				response.End
			End If
		End If
		rsget.close


		'======================================================================
		'신규 방식(tbl_taxSheet 이용)
		set oTax = new CTax

		oTax.FRectsearchKey = " t.orderidx "
		oTax.FRectsearchString = CStr(request("idx"))
		oTax.FRectDelYn = "N"

		oTax.GetTaxList

		if oTax.FResultCount > 0 then
			if oTax.FTaxList(0).FisueYn="Y" then
				response.write "<script type='text/javascript'>alert('이미 발행된 세금계산서가 있습니다.\n\n변경 하시려면 거래처에 [취소요청]후 매출 세금계산서 목록에서 [삭제]후 발행하셔야 합니다');history.back();</script>"
			else
				response.write "<script type='text/javascript'>alert('발행대기중인 세금계산서가 있습니다.\n\n변경 하시려면 매출 세금계산서 목록에서 [삭제]후 발행하셔야 합니다.');history.back();</script>"
			end if

			response.End
		end if

	End If

	' 정산테이블 상태변경
	sqlStr = " update [db_shop].[dbo].tbl_fran_meachuljungsan_master SET " + VbCrlf
	sqlStr = sqlStr + "  statecd='" + request("statecd") + "'"
	sqlStr = sqlStr + " ,etcstr='" + html2db(request("etcstr")) + "'"
	sqlStr = sqlStr + " ,finishuserid='" + session("ssBctId") + "'"
	sqlStr = sqlStr + " ,finishusername='" + session("ssBctCname") + "'"
	' 입금완료시 입금일은 업데이트 하나 타이틀은 수정할 수 없다.
	If request("statecd") = "7" And Len(request("ipkumdate")) = 10 Then
		sqlStr = sqlStr + " ,ipkumDate='" + request("ipkumdate") + "'"	' 입금일
	Else
		sqlStr = sqlStr + " ,title='" + html2db(request("title")) + "'"
	End If
	sqlStr = sqlStr + " where idx=" + CStr(idx)  + VbCrlf

	dbget.Execute(sqlStr)

	If request("statecd") = "7" And Len(request("ipkumdate")) = 10 Then

		' 주문서 테이블 입금일자 업데이트
		sqlStr = " UPDATE A " + vbCrlf
		sqlStr = sqlStr + " SET a.ipkumDate = '" + request("ipkumdate") + "'"  + vbCrlf
		sqlStr = sqlStr + " FROM db_storage.dbo.tbl_ordersheet_master a " & vbCrLf
		sqlStr = sqlStr + " INNER JOIN [db_shop].[dbo].tbl_fran_meachuljungsan_submaster b " & vbCrLf
		sqlStr = sqlStr + " ON a.baljucode = b.code02 " & vbCrLf
		sqlStr = sqlStr + " WHERE b.masterIdx = " + CStr(idx)

		dbget.Execute(sqlStr)
	End If

elseif mode="changeIssueState" then

	sqlStr = " update [db_shop].[dbo].tbl_fran_meachuljungsan_master SET " + VbCrlf

	if (request("issuestatecd") = "NULL") then
		sqlStr = sqlStr + "  issuestatecd=NULL, taxdate=NULL, taxregdate=NULL, eserotaxkey=NULL  "
	else
		sqlStr = sqlStr + "  issuestatecd='" + request("issuestatecd") + "'"

		If request("issuestatecd") = "9" And Len(request("taxdate")) = 10 Then
			sqlStr = sqlStr + "  , taxdate='" + request("taxdate") + "'"
		end if
	end if

	sqlStr = sqlStr + " where idx=" + CStr(idx)  + VbCrlf

	dbget.Execute(sqlStr)

	if (paperissuetype = "1") and (request("issuestatecd") = "NULL") then
		sqlStr = " exec [db_partner].[dbo].[usp_Ten_etcMeachul_DelTaxInfoForce] " + CStr(idx) + " "
		dbget.Execute(sqlStr)
	end if

	if (request("issuestatecd") <> "NULL") and (papertype = "110") and (eserotaxkey <> "") then
		'// 세금계산서(역발행)

		sqlStr = " update m "
		sqlStr = sqlStr + " 	set m.issuestatecd = 9, m.taxdate = e.appdate, m.taxregdate = e.regdate, eserotaxkey = e.taxkey "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	[db_shop].[dbo].tbl_fran_meachuljungsan_master m "
		sqlStr = sqlStr + " 	join db_Partner.dbo.tbl_Esero_Tax e "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		1 = 1 "
		sqlStr = sqlStr + " 		and m.idx = " + CStr(idx) + " "
		sqlStr = sqlStr + " 		and e.taxkey = '" + CStr(eserotaxkey) + "' "
		rsget.Open sqlStr, dbget, 1

	end if

elseif mode="changeIpkumState" then

	sqlStr = " update [db_shop].[dbo].tbl_fran_meachuljungsan_master SET " + VbCrlf

	if (request("ipkumstatecd") = "NULL") then
		sqlStr = sqlStr + "  ipkumstatecd=NULL "
		sqlStr = sqlStr + "  , ipkumDate=NULL "
	else
		sqlStr = sqlStr + "  ipkumstatecd='" + request("ipkumstatecd") + "'"
		If request("ipkumstatecd") = "9" And Len(request("ipkumDate")) = 10 Then
			sqlStr = sqlStr + "  , ipkumDate='" + request("ipkumDate") + "'"
		end if
	end if

	sqlStr = sqlStr + " where idx=" + CStr(idx)  + VbCrlf

	dbget.Execute(sqlStr)

elseif mode="etcsubadd" then
    cnt=0
	sqlStr = " select count(idx) as cnt from [db_shop].[dbo].tbl_fran_meachuljungsan_master"
	sqlStr = sqlStr + " where idx=" + CStr(topidx)  + VbCrlf
    sqlStr = sqlStr + " and statecd=0"

	rsget.Open sqlStr, dbget, 1
		cnt = rsget("cnt")
	rsget.Close

	if cnt<1 then
		response.write "<script type='text/javascript'>alert('수정중 상태에서만 추가 가능합니다.');</script>"
		response.write "<script type='text/javascript'>window.close();</script>"
		dbget.close()	:	response.End
	end if

	sqlStr = " select * from [db_shop].[dbo].tbl_fran_meachuljungsan_submaster where 1=0"
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew

	rsget("masteridx") = topidx
	rsget("linkidx") = 0
	rsget("shopid") = shopid

	if (dd = "") then
		rsget("code01") = yyyy + "-" + mm
		rsget("execdate") = yyyy + "-" + mm + "-01"
	else
		rsget("code01") = yyyy + "-" + mm + "-" + dd
		rsget("execdate") = yyyy + "-" + mm + "-" + dd
	end if


	rsget("code02") = makerid
	rsget("totalcount") = 0
	rsget("totalsellcash") = 0
	rsget("totalbuycash") = 0
	rsget("totalsuplycash") = 0
	rsget("totalorgsellcash") = 0
	rsget.update
	iid = rsget("idx")
	rsget.close

elseif mode="etcsubdetailadd" then
    ''현재상태 0 수정중인 경우만 삭제 가능
	cnt=0
	sqlStr = " select count(idx) as cnt from [db_shop].[dbo].tbl_fran_meachuljungsan_master"
	sqlStr = sqlStr + " where idx=" + CStr(topidx)  + VbCrlf
    sqlStr = sqlStr + " and statecd=0"

	rsget.Open sqlStr, dbget, 1
		cnt = rsget("cnt")
	rsget.Close

	if cnt<1 then
		response.write "<script type='text/javascript'>alert('수정중 상태에서만 추가 가능합니다.');</script>"
		response.write "<script type='text/javascript'>window.close();</script>"
		dbget.close()	:	response.End
	end if


	sqlStr = " select * from [db_shop].[dbo].tbl_fran_meachuljungsan_subdetail where 1=0"
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew

	rsget("masteridx") = idx
	rsget("topmasteridx") = topidx
	rsget("linkbaljucode") = request("linkbaljucode")
	rsget("linkmastercode") = "0"
	rsget("linkdetailidx") = 0
	rsget("itemgubun") = request("itemgubun")
	rsget("itemid") = request("itemid")
	rsget("itemoption") = request("itemoption")
	rsget("itemname") = html2Db(request("itemname"))
	rsget("itemoptionname") = html2Db(request("itemoptionname"))
	rsget("makerid") = request("makerid")
	rsget("itemno") = request("itemno")
	rsget("sellcash") = request("sellcash")
	rsget("suplycash") = request("suplycash")
	rsget("buycash") = request("buycash")
	rsget("orgsellcash") = request("orgsellcash")

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

elseif mode="deldetail" then
    ''현재상태 0 수정중인 경우만 삭제 가능
	cnt=0
	sqlStr = " select count(idx) as cnt from [db_shop].[dbo].tbl_fran_meachuljungsan_master"
	sqlStr = sqlStr + " where idx=" + CStr(topidx)  + VbCrlf
    sqlStr = sqlStr + " and statecd=0"

	rsget.Open sqlStr, dbget, 1
		cnt = rsget("cnt")
	rsget.Close

	if cnt<1 then
		response.write "<script type='text/javascript'>alert('수정중 상태에서만 삭제 가능합니다.');</script>"
		response.write "<script type='text/javascript'>window.close();</script>"
		dbget.close()	:	response.End
	end if


	ckidx = trim(request.form("ckidx") + ",")

	if Right(ckidx,1)="," then
		ckidx = Left(ckidx,Len(ckidx)-1)
	end if

	''response.write ckidx


	sqlStr = " delete from [db_shop].[dbo].tbl_fran_meachuljungsan_subdetail" + VbCrlf
	sqlStr = sqlStr + " where idx in (" + trim(ckidx) + ")"
	''response.write sqlStr
	''dbget.close()	:	response.End
	rsget.Open sqlStr, dbget, 1


	'' update Sub master ''전체 삭제시 고려.
	sqlStr = " update [db_shop].[dbo].tbl_fran_meachuljungsan_submaster"
	sqlStr = sqlStr + " set totalsellcash=0"
	sqlStr = sqlStr + " ,totalbuycash=0"
	sqlStr = sqlStr + " ,totalsuplycash=0"
	sqlStr = sqlStr + " ,totalorgsellcash=0"
	sqlStr = sqlStr + " where masteridx=" + CStr(topidx)

	rsget.Open sqlStr, dbget, 1

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

elseif mode="modidetail" then
    ''현재상태 0 수정중인 경우만 수정 가능
	cnt=0
	sqlStr = " select count(idx) as cnt from [db_shop].[dbo].tbl_fran_meachuljungsan_master"
	sqlStr = sqlStr + " where idx=" + CStr(topidx)  + VbCrlf
    sqlStr = sqlStr + " and statecd=0"

	rsget.Open sqlStr, dbget, 1
		cnt = rsget("cnt")
	rsget.Close

	if cnt<1 then
		response.write "<script type='text/javascript'>alert('수정중 상태에서만 수정 가능합니다.');</script>"
		response.write "<script type='text/javascript'>window.close();</script>"
		dbget.close()	:	response.End
	end if

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
	sqlStr = sqlStr + " set totalsellcash=0"
	sqlStr = sqlStr + " ,totalbuycash=0"
	sqlStr = sqlStr + " ,totalsuplycash=0"
	sqlStr = sqlStr + " ,totalorgsellcash=0"
	sqlStr = sqlStr + " where masteridx=" + CStr(topidx)

	rsget.Open sqlStr, dbget, 1

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

else

	response.write "잘못된 접근입니다." & mode
	response.end

end if

'// ===========================================================================
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
    if (shopdiv="12") then shopdiv="11"

    ''iTs
    ''if (shopid="streetshop874") then shopdiv="1"
    ''if (shopid="streetshop884") then shopdiv="1"
    ''29cm
    ''if (shopid="streetshop878") then shopdiv="1"
    ''텐텐
    if (shopid="cafe003") then shopdiv="1"
end function

function fnGetOnIpjumShopName(shopid, byREF shopdiv)
    sqlStr = " select socname as shopname, '13' as shopdiv "
    sqlStr = sqlStr + " from db_user.dbo.tbl_user_c "
    sqlStr = sqlStr + " where userid='"&shopid&"'"

	'response.write sqlStr &"<Br>"
    rsget.Open sqlStr, dbget, 1
    If Not rsget.EOF then
        fnGetOnIpjumShopName 	= rsget("shopname")
        shopdiv       			= rsget("shopdiv")
    end if
    rsget.close

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

function Is3PLShopid(shopid)
	dim sqlStr

	Is3PLShopid = False

	sqlStr = " select top 1 p.id as shopid "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_partner.dbo.tbl_partner p "
	sqlStr = sqlStr + " 	left join db_shop.dbo.tbl_shop_user u "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		u.userid = p.id "
	sqlStr = sqlStr + " 	join db_partner.dbo.tbl_partner_tpl t "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		p.tplcompanyid = t.tplcompanyid "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and p.id = '" + CStr(shopid) + "' "
	sqlStr = sqlStr + " 	and IsNull(p.tplcompanyid, '') <> '' "
	'' response.write sqlStr
	rsget.Open sqlStr,dbget,1
	if not rsget.EOF then
		Is3PLShopid = True
	end if
	rsget.close
end function

function IsEtcShopid(shopid)
	dim sqlStr

	IsEtcShopid = False

	sqlStr = " select top 1 p.id as shopid "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_partner.dbo.tbl_partner p "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and p.id = '" + CStr(shopid) + "' "
	sqlStr = sqlStr + " 	and p.userdiv in ('501', '503', '900', '903', '999') "
	'' response.write sqlStr
	rsget.Open sqlStr,dbget,1
	if not rsget.EOF then
		IsEtcShopid = True
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
	<% elseif (mode="etcsubadd") or (mode="etcsubdetailadd") or (mode="regpapertype") then %>
		alert('저장 되었습니다.');
		opener.location.reload();
		window.close();
	<% elseif mode="modidetail" or mode="deldetail" then %>
		alert('수정 되었습니다.');
		location.replace('<%= refer %>');
	<% else %>
		alert('저장 되었습니다.');
		// opener.popMasterEdit('<%= iid %>');
		opener.location.reload();
		window.close();
	<% end if %>

</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->
