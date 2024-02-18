<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  해외 출고관리
' History : 2009.07.01 이상구 생성
'			2017.11.03 한용민 수정
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<%
dim refer
refer = request.ServerVariables("HTTP_REFERER")

dim menupos, enclistpageurl, ArrShopInfo, currencyunit, currencyChar, loginsite, shopdiv
dim mode, submode, masteridx, detailidx, orgcartoonboxno
dim title, baljudate, shopid, workstate, delivermethod, deliverpay, requestdt, deliverdt, comment, reguserid
dim cartoonboxno, cartonboxsongjangdiv, cartonboxsongjangno, cartoonboxweight, cartoonboxType, innerboxno, innerboxsongjangno, innerboxweight
dim detailidxarr, cartoonboxnoarr, cartoonboxweightarr, cartoonboxTypearr, cartonboxsongjangnoarr, innerboxnoarr, innerboxweightarr, baljudatearr, shopidarr, baljunum
dim tmpdetailidxarr, tmpcartoonboxnoarr, affectedRows, errMsg, maybediffkey, maybediffkeyENname
	menupos		= request("menupos")
	mode		= request("mode")
	submode		= request("submode")
	enclistpageurl		= request("enclistpageurl")

if (enclistpageurl <> "") and (InStr(refer, "enclistpageurl") = 0) then
	refer = refer & "&enclistpageurl=" + enclistpageurl
end if

	masteridx	= request("masteridx")
	detailidx	= request("detailidx")
	title		= html2db(request("title"))
	baljudate	= request("baljudate")
	shopid		= request("shopid")
	workstate	= request("workstate")
	delivermethod	= request("delivermethod")
	deliverpay	= request("deliverpay")
	requestdt	= request("requestdt")
	deliverdt	= request("deliverdt")
	comment		= html2db(request("comment"))
	reguserid	= session("ssBctid")
	cartoonboxno		= request("cartoonboxno")
	cartoonboxweight	= request("cartoonboxweight")
	innerboxno			= request("innerboxno")
	innerboxweight		= request("innerboxweight")
	orgcartoonboxno		= request("orgcartoonboxno")
	detailidxarr		= request("detailidxarr")
	cartoonboxnoarr		= request("cartoonboxnoarr")
	cartoonboxweightarr	= request("cartoonboxweightarr")
	cartoonboxTypearr	= request("cartoonboxTypearr")
	cartonboxsongjangnoarr	= request("cartonboxsongjangnoarr")
	innerboxnoarr		= request("innerboxnoarr")
	innerboxweightarr	= request("innerboxweightarr")
	baljudatearr		= request("baljudatearr")
	shopidarr			= request("shopidarr")
	cartonboxsongjangdiv	= request("cartonboxsongjangdiv")
	cartonboxsongjangno		= request("cartonboxsongjangno")
	innerboxsongjangno		= request("innerboxsongjangno")
	baljunum				= request("baljunum")

if (deliverpay = "") then
	deliverpay = 0
end if

'==============================================================================
dim sqlStr,i, iid

Function UpdateCartonBoxNWeight(idx)
	dim sql

	if (idx = "") then
		Exit Function
	end if

	sql = " update d "
	sql = sql + " set "
	sql = sql + " 	d.cartoonboxNweight = T.cartoonboxNweight "
	sql = sql + " from "
	sql = sql + " 	[db_storage].[dbo].tbl_cartoonbox_detail d "
	sql = sql + " 	join ( "
	sql = sql + " 		select d.masteridx, d.cartoonboxno, sum(d.innerboxweight) as cartoonboxNweight "
	sql = sql + " 		from [db_storage].[dbo].tbl_cartoonbox_detail d "
	sql = sql + " 		where masteridx = " + CStr(idx) + " "
	sql = sql + " 		group by d.masteridx, d.cartoonboxno "
	sql = sql + " 	) T "
	sql = sql + " 	on "
	sql = sql + " 		1 = 1 "
	sql = sql + " 		and d.masteridx = T.masteridx "
	sql = sql + " 		and d.cartoonboxno = T.cartoonboxno "
	sql = sql + " where "
	sql = sql + " 	d.masteridx = " + CStr(idx) + " "
	dbget.Execute sql

End Function

function fnGetMayYYYYMM(idx)
    ''정산년월 계산.
    sqlStr = " select top 1 convert(varchar(7), regdate,21) as MaybeYYYYMM "
    sqlStr = sqlStr + " from [db_storage].[dbo].tbl_cartoonbox_master "
    sqlStr = sqlStr + " where idx=" + CStr(idx)

    rsget.Open sqlStr, dbget, 1
    If Not rsget.EOF then
        fnGetMayYYYYMM = rsget("MaybeYYYYMM")
    end if
    rsget.close
end function

'브랜드 영문명 가져옴	'/2017.11.03 한용민 생성
function fnGetbrandName(userid)
	dim sqlStr, tmpsocname

	if isnull(userid) or userid="" then exit function

	sqlStr = "select top 1 socname, socname_kor"
	sqlStr = sqlStr & " from db_user.dbo.tbl_user_c"
	sqlStr = sqlStr & " where userid = '"& userid &"'"

	'response.write sqlStr & "<br>"
    rsget.Open sqlStr, dbget, 1
    If Not rsget.EOF then
        tmpsocname = rsget("socname")
    end if
    rsget.close

	fnGetbrandName = tmpsocname
end function

function fnGetShopName(shopid)
    sqlStr = " select shopname, shopdiv "
    sqlStr = sqlStr + " from db_shop.dbo.tbl_shop_user "
    sqlStr = sqlStr + " where userid='"&shopid&"'"
    rsget.Open sqlStr, dbget, 1
    If Not rsget.EOF then
        fnGetShopName = rsget("shopname")
    end if
    rsget.close

end function

function fnGetMayDiffKey(shopid, MaybeYYYYMM)
    fnGetMayDiffKey = 1

    sqlStr = " select count(*) as Maydiffkey from [db_storage].[dbo].tbl_cartoonbox_master where shopid='"&shopid&"' and convert(varchar(7), regdate,21) = '"&MaybeYYYYMM&"'"

    rsget.Open sqlStr, dbget, 1
    If Not rsget.EOF then
        fnGetMayDiffKey = rsget("Maydiffkey")
    end if
    rsget.close
end function

dim MaybeYYYYMM

if (mode = "newmaster") then

	ArrShopInfo = getoffshopuser(shopid)

	IF isArray(ArrShopInfo) then
		currencyunit = ArrShopInfo(1,0)
		currencyChar = ArrShopInfo(3,0)
		loginsite = ArrShopInfo(2,0)
		shopdiv = ArrShopInfo(12,0)
    END IF

	sqlStr = " select * from [db_storage].[dbo].tbl_cartoonbox_master where 1=0"

	'response.write sqlStr &"<Br>"
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew
	rsget("title") = title
	rsget("shopid") = shopid
	rsget("workstate") = workstate
	rsget("delivermethod") = delivermethod
	rsget("deliverpay") = deliverpay
	if (requestdt <> "") then
		rsget("requestdt") = requestdt
	end if
	if (deliverdt <> "") then
		rsget("deliverdt") = deliverdt
	end if
	rsget("comment") = comment
	rsget("reguserid") = reguserid

	rsget.update
		iid = rsget("idx")
	rsget.close

	if (title = "") then
		MaybeYYYYMM = fnGetMayYYYYMM(iid)
		maybediffkey = fnGetMayDiffKey(shopid, MaybeYYYYMM)		' 차수
		if maybediffkey = "1" then
			maybediffkeyENname = "st"
		elseif maybediffkey = "2" then
			maybediffkeyENname = "nd"
		elseif maybediffkey = "3" then
			maybediffkeyENname = "rd"
		else
			maybediffkeyENname = "th"
		end if

		'if loginsite="WSLWEB" then
			title = fnGetbrandName(shopid) & " " & MaybeYYYYMM & " " & maybediffkey & maybediffkeyENname
		'else
		'	title = fnGetShopName(shopid) + " " + MaybeYYYYMM + " " + maybediffkey + "차 출고분"
		'end if

		sqlStr = " update "
		sqlStr = sqlStr + " 	[db_storage].[dbo].tbl_cartoonbox_master "
		sqlStr = sqlStr + " set "
		sqlStr = sqlStr + " 	title = '" + CStr(title) + "' "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	idx = " + CStr(iid) + "	 "
		dbget.Execute sqlStr
	end if

	if (detailidxarr <> "") then
		sqlStr = " update [db_storage].[dbo].tbl_cartoonbox_detail "
		sqlStr = sqlStr + " set "
		sqlStr = sqlStr + " 	masteridx = " + CStr(iid) + " "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + " 	and idx in (" + CStr(detailidxarr) + ") "
		'response.write "aaaaaaaaaaaa" & sqlStr
		dbget.Execute sqlStr
	end if

	refer = refer + "&idx=" + CStr(iid)

elseif (mode = "savemaster") then

	if (CStr(workstate) <> "7") then
		deliverdt = ""
	end if

	sqlStr = " update "
	sqlStr = sqlStr + " 	[db_storage].[dbo].tbl_cartoonbox_master "
	sqlStr = sqlStr + " set "
	sqlStr = sqlStr + " 	title = '" + CStr(title) + "' "
	sqlStr = sqlStr + " 	, workstate = '" + CStr(workstate) + "' "
	sqlStr = sqlStr + " 	, delivermethod = '" + CStr(delivermethod) + "' "
	sqlStr = sqlStr + " 	, deliverpay = " + CStr(deliverpay) + " "

	if (requestdt <> "") then
		sqlStr = sqlStr + " 	, requestdt = '" + CStr(requestdt) + "' "
	else
		sqlStr = sqlStr + " 	, requestdt = null "
	end if
	if (deliverdt <> "") then
		sqlStr = sqlStr + " 	, deliverdt = '" + CStr(deliverdt) + "' "
	else
		sqlStr = sqlStr + " 	, deliverdt = null "
	end if
	sqlStr = sqlStr + " 	, comment = '" + CStr(comment) + "' "
	sqlStr = sqlStr + " 	, reguserid = '" + CStr(reguserid) + "' "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	idx = " + CStr(masteridx) + "	 "
	dbget.Execute sqlStr

elseif (mode = "delmaster") then

	sqlStr = " delete from [db_storage].[dbo].tbl_cartoonbox_master "
	sqlStr = sqlStr + " where idx = " + CStr(masteridx) + " "
	dbget.Execute sqlStr

	sqlStr = " update [db_storage].[dbo].tbl_cartoonbox_detail "
	sqlStr = sqlStr + " set masteridx = null "
	sqlStr = sqlStr + " where masteridx = " + CStr(masteridx) + " "
	dbget.Execute sqlStr

	sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master "
	sqlStr = sqlStr + " set workidx = null "
	sqlStr = sqlStr + " where workidx = " + CStr(masteridx) + " "
	dbget.Execute sqlStr

	refer = "/admin/fran/cartoonbox_list.asp?menupos=" + CStr(menupos)

elseif (mode = "modifybox") then

	response.write "Error"
	response.end
	'TODO : 검증

	sqlStr = " update [db_storage].[dbo].tbl_cartoonbox_detail "
	sqlStr = sqlStr + " set "
	sqlStr = sqlStr + " 	cartoonboxno = " + CStr(cartoonboxno) + " "
	sqlStr = sqlStr + " 	, cartoonboxweight = " + CStr(cartoonboxweight) + " "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and masteridx = " + CStr(masteridx) + " "
	sqlStr = sqlStr + " 	and cartoonboxno = " + CStr(orgcartoonboxno) + " "
	'response.write "aaaaaaaaaaaa" & sqlStr
	dbget.Execute sqlStr

	sqlStr = " update [db_storage].[dbo].tbl_cartoonbox_detail "
	sqlStr = sqlStr + " set "
	sqlStr = sqlStr + " 	cartoonboxno = " + CStr(cartoonboxno) + " "
	sqlStr = sqlStr + " 	, cartoonboxweight = " + CStr(cartoonboxweight) + " "
	sqlStr = sqlStr + " 	, innerboxno = " + CStr(innerboxno) + " "
	sqlStr = sqlStr + " 	, innerboxweight = " + CStr(innerboxweight) + " "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and masteridx = " + CStr(masteridx) + " "
	sqlStr = sqlStr + " 	and idx = " + CStr(detailidx) + " "
	'response.write "aaaaaaaaaaaa" & sqlStr
	dbget.Execute sqlStr

elseif (mode = "modifycartoondetail") then

	'샵별패킹내역(박스별)

	if (detailidx = "") then
		'수기생성

		'박스는 생성할 수 없다. 로직스에서 입력된 데이타를 바탕으로 자동생성된다.(skyer9)
		'response.write "Error"
		'response.end

		sqlStr = " insert into [db_storage].[dbo].tbl_cartoonbox_detail( "
		sqlStr = sqlStr + " 	baljudate "
		sqlStr = sqlStr + " 	, shopid "
		sqlStr = sqlStr + " 	, cartoonboxno "
		sqlStr = sqlStr + " 	, cartoonboxweight "
		sqlStr = sqlStr + " 	, innerboxno "
		sqlStr = sqlStr + " 	, innerboxweight "
		sqlStr = sqlStr + " ) "
		sqlStr = sqlStr + " values( "
		sqlStr = sqlStr + " 	'" + CStr(baljudate) + "' "
		sqlStr = sqlStr + " 	, '" + CStr(shopid) + "' "
		sqlStr = sqlStr + " 	, " + CStr(cartoonboxno) + " "
		sqlStr = sqlStr + " 	, " + CStr(cartoonboxweight) + " "
		sqlStr = sqlStr + " 	, " + CStr(innerboxno) + " "
		sqlStr = sqlStr + " 	, " + CStr(innerboxweight) + " "
		sqlStr = sqlStr + " ) "

		'response.write sqlStr & "<Br>"
		'response.end
		dbget.Execute sqlStr
	else
		sqlStr = " update [db_storage].[dbo].tbl_cartoonbox_detail "
		sqlStr = sqlStr + " set "
		sqlStr = sqlStr + " 	cartoonboxno = " + CStr(cartoonboxno) + " "
		sqlStr = sqlStr + " 	, cartoonboxweight = " + CStr(cartoonboxweight) + " "
		''sqlStr = sqlStr + " 	, innerboxno = " + CStr(innerboxno) + " "
		sqlStr = sqlStr + " 	, innerboxweight = " + CStr(innerboxweight) + " "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "

		if (masteridx <> "") then
			sqlStr = sqlStr + " 	and masteridx = " + CStr(masteridx) + " "
		end if

		sqlStr = sqlStr + " 	and idx = " + CStr(detailidx) + " "
		'response.write "aaaaaaaaaaaa" & sqlStr
		dbget.Execute sqlStr
	end if

	if (innerboxsongjangno <> "") then
		sqlStr = " update d "
		sqlStr = sqlStr + " set "
		sqlStr = sqlStr + " 	d.boxsongjangno = '" + CStr(innerboxsongjangno) + "' "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	[db_storage].[dbo].tbl_shopbalju b "
		sqlStr = sqlStr + " 	join [db_storage].[dbo].tbl_ordersheet_master m "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		1 = 1 "
		sqlStr = sqlStr + " 		and b.baljucode = m.baljucode "
		sqlStr = sqlStr + " 		and b.baljuid = m.baljuid "
		sqlStr = sqlStr + " 	join [db_storage].[dbo].tbl_ordersheet_detail d "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		m.idx = d.masteridx "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + " 	and convert(varchar(10), b.baljudate, 21) = '" + CStr(baljudate) + "' "
		sqlStr = sqlStr + " 	and b.baljuid = '" + CStr(shopid) + "' "
		sqlStr = sqlStr + " 	and IsNull(d.packingstate, '0') = '" + CStr(innerboxno) + "' "
		dbget.Execute sqlStr
	end if

	sqlStr = " update [db_storage].[dbo].tbl_cartoonbox_detail "
	sqlStr = sqlStr + " set "
	sqlStr = sqlStr + " 	cartoonboxweight = " + CStr(cartoonboxweight) + " "
	if (cartonboxsongjangdiv <> "") then
		sqlStr = sqlStr + " 	, cartonboxsongjangdiv = '" + CStr(cartonboxsongjangdiv) + "' "
	end if
	if (cartonboxsongjangno <> "") then
		sqlStr = sqlStr + " 	, cartonboxsongjangno = '" + CStr(cartonboxsongjangno) + "' "
	end if
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and shopid = '" + CStr(shopid) + "' "
	sqlStr = sqlStr + " 	and convert(varchar(10), baljudate, 21) = '" + CStr(baljudate) + "' "
	sqlStr = sqlStr + " 	and cartoonboxno = " + CStr(cartoonboxno) + " "
	'response.write "aaaaaaaaaaaa" & sqlStr
	dbget.Execute sqlStr

elseif (mode = "setrecv") Then

	sqlStr = " update "
	sqlStr = sqlStr + " [db_storage].[dbo].tbl_cartoonbox_detail "
	sqlStr = sqlStr + " set shopReceive = 'Y', shopReceiveUserID = '" & reguserid & "' "
	sqlStr = sqlStr + " where shopid = '" & shopid & "' and baljudate = '" & baljudate & "' and innerboxno = " & innerboxno & " and shopReceive = 'N' "
	''response.write "aaaaaaaaaaaa" & sqlStr
	dbget.Execute sqlStr, affectedRows

	If (affectedRows = 1) Then

		sqlStr = " update c "
		sqlStr = sqlStr + " set c.logischulgo = c.logischulgo + T.itemno "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	[db_summary].[dbo].[tbl_current_shopstock_summary] c "
		sqlStr = sqlStr + " 	join ( "
		sqlStr = sqlStr + " 		select b.baljuid as shopid, d.itemgubun, d.itemid, d.itemoption, sum(d.realitemno) as itemno "
		sqlStr = sqlStr + " 		from "
		sqlStr = sqlStr + " 			[db_storage].[dbo].tbl_shopbalju b "
		sqlStr = sqlStr + " 			join [db_storage].[dbo].tbl_ordersheet_master m "
		sqlStr = sqlStr + " 			on "
		sqlStr = sqlStr + " 				b.baljucode = m.baljucode "
		sqlStr = sqlStr + " 			join [db_storage].[dbo].tbl_ordersheet_detail d "
		sqlStr = sqlStr + " 			on "
		sqlStr = sqlStr + " 				m.idx = d.masteridx "
		sqlStr = sqlStr + " 		where "
		sqlStr = sqlStr + " 			1 = 1 "
		sqlStr = sqlStr + " 			and DateDiff(day, b.baljudate, '" & baljudate & "') = 0 "
		sqlStr = sqlStr + " 			and b.baljuid = '" & shopid & "' "
		sqlStr = sqlStr + " 			and d.packingstate = " & innerboxno
		sqlStr = sqlStr + " 		group by "
		sqlStr = sqlStr + " 			b.baljuid, d.itemgubun, d.itemid, d.itemoption "
		sqlStr = sqlStr + " 	) T "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		1 = 1 "
		sqlStr = sqlStr + " 		and c.shopid = T.shopid "
		sqlStr = sqlStr + " 		and c.itemgubun = T.itemgubun "
		sqlStr = sqlStr + " 		and c.itemid = T.itemid "
		sqlStr = sqlStr + " 		and c.itemoption = T.itemoption "
		''response.write "aaaaaaaaaaaa" & sqlStr
		dbget.Execute sqlStr, affectedRows
	End If

	''response.end
elseif (mode = "saveselectedbox") then

	tmpdetailidxarr = "0" + Replace(detailidxarr, "|", ",")
	tmpcartoonboxnoarr = "0" + Replace(cartoonboxnoarr, "|", ",")

	detailidxarr 		= split(detailidxarr,"|")
	cartoonboxnoarr 	= split(cartoonboxnoarr,"|")
	innerboxnoarr 		= split(innerboxnoarr,"|")
	innerboxweightarr 	= split(innerboxweightarr,"|")

	baljudatearr 		= split(baljudatearr,"|")
	shopidarr 			= split(shopidarr,"|")

	'하나의 주문서는 하나의 작업에 전부 지정되어야한다.
	'하나의 주문서가 서로 다른 작업에 지정되어서는 않된다.
	sqlStr = " select top 1 b.baljuid, convert(varchar(10),b.baljudate,21) as baljudate, b.baljucode, d.packingstate as boxno, m.workidx "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_storage.dbo.tbl_shopbalju b "
	sqlStr = sqlStr + " 	join db_storage.dbo.tbl_ordersheet_master m "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and b.baljuid = '" + CStr(shopid) + "' "
	sqlStr = sqlStr + " 		and b.baljucode = m.baljucode "
	sqlStr = sqlStr + " 	join db_storage.dbo.tbl_ordersheet_detail d "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		m.idx = d.masteridx "
	sqlStr = sqlStr + " 	join db_storage.dbo.tbl_cartoonbox_detail cd "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and cd.shopid = b.baljuid "
	sqlStr = sqlStr + " 		and convert(varchar(10),cd.baljudate,21) = convert(varchar(10),b.baljudate,21) "
	sqlStr = sqlStr + " 		and d.packingstate = cd.innerboxno "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and m.deldt is null "
	sqlStr = sqlStr + " 	and d.deldt is null "
	sqlStr = sqlStr + " 	and IsNull(d.packingstate, 0) <> 0 "
	sqlStr = sqlStr + " 	and IsNull(m.workidx, 0) <> 0 "
	sqlStr = sqlStr + " 	and IsNull(m.workidx, 0) <> " + CStr(masteridx) + " "
	sqlStr = sqlStr + " 	and cd.idx in (" + CStr(tmpdetailidxarr) + ") "
	'response.write "aaaaaaaaaaaa" & sqlStr
	rsget.Open sqlStr, dbget, 1

	errMsg = ""
	if not rsget.EOF  then
		errMsg = "" + CStr(rsget("baljuid")) + " 삽의 " + CStr(rsget("baljudate")) + " 일자 " + CStr(rsget("boxno")) + " 번 박스(주문번호 : " + CStr(rsget("baljucode")) + ")가 이미 " + CStr(rsget("workidx")) + " 번 작업에 등록되어 있습니다."
		errMsg = errMsg + "\n\n하나의 주문서는 하나의 작업에 모두 등록되어야 합니다."
	end if
	rsget.close

	if (errMsg <> "") then
		response.write "<script>alert('" + CStr(errMsg) + "');</script>"
		response.Write errMsg
		response.end
	end if

	sqlStr = " update m "
	sqlStr = sqlStr + " set m.workidx = " + CStr(masteridx) + " "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_storage.dbo.tbl_shopbalju b "
	sqlStr = sqlStr + " 	join db_storage.dbo.tbl_ordersheet_master m "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and b.baljuid = '" + CStr(shopid) + "' "
	sqlStr = sqlStr + " 		and b.baljucode = m.baljucode "
	sqlStr = sqlStr + " 	join db_storage.dbo.tbl_ordersheet_detail d "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		m.idx = d.masteridx "
	sqlStr = sqlStr + " 	join db_storage.dbo.tbl_cartoonbox_detail cd "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and cd.shopid = b.baljuid "
	sqlStr = sqlStr + " 		and convert(varchar(10),cd.baljudate,21) = convert(varchar(10),b.baljudate,21) "
	sqlStr = sqlStr + " 		and d.packingstate = cd.innerboxno "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and m.deldt is null "
	sqlStr = sqlStr + " 	and d.deldt is null "
	sqlStr = sqlStr + " 	and IsNull(d.packingstate, 0) <> 0 "
	sqlStr = sqlStr + " 	and cd.idx in (" + CStr(tmpdetailidxarr) + ") "
	dbget.Execute sqlStr

	for i = 0 to UBound(detailidxarr)

		if (Trim(detailidxarr(i)) <> "") then

			if (Trim(detailidxarr(i))*1 = 0) then

				'박스는 생성할 수 없다. 로직스에서 입력된 데이타를 바탕으로 자동생성된다.(skyer9)
				response.write "Error"
				response.end

				sqlStr = " insert into [db_storage].[dbo].tbl_cartoonbox_detail( "
				sqlStr = sqlStr + " 	masteridx "
				sqlStr = sqlStr + " 	, baljudate "
				sqlStr = sqlStr + " 	, shopid "
				sqlStr = sqlStr + " 	, cartoonboxno "
				sqlStr = sqlStr + " 	, cartoonboxweight "
				sqlStr = sqlStr + " 	, innerboxno "
				sqlStr = sqlStr + " 	, innerboxweight "
				sqlStr = sqlStr + " ) "
				sqlStr = sqlStr + " values( "
				sqlStr = sqlStr + " 	" + CStr(masteridx) + " "
				sqlStr = sqlStr + " 	, '" + CStr(baljudatearr(i)) + "' "
				sqlStr = sqlStr + " 	, '" + CStr(shopidarr(i)) + "' "
				sqlStr = sqlStr + " 	, " + CStr(cartoonboxnoarr(i)) + " "
				sqlStr = sqlStr + " 	, 0 "
				sqlStr = sqlStr + " 	, " + CStr(innerboxnoarr(i)) + " "
				sqlStr = sqlStr + " 	, " + CStr(innerboxweightarr(i)) + " "
				sqlStr = sqlStr + " ) "
				'response.write "aaaaaaaaaaaa" & sqlStr
				dbget.Execute sqlStr

			else

				sqlStr = " update [db_storage].[dbo].tbl_cartoonbox_detail "
				sqlStr = sqlStr + " set "
				sqlStr = sqlStr + " 	masteridx = " + CStr(masteridx) + " "
				sqlStr = sqlStr + " 	, cartoonboxno = " + CStr(cartoonboxnoarr(i)) + " "
				sqlStr = sqlStr + " 	, cartoonboxweight = 0 "
				sqlStr = sqlStr + " 	, innerboxno = " + CStr(innerboxnoarr(i)) + " "
				sqlStr = sqlStr + " 	, innerboxweight = " + CStr(innerboxweightarr(i)) + " "
				sqlStr = sqlStr + " where "
				sqlStr = sqlStr + " 	1 = 1 "
				sqlStr = sqlStr + " 	and idx = " + CStr(detailidxarr(i)) + " "
				'response.write "aaaaaaaaaaaa" & sqlStr
				dbget.Execute sqlStr

			end if

		end if

	next

	sqlStr = " update [db_storage].[dbo].tbl_cartoonbox_detail "
	sqlStr = sqlStr + " set "
	sqlStr = sqlStr + " 	cartoonboxweight = 0 "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and masteridx = " + CStr(masteridx) + " "
	sqlStr = sqlStr + " 	and cartoonboxno in (" + CStr(tmpcartoonboxnoarr) + ") "
	'response.write "aaaaaaaaaaaa" & sqlStr
	dbget.Execute sqlStr

	if (submode = "popup") then
		Call UpdateCartonBoxNWeight(masteridx)

		response.write "<script>alert('저장되었습니다.');</script>"
		response.write "<script>opener.location.reload(); opener.focus(); window.close();</script>"
		dbget.close
		response.end
	end if

elseif (mode = "deselectbox") then

	sqlStr = " update m "
	sqlStr = sqlStr + " set m.workidx = NULL "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_storage.dbo.tbl_shopbalju b "
	sqlStr = sqlStr + " 	join db_storage.dbo.tbl_ordersheet_master m "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and b.baljuid = '" + CStr(shopid) + "' "
	sqlStr = sqlStr + " 		and b.baljucode = m.baljucode "
	sqlStr = sqlStr + " 	join db_storage.dbo.tbl_ordersheet_detail d "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		m.idx = d.masteridx "
	sqlStr = sqlStr + " 	join db_storage.dbo.tbl_cartoonbox_detail cd "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and cd.shopid = b.baljuid "
	sqlStr = sqlStr + " 		and convert(varchar(10),cd.baljudate,21) = convert(varchar(10),b.baljudate,21) "
	sqlStr = sqlStr + " 		and d.packingstate = cd.innerboxno "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and m.deldt is null "
	sqlStr = sqlStr + " 	and d.deldt is null "
	sqlStr = sqlStr + " 	and IsNull(d.packingstate, 0) <> 0 "
	sqlStr = sqlStr + " 	and cd.masteridx = " + CStr(masteridx) + " "
	dbget.Execute sqlStr

	sqlStr = " update m "
	sqlStr = sqlStr + " set m.workidx = " + CStr(masteridx) + " "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_storage.dbo.tbl_shopbalju b "
	sqlStr = sqlStr + " 	join db_storage.dbo.tbl_ordersheet_master m "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and b.baljuid = '" + CStr(shopid) + "' "
	sqlStr = sqlStr + " 		and b.baljucode = m.baljucode "
	sqlStr = sqlStr + " 	join db_storage.dbo.tbl_ordersheet_detail d "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		m.idx = d.masteridx "
	sqlStr = sqlStr + " 	join db_storage.dbo.tbl_cartoonbox_detail cd "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and cd.shopid = b.baljuid "
	sqlStr = sqlStr + " 		and convert(varchar(10),cd.baljudate,21) = convert(varchar(10),b.baljudate,21) "
	sqlStr = sqlStr + " 		and d.packingstate = cd.innerboxno "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and m.deldt is null "
	sqlStr = sqlStr + " 	and d.deldt is null "
	sqlStr = sqlStr + " 	and IsNull(d.packingstate, 0) <> 0 "
	sqlStr = sqlStr + " 	and cd.idx not in (" + CStr(detailidxarr) + ") "
	sqlStr = sqlStr + " 	and cd.masteridx = " + CStr(masteridx) + " "
	dbget.Execute sqlStr

	sqlStr = " update [db_storage].[dbo].tbl_cartoonbox_detail "
	sqlStr = sqlStr + " set masteridx = null "
	sqlStr = sqlStr + " where masteridx = " + CStr(masteridx) + " "
	sqlStr = sqlStr + " 	and idx in (" + CStr(detailidxarr) + ") "
	dbget.Execute sqlStr

elseif (mode = "savedetailarr") then

	detailidxarr 		= split(detailidxarr,"|")
	cartoonboxnoarr 	= split(cartoonboxnoarr,"|")
	cartoonboxweightarr = split(cartoonboxweightarr,"|")
	cartoonboxTypearr	= split(cartoonboxTypearr,"|")
	cartonboxsongjangnoarr 	= split(cartonboxsongjangnoarr,"|")
	innerboxnoarr 		= split(innerboxnoarr,"|")
	innerboxweightarr 	= split(innerboxweightarr,"|")

	for i = 0 to UBound(detailidxarr)

		if (Trim(detailidxarr(i)) <> "") then

			sqlStr = " update [db_storage].[dbo].tbl_cartoonbox_detail "
			sqlStr = sqlStr + " set "
			sqlStr = sqlStr + " 	innerboxno = " + CStr(innerboxnoarr(i)) + " "
			sqlStr = sqlStr + " 	, innerboxweight = " + CStr(innerboxweightarr(i)) + " "
			sqlStr = sqlStr + " 	, cartoonboxno = " + CStr(cartoonboxnoarr(i)) + " "        ''2016/09/12 추가. 이문재 요청
			sqlStr = sqlStr + " where "
			sqlStr = sqlStr + " 	1 = 1 "
			sqlStr = sqlStr + " 	and masteridx = " + CStr(masteridx) + " "
			sqlStr = sqlStr + " 	and idx = " + CStr(detailidxarr(i)) + " "
			'response.write "aaaaaaaaaaaa" & sqlStr
			dbget.Execute sqlStr

			if (cartoonboxweightarr(i)*1 <> -1) then
				sqlStr = " update [db_storage].[dbo].tbl_cartoonbox_detail "
				sqlStr = sqlStr + " set "
				sqlStr = sqlStr + " 	cartoonboxweight = " + CStr(cartoonboxweightarr(i)) + " "
				sqlStr = sqlStr + " 	, cartoonboxType = '" + CStr(cartoonboxTypearr(i)) + "' "
				sqlStr = sqlStr + " 	, cartonboxsongjangno = '" + CStr(cartonboxsongjangnoarr(i)) + "' "
				sqlStr = sqlStr + " where "
				sqlStr = sqlStr + " 	1 = 1 "
				sqlStr = sqlStr + " 	and masteridx = " + CStr(masteridx) + " "
				sqlStr = sqlStr + " 	and cartoonboxno = " + CStr(cartoonboxnoarr(i)) + " "
				''sqlStr = sqlStr + " 	and baljudate = (select baljudate from [db_storage].[dbo].tbl_cartoonbox_detail where idx = " & CStr(detailidxarr(i)) & " ) "
				''response.write "aaaaaaaaaaaa" & sqlStr
				dbget.Execute sqlStr
			end if

		end if

	next

elseif (mode = "refreshsupplycash") then

	sqlStr = " update c "
	sqlStr = sqlStr + " set c.totsuplycash = T.totsuplycash, c.totforeign_suplycash = T.totforeign_suplycash, c.currencyUnit = T.currencyUnit "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	[db_storage].[dbo].tbl_cartoonbox_master c "
	sqlStr = sqlStr + " 	join ( "
	sqlStr = sqlStr + " 		select "
	sqlStr = sqlStr + " 			T.cmasteridx "
	sqlStr = sqlStr + " 			, sum(d.suplycash * d.realitemno) as totsuplycash "
	sqlStr = sqlStr + " 			, sum(d.foreign_suplycash * d.realitemno) as totforeign_suplycash "
	sqlStr = sqlStr + " 			, max(m.currencyUnit) as currencyUnit "
	sqlStr = sqlStr + " 		from "
	sqlStr = sqlStr + " 			[db_storage].[dbo].tbl_shopbalju b "
	sqlStr = sqlStr + " 			JOIN [db_storage].[dbo].tbl_ordersheet_master m "
	sqlStr = sqlStr + " 			on "
	sqlStr = sqlStr + " 				b.baljucode = m.baljucode "
	sqlStr = sqlStr + " 			JOIN [db_storage].[dbo].tbl_ordersheet_detail d "
	sqlStr = sqlStr + " 			on "
	sqlStr = sqlStr + " 				1 = 1 "
	sqlStr = sqlStr + " 				and d.masteridx = m.idx "
	sqlStr = sqlStr + " 			join 	( "
	sqlStr = sqlStr + " 				select convert(varchar(10), cd.baljudate, 121) as baljudate, cd.shopid, cd.innerboxno, c.delivermethod, c.idx as cmasteridx "
	sqlStr = sqlStr + " 				from "
	sqlStr = sqlStr + " 					[db_storage].[dbo].tbl_cartoonbox_master c "
	sqlStr = sqlStr + " 					left join [db_storage].[dbo].tbl_cartoonbox_detail cd "
	sqlStr = sqlStr + " 					on "
	sqlStr = sqlStr + " 						c.idx = cd.masteridx "
	sqlStr = sqlStr + " 				where "
	sqlStr = sqlStr + " 					1 = 1 "
	sqlStr = sqlStr + " 					and c.idx = " & masteridx
	sqlStr = sqlStr + " 			) T "
	sqlStr = sqlStr + " 			on "
	sqlStr = sqlStr + " 				1 = 1 "
	sqlStr = sqlStr + " 				and b.baljuid = T.shopid  "
	sqlStr = sqlStr + " 				and b.baljudate >= DateAdd(d, 0, T.baljudate) "
	sqlStr = sqlStr + " 				and b.baljudate < DateAdd(d, 1, T.baljudate) "
	sqlStr = sqlStr + " 				and d.packingstate = T.innerboxno "
	sqlStr = sqlStr + " 		group by "
	sqlStr = sqlStr + " 			T.cmasteridx "
	sqlStr = sqlStr + " 	) T "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		c.idx = T.cmasteridx "
	''response.write "aaaaaaaaaaaa" & sqlStr
	dbget.Execute sqlStr

end if

Call UpdateCartonBoxNWeight(masteridx)

%>
<%= masteridx %>
<script language="javascript">
	alert('저장 되었습니다.');
	location.replace('<%= refer %>');
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->
