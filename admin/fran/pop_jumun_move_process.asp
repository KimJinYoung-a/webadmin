<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  재고이동
' History : 2022.07.01 이상구 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%

dim PriceEditEnable
PriceEditEnable = false

dim mode
dim masteridx, shopid, moveshopid, makerid, itemgubunarr, itemidarr, itemoptionarr, sellcasharr, buycasharr, suplycasharr, itemnoarr
dim itemgubun, itemid, itemoption, sellcash, buycash, suplycash, itemno
dim scheduledt, songjangdiv, songjangno, comment
dim IsForeignShop
dim newidx, baljucode, brandlist, targetid, targetname
dim newidxTwo, idx, idxArr, found, statecd

mode = request("mode")

masteridx  = requestCheckVar(request("masteridx"), 32)
shopid  = requestCheckVar(request("shopid"), 32)
moveshopid  = requestCheckVar(request("moveshopid"), 32)
makerid  = requestCheckVar(request("makerid"), 32)

itemgubunarr = request("itemgubunarr")
itemidarr = request("itemidarr")
itemoptionarr = request("itemoptionarr")
sellcasharr = request("sellcasharr")
buycasharr = request("buycasharr")
suplycasharr = request("suplycasharr")
itemnoarr = request("itemnoarr")

scheduledt  = requestCheckVar(request("scheduledt"), 32)
songjangdiv  = requestCheckVar(request("songjangdiv"), 32)
songjangno  = requestCheckVar(request("songjangno"), 32)
comment  = html2db(requestCheckVar(request("comment"), 3200))
idx  = requestCheckVar(request("idx"), 3200)


dim refer
refer = request.ServerVariables("HTTP_REFERER")

dim sqlStr, i, cnt, buf


''response.write itemgubunarr
''dbget.close : response.end

function fnChulgoProc(masteridx)
    dim sqlStr
    dim shopid, baljuname, reguser, finishname, orgbaljucode, regname
    dim baljucode, iid, ipgodate
    dim AssignedRows
    dim divcode, ipchulflag

    finishname = html2db(session("ssBctCname"))
	reguser = session("ssBctId")
	regname = session("ssBctCname")

    sqlStr = " select top 1 * "
    sqlStr = sqlStr & " from "
    sqlStr = sqlStr & " [db_storage].[dbo].tbl_ordersheet_master "
    sqlStr = sqlStr & " where idx = " & masteridx
	rsget.Open sqlStr,dbget,1
	if not rsget.EOF  then
		found = True
        shopid = rsget("baljuid")
        baljuname = rsget("baljuname")
        orgbaljucode = rsget("baljucode")
        ipgodate = rsget("ipgodate")

        ipgodate = Left(ipgodate, 10)
	end if
	rsget.close

    divcode = "006"
    ipchulflag = "S"

    sqlStr = " select top 1 userdiv "
    sqlStr = sqlStr & " from "
    sqlStr = sqlStr & " [db_partner].[dbo].tbl_partner "
    sqlStr = sqlStr & " where id = '" & shopid & "' "
	rsget.Open sqlStr,dbget,1
	if not rsget.EOF  then
        if rsget("userdiv") = "900" then
            divcode = "999"
            ipchulflag = "E"
        end if
	end if
	rsget.close

	'1.온라인 출고 마스타
	sqlStr = " select * from [db_storage].[dbo].tbl_acount_storage_master where 1=0"

	'response.write sqlStr &"<Br>"
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew
	rsget("code") = ""
	''출고처
	rsget("socid") = shopid
	rsget("socname") = baljuname
	rsget("chargeid") = reguser
	rsget("finishname") = finishname
	rsget("divcode") = divcode
	rsget("vatcode") = "008"
	rsget("comment") = orgbaljucode + " 주문 자동출고처리"
	rsget("chargename") = regname
	rsget("ipchulflag") = ipchulflag

	rsget.update
	    iid = rsget("id")
	rsget.close

	baljucode = "SO" + Format00(6,Right(CStr(iid),6))

	sqlStr = " update [db_storage].[dbo].tbl_acount_storage_master " + VBCrlf
	sqlStr = sqlStr + " set code='" + baljucode + "'" + VBCrlf
	sqlStr = sqlStr + " where id=" + CStr(iid)

	'response.write sqlStr &"<Br>"
	rsget.Open sqlStr,dbget,1

	'''2.온라인 출고 디테일 입력
	sqlStr = " insert into [db_storage].[dbo].tbl_acount_storage_detail " + VBCrlf
	sqlStr = sqlStr + " (mastercode,itemid,itemoption,sellcash,suplycash,"
	sqlStr = sqlStr + " itemno,indt,updt,buycash,mwgubun,iitemgubun,iitemname,iitemoptionname,imakerid)"
	sqlStr = sqlStr + " select '" + baljucode + "',d.itemid, d.itemoption, d.sellcash, d.suplycash,"
	sqlStr = sqlStr + " d.realitemno*-1, getdate(),getdate(),d.buycash,d.ipgoflag,d.itemgubun,d.itemname,d.itemoptionname,d.makerid"
	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_detail d"
	''sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item i"
	''sqlStr = sqlStr + " on d.itemgubun='10' and d.itemid=i.itemid"
	sqlStr = sqlStr + " where d.masteridx=" + CStr(masteridx)
	sqlStr = sqlStr + " and deldt is null"
	sqlStr = sqlStr + " and d.realitemno<>0"

	'response.write sqlStr &"<Br>"
	rsget.Open sqlStr,dbget,1

	'// 평균매입가 => 출고내역매입가
	sqlStr = " exec [db_storage].[dbo].[usp_Ten_AvgIpgoPriceToAccoundStorageBuycash] '" & baljucode & "' "
	dbget.Execute sqlStr

	'''2.온라인 출고 마스타 업데이트
	sqlStr = " update [db_storage].[dbo].tbl_acount_storage_master " + VBCrlf
	sqlStr = sqlStr + " set executedt='" + ipgodate + "'" + VBCrlf
	sqlStr = sqlStr + " ,scheduledt='" + ipgodate + "'" + VBCrlf
	sqlStr = sqlStr + " ,totalsellcash=IsNULL(T.totsell,0)" + VBCrlf
	sqlStr = sqlStr + " ,totalsuplycash=IsNULL(T.totsupp,0)" + VBCrlf
	sqlStr = sqlStr + " ,totalbuycash=IsNULL(T.totbuy,0)" + VBCrlf
	sqlStr = sqlStr + " ,indt=getdate()" + VBCrlf
	sqlStr = sqlStr + " ,updt=getdate()" + VBCrlf
	sqlStr = sqlStr + " from (select sum(sellcash*itemno) as totsell, " + vbCrlf
	sqlStr = sqlStr + " sum(suplycash*itemno) as totsupp, " + vbCrlf
	sqlStr = sqlStr + " sum(buycash*itemno) as totbuy " + vbCrlf
	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_detail" + vbCrlf
	sqlStr = sqlStr + " where mastercode='"  + CStr(baljucode) + "'" + vbCrlf
	sqlStr = sqlStr + " and deldt is null" + vbCrlf
	sqlStr = sqlStr + " ) as T"
	sqlStr = sqlStr + " where id=" + CStr(iid)

	'response.write sqlStr &"<Br>"
	rsget.Open sqlStr,dbget,1

	''출고된 내역 한정판매설정 : 안함

	''상태변경
	sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master" + vbCrlf
	sqlStr = sqlStr + " set statecd='7'" + vbCrlf
	sqlStr = sqlStr + " ,ipgodate='" + ipgodate + "'" + vbCrlf
	sqlStr = sqlStr + " ,alinkcode='" + baljucode + "'" + vbCrlf
	sqlStr = sqlStr + " where idx=" + CStr(masteridx)

	'response.write sqlStr &"<Br>"
	rsget.Open sqlStr, dbget, 1

    ''재고반영 ''쿼리 확인
    if (baljucode<>"") then  ''2016/05/31
        sqlStr = "exec db_summary.dbo.sp_ten_recentIpChul_Update '" & baljucode & "','','',0,'',''"

        'response.write sqlStr &"<Br>"
    	dbget.Execute sqlStr, AssignedRows

		'// 매장재고 반영
		sqlStr = "exec [db_summary].[dbo].[sp_Ten_Shop_Stock_RecentLogicsIpChul_Update] '" & shopid & "', '" & baljucode & "', 'N' "
		'response.write sqlStr & "<Br>"
		dbget.Execute sqlStr
    end if
end function

if (mode = "additem") then
	if (masteridx = "") then
		buf = 0
		sqlStr = ""
		sqlStr = sqlStr + " select count(d.shopid) as cnt "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	[db_shop].[dbo].tbl_shop_designer d "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + " 	and d.shopid in ('" & shopid & "', '" & moveshopid & "') "
		sqlStr = sqlStr + " 	and d.makerid = '" & makerid & "' "
		rsget.Open sqlStr, dbget, adOpenDynamic, adLockOptimistic, adCmdText
		if Not rsget.Eof then
			buf = rsget("cnt")
		end if
		rsget.Close

		if (buf <> 2) then
			response.write "출발매장/도착매장 또는 브랜드가 잘못 입력되었습니다."
			dbget.close : response.end
		end if


		sqlStr = ""
		sqlStr = sqlStr + " INSERT INTO [db_storage].[dbo].[tbl_ordersheet_master_before_save](targetid, baljuid, scheduledate, songjangdiv, songjangno, comment) "
		sqlStr = sqlStr + " VALUES('" & shopid & "', '" & moveshopid & "', '" & scheduledt & "', '" & songjangdiv & "', '" & songjangno & "', '" & comment & "'); "
		dbget.Execute sqlStr

		sqlStr = " SELECT @@IDENTITY as masteridx "
		rsget.Open sqlStr, dbget, adOpenDynamic, adLockOptimistic, adCmdText
		if Not rsget.Eof then
			masteridx = rsget("masteridx")
		end if
		rsget.Close

		if (masteridx = "") then
			response.write "잘못된 접속입니다.[0]"
			dbget.close : response.end
		end if
	else
		sqlStr = ""
		sqlStr = sqlStr + " UPDATE [db_storage].[dbo].[tbl_ordersheet_master_before_save] "
		sqlStr = sqlStr + " set updt = getdate() "
		sqlStr = sqlStr + " 		, scheduledate = '" & scheduledt & "', songjangdiv = '" & songjangdiv & "', songjangno = '" & songjangno & "', comment = '" & comment & "' "
		sqlStr = sqlStr + " 		where idx = " & masteridx
		dbget.Execute sqlStr
	end if

	itemgubunarr = Split(itemgubunarr, "|")
	itemidarr = Split(itemidarr, "|")
	itemoptionarr = Split(itemoptionarr, "|")
	sellcasharr = Split(sellcasharr, "|")
	buycasharr = Split(buycasharr, "|")
	suplycasharr = Split(suplycasharr, "|")
	itemnoarr = Split(itemnoarr, "|")

	for i = 0 to UBound(itemgubunarr) - 1
		if Trim(itemgubunarr(i)) <> "" then
			itemgubun = Trim(itemgubunarr(i))
			itemid = Trim(itemidarr(i))
			itemoption = Trim(itemoptionarr(i))
			sellcash = Trim(sellcasharr(i))
			buycash = Trim(buycasharr(i))
			suplycash = Trim(suplycasharr(i))
			itemno = Trim(itemnoarr(i))

			sqlStr = ""
			sqlStr = sqlStr + " IF EXISTS(select top 1 idx from [db_storage].[dbo].[tbl_ordersheet_detail_before_save] where masteridx = " & masteridx & " and itemgubun = '" & itemgubun & "' and itemid = " & itemid & " and itemoption = '" & itemoption & "') "
			sqlStr = sqlStr + " BEGIN "
			sqlStr = sqlStr + "		UPDATE [db_storage].[dbo].[tbl_ordersheet_detail_before_save] "
			sqlStr = sqlStr + "		SET updt = getdate()"
			if (PriceEditEnable = True) then
				sqlStr = sqlStr + "			, sellcash = " & sellcash
				sqlStr = sqlStr + "			, buycash = " & buycash
				sqlStr = sqlStr + "			, suplycash = " & suplycash
			end if

			sqlStr = sqlStr + "			, itemno = " & itemno
			sqlStr = sqlStr + "		WHERE masteridx = " & masteridx & " and itemgubun = '" & itemgubun & "' and itemid = " & itemid & " and itemoption = '" & itemoption & "' "
			sqlStr = sqlStr + " END "
			sqlStr = sqlStr + " ELSE "
			sqlStr = sqlStr + " BEGIN "
			sqlStr = sqlStr + "		INSERT INTO [db_storage].[dbo].[tbl_ordersheet_detail_before_save](masteridx, itemgubun, itemid, itemoption, sellcash, buycash, suplycash, itemno) "
			sqlStr = sqlStr + "		VALUES(" & masteridx & ", '" & itemgubun & "', " & itemid & ", '" & itemoption & "'"
			if (PriceEditEnable = True) then
				sqlStr = sqlStr + "			, " & sellcash & ", " & buycash & ", " & suplycash
			else
				sqlStr = sqlStr + "			, 0, 0, 0"
			end if

			sqlStr = sqlStr + "			, " & itemno & ")"
			sqlStr = sqlStr + " END "
			dbget.Execute sqlStr

		end if
	next

	sqlStr = ""
	sqlStr = sqlStr + " IF EXISTS(select top 1 idx from [db_storage].[dbo].[tbl_ordersheet_detail_before_save] where masteridx = " & masteridx & " and sellcash = 0 and suplycash = 0 and buycash = 0) "
	sqlStr = sqlStr + " BEGIN "
	sqlStr = sqlStr + "		update b "
	sqlStr = sqlStr + "		set b.sellcash = T.shopitemprice, b.suplycash = (case when T.shopbuyprice < T.shopsuplycash then T.shopsuplycash else T.shopbuyprice end), b.buycash = T.shopsuplycash "
	sqlStr = sqlStr + "		from "
	sqlStr = sqlStr + "			[db_storage].[dbo].[tbl_ordersheet_detail_before_save] b "
	sqlStr = sqlStr + "			join ( "
	sqlStr = sqlStr + "				select "
	sqlStr = sqlStr + "					b.masteridx, b.itemgubun, b.itemid, b.itemoption "
	sqlStr = sqlStr + "					, s.shopitemprice "
	sqlStr = sqlStr + "					, (case "
	sqlStr = sqlStr + "							when s.shopbuyprice <> 0 then s.shopbuyprice "
	sqlStr = sqlStr + "							when IsNull(m.defaultsuplymargin,0) = 0 and IsNull(m.defaultmargin,0) = 0 then Round(s.shopitemprice * (100.0 - (35 - 5))/100, 0) "
	sqlStr = sqlStr + "							when IsNull(m.defaultsuplymargin,0) = 0 and IsNull(m.defaultmargin,0) <> 0 then Round(s.shopitemprice * (100.0 - (m.defaultmargin - 5))/100, 0) "
	sqlStr = sqlStr + "							when IsNull(m.defaultsuplymargin,0) <> 0 then Round(s.shopitemprice * (100.0 - m.defaultsuplymargin)/100, 0) "
	sqlStr = sqlStr + "							else s.shopitemprice end) as shopbuyprice "
	sqlStr = sqlStr + "					, (case "
	sqlStr = sqlStr + "							when s.shopsuplycash <> 0 then s.shopsuplycash "
	sqlStr = sqlStr + "							when IsNull(i.mwdiv, '') = 'M' and IsNull(i.buycash, 0) <> 0 and IsNull(m.comm_cd,'') <> 'B012' and IsNull(m.comm_cd,'') <> 'B022' then Round(IsNull(i.buycash,0),0) + Round(IsNull(o.optaddprice,0),0) "
	sqlStr = sqlStr + "							when IsNull(m.defaultmargin,0) = 0 then Round(s.shopitemprice * (100.0 - 35)/100, 0) "
	sqlStr = sqlStr + "							when IsNull(m.defaultmargin,0) <> 0 then Round(s.shopitemprice * (100.0 - IsNull(m.defaultmargin,0))/100, 0) "
	sqlStr = sqlStr + "							else s.shopitemprice end) as shopsuplycash "
	sqlStr = sqlStr + "				from "
	sqlStr = sqlStr + "					[db_storage].[dbo].[tbl_ordersheet_detail_before_save] b "
	sqlStr = sqlStr + "					join [db_shop].[dbo].tbl_shop_item s "
	sqlStr = sqlStr + "					on "
	sqlStr = sqlStr + "						1 = 1 "
	sqlStr = sqlStr + "						and b.masteridx = " & masteridx & " "
	sqlStr = sqlStr + "						and b.itemgubun = s.itemgubun "
	sqlStr = sqlStr + "						and b.itemid = s.shopitemid "
	sqlStr = sqlStr + "						and b.itemoption = s.itemoption "
	sqlStr = sqlStr + "					left join [db_shop].[dbo].tbl_shop_designer m "
	sqlStr = sqlStr + "					on	 "
	sqlStr = sqlStr + "						1 = 1 "
	sqlStr = sqlStr + "						and m.shopid = '" & shopid & "' "
	sqlStr = sqlStr + "						and m.makerid = s.makerid "
	sqlStr = sqlStr + "					left join [db_item].[dbo].tbl_item i "
	sqlStr = sqlStr + "					on "
	sqlStr = sqlStr + "						1 = 1 "
	sqlStr = sqlStr + "						and s.itemgubun = '10' "
	sqlStr = sqlStr + "						and s.shopitemid = i.itemid "
	sqlStr = sqlStr + "					left join [db_item].[dbo].tbl_item_option o "
	sqlStr = sqlStr + "					on "
	sqlStr = sqlStr + "						1 = 1 "
	sqlStr = sqlStr + "						and s.itemgubun='10' "
	sqlStr = sqlStr + "						and s.shopitemid = o.itemid "
	sqlStr = sqlStr + "						and s.itemoption=o.itemoption "
	sqlStr = sqlStr + "			) T "
	sqlStr = sqlStr + "			on "
	sqlStr = sqlStr + "				1 = 1 "
	sqlStr = sqlStr + "				and b.masteridx = T.masteridx "
	sqlStr = sqlStr + "				and b.itemgubun = T.itemgubun "
	sqlStr = sqlStr + "				and b.itemid = T.itemid "
	sqlStr = sqlStr + "				and b.itemoption = T.itemoption "
	sqlStr = sqlStr + "		where "
	sqlStr = sqlStr + "			(b.sellcash = 0) and (b.suplycash = 0) and (b.buycash = 0) "
	sqlStr = sqlStr + " END "
	''response.write sqlStr
	''response.end
	dbget.Execute sqlStr

	refer = "/admin/fran/pop_jumun_move.asp?masteridx=" & masteridx & "&makerid=" & makerid
elseif (mode = "saveorder") then

	'// ========================================================================
	'// 0. 검증
	'// ========================================================================

	IsForeignShop = False

	sqlStr = "select top 1"
	sqlStr = sqlStr & " u.userid, u.shopname, isNULL(u.currencyUnit,'USD') as currencyUnit, isnull(u.countrylangcd,'EN') as countrylangcd"
	sqlStr = sqlStr & " , loginsite, isNULL(r.exchangeRate,1120) as exchangeRate"
	sqlStr = sqlStr & " from db_shop.dbo.tbl_shop_user u"
	sqlStr = sqlStr & " join db_item.dbo.tbl_exchangeRate r"
	sqlStr = sqlStr & " 	on u.currencyUnit = r.currencyUnit"
	sqlStr = sqlStr & " 	and u.countrylangcd = r.countrylangcd"
	sqlStr = sqlStr & " 	and r.sitename='WSLWEB'"
	sqlStr = sqlStr & " where u.isusing = 'Y' and u.userid in ('" & shopid & "', '" & moveshopid & "') "
	sqlStr = sqlStr & " 	and (isNULL(u.currencyUnit,'USD') <> 'KRW' or isnull(u.countrylangcd,'EN') <> 'KR') "

	''response.write sqlStr &"<br>"
	rsget.Open sqlStr,dbget,1
	if not rsget.EOF  then
		IsForeignShop = True
	end if
	rsget.close

	if IsForeignShop then
		response.write "에러 : 해외샵은 등록불가입니다[작업 안되어 있음]."
		dbget.close : response.end
	end if

	targetid = "10x10"
	targetname = "텐바이텐"

	'// ========================================================================
	'// 1. 반품주문서 작성
	'// ========================================================================
	sqlStr = "insert into [db_storage].[dbo].tbl_ordersheet_master("
	sqlStr = sqlStr & " 	targetid, baljuid, reguser, targetname, baljuname, regname"
	sqlStr = sqlStr & " 	, divcode"
	sqlStr = sqlStr & " 	, totalsellcash, totalsuplycash, totalbuycash, jumunsellcash, jumunsuplycash, jumunbuycash, vatinclude, regdate, updt, scheduledate, songjangdiv, songjangno"
	sqlStr = sqlStr & " 	, statecd, comment, brandlist, cwFlag, currencyUnit"
	sqlStr = sqlStr & " )"
	sqlStr = sqlStr & " select top 1"
	sqlStr = sqlStr & " 	'10x10', b.targetid as baljuid, '" & session("ssBctId") & "', '텐바이텐', su.shopname, '" & html2db(session("ssBctCname")) & "'"
	sqlStr = sqlStr & " 	, (case when su.shopdiv = 1 then '501' else '503' end) as divcode"
	sqlStr = sqlStr & " 	, 0, 0, 0, 0, 0, 0, 'Y', getdate(), getdate(), b.scheduledate, b.songjangdiv, b.songjangno"
	sqlStr = sqlStr & " 	, ' ' as statecd, b.comment, '' as brandlist, 0, 'KRW' as currencyUnit"
	sqlStr = sqlStr & " from"
	sqlStr = sqlStr & " 	[db_storage].[dbo].[tbl_ordersheet_master_before_save] b"
	sqlStr = sqlStr & " 	join [db_shop].[dbo].tbl_shop_user su"
	sqlStr = sqlStr & " 	on"
	sqlStr = sqlStr & " 		b.targetid = su.userid"
	sqlStr = sqlStr & " where b.idx = " & masteridx
	dbget.Execute sqlStr

	newidx = ""
	sqlStr = " SELECT @@IDENTITY as newidx "
	rsget.Open sqlStr, dbget, adOpenDynamic, adLockOptimistic, adCmdText
	if Not rsget.Eof then
		newidx = rsget("newidx")
	end if
	rsget.Close

	if newidx = "" then
		response.write "에러 : 알 수 없는 에러[0]."
		dbget.close : response.end
	end if

	baljucode = "SJ" + Format00(6,Right(CStr(newidx),6))

	sqlStr = "update [db_storage].[dbo].tbl_ordersheet_master"
	sqlStr = sqlStr & " set updt = getdate() "
	sqlStr = sqlStr & " , baljucode = '" & baljucode & "' "
	sqlStr = sqlStr & " where idx = " & newidx
	dbget.Execute sqlStr

	sqlStr = "insert into [db_storage].[dbo].tbl_ordersheet_detail("
	sqlStr = sqlStr & " 	masteridx,makerid,itemgubun,itemid,itemoption"
	sqlStr = sqlStr & " 	,itemname,itemoptionname"
	sqlStr = sqlStr & " 	,sellcash,suplycash,buycash,baljuitemno,realitemno,baljudiv"
	sqlStr = sqlStr & " )"
	sqlStr = sqlStr & " select"
	sqlStr = sqlStr & " 	" & newidx & ", IsNull(si.makerid, i.makerid), b.itemgubun, b.itemid, b.itemoption"
	sqlStr = sqlStr & " 	, IsNull(si.shopitemname, i.itemname) AS itemname,IsNull(si.shopitemoptionname, IsNull(o.optionname, '')) AS itemoptionname"
	sqlStr = sqlStr & " 	,b.sellcash, b.suplycash, b.buycash, b.itemno*-1, b.itemno*-1, '0'"
	sqlStr = sqlStr & " from"
	sqlStr = sqlStr & " 	[db_storage].[dbo].[tbl_ordersheet_detail_before_save] b"
	sqlStr = sqlStr & " 	LEFT JOIN [db_shop].[dbo].tbl_shop_item si "
	sqlStr = sqlStr & " 	ON "
	sqlStr = sqlStr & " 		1 = 1 "
	sqlStr = sqlStr & " 		and b.itemgubun <> '10' "
	sqlStr = sqlStr & " 		AND b.itemgubun = si.itemgubun "
	sqlStr = sqlStr & " 		AND b.itemid = si.shopitemid "
	sqlStr = sqlStr & " 		AND b.itemoption = si.itemoption "
	sqlStr = sqlStr & " 	LEFT JOIN [db_item].[dbo].tbl_item i "
	sqlStr = sqlStr & " 	ON "
	sqlStr = sqlStr & " 		1 = 1 "
	sqlStr = sqlStr & " 		and b.itemgubun = '10' "
	sqlStr = sqlStr & " 		AND b.itemid = i.itemid "
	sqlStr = sqlStr & " 	LEFT JOIN [db_item].[dbo].tbl_item_option o "
	sqlStr = sqlStr & " 	ON "
	sqlStr = sqlStr & " 		1 = 1 "
	sqlStr = sqlStr & " 		and b.itemgubun = '10' "
	sqlStr = sqlStr & " 		AND b.itemid = i.itemid "
	sqlStr = sqlStr & " 		AND o.itemid = i.itemid "
	sqlStr = sqlStr & " 		AND b.itemoption = o.itemoption"
	sqlStr = sqlStr & " where"
	sqlStr = sqlStr & " 	b.masteridx = " & masteridx & " and b.itemno <> 0 "
	dbget.Execute sqlStr

	sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master" + vbCrlf
	sqlStr = sqlStr + " set jumunsellcash=IsNULL(T.totsell,0)" + vbCrlf
	sqlStr = sqlStr + " ,jumunsuplycash=IsNULL(T.totsupp,0)" + vbCrlf
	sqlStr = sqlStr + " ,jumunbuycash=IsNULL(T.totbuy,0)" + vbCrlf
	sqlStr = sqlStr + " ,totalsellcash=IsNULL(T.realtotsell,0)" + vbCrlf
	sqlStr = sqlStr + " ,totalsuplycash=IsNULL(T.realtotsupp,0)" + vbCrlf
	sqlStr = sqlStr + " ,totalbuycash=IsNULL(T.realtotbuy,0)" + vbCrlf
	sqlStr = sqlStr + " ,jumunforeign_sellcash=IsNULL(T.totforeign_sellcash,0)" + vbCrlf			'// 발주시 해외 소비자가
	sqlStr = sqlStr + " ,jumunforeign_suplycash=IsNULL(T.totforeign_suplycash,0)" + vbCrlf			'// 발주시 해외 공급가
	sqlStr = sqlStr + " ,totalforeign_sellcash=IsNULL(T.realforeign_sellcash,0)" + vbCrlf			'// 확정 해외 소비자가
	sqlStr = sqlStr + " ,totalforeign_suplycash	=IsNULL(T.realforeign_suplycash,0)" + vbCrlf		'// 확정 해외 공급가
	sqlStr = sqlStr + " from (" + vbCrlf
	sqlStr = sqlStr + " 	select sum(sellcash*baljuitemno) as totsell" + vbCrlf
	sqlStr = sqlStr + " 	,sum(suplycash*baljuitemno) as totsupp" + vbCrlf
	sqlStr = sqlStr + " 	,sum(buycash*baljuitemno) as totbuy" + vbCrlf
	sqlStr = sqlStr + " 	,sum(sellcash*realitemno) as realtotsell" + vbCrlf
	sqlStr = sqlStr + " 	,sum(suplycash*realitemno) as realtotsupp" + vbCrlf
	sqlStr = sqlStr + " 	,sum(buycash*realitemno) as realtotbuy" + vbCrlf
	sqlStr = sqlStr + " 	,sum(foreign_sellcash*baljuitemno) as totforeign_sellcash" + vbCrlf
	sqlStr = sqlStr + " 	,sum(foreign_suplycash*baljuitemno) as totforeign_suplycash" + vbCrlf
	sqlStr = sqlStr + " 	,sum(foreign_sellcash*realitemno) as realforeign_sellcash" + vbCrlf
	sqlStr = sqlStr + " 	,sum(foreign_suplycash*realitemno) as realforeign_suplycash" + vbCrlf
	sqlStr = sqlStr + "  	from [db_storage].[dbo].tbl_ordersheet_detail" + vbCrlf
	sqlStr = sqlStr + "  	where masteridx="  + CStr(newidx) + vbCrlf
	sqlStr = sqlStr + " 	and deldt is null" + vbCrlf
	sqlStr = sqlStr + " ) as T" + vbCrlf
	sqlStr = sqlStr + " where [db_storage].[dbo].tbl_ordersheet_master.idx=" + CStr(newidx)
	dbget.Execute sqlStr

	brandlist = ""
	sqlStr = " select distinct makerid from [db_storage].[dbo].tbl_ordersheet_detail"
	sqlStr = sqlStr + " where masteridx=" + CStr(newidx)

	'response.write sqlStr &"<Br>"
	rsget.Open sqlStr, dbget, 1
		do until rsget.eof
			brandlist = brandlist + rsget("makerid") + ","
			rsget.movenext
		loop
	rsget.close

	if brandlist<>"" then
		brandlist = Left(brandlist,Len(brandlist)-1)
		brandlist = Left(brandlist,255)
	end if

	sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master"
	sqlStr = sqlStr + " set brandlist='" + brandlist + "'"
	sqlStr = sqlStr + " where idx=" + CStr(newidx)
	dbget.Execute sqlStr


	'// ========================================================================
	'// 2. 출고주문서 작성
	'// ========================================================================

	sqlStr = " insert into [db_storage].[dbo].tbl_ordersheet_master("
	sqlStr = sqlStr + " 	targetid, baljuid, reguser, targetname, baljuname, regname, divcode, totalsellcash, totalsuplycash, totalbuycash, jumunsellcash, jumunsuplycash, jumunbuycash, vatinclude, regdate, updt, scheduledate, songjangdiv, songjangno, statecd, comment, brandlist, cwFlag"
	sqlStr = sqlStr + " )"
	sqlStr = sqlStr + " select m.targetid, b.baljuid, reguser, targetname, su.shopname, regname, (case when su.shopdiv = 1 then '501' else '503' end), totalsellcash*-1, totalsuplycash*-1, totalbuycash*-1, jumunsellcash*-1, jumunsuplycash*-1, jumunbuycash*-1, vatinclude, getdate(), getdate(), m.scheduledate, m.songjangdiv, m.songjangno, statecd, m.comment, brandlist, cwFlag"
	sqlStr = sqlStr + " from"
	sqlStr = sqlStr + " 	[db_storage].[dbo].tbl_ordersheet_master m"
	sqlStr = sqlStr + " 	join [db_storage].[dbo].[tbl_ordersheet_master_before_save] b on b.idx = " & masteridx
	sqlStr = sqlStr + " 	join [db_shop].[dbo].tbl_shop_user su on b.baljuid = su.userid"
	sqlStr = sqlStr + " where m.idx = " & newidx
	dbget.Execute sqlStr

	newidxTwo = ""
	sqlStr = " SELECT @@IDENTITY as newidxTwo "
	rsget.Open sqlStr, dbget, adOpenDynamic, adLockOptimistic, adCmdText
	if Not rsget.Eof then
		newidxTwo = rsget("newidxTwo")
	end if
	rsget.Close


	baljucode = "SJ" + Format00(6,Right(CStr(newidxTwo),6))

	sqlStr = "update [db_storage].[dbo].tbl_ordersheet_master"
	sqlStr = sqlStr & " set updt = getdate() "
	sqlStr = sqlStr & " , baljucode = '" & baljucode & "' "
	sqlStr = sqlStr & " where idx = " & newidxTwo
	''response.write sqlStr &"<br>"
	''dbget.close : response.end
	dbget.Execute sqlStr

	sqlStr = " insert into [db_storage].[dbo].tbl_ordersheet_detail("
	sqlStr = sqlStr + " 	masteridx, itemgubun, makerid, itemid, itemoption, itemname, itemoptionname, sellcash, suplycash, buycash, baljuitemno, realitemno, regdate, baljudiv"
	sqlStr = sqlStr + " )"
	sqlStr = sqlStr + " select " & newidxTwo & ", itemgubun, makerid, itemid, itemoption, itemname, itemoptionname, sellcash, suplycash, buycash, baljuitemno*-1, realitemno*-1, getdate(), baljudiv"
	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_detail"
	sqlStr = sqlStr + " where masteridx = " & newidx
	dbget.Execute sqlStr

	response.write "<script>alert('저장되었습니다.'); opener.location.href='/admin/fran/jumunlist.asp?menupos=520'; opener.focus(); window.close();</script>"
	dbget.close : response.end
elseif (mode = "saveorderbysheet") then
    idxArr = Split(idx, ",")

    for i = 0 to UBound(idxArr)
        idx = Trim(idxArr(i))

        if idx <> "" then
            found = False

            sqlStr = " select top 1 statecd "
            sqlStr = sqlStr & " from "
            sqlStr = sqlStr & " [db_storage].[dbo].tbl_ordersheet_master "
            sqlStr = sqlStr & " where idx = " & idx & " and deldt is NULL "
	        rsget.Open sqlStr,dbget,1
	        if not rsget.EOF  then
		        found = True
                statecd = rsget("statecd")
	        end if
	        rsget.close

            if (found = False) then
                Response.write "삭제된 주문입니다."
                dbget.close : response.end
            end if

            if (statecd >= "7") then
                Response.write "출고완료된 주문입니다."
                dbget.close : response.end
            end if

            '// ============================================
            '// 1. 재고이동 주문 생성
            '// ============================================
	        sqlStr = " insert into [db_storage].[dbo].tbl_ordersheet_master("
	        sqlStr = sqlStr + " 	targetid, baljuid, reguser, targetname, baljuname, regname, divcode, totalsellcash, totalsuplycash, totalbuycash, jumunsellcash, jumunsuplycash, jumunbuycash, vatinclude, regdate, updt, scheduledate, songjangdiv, songjangno, statecd, comment, brandlist, cwFlag"
	        sqlStr = sqlStr + " )"
	        sqlStr = sqlStr + " select m.targetid, '" & moveshopid & "', reguser, targetname, IsNull(su.shopname,'" & moveshopid & "'), regname, divcode, totalsellcash*-1, totalsuplycash*-1, totalbuycash*-1, jumunsellcash*-1, jumunsuplycash*-1, jumunbuycash*-1, vatinclude, getdate(), getdate(), m.scheduledate, m.songjangdiv, m.songjangno, statecd, m.comment, brandlist, cwFlag"
	        sqlStr = sqlStr + " from"
	        sqlStr = sqlStr + " 	[db_storage].[dbo].tbl_ordersheet_master m"
            sqlStr = sqlStr + " 	left join [db_shop].[dbo].tbl_shop_user su on '" & moveshopid & "' = su.userid"
	        sqlStr = sqlStr + " where m.idx = " & idx
            ''Response.write sqlStr
            ''dbget.close : Response.end
	        dbget.Execute sqlStr

	        newidxTwo = ""
	        sqlStr = " SELECT @@IDENTITY as newidxTwo "
	        rsget.Open sqlStr, dbget, adOpenDynamic, adLockOptimistic, adCmdText
	        if Not rsget.Eof then
		        newidxTwo = rsget("newidxTwo")
	        end if
	        rsget.Close

	        baljucode = "SJ" + Format00(6,Right(CStr(newidxTwo),6))

	        sqlStr = "update [db_storage].[dbo].tbl_ordersheet_master"
	        sqlStr = sqlStr & " set updt = getdate() "
	        sqlStr = sqlStr & " , baljucode = '" & baljucode & "' "
	        sqlStr = sqlStr & " where idx = " & newidxTwo
	        ''response.write sqlStr &"<br>"
	        ''dbget.close : response.end
	        dbget.Execute sqlStr

	        sqlStr = " insert into [db_storage].[dbo].tbl_ordersheet_detail("
	        sqlStr = sqlStr + " 	masteridx, itemgubun, makerid, itemid, itemoption, itemname, itemoptionname, sellcash, suplycash, buycash, baljuitemno, realitemno, regdate, baljudiv"
	        sqlStr = sqlStr + " )"
	        sqlStr = sqlStr + " select " & newidxTwo & ", itemgubun, makerid, itemid, itemoption, itemname, itemoptionname, sellcash, suplycash, buycash, baljuitemno*-1, realitemno*-1, getdate(), baljudiv"
	        sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_detail"
	        sqlStr = sqlStr + " where masteridx = " & idx
	        dbget.Execute sqlStr

            '// 2. 반품주문 / 재고이동주문 출고완료 전환(물류 재고변동은 없으므로 재고보정 불필요)
	        sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master"
	        sqlStr = sqlStr + " set ipgodate = '" & scheduledt & "', songjangdiv = '" & songjangdiv & "', songjangno = '" & songjangno & "'  "
	        sqlStr = sqlStr + " where idx in (" & idx & ", " & newidxTwo & ")"
	        dbget.Execute sqlStr

            Call fnChulgoProc(idx)
            Call fnChulgoProc(newidxTwo)
        end if
    next

	response.write "<script>alert('저장되었습니다.'); opener.location.reload(); opener.focus(); window.close();</script>"
	dbget.close : response.end
else
	response.write "잘못된 접속입니다."
	dbget.close : response.end
end if

%>
<script language="javascript">
// alert('저장 되었습니다.');
location.replace('<%= refer %>');
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->
