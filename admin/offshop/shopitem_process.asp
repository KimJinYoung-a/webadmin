<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'####################################################
' Description :  오프라인 상품 등록
' History : 2009.04.07 서동석 생성
'			2011.07.07 한용민 수정
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
''****
''  discountsellprice 삭제, orgsellprice (소비자가) 추가, shopitemprice : 실판매가.
''****
dim chargeid,shopid,divcode,vatcode ,cksel ,itemlinktypearr ,onofflinkynarr
dim itemoptionarr,itempricearr,isusingarr,extbarcodearr, itemsuplyarr
dim designer,itemgubun,itemname,sellcash,suplycash ,mode,itemgubunarr,itemarr, itemidarr
dim detailidxarr, currjungsanidarr ,shopbuyprice ,orgsellpricearr ,idx , menupos
dim sellcasharr,suplycasharr,itemnoarr,designerarr ,shopbuypricearr, centermwdivarr
dim extbarcodeAlreadyExistsString ,stockitemexists ,i,cnt,sqlStr ,extbarcodeAlreadyExists
	menupos    = requestCheckVar(request("menupos"),10)
	chargeid    = requestCheckVar(request("chargeid"),32)
	shopid      = requestCheckVar(request("shopid"),32)
	divcode     = requestCheckVar(request("divcode"),3)
	vatcode     = requestCheckVar(request("vatcode"),3)
	idx         = requestCheckVar(request("idx"),10)
	cksel = request("cksel")
	mode            = requestCheckVar(request("mode"),32)
	itemgubunarr    = request("itemgubunarr")
	itemarr         = request("itemarr")
	itemidarr       = request("itemidarr")
	itemoptionarr   = request("itemoptionarr")
	orgsellpricearr = request("orgsellpricearr")   ''소비자가 추가
	itempricearr    = request("itempricearr")
	itemsuplyarr    = request("itemsuplyarr")
	isusingarr      = request("isusingarr")
	extbarcodearr   = request("extbarcodearr")
	detailidxarr    = request("detailidxarr")
	currjungsanidarr = request("currjungsanidarr")
	sellcasharr     = request("sellcasharr")
	suplycasharr    = request("suplycasharr")
	itemnoarr       = request("itemnoarr")
	designerarr     = request("designerarr")
	''discountsellpricearr = request("discountsellpricearr")
	shopbuypricearr = request("shopbuypricearr")
	itemlinktypearr         = request("itemlinktypearr")
	onofflinkynarr         = request("onofflinkynarr")
	centermwdivarr    = request("centermwdivarr")

	''개별등록(오프상품)
	designer        = requestCheckVar(request("designer"),32)
	itemgubun       = requestCheckVar(request("itemgubun"),2)
	itemname        = requestCheckVar(request("itemname"),124)
	sellcash        = requestCheckVar(request("sellcash"),20)
	suplycash       = requestCheckVar(request("suplycash"),20)
	shopbuyprice    = requestCheckVar(request("shopbuyprice"),20)

dim refer
	refer = request.ServerVariables("HTTP_REFERER")

dim shopitemNotExists

'' 일괄수정
if mode ="arrmodi" then
	itemgubunarr = Left(itemgubunarr,Len(itemgubunarr)-1)
	itemgubunarr = split(itemgubunarr,"|")
	itemarr = Left(itemarr,Len(itemarr)-1)
	itemarr = split(itemarr,"|")
	itemoptionarr = Left(itemoptionarr,Len(itemoptionarr)-1)
	itemoptionarr = split(itemoptionarr,"|")
	extbarcodearr = split(extbarcodearr,"|")
	centermwdivarr = split(centermwdivarr,"|")

	if C_ADMIN_USER then
		orgsellpricearr = Left(orgsellpricearr,Len(orgsellpricearr)-1)
		orgsellpricearr = split(orgsellpricearr,"|")
		itempricearr = Left(itempricearr,Len(itempricearr)-1)
		itempricearr = split(itempricearr,"|")
		itemsuplyarr = Left(itemsuplyarr,Len(itemsuplyarr)-1)
		itemsuplyarr = split(itemsuplyarr,"|")
		shopbuypricearr = Left(shopbuypricearr,Len(shopbuypricearr)-1)
		shopbuypricearr = split(shopbuypricearr,"|")
		onofflinkynarr = Left(onofflinkynarr,Len(onofflinkynarr)-1)
		onofflinkynarr = split(onofflinkynarr,"|")
	end if
	if C_ADMIN_USER or C_IS_Maker_Upche then
		isusingarr = Left(isusingarr,Len(isusingarr)-1)
		isusingarr = split(isusingarr,"|")
	end if

	cnt = ubound(itemarr)

	for i=0 to cnt
		''CheckBarCode Already Exists
		extbarcodeAlreadyExists = false
		if extbarcodearr(i)<>"" then
			sqlStr = " select top 1 * from [db_item].[dbo].tbl_item_option_stock" + VbCrlf
			sqlStr = sqlStr + " where barcode='" + CStr(requestCheckVar(trim(extbarcodearr(i)),32)) + "'" + VbCrlf
			sqlStr = sqlStr + " and not ("
			sqlStr = sqlStr + " 	itemgubun='" + requestCheckVar(itemgubunarr(i),2) + "'" + VbCrlf
			sqlStr = sqlStr + " 	and itemid=" + CStr(requestCheckVar(itemarr(i),10)) + "" + VbCrlf
			sqlStr = sqlStr + " 	and itemoption='" + CStr(requestCheckVar(itemoptionarr(i),4)) + "'" + VbCrlf
			sqlStr = sqlStr + " ) "

			'response.write sqlStr &"<br>"
			rsget.Open sqlStr,dbget,1

			if Not rsget.EOF then
				extbarcodeAlreadyExists = true
				extbarcodeAlreadyExistsString = extbarcodeAlreadyExistsString + requestCheckVar(extbarcodearr(i),32) + ","
			end if

			rsget.close
		end if

		if Not extbarcodeAlreadyExists then

			sqlStr = " update [db_shop].[dbo].tbl_shop_item set"
			sqlStr = sqlStr + " extbarcode='" + CStr(requestCheckVar(extbarcodearr(i),32)) + "'"
			sqlStr = sqlStr + " ,updt=getdate()"

			if C_ADMIN_USER then
				sqlStr = sqlStr + " ,shopitemprice=" + CStr(requestCheckVar(itempricearr(i),20)) + ""
				sqlStr = sqlStr + " ,orgsellprice=" + CStr(requestCheckVar(orgsellpricearr(i),20)) + ""     '' 소비자가 추가
				sqlStr = sqlStr + " ,shopsuplycash=" + CStr(requestCheckVar(itemsuplyarr(i),20)) + ""
				sqlStr = sqlStr + " ,shopbuyprice=" + requestCheckVar(shopbuypricearr(i),20) + ""
				sqlStr = sqlStr + " ,onofflinkyn='" + CStr(requestCheckVar(onofflinkynarr(i),1)) + "'"
			end if
			if C_ADMIN_USER or C_IS_Maker_Upche then
				sqlStr = sqlStr + " ,isusing='" + CStr(requestCheckVar(isusingarr(i),1)) + "'"
			end if

			sqlStr = sqlStr + " ,centermwdiv='" + CStr(requestCheckVar(centermwdivarr(i),1)) + "'"
			sqlStr = sqlStr + " where itemgubun='" + requestCheckVar(itemgubunarr(i),2) + "'"
			sqlStr = sqlStr + " and shopitemid=" + CStr(requestCheckVar(itemarr(i),10)) + ""
			sqlStr = sqlStr + " and itemoption='" + CStr(requestCheckVar(itemoptionarr(i),4)) + "'"

			'response.write sqlStr &"<br>"
			dbget.Execute sqlStr

			' 10코드 일때만 해외가격 엎어침.
			if requestCheckVar(itemgubunarr(i),2)="10" then
				'/ 사이트별 화폐단위
				sqlStr = "select" & vbcrlf
				sqlStr = sqlStr & "	e.sitename, e.currencyunit" & vbcrlf
				sqlStr = sqlStr & "	, (select top 1 exchangeRate from db_item.dbo.tbl_exchangeRate where e.sitename=sitename and e.currencyunit=currencyunit) as exchangeRate" & vbcrlf
				sqlStr = sqlStr & "	, (select top 1 multiplerate from db_item.dbo.tbl_exchangeRate where e.sitename=sitename and e.currencyunit=currencyunit) as multiplerate" & vbcrlf
				sqlStr = sqlStr & "	, (select top 1 linkPriceType from db_item.dbo.tbl_exchangeRate where e.sitename=sitename and e.currencyunit=currencyunit) as linkPriceType" & vbcrlf
				sqlStr = sqlStr & "	into #tmp_exchangeRatecurrencyunitgroup" & vbcrlf
				sqlStr = sqlStr & "	from db_item.dbo.tbl_exchangeRate e" & vbcrlf
				sqlStr = sqlStr & "	where e.sitename='WSLWEB'" & vbcrlf
				sqlStr = sqlStr & "	group by e.sitename, e.currencyunit" & vbcrlf

				'response.write sqlStr & "<br>"
				dbget.execute sqlStr

				'온라인 해외판매상품 가격 변경
				sqlstr = " if exists(" + vbcrlf
				sqlstr = sqlstr & "		select *" + vbcrlf
				sqlstr = sqlstr & "		from db_item.dbo.tbl_item_multiLang_price" + vbcrlf
				sqlstr = sqlstr & "		where itemid=N'" & CStr(requestCheckVar(itemarr(i),10)) & "'" + vbcrlf
				sqlstr = sqlstr & "		and sitename=N'WSLWEB'" + vbcrlf
				sqlstr = sqlstr & "		and currencyUnit=N'KRW'" + vbcrlf
				sqlstr = sqlstr & "	)" + vbcrlf
				sqlstr = sqlstr & "		update P " + vbcrlf
				sqlstr = sqlstr & "		set orgprice=((case when ee.linkPriceType='1' then "& requestCheckVar(itempricearr(i),20) &" else "& requestCheckVar(orgsellpricearr(i),20) &" end) * ee.multiplerate)" + vbcrlf
				sqlstr = sqlstr & "		,wonprice=((case when ee.linkPriceType='1' then "& requestCheckVar(itempricearr(i),20) &" else "& requestCheckVar(orgsellpricearr(i),20) &" end) * ee.multiplerate)" + vbcrlf
				sqlstr = sqlstr & "		,lastupdate=getdate()" + vbcrlf
				sqlstr = sqlstr & "		,lastuserid=N'"& session("ssBctId") &"'" + vbcrlf
				sqlstr = sqlstr & "		From db_item.dbo.tbl_item_multiLang_price P"+ vbcrlf
				sqlstr = sqlstr & " 	join #tmp_exchangeRatecurrencyunitgroup ee" & vbcrlf
				sqlstr = sqlstr & " 		on p.sitename = ee.sitename" & vbcrlf
				sqlstr = sqlstr & " 		and p.currencyUnit = ee.currencyUnit" & vbcrlf
				sqlstr = sqlstr & "		where P.itemid=N'" & CStr(requestCheckVar(itemarr(i),10)) & "'" + vbcrlf
				sqlstr = sqlstr & "		and P.sitename=N'WSLWEB'" + vbcrlf
				sqlstr = sqlstr & "		and P.currencyUnit=N'KRW'" + vbcrlf

				'response.write sqlStr &"<br>"
				dbget.execute sqlstr

				'임시테이블 삭제
				dbget.execute "DROP TABLE #tmp_exchangeRatecurrencyunitgroup "
			end if

			''바코드 테이블 확인

			if trim(CStr(itemarr(i)))<>"" then

				sqlStr = " select top 1 * from [db_item].[dbo].tbl_item_option_stock" + VbCrlf
				sqlStr = sqlStr + " where itemgubun='" + requestCheckVar(itemgubunarr(i),2) + "'" + VbCrlf
				sqlStr = sqlStr + " and itemid=" + CStr(requestCheckVar(itemarr(i),10)) + "" + VbCrlf
				sqlStr = sqlStr + " and itemoption='" + CStr(requestCheckVar(itemoptionarr(i),4)) + "'" + VbCrlf

				'response.write sqlStr &"<br>"
				rsget.Open sqlStr,dbget,1
					stockitemexists = (not rsget.Eof)
				rsget.close

				if (stockitemexists) then
					sqlStr = " update [db_item].[dbo].tbl_item_option_stock" + VbCrlf
					sqlStr = sqlStr + " set barcode='" + CStr(requestCheckVar(trim(extbarcodearr(i)),32)) + "'" + VbCrlf
					sqlStr = sqlStr + " where itemgubun='" + requestCheckVar(itemgubunarr(i),2) + "'" + VbCrlf
					sqlStr = sqlStr + " and itemid=" + CStr(requestCheckVar(itemarr(i),10)) + "" + VbCrlf
					sqlStr = sqlStr + " and itemoption='" + CStr(requestCheckVar(itemoptionarr(i),4)) + "'" + VbCrlf

					'response.write sqlStr &"<br>"
					dbget.Execute sqlStr
				else
					sqlStr = " insert into [db_item].[dbo].tbl_item_option_stock" + VbCrlf
					sqlStr = sqlStr + " (itemgubun,itemid,itemoption,barcode)" + VbCrlf
					sqlStr = sqlStr + " values("
					sqlStr = sqlStr + " '" + requestCheckVar(itemgubunarr(i),2) + "'," + VbCrlf
					sqlStr = sqlStr + " " + CStr(requestCheckVar(itemarr(i),10)) + "," + VbCrlf
					sqlStr = sqlStr + " '" + requestCheckVar(itemoptionarr(i),4) + "'," + VbCrlf
					sqlStr = sqlStr + " '" + requestCheckVar(trim(extbarcodearr(i)),32) + "'" + VbCrlf
					sqlStr = sqlStr + " )" + VbCrlf

					'response.write sqlStr &"<br>"
					dbget.Execute sqlStr
				end if
			end if
		end if
	next

elseif mode ="arradd" then
	itemgubunarr = Left(itemgubunarr,Len(itemgubunarr)-1)
	itemarr = Left(itemarr,Len(itemarr)-1)
	itemoptionarr = Left(itemoptionarr,Len(itemoptionarr)-1)
	itemlinktypearr = Left(itemlinktypearr,Len(itemlinktypearr)-1)
	itemgubunarr = split(itemgubunarr,"|")
	itemarr = split(itemarr,"|")
	itemoptionarr = split(itemoptionarr,"|")
	itemlinktypearr = split(itemlinktypearr,"|")

	cnt = ubound(itemarr)

	for i=0 to cnt
		shopitemNotExists = false

		sqlStr = " select count(s.shopitemid) as cnt" + VbCrlf
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item s " + VbCrlf
		sqlStr = sqlStr + " where shopitemid=" + requestCheckVar(itemarr(i),10) + VbCrlf
		sqlStr = sqlStr + " and itemoption='" + requestCheckVar(itemoptionarr(i),4) + "'" + VbCrlf
		sqlStr = sqlStr + " and itemgubun='" + requestCheckVar(itemgubunarr(i),2) + "'" + VbCrlf

		'response.write sqlStr &"<Br>"
		rsget.Open sqlStr,dbget,1
			shopitemNotExists = rsget("cnt")<1
		rsget.close

		if shopitemNotExists then
			sqlStr = " insert into [db_shop].[dbo].tbl_shop_item" + VbCrlf
			sqlStr = sqlStr + " (itemgubun,shopitemid,itemoption," + VbCrlf
			sqlStr = sqlStr + " makerid,shopitemname,shopitemoptionname,shopitemprice,orgsellprice,shopsuplycash,shopbuyprice," + VbCrlf
			sqlStr = sqlStr + " extbarcode, vatinclude,catecdl,catecdm,catecdn," + VbCrlf
			sqlStr = sqlStr + " centermwdiv"
			sqlStr = sqlStr + " )" + VbCrlf
			sqlStr = sqlStr + " select top 1 '10', i.itemid, '" + requestCheckVar(itemoptionarr(i),4) + "'"
			sqlStr = sqlStr + " , i.makerid, i.itemname, IsNull(v.optionname,'') as optname" + VbCrlf

			if itemlinktypearr(i) = "S" then
				sqlStr = sqlStr + " , i.sellcash + IsNULL(v.optaddprice,0)" + VbCrlf
			elseif itemlinktypearr(i) = "O" then
				sqlStr = sqlStr + " , i.orgprice + IsNULL(v.optaddprice,0)" + VbCrlf
			end if

			sqlStr = sqlStr + " , i.orgprice + IsNULL(v.optaddprice,0)" + VbCrlf
			sqlStr = sqlStr + " , 0, 0" + VbCrlf
			sqlStr = sqlStr + " , IsNULL(s.barcode,'') as barcode ,i.vatinclude, i.cate_large, i.cate_mid, i.cate_small " + VbCrlf
			sqlStr = sqlStr + " , Case When i.mwdiv<>'U' then i.mwdiv else isNULL(sd.defaultCenterMwdiv,'W') end " + VbCrlf
			sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i " + VbCrlf
			sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option v"
			sqlStr = sqlStr + " 	on i.itemid=v.itemid"
			sqlStr = sqlStr + " 	and v.itemoption='" + requestCheckVar(itemoptionarr(i),4) + "'" + VbCrlf
			sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option_stock s"
			sqlStr = sqlStr + " 	on s.itemgubun='10' and i.itemid=s.itemid"
			sqlStr = sqlStr + " 	and s.itemoption='" + requestCheckVar(itemoptionarr(i),4) + "'" + VbCrlf
			sqlStr = sqlStr + " left join db_shop.dbo.tbl_shop_designer sd" + VbCrlf           '''업체배송인경우 센터매입구분 추가 2014/07/16
			sqlStr = sqlStr + " 	on sd.shopid='streetshop000'"+ VbCrlf
			sqlStr = sqlStr + " 	and i.makerid=sd.makerid"+ VbCrlf
			sqlStr = sqlStr + " where i.itemid=" + requestCheckVar(itemarr(i),10) + VbCrlf

			'response.write sqlStr &"<Br>"
			dbget.Execute sqlStr
			
			''2017/05/18 추가 물류바코드.
			sqlStr = " exec  db_shop.[dbo].[sp_ten_shop_tnbarcode_update] '10',"&requestCheckVar(itemarr(i),10)&",'"&requestCheckVar(itemoptionarr(i),4)&"'"
			dbget.Execute sqlStr
		end if
	next

elseif mode ="arraddACA" then
	itemgubunarr = Left(itemgubunarr,Len(itemgubunarr)-1)
	itemidarr = Left(itemidarr,Len(itemidarr)-1)
	itemoptionarr = Left(itemoptionarr,Len(itemoptionarr)-1)
	itemgubunarr = split(itemgubunarr,"|")
	itemidarr = split(itemidarr,"|")
	itemoptionarr = split(itemoptionarr,"|")

	cnt = ubound(itemidarr)

	sqlStr = " CREATE TABLE #tmpTableInput(itemgubun char(2), itemid int, itemoption char(4)) "
	dbget.Execute sqlStr

	for i = 0 to cnt
		shopitemNotExists = false

		sqlStr = " select count(s.shopitemid) as cnt" + VbCrlf
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item s " + VbCrlf
		sqlStr = sqlStr + " where shopitemid=" + requestCheckVar(itemidarr(i),10) + VbCrlf
		sqlStr = sqlStr + " and itemoption='" + requestCheckVar(itemoptionarr(i),4) + "'" + VbCrlf
		sqlStr = sqlStr + " and itemgubun='" + requestCheckVar(itemgubunarr(i),2) + "'" + VbCrlf

		'response.write sqlStr &"<Br>"
		rsget.Open sqlStr,dbget,1
		shopitemNotExists = rsget("cnt")<1
		rsget.close

		if shopitemNotExists then
			sqlStr = " insert into #tmpTableInput(itemgubun, itemid, itemoption) "
			sqlStr = sqlStr + " values('" + requestCheckVar(itemgubunarr(i),2) + "', " & requestCheckVar(itemidarr(i),10) & ", '" + requestCheckVar(itemoptionarr(i),4) + "')"
			''response.write sqlStr &"<Br>"
			dbget.Execute sqlStr
		end if
	next

	sqlStr = " select "
	sqlStr = sqlStr + " 	'98' as itemgubun, "
	sqlStr = sqlStr + " 	i.itemid, "
	sqlStr = sqlStr + " 	IsNull(o.itemoption, '0000') as itemoption, "
	sqlStr = sqlStr + " 	i.makerid, "
	sqlStr = sqlStr + " 	i.cate_large, "
	sqlStr = sqlStr + " 	i.cate_mid, "
	sqlStr = sqlStr + " 	i.cate_small, "
	sqlStr = sqlStr + " 	i.itemname, "
	sqlStr = sqlStr + " 	IsNull(o.optionname, '') as itemoptionname, "
	sqlStr = sqlStr + " 	i.sellcash, "
	sqlStr = sqlStr + " 	i.buycash, "
	sqlStr = sqlStr + " 	i.orgprice, "
	sqlStr = sqlStr + " 	i.orgsuplycash, "
	sqlStr = sqlStr + " 	i.sailprice, "
	sqlStr = sqlStr + " 	i.sailsuplycash, "
	sqlStr = sqlStr + " 	i.sellyn, "
	sqlStr = sqlStr + " 	i.isusing, "
	sqlStr = sqlStr + " 	i.mwdiv, "
	sqlStr = sqlStr + " 	i.vatyn, "
	sqlStr = sqlStr + " 	i.deliverytype, "
	sqlStr = sqlStr + " 	IsNull(o.optaddprice,0) as optaddprice "
	sqlStr = sqlStr + " into #tmpTable "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	[ACADEMYDB].[db_academy].[dbo].[tbl_diy_item] i "
	sqlStr = sqlStr + " 	left join [ACADEMYDB].[db_academy].[dbo].[tbl_diy_item_option] o "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		i.itemid = o.itemid "
	''sqlStr = sqlStr + " 	where i.deliverytype in (1,4) and i.mwdiv <> 'U' "
	dbget.Execute sqlStr

	sqlStr = " alter table #tmpTable add primary key (itemgubun, itemid, itemoption) "
	''dbget.Execute sqlStr

	sqlStr = " alter table #tmpTableInput add primary key (itemgubun, itemid, itemoption) "
	''dbget.Execute sqlStr

	sqlStr = " insert into [db_shop].[dbo].tbl_shop_item" + VbCrlf
	sqlStr = sqlStr + " (itemgubun,shopitemid,itemoption," + VbCrlf
	sqlStr = sqlStr + " makerid,shopitemname,shopitemoptionname,shopitemprice,orgsellprice,shopsuplycash,shopbuyprice," + VbCrlf
	sqlStr = sqlStr + " extbarcode, vatinclude,catecdl,catecdm,catecdn," + VbCrlf
	sqlStr = sqlStr + " centermwdiv"
	sqlStr = sqlStr + " )" + VbCrlf
	sqlStr = sqlStr + " select T.itemgubun, T.itemid, T.itemoption "
	sqlStr = sqlStr + " , T.makerid, T.itemname, T.itemoptionname" + VbCrlf
	sqlStr = sqlStr + " , T.orgprice + IsNULL(T.optaddprice,0)" + VbCrlf
	sqlStr = sqlStr + " , T.sellcash + IsNULL(T.optaddprice,0)" + VbCrlf
	sqlStr = sqlStr + " , 0, 0" + VbCrlf
	sqlStr = sqlStr + " , NULL ,T.vatyn, T.cate_large, T.cate_mid, T.cate_small " + VbCrlf
	sqlStr = sqlStr + " , T.mwdiv " + VbCrlf
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	#tmpTableInput TI "
	sqlStr = sqlStr + " 	join #tmpTable T "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and TI.itemgubun = T.itemgubun "
	sqlStr = sqlStr + " 		and TI.itemid = T.itemid "
	sqlStr = sqlStr + " 		and TI.itemoption = T.itemoption "
	sqlStr = sqlStr + " 	left join [db_shop].[dbo].tbl_shop_item s "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and s.itemgubun = T.itemgubun "
	sqlStr = sqlStr + " 		and s.shopitemid = T.itemid "
	sqlStr = sqlStr + " 		and s.itemoption = T.itemoption "
	sqlStr = sqlStr + " 	where "
	sqlStr = sqlStr + " 		s.shopitemid is NULL "
	dbget.Execute sqlStr

	sqlStr = " drop table #tmpTableInput "
	dbget.Execute sqlStr

	sqlStr = " drop table #tmpTable "
	dbget.Execute sqlStr
	

elseif (mode="addetcoffitem") then
    dim itemid,itemoption,orgsellprice ''itemgubun, shopbuyprice
    dim shopitemprice, discountsellprice, shopsuplycash
    dim extbarcode, isusing, shopitemname, shopitemoptionname, vatinclude, makerid, centermwdiv
    dim cd1, cd2, cd3
    itemgubun   = requestCheckVar(request("itemgubun"),2)
    itemid      = requestCheckVar(request("itemid"),10)
    itemoption  = requestCheckVar(request("itemoption"),4)

    orgsellprice = requestCheckVar(request("orgsellprice"),20)
    shopitemprice = requestCheckVar(request("shopitemprice"),20)
    discountsellprice = requestCheckVar(request("discountsellprice"),20)
    shopsuplycash = requestCheckVar(request("shopsuplycash"),20)
    shopbuyprice = requestCheckVar(request("shopbuyprice"),20)
    extbarcode = requestCheckVar(request("extbarcode"),32)
    isusing = requestCheckVar(request("isusing"),1)
    shopitemname = requestCheckVar(html2db(request("shopitemname")),124)
    shopitemoptionname = requestCheckVar(html2db(request("shopitemoptionname")),96)
    vatinclude = requestCheckVar(request("vatinclude"),1)
    makerid = requestCheckVar(request("makerid"),32)
    centermwdiv = requestCheckVar(request("centermwdiv"),1)

    cd1 = requestCheckVar(request("cd1"),3)
    cd2 = requestCheckVar(request("cd2"),3)
    cd3 = requestCheckVar(request("cd3"),3)

    if (Not IsNumeric(orgsellprice)) or (orgsellprice="") then orgsellprice =0
    if (Not IsNumeric(discountsellprice)) or (discountsellprice="") then discountsellprice =0
    if (Not IsNumeric(shopsuplycash)) or (shopsuplycash="") then shopsuplycash =0
    if (Not IsNumeric(shopbuyprice)) or (shopbuyprice="") then shopbuyprice =0

    if CStr(orgsellprice)="0" then orgsellprice=shopitemprice

	sqlStr = " select top 1 shopitemid"
	sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item"
	sqlStr = sqlStr + " where itemgubun='" + itemgubun + "'"
	sqlStr = sqlStr + " order by shopitemid desc"

	rsget.Open sqlStr,dbget,1
		if not rsget.Eof then
			itemid = rsget("shopitemid")+1
		else
			itemid = 1
		end if
	rsget.close

	itemoption = "0000"

	sqlStr = " insert into [db_shop].[dbo].tbl_shop_item" + vbCrlf
	sqlStr = sqlStr + " (itemgubun,shopitemid,itemoption," + vbCrlf
	sqlStr = sqlStr + " makerid,shopitemname,shopitemoptionname,orgsellprice,shopitemprice," + vbCrlf
	sqlStr = sqlStr + " shopsuplycash,shopbuyprice, discountsellprice,"
	if cd1<>"" then
		sqlStr = sqlStr + "  catecdl," + vbCrlf
	end if

	if cd2<>"" then
		sqlStr = sqlStr + " catecdm," + vbCrlf
	end if

	if cd3<>"" then
		sqlStr = sqlStr + " catecdn," + vbCrlf
	end if

    if (centermwdiv<>"") then
        sqlStr = sqlStr + " centermwdiv," + vbCrlf
    end if

	sqlStr = sqlStr + " vatinclude)" + vbCrlf

	sqlStr = sqlStr + " values(" + vbCrlf
	sqlStr = sqlStr + " '" + itemgubun + "'" + vbCrlf
	sqlStr = sqlStr + " ," + CStr(itemid) + "" + vbCrlf
	sqlStr = sqlStr + " ,'0000'" + vbCrlf
	sqlStr = sqlStr + " ,'" + makerid + "'" + vbCrlf
	sqlStr = sqlStr + " ,'" + shopitemname + "'" + vbCrlf
	sqlStr = sqlStr + " ,'" + shopitemoptionname + "'" + vbCrlf
	sqlStr = sqlStr + " ," + CStr(orgsellprice) + "" + vbCrlf
	sqlStr = sqlStr + " ," + CStr(shopitemprice) + "" + vbCrlf
	sqlStr = sqlStr + " ," + CStr(shopsuplycash) + "" + vbCrlf
	sqlStr = sqlStr + " ," + CStr(shopbuyprice) + "" + vbCrlf
	sqlStr = sqlStr + " ,0" + vbCrlf

	if cd1<>"" then
		sqlStr = sqlStr + " ,'" + cd1 + "'" + vbCrlf
	end if

	if cd2<>"" then
		sqlStr = sqlStr + " ,'" + cd2 + "'" + vbCrlf
	end if

	if cd3<>"" then
		sqlStr = sqlStr + " ,'" + cd3 + "'" + vbCrlf
	end if

    if (centermwdiv<>"") then
        sqlStr = sqlStr + " ,'" + centermwdiv + "'" + vbCrlf
    end if

	sqlStr = sqlStr + " ,'" + vatinclude + "'" + vbCrlf
	sqlStr = sqlStr + " )" + vbCrlf

	dbget.Execute sqlStr

    ''2017/05/18 추가 물류바코드.
	sqlStr = " exec  db_shop.[dbo].[sp_ten_shop_tnbarcode_update] '"&itemgubun&"',"&itemid&",'0000'"
	dbget.Execute sqlStr
	
elseif mode="arrins" then

	sqlStr = " select top 1 statecd from [db_shop].[dbo].tbl_shop_ipchul_master"
	sqlStr = sqlStr + " where idx=" + CStr(idx)
	rsget.Open sqlStr, dbget, 1
	if not rsget.Eof then
		if rsget("statecd")<>"0" then
			response.write "<script type='text/javascript'>alert('현재 입고대기 상태가 아닙니다.');</script>"
			response.write "<script type='text/javascript'>location.replace('" + refer + "');</script>"
			dbget.close()	:	response.End
		end if
	end if
	rsget.Close

	itemgubunarr = Left(itemgubunarr,Len(itemgubunarr)-1)
	itemarr = Left(itemarr,Len(itemarr)-1)
	itemoptionarr = Left(itemoptionarr,Len(itemoptionarr)-1)
	sellcasharr = Left(sellcasharr,Len(sellcasharr)-1)
	suplycasharr = Left(suplycasharr,Len(suplycasharr)-1)
	itemnoarr = Left(itemnoarr,Len(itemnoarr)-1)
	designerarr = Left(designerarr,Len(designerarr)-1)

	itemgubunarr = split(itemgubunarr,"|")
	itemarr = split(itemarr,"|")
	itemoptionarr = split(itemoptionarr,"|")
	sellcasharr = split(sellcasharr,"|")
	suplycasharr = split(suplycasharr,"|")
	itemnoarr = split(itemnoarr,"|")
	designerarr = split(designerarr,"|")

	cnt = ubound(itemarr)

	for i=0 to cnt
		sqlStr = " insert into [db_shop].[dbo].tbl_shop_ipchul_detail" + vbCrlf
		sqlStr = sqlStr + " (masteridx,itemgubun,shopitemid,itemoption," + vbCrlf
		sqlStr = sqlStr + " designerid,sellcash,suplycash,itemno)"  + vbCrlf
		sqlStr = sqlStr + " values(" + CStr(idx)  + "," + vbCrlf
		sqlStr = sqlStr + "'" + requestCheckVar(itemgubunarr(i),2) + "'," + vbCrlf
		sqlStr = sqlStr + "" + requestCheckVar(itemarr(i),10) + "," + vbCrlf
		sqlStr = sqlStr + "'" + requestCheckVar(itemoptionarr(i),4) + "'," + vbCrlf
		sqlStr = sqlStr + "'" + requestCheckVar(designerarr(i),32) + "'," + vbCrlf
		sqlStr = sqlStr + "" + requestCheckVar(sellcasharr(i),20) + "," + vbCrlf
		sqlStr = sqlStr + "" + requestCheckVar(suplycasharr(i),20) + "," + vbCrlf
		sqlStr = sqlStr + "" + requestCheckVar(itemnoarr(i),10) + ")"

		rsget.Open sqlStr, dbget, 1
	next

	sqlStr = " update [db_shop].[dbo].tbl_shop_ipchul_master" + vbCrlf
	sqlStr = sqlStr + " set totalsellcash=T.totsell," + vbCrlf
	sqlStr = sqlStr + " totalsuplycash=T.totsupp" + vbCrlf
	sqlStr = sqlStr + " from (select sum(sellcash*itemno) as totsell, " + vbCrlf
	sqlStr = sqlStr + " sum(suplycash*itemno) as totsupp from " + vbCrlf
	sqlStr = sqlStr + " [db_shop].[dbo].tbl_shop_ipchul_detail" + vbCrlf
	sqlStr = sqlStr + " where masteridx="  + CStr(idx) + vbCrlf
	sqlStr = sqlStr + " and deleteyn='N'" + vbCrlf
	sqlStr = sqlStr + " ) as T" + vbCrlf
	sqlStr = sqlStr + " where [db_shop].[dbo].tbl_shop_ipchul_master.idx=" + CStr(idx)

	rsget.Open sqlStr, dbget, 1

elseif mode="addipchullist" then
	dim scheduledt
	scheduledt = Cstr(dateserial(request("yyyy1"),request("mm1"),request("dd1")))

	sqlStr = " insert into [db_shop].[dbo].tbl_shop_ipchul_master"
	sqlStr = sqlStr + " (chargeid,shopid,divcode,vatcode,scheduledate,statecd)"
	sqlStr = sqlStr + " values('" + chargeid + "',"
	sqlStr = sqlStr + " '" + shopid + "',"
	sqlStr = sqlStr + " '" + divcode + "',"
	sqlStr = sqlStr + " '" + vatcode + "',"
	sqlStr = sqlStr + " '" + scheduledt + "',"
	sqlStr = sqlStr + " '0')"

	rsget.Open sqlStr, dbget, 1

	sqlStr = " select ident_current('[db_shop].[dbo].tbl_shop_ipchul_master') as idx "

	rsget.Open sqlStr, dbget, 1
		idx = rsget("idx")
	rsget.close

	itemgubunarr = Left(itemgubunarr,Len(itemgubunarr)-1)
	itemarr = Left(itemarr,Len(itemarr)-1)
	itemoptionarr = Left(itemoptionarr,Len(itemoptionarr)-1)
	sellcasharr = Left(sellcasharr,Len(sellcasharr)-1)
	suplycasharr = Left(suplycasharr,Len(suplycasharr)-1)
	itemnoarr = Left(itemnoarr,Len(itemnoarr)-1)
	designerarr = Left(designerarr,Len(designerarr)-1)

	itemgubunarr = split(itemgubunarr,"|")
	itemarr = split(itemarr,"|")
	itemoptionarr = split(itemoptionarr,"|")
	sellcasharr = split(sellcasharr,"|")
	suplycasharr = split(suplycasharr,"|")
	itemnoarr = split(itemnoarr,"|")
	designerarr = split(designerarr,"|")

	cnt = ubound(itemarr)

	for i=0 to cnt
		sqlStr = " insert into [db_shop].[dbo].tbl_shop_ipchul_detail" + vbCrlf
		sqlStr = sqlStr + " (masteridx,itemgubun,shopitemid,itemoption," + vbCrlf
		sqlStr = sqlStr + " designerid,sellcash,suplycash,itemno)"  + vbCrlf
		sqlStr = sqlStr + " values(" + CStr(idx)  + "," + vbCrlf
		sqlStr = sqlStr + "'" + requestCheckVar(itemgubunarr(i),2) + "'," + vbCrlf
		sqlStr = sqlStr + "" + requestCheckVar(itemarr(i),10) + "," + vbCrlf
		sqlStr = sqlStr + "'" + requestCheckVar(itemoptionarr(i),4) + "'," + vbCrlf
		sqlStr = sqlStr + "'" + requestCheckVar(designerarr(i),32) + "'," + vbCrlf
		sqlStr = sqlStr + "" + requestCheckVar(sellcasharr(i),20) + "," + vbCrlf
		sqlStr = sqlStr + "" + requestCheckVar(suplycasharr(i),20) + "," + vbCrlf
		sqlStr = sqlStr + "" + requestCheckVar(itemnoarr(i),10) + ")"

		rsget.Open sqlStr, dbget, 1
	next

	sqlStr = " update [db_shop].[dbo].tbl_shop_ipchul_master" + vbCrlf
	sqlStr = sqlStr + " set totalsellcash=T.totsell," + vbCrlf
	sqlStr = sqlStr + " totalsuplycash=T.totsupp" + vbCrlf
	sqlStr = sqlStr + " from (select sum(sellcash*itemno) as totsell, " + vbCrlf
	sqlStr = sqlStr + " sum(suplycash*itemno) as totsupp from " + vbCrlf
	sqlStr = sqlStr + " [db_shop].[dbo].tbl_shop_ipchul_detail" + vbCrlf
	sqlStr = sqlStr + " where masteridx="  + CStr(idx) + vbCrlf
	sqlStr = sqlStr + " and deleteyn='N'" + vbCrlf
	sqlStr = sqlStr + " ) as T" + vbCrlf
	sqlStr = sqlStr + " where [db_shop].[dbo].tbl_shop_ipchul_master.idx=" + CStr(idx)

	rsget.Open sqlStr, dbget, 1

elseif mode="editselljungsanid" then
	detailidxarr  = Left(detailidxarr,Len(detailidxarr)-1)
	currjungsanidarr = Left(currjungsanidarr,Len(currjungsanidarr)-1)

	detailidxarr = split(detailidxarr,"|")
	currjungsanidarr = split(currjungsanidarr,"|")

	cnt = ubound(detailidxarr)

	for i=0 to cnt
		sqlStr = " update [db_shop].[dbo].tbl_shopjumun_detail"
		sqlStr = sqlStr + " set jungsanid='" + requestCheckVar(currjungsanidarr(i),32) + "'"
		sqlStr = sqlStr + " where idx=" + requestCheckVar(detailidxarr(i),10) + ""

		rsget.Open sqlStr, dbget, 1
	next

elseif mode="itemnamemodiarr" then
	'sqlStr = " update [db_shop].[dbo].tbl_shop_item"
	'sqlStr = sqlStr + " set shopitemname=T.itemname"
	'sqlStr = sqlStr + " , updt=getdate()"
	'sqlStr = sqlStr + " from [db_item].[dbo].tbl_item T"
	'sqlStr = sqlStr + " where [db_shop].[dbo].tbl_shop_item.itemgubun='10'"
	'sqlStr = sqlStr + " and [db_shop].[dbo].tbl_shop_item.shopitemid=T.itemid"
	'sqlStr = sqlStr + " and [db_shop].[dbo].tbl_shop_item.shopitemname<>T.itemname"
	'sqlStr = sqlStr + " and T.itemid in (" + cksel + ")"

	'rsget.Open sqlStr, dbget, 1

	sqlStr = " update S"
    sqlStr = sqlStr + " set s.shopitemname=i.itemname"
    sqlStr = sqlStr + " ,s.shopitemoptionname=IsNULL(o.optionName,'')"
    sqlStr = sqlStr + " ,s.updt=getdate()"
    sqlStr = sqlStr + " from db_shop.dbo.tbl_shop_item s"
    sqlStr = sqlStr + " 	 join db_item.dbo.tbl_item i"
    sqlStr = sqlStr + " 	on s.itemgubun='10'"
    sqlStr = sqlStr + " 	and s.shopitemid=i.itemid"
    sqlStr = sqlStr + " 	 join db_item.dbo.tbl_item_option o"
    sqlStr = sqlStr + " 	on s.itemgubun='10'"
    sqlStr = sqlStr + " 	and s.shopitemid=o.itemid"
    sqlStr = sqlStr + " 	and s.itemoption=IsNULL(o.itemoption,'0000')"
    sqlStr = sqlStr + " where (s.shopitemname<>i.itemname or s.shopitemoptionname<>IsNULL(o.optionName,''))"
    sqlStr = sqlStr + " and s.shopitemid in (" + cksel + ")"

    if (cksel<>"") then
        dbget.Execute sqlStr
    end if

elseif mode="makeridmodiarr" then
	sqlStr = " update [db_shop].[dbo].tbl_shop_item"
	sqlStr = sqlStr + " set makerid=T.makerid"
	sqlStr = sqlStr + " , updt=getdate()"
	sqlStr = sqlStr + " from [db_item].[dbo].tbl_item T"
	sqlStr = sqlStr + " where [db_shop].[dbo].tbl_shop_item.itemgubun='10'"
	sqlStr = sqlStr + " and [db_shop].[dbo].tbl_shop_item.shopitemid=T.itemid"
	sqlStr = sqlStr + " and [db_shop].[dbo].tbl_shop_item.makerid<>T.makerid"
	sqlStr = sqlStr + " and T.itemid in (" + cksel + ")"

	rsget.Open sqlStr, dbget, 1

end if

if (mode ="offitemreg") or (mode="arrins") then
	refer = refer + "&react=true"
elseif mode="addipchullist" then
	refer = "/common/offshop/shop_ipchullist.asp?menupos="&menupos&""
end if
%>

<script type='text/javascript'>

	<% if mode="addetcoffitem" then %>
		alert('저장 되었습니다.');
		opener.location.reload();
	<% else %>
		alert('저장 되었습니다.');
	<% end if %>

	<% if extbarcodeAlreadyExistsString<>"" then %>
		alert('<%= "일부 상품은 저장되지 않았습니다. 이미등록된 바코드 - " + extbarcodeAlreadyExistsString %>');
	<% end if %>

	location.replace('<%= refer %>');

</script>

<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
