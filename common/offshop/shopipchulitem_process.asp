<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프샵 주문서 작성
' History : 2009.04.07 서동석 생성
'			2011.05.16 한용민 수정
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopipchulcls.asp"-->
<%
dim IS_HIDE_BUYCASH : IS_HIDE_BUYCASH = False
dim mode,itemgubunarr,itemarr ,idx ,shopbuyprice ,scheduledt ,comment ,TotalSellcash,ipchulmoveidx
dim itemoptionarr,itempricearr,chargeidarr,isusingarr,extbarcodearr, itemsuplyarr , moveshopid ,firstshopid
dim designer,itemgubun,itemname,sellcash,suplycash ,chargeid,shopid,divcode,vatcode ,moveidx , comm_cd
dim detailidxarr, currjungsanidarr, shopitemnamearr ,cksel ,songjangdiv,songjangno, isreq
dim franitempricearr, fransuplycasharr, cmsitempricearr, cmssuplycasharr ,currState , tmpcontractyn
dim sellcasharr,suplycasharr,itemnoarr,designerarr ,discountsellpricearr, shopbuypricearr
dim extbarcodeAlreadyExistsString ,stockitemexists ,extbarcodeAlreadyExists ,i,cnt,sqlStr ,menupos, waitflag
dim addshopid, newiid,newtargetid,newbaljuid
	comment = request("comment")
	menupos = requestCheckVar(request("menupos"),10)
	mode = requestCheckVar(request("mode"),32)
	itemgubunarr = request("itemgubunarr")
	itemarr = request("itemarr")
	itemoptionarr = request("itemoptionarr")
	itempricearr = request("itempricearr")
	itemsuplyarr = request("itemsuplyarr")
	chargeidarr = request("chargeidarr")
	isusingarr = request("isusingarr")
	extbarcodearr = request("extbarcodearr")
	detailidxarr  = request("detailidxarr")
	currjungsanidarr = request("currjungsanidarr")
	shopitemnamearr = (request("shopitemnamearr"))
	franitempricearr = request("franitempricearr")
	fransuplycasharr = request("fransuplycasharr")
	cmsitempricearr = request("cmsitempricearr")
	cmssuplycasharr = request("cmssuplycasharr")
	designer = requestCheckVar(request("designer"),32)
	itemgubun = requestCheckVar(request("itemgubun"),2)
	itemname = requestCheckVar(request("itemname"),124)
	sellcash = requestCheckVar(request("sellcash"),20)
	suplycash = requestCheckVar(request("suplycash"),20)
	shopbuyprice = requestCheckVar(request("shopbuyprice"),20)
	sellcasharr = request("sellcasharr")
	suplycasharr = request("suplycasharr")
	itemnoarr  = request("itemnoarr")
	designerarr = request("designerarr")
	discountsellpricearr = request("discountsellpricearr")
	shopbuypricearr = request("shopbuypricearr")
	fransuplycasharr = request("fransuplycasharr")
	chargeid = requestCheckVar(request("chargeid"),32)
	shopid = requestCheckVar(request("shopid"),32)
	divcode = requestCheckVar(request("divcode"),3)
	vatcode = requestCheckVar(request("vatcode"),3)
	idx = requestCheckVar(request("idx"),10)
	cksel = request("cksel")
	songjangdiv = requestCheckVar(request("songjangdiv"),2)
	songjangno = Left(html2db(request("songjangno")),32)
	isreq      = requestCheckVar(request("isreq"),10)
	moveshopid = requestCheckVar(request("moveshopid"),32)
	scheduledt = requestCheckVar(request("scheduledt"),30)
	firstshopid = requestCheckVar(request("firstshopid"),32)

	''작성중인경우.
	waitflag = requestCheckVar(request("waitflag"),10)

	addshopid = request("addshopid")
	''response.write addshopid : response.end

if C_ADMIN_USER or C_IS_OWN_SHOP then

'' 매장인경우
elseif (C_IS_SHOP) then
	IS_HIDE_BUYCASH = True
end if

tmpcontractyn = false

dim refer
	refer = request.ServerVariables("HTTP_REFERER")

if mode="arrins" then

	sqlStr = " select top 1 statecd from [db_shop].[dbo].tbl_shop_ipchul_master"
	sqlStr = sqlStr + " where idx=" + CStr(idx)

	'response.write sqlStr &"<br>"
	rsget.Open sqlStr, dbget, 1
	if not rsget.Eof then
	    currState = rsget("statecd")
		if currState>0 then
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
	shopbuypricearr = Left(shopbuypricearr,Len(shopbuypricearr)-1)
	itemnoarr = Left(itemnoarr,Len(itemnoarr)-1)
	designerarr = Left(designerarr,Len(designerarr)-1)

	itemgubunarr = split(itemgubunarr,"|")
	itemarr = split(itemarr,"|")
	itemoptionarr = split(itemoptionarr,"|")
	sellcasharr = split(sellcasharr,"|")
	suplycasharr = split(suplycasharr,"|")
	shopbuypricearr = split(shopbuypricearr,"|")
	itemnoarr = split(itemnoarr,"|")
	designerarr = split(designerarr,"|")

	cnt = ubound(itemarr)

	for i=0 to cnt
		sqlStr = " insert into [db_shop].[dbo].tbl_shop_ipchul_detail" + vbCrlf
		sqlStr = sqlStr + " (masteridx,itemgubun,shopitemid,itemoption," + vbCrlf
		sqlStr = sqlStr + " designerid,sellcash,suplycash,shopbuyprice,itemno, reqno)"  + vbCrlf
		sqlStr = sqlStr + " values(" + CStr(idx)  + "," + vbCrlf
		sqlStr = sqlStr + "'" + requestCheckVar(itemgubunarr(i),2) + "'," + vbCrlf
		sqlStr = sqlStr + "" + requestCheckVar(itemarr(i),10) + "," + vbCrlf
		sqlStr = sqlStr + "'" + requestCheckVar(itemoptionarr(i),4) + "'," + vbCrlf
		sqlStr = sqlStr + "'" + requestCheckVar(designerarr(i),32) + "'," + vbCrlf
		sqlStr = sqlStr + "" + requestCheckVar(sellcasharr(i),20) + "," + vbCrlf
		sqlStr = sqlStr + "" + requestCheckVar(suplycasharr(i),20) + "," + vbCrlf
		sqlStr = sqlStr + "" + requestCheckVar(shopbuypricearr(i),20) + "," + vbCrlf
		sqlStr = sqlStr + "" + requestCheckVar(itemnoarr(i),10) + + "," + vbCrlf
		if (currState=-2) then
            sqlStr = sqlStr + "" + itemnoarr(i) + vbCrlf
        else
            sqlStr = sqlStr + "0" + vbCrlf
        end if
        sqlStr = sqlStr + "" + ")"

        'response.write sqlStr &"<br>"
		dbget.Execute(sqlStr)
	next

	if IS_HIDE_BUYCASH = True and shopid <> "" then
		sqlStr = " IF EXISTS(select top 1 idx from [db_shop].[dbo].tbl_shop_ipchul_detail where masteridx = " + CStr(idx)  + " and suplycash < 0) "
		sqlStr = sqlStr + " BEGIN "
		sqlStr = sqlStr + " 	update d "
		''sqlStr = sqlStr + " 	set d.sellcash = T.sellcash, d.shopbuyprice = (case when T.suplycash < T.buycash then T.buycash else T.suplycash end), d.suplycash = T.buycash "
		sqlStr = sqlStr + " 	set d.suplycash = T.buycash "
		sqlStr = sqlStr + " 	FROM "
		sqlStr = sqlStr + " 		[db_shop].[dbo].tbl_shop_ipchul_detail d "
		sqlStr = sqlStr + " 		join ( "
		sqlStr = sqlStr + " 			select "
		sqlStr = sqlStr + " 				d.masteridx, d.itemgubun, d.shopitemid, d.itemoption "
		sqlStr = sqlStr + " 				, s.shopitemprice as sellcash "
		sqlStr = sqlStr + " 				, (case "
		sqlStr = sqlStr + " 						when s.shopbuyprice <> 0 then s.shopbuyprice "
		sqlStr = sqlStr + " 						when IsNull(m.defaultsuplymargin,0) = 0 and IsNull(m.defaultmargin,0) = 0 then Round(s.shopitemprice * (100.0 - (35 - 5))/100, 0) "
		sqlStr = sqlStr + " 						when IsNull(m.defaultsuplymargin,0) = 0 and IsNull(m.defaultmargin,0) <> 0 then Round(s.shopitemprice * (100.0 - (m.defaultmargin - 5))/100, 0) "
		sqlStr = sqlStr + " 						when IsNull(m.defaultsuplymargin,0) <> 0 then Round(s.shopitemprice * (100.0 - m.defaultsuplymargin)/100, 0) "
		sqlStr = sqlStr + " 						else s.shopitemprice end) as suplycash "
		sqlStr = sqlStr + " 				, (case "
		sqlStr = sqlStr + " 						when s.shopsuplycash <> 0 then s.shopsuplycash "
		sqlStr = sqlStr + " 						when IsNull(i.mwdiv, '') = 'M' and IsNull(i.buycash, 0) <> 0 and IsNull(m.comm_cd,'') <> 'B012' and IsNull(m.comm_cd,'') <> 'B022' then Round(IsNull(i.buycash,0),0) + Round(IsNull(o.optaddprice,0),0) "
		sqlStr = sqlStr + " 						when IsNull(m.defaultmargin,0) = 0 then Round(s.shopitemprice * (100.0 - 35)/100, 0) "
		sqlStr = sqlStr + " 						when IsNull(m.defaultmargin,0) <> 0 then Round(s.shopitemprice * (100.0 - IsNull(m.defaultmargin,0))/100, 0) "
		sqlStr = sqlStr + " 						else s.shopitemprice end) as buycash "
		sqlStr = sqlStr + " 			from "
		sqlStr = sqlStr + " 				[db_shop].[dbo].tbl_shop_ipchul_detail d "
		sqlStr = sqlStr + " 				join [db_shop].[dbo].tbl_shop_item s "
		sqlStr = sqlStr + " 				on "
		sqlStr = sqlStr + " 					1 = 1 "
		sqlStr = sqlStr + " 					and d.masteridx = " + CStr(idx)  + " "
		sqlStr = sqlStr + " 					and d.itemgubun = s.itemgubun "
		sqlStr = sqlStr + " 					and d.shopitemid = s.shopitemid "
		sqlStr = sqlStr + " 					and d.itemoption = s.itemoption "
		sqlStr = sqlStr + " 				left join [db_shop].[dbo].tbl_shop_designer m "
		sqlStr = sqlStr + " 				on	 "
		sqlStr = sqlStr + " 					1 = 1 "
		sqlStr = sqlStr + " 					and m.shopid = '" & shopid & "' "
		sqlStr = sqlStr + " 					and m.makerid = s.makerid "
		sqlStr = sqlStr + " 				left join [db_item].[dbo].tbl_item i "
		sqlStr = sqlStr + " 				on "
		sqlStr = sqlStr + " 					1 = 1 "
		sqlStr = sqlStr + " 					and s.itemgubun = '10' "
		sqlStr = sqlStr + " 					and s.shopitemid = i.itemid "
		sqlStr = sqlStr + " 				left join [db_item].[dbo].tbl_item_option o "
		sqlStr = sqlStr + " 				on "
		sqlStr = sqlStr + " 					1 = 1 "
		sqlStr = sqlStr + " 					and s.itemgubun='10' "
		sqlStr = sqlStr + " 					and s.shopitemid = o.itemid "
		sqlStr = sqlStr + " 					and s.itemoption=o.itemoption "
		sqlStr = sqlStr + " 		) T "
		sqlStr = sqlStr + " 		on "
		sqlStr = sqlStr + " 			1 = 1 "
		sqlStr = sqlStr + " 			and d.masteridx = T.masteridx "
		sqlStr = sqlStr + " 			and d.itemgubun = T.itemgubun "
		sqlStr = sqlStr + " 			and d.shopitemid = T.shopitemid "
		sqlStr = sqlStr + " 			and d.itemoption = T.itemoption "
		sqlStr = sqlStr + " 	WHERE "
		sqlStr = sqlStr + " 		d.suplycash < 0 "
		sqlStr = sqlStr + " END "
		rsget.Open sqlStr, dbget, 1
	end if

	sqlStr = " update [db_shop].[dbo].tbl_shop_ipchul_detail" + vbCrlf
	sqlStr = sqlStr + " set itemname=T.shopitemname" + vbCrlf
	sqlStr = sqlStr + " ,itemoptionname=T.shopitemoptionname" + vbCrlf
	sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item T" + vbCrlf
	sqlStr = sqlStr + " where [db_shop].[dbo].tbl_shop_ipchul_detail.masteridx=" + CStr(idx)
	sqlStr = sqlStr + " and [db_shop].[dbo].tbl_shop_ipchul_detail.itemgubun=T.itemgubun"
	sqlStr = sqlStr + " and [db_shop].[dbo].tbl_shop_ipchul_detail.shopitemid=T.shopitemid"
	sqlStr = sqlStr + " and [db_shop].[dbo].tbl_shop_ipchul_detail.itemoption=T.itemoption"

	'response.write sqlStr &"<br>"
	dbget.Execute(sqlStr)

	sqlStr = " update [db_shop].[dbo].tbl_shop_ipchul_master" + vbCrlf
	sqlStr = sqlStr + " set totalsellcash=IsNULL(T.totsell,0)" + vbCrlf
	sqlStr = sqlStr + " ,totalsuplycash=IsNULL(T.totsupp,0)" + vbCrlf
	sqlStr = sqlStr + " ,totalshopbuyprice=IsNULL(T.totshopbuy,0)" + vbCrlf
	sqlStr = sqlStr + " from ("
	sqlStr = sqlStr + "		select sum(sellcash*itemno) as totsell " + vbCrlf
	sqlStr = sqlStr + " 	,sum(suplycash*itemno) as totsupp " + vbCrlf
	sqlStr = sqlStr + " 	,sum(shopbuyprice*itemno) as totshopbuy " + vbCrlf
	sqlStr = sqlStr + " 	from " + vbCrlf
	sqlStr = sqlStr + " 	[db_shop].[dbo].tbl_shop_ipchul_detail" + vbCrlf
	sqlStr = sqlStr + " 	where masteridx="  + CStr(idx) + vbCrlf
	sqlStr = sqlStr + " 	and deleteyn='N'" + vbCrlf
	sqlStr = sqlStr + " ) as T" + vbCrlf
	sqlStr = sqlStr + " where [db_shop].[dbo].tbl_shop_ipchul_master.idx=" + CStr(idx)

	'response.write sqlStr &"<br>"
	dbget.Execute(sqlStr)

	'//기주문 업데이트
	PreOrderUpdateByBrand_off idx,chargeid,shopid

'/매장재고이동
elseif mode="ipchulmove" then

	if firstshopid="" or moveshopid ="" then
		response.write "<script type='text/javascript'>"
		response.write "	alert('매장이 지정되지 않았습니다');"
		response.write "	self.close();"
		response.write "</script>"
		dbget.close() : response.end
	end if

	if chargeid = "" then
		response.write "<script type='text/javascript'>"
		response.write "	alert('공급처(브랜드)가 지정되지 않았습니다');"
		response.write "	self.close();"
		response.write "</script>"
		dbget.close() : response.end
	end if

	if (C_IS_Maker_Upche) or getoffshopdiv(firstshopid) <> "1" or getoffshopdiv(moveshopid) <> "1" then
		response.write "<script type='text/javascript'>"
		response.write "	alert('직영매장만 이용가능한 매뉴입니다');"
		response.write "	self.close();"
		response.write "</script>"
		dbget.close() : response.end
	end if

	if isreq<>"M" then
		response.write "<script type='text/javascript'>"
		response.write "	alert('구분이 매장재고이동이 아닙니다');"
		response.write "	self.close();"
		response.write "</script>"
		dbget.close() : response.end
	end if

	if comment <> "" then
		if checkNotValidHTML(comment) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
		response.write "	self.close();"
		response.write "</script>"
		dbget.close() : response.end
		end if
	end if

	sqlStr = " select"
	sqlStr = sqlStr & " u.userid,u.shopname"
	sqlStr = sqlStr & " from [db_shop].[dbo].tbl_shop_user u"
	sqlStr = sqlStr & " Join [db_shop].[dbo].tbl_shop_designer d"
	sqlStr = sqlStr & " 	on u.userid=d.shopid"
	sqlStr = sqlStr & " left join ("
	sqlStr = sqlStr & " 	select top 1 "
	sqlStr = sqlStr & " 	shopid , defaultmargin,defaultsuplymargin"
	sqlStr = sqlStr & " 	from [db_shop].[dbo].tbl_shop_designer"
	sqlStr = sqlStr & " 	where shopid = '"&moveshopid&"'"
	sqlStr = sqlStr & " 	and makerid = '"&chargeid&"'"
	sqlStr = sqlStr & "	) as t"
	sqlStr = sqlStr & "		on d.defaultmargin = t.defaultmargin"
	sqlStr = sqlStr & "		and d.defaultsuplymargin = t.defaultsuplymargin"
	sqlStr = sqlStr & " where u.isusing='Y'"
	sqlStr = sqlStr & " and d.makerid='" + chargeid + "'"
	sqlStr = sqlStr & " and u.userid <> '"&firstshopid&"'"
	sqlStr = sqlStr & " and t.shopid is not null"

	'response.write sqlStr &"<br>"
	rsget.Open sqlStr, dbget, 1
	if not rsget.Eof then
	    tmpcontractyn = true
	end if
	rsget.Close

	if not(tmpcontractyn) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('출발매장과 도착매장이 계약조건이 틀립니다');"
		response.write "	self.close();"
		response.write "</script>"
		dbget.close() : response.end
	end if

	if itemarr = "" then
		response.write "<script type='text/javascript'>"
		response.write "	alert('상품이 선택되지 않았습니다');"
		response.write "	self.close();"
		response.write "</script>"
		dbget.close() : response.end
	end if

	itemgubunarr = Left(itemgubunarr,Len(itemgubunarr)-1)
	itemarr = Left(itemarr,Len(itemarr)-1)
	itemoptionarr = Left(itemoptionarr,Len(itemoptionarr)-1)
	sellcasharr = Left(sellcasharr,Len(sellcasharr)-1)
	suplycasharr = Left(suplycasharr,Len(suplycasharr)-1)
	shopbuypricearr = Left(shopbuypricearr,Len(shopbuypricearr)-1)
	itemnoarr = Left(itemnoarr,Len(itemnoarr)-1)
	designerarr = Left(designerarr,Len(designerarr)-1)

	itemgubunarr = split(itemgubunarr,"|")
	itemarr = split(itemarr,"|")
	itemoptionarr = split(itemoptionarr,"|")
	sellcasharr = split(sellcasharr,"|")
	suplycasharr = split(suplycasharr,"|")
	shopbuypricearr = split(shopbuypricearr,"|")
	itemnoarr = split(itemnoarr,"|")
	designerarr = split(designerarr,"|")

	cnt = ubound(itemarr)

	idx = ""

	if isarray(getupcheshopcontractinfo(firstshopid,chargeid)) then
		comm_cd = getupcheshopcontractinfo(firstshopid,chargeid)(7,0)
	else
		response.write "<script type='text/javascript'>"
		response.write "	alert('업체와 출발 매장간 계약이 없습니다');"
		response.write "	self.close();"
		response.write "</script>"
		dbget.close() : response.end
	end if

	'------------------------------출발매장
	sqlStr = " insert into [db_shop].[dbo].tbl_shop_ipchul_master"
	sqlStr = sqlStr + " (chargeid,shopid,divcode,vatcode,scheduledate,statecd,songjangdiv,songjangno,reguserid,isbaljuExists,comment ,comm_cd)"
	sqlStr = sqlStr + " values('" + chargeid + "'"
	sqlStr = sqlStr + " ,'" + firstshopid + "'"
	sqlStr = sqlStr + " ,'" + divcode + "'"
	sqlStr = sqlStr + " ,'" + vatcode + "'"
	sqlStr = sqlStr + " ,'" + scheduledt + "'"
	sqlStr = sqlStr + " ,0"
	sqlStr = sqlStr + " ,'" + songjangdiv + "'"
	sqlStr = sqlStr + " ,'" + songjangno + "'"
	sqlStr = sqlStr + " ,'" + session("ssBctId") + "'"
	sqlStr = sqlStr + " ,'M'"
	sqlStr = sqlStr + " ,'" + html2db(comment) + "'"
	sqlStr = sqlStr + " ,'" + comm_cd + "'"
	sqlStr = sqlStr + " )"

	'response.write sqlStr &"<br>"
	dbget.Execute(sqlStr)

	sqlStr = " select ident_current('[db_shop].[dbo].tbl_shop_ipchul_master') as idx "

	'response.write sqlStr &"<br>"
	rsget.Open sqlStr, dbget, 1
		idx = rsget("idx")
	rsget.close

	for i=0 to cnt
		sqlStr = " insert into [db_shop].[dbo].tbl_shop_ipchul_detail" + vbCrlf
		sqlStr = sqlStr + " (masteridx,itemgubun,shopitemid,itemoption," + vbCrlf
		sqlStr = sqlStr + " designerid,sellcash,suplycash,shopbuyprice,itemno,reqno)"  + vbCrlf
		sqlStr = sqlStr + " values(" + CStr(idx)  + "," + vbCrlf
		sqlStr = sqlStr + "'" + requestCheckVar(itemgubunarr(i),2) + "'," + vbCrlf
		sqlStr = sqlStr + "" + requestCheckVar(itemarr(i),10) + "," + vbCrlf
		sqlStr = sqlStr + "'" + requestCheckVar(itemoptionarr(i),4) + "'," + vbCrlf
		sqlStr = sqlStr + "'" + requestCheckVar(designerarr(i),32) + "'," + vbCrlf
		sqlStr = sqlStr + "" + requestCheckVar(sellcasharr(i),20) + "," + vbCrlf
		sqlStr = sqlStr + "" + requestCheckVar(suplycasharr(i),20) + "," + vbCrlf
		sqlStr = sqlStr + "" + requestCheckVar(shopbuypricearr(i),20) + "," + vbCrlf
		sqlStr = sqlStr + "" + cstr(requestCheckVar(itemnoarr(i) * -1,10)) + "," + vbCrlf
		sqlStr = sqlStr + "" + cstr(requestCheckVar(itemnoarr(i) * -1,10)) + vbCrlf
		sqlStr = sqlStr + "" + ")"

        'response.write sqlStr & "<Br>"
		dbget.Execute(sqlStr)
	next

	sqlStr = " update [db_shop].[dbo].tbl_shop_ipchul_detail" + vbCrlf
	sqlStr = sqlStr + " set itemname=T.shopitemname" + vbCrlf
	sqlStr = sqlStr + " ,itemoptionname=T.shopitemoptionname" + vbCrlf
	sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item T" + vbCrlf
	sqlStr = sqlStr + " where [db_shop].[dbo].tbl_shop_ipchul_detail.masteridx=" + CStr(idx)
	sqlStr = sqlStr + " and [db_shop].[dbo].tbl_shop_ipchul_detail.itemgubun=T.itemgubun"
	sqlStr = sqlStr + " and [db_shop].[dbo].tbl_shop_ipchul_detail.shopitemid=T.shopitemid"
	sqlStr = sqlStr + " and [db_shop].[dbo].tbl_shop_ipchul_detail.itemoption=T.itemoption"

	'response.write & "<Br>"
	dbget.Execute(sqlStr)

	'' Master Summary
	sqlStr = " update [db_shop].[dbo].tbl_shop_ipchul_master" + vbCrlf
	sqlStr = sqlStr + " set totalsellcash=IsNULL(T.totsell,0)" + vbCrlf
	sqlStr = sqlStr + " ,totalsuplycash=IsNULL(T.totsupp,0)" + vbCrlf
	sqlStr = sqlStr + " ,totalshopbuyprice=IsNULL(T.totshopbuy,0)" + vbCrlf
	sqlStr = sqlStr + " from (" + vbCrlf
	sqlStr = sqlStr + " 	select" + vbCrlf
	sqlStr = sqlStr + " 	sum(sellcash*itemno) as totsell " + vbCrlf
	sqlStr = sqlStr + " 	,sum(suplycash*itemno) as totsupp " + vbCrlf
	sqlStr = sqlStr + " 	,sum(shopbuyprice*itemno) as totshopbuy " + vbCrlf
	sqlStr = sqlStr + " 	from " + vbCrlf
	sqlStr = sqlStr + " 	[db_shop].[dbo].tbl_shop_ipchul_detail" + vbCrlf
	sqlStr = sqlStr + " 	where masteridx="  + CStr(idx) + vbCrlf
	sqlStr = sqlStr + " 	and deleteyn='N'" + vbCrlf
	sqlStr = sqlStr + " ) as T" + vbCrlf
	sqlStr = sqlStr + " where [db_shop].[dbo].tbl_shop_ipchul_master.idx=" + CStr(idx)

	'response.write sqlStr & "<Br>"
	dbget.Execute(sqlStr)

	'//기주문 업데이트
	PreOrderUpdateByBrand_off idx,chargeid,firstshopid
	'------------------------------출발매장

	'-------------------------------도착매장
	comm_cd = ""

	if isarray(getupcheshopcontractinfo(moveshopid,chargeid)) then
		comm_cd = getupcheshopcontractinfo(moveshopid,chargeid)(7,0)
	else
		response.write "<script type='text/javascript'>"
		response.write "	alert('업체와 도착매장간 계약이 없습니다');"
		response.write "	self.close();"
		response.write "</script>"
		dbget.close() : response.end
	end if

    ''isreq 입고요청. Flag , isbaljuExists 'Y'
	sqlStr = " insert into [db_shop].[dbo].tbl_shop_ipchul_master"
	sqlStr = sqlStr + " (chargeid,shopid,divcode,vatcode,scheduledate,statecd,songjangdiv,songjangno,reguserid,isbaljuExists,comment ,ipchulmoveidx ,comm_cd)"
	sqlStr = sqlStr + " values('" + chargeid + "'"
	sqlStr = sqlStr + " ,'" + moveshopid + "'"
	sqlStr = sqlStr + " ,'" + divcode + "'"
	sqlStr = sqlStr + " ,'" + vatcode + "'"
	sqlStr = sqlStr + " ,'" + scheduledt + "'"
	sqlStr = sqlStr + " ,0"
	sqlStr = sqlStr + " ,'" + songjangdiv + "'"
	sqlStr = sqlStr + " ,'" + songjangno + "'"
	sqlStr = sqlStr + " ,'" + session("ssBctId") + "'"
	sqlStr = sqlStr + " ,'M'"
	sqlStr = sqlStr + " ,'" + html2db(comment) + "'"
	sqlStr = sqlStr + " ," + cstr(idx) + ""
	sqlStr = sqlStr + " ,'" + comm_cd + "'"
	sqlStr = sqlStr + " )"

	'response.write sqlStr &"<br>"
	dbget.Execute(sqlStr)

	sqlStr = " select ident_current('[db_shop].[dbo].tbl_shop_ipchul_master') as idx "

	'response.write sqlStr &"<br>"
	rsget.Open sqlStr, dbget, 1
		moveidx = rsget("idx")
	rsget.close

	for i=0 to cnt
		sqlStr = " insert into [db_shop].[dbo].tbl_shop_ipchul_detail" + vbCrlf
		sqlStr = sqlStr + " (masteridx,itemgubun,shopitemid,itemoption," + vbCrlf
		sqlStr = sqlStr + " designerid,sellcash,suplycash,shopbuyprice,itemno,reqno)"  + vbCrlf
		sqlStr = sqlStr + " values(" + CStr(moveidx)  + "," + vbCrlf
		sqlStr = sqlStr + "'" + requestCheckVar(itemgubunarr(i),2) + "'," + vbCrlf
		sqlStr = sqlStr + "" + requestCheckVar(itemarr(i),10) + "," + vbCrlf
		sqlStr = sqlStr + "'" + requestCheckVar(itemoptionarr(i),4) + "'," + vbCrlf
		sqlStr = sqlStr + "'" + requestCheckVar(designerarr(i),32) + "'," + vbCrlf
		sqlStr = sqlStr + "" + requestCheckVar(sellcasharr(i),20) + "," + vbCrlf
		sqlStr = sqlStr + "" + requestCheckVar(suplycasharr(i),20) + "," + vbCrlf
		sqlStr = sqlStr + "" + requestCheckVar(shopbuypricearr(i),20) + "," + vbCrlf
		sqlStr = sqlStr + "" + requestCheckVar(itemnoarr(i),10) + "," + vbCrlf
		sqlStr = sqlStr + "" + requestCheckVar(itemnoarr(i),10) + vbCrlf
		sqlStr = sqlStr + "" + ")"

        'response.write & "<Br>"
		dbget.Execute(sqlStr)
	next

	sqlStr = " update [db_shop].[dbo].tbl_shop_ipchul_detail" + vbCrlf
	sqlStr = sqlStr + " set itemname=T.shopitemname" + vbCrlf
	sqlStr = sqlStr + " ,itemoptionname=T.shopitemoptionname" + vbCrlf
	sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item T" + vbCrlf
	sqlStr = sqlStr + " where [db_shop].[dbo].tbl_shop_ipchul_detail.masteridx=" + CStr(moveidx)
	sqlStr = sqlStr + " and [db_shop].[dbo].tbl_shop_ipchul_detail.itemgubun=T.itemgubun"
	sqlStr = sqlStr + " and [db_shop].[dbo].tbl_shop_ipchul_detail.shopitemid=T.shopitemid"
	sqlStr = sqlStr + " and [db_shop].[dbo].tbl_shop_ipchul_detail.itemoption=T.itemoption"

	'response.write & "<Br>"
	dbget.Execute(sqlStr)

	'' Master Summary
	sqlStr = " update [db_shop].[dbo].tbl_shop_ipchul_master" + vbCrlf
	sqlStr = sqlStr + " set totalsellcash=IsNULL(T.totsell,0)" + vbCrlf
	sqlStr = sqlStr + " ,totalsuplycash=IsNULL(T.totsupp,0)" + vbCrlf
	sqlStr = sqlStr + " ,totalshopbuyprice=IsNULL(T.totshopbuy,0)" + vbCrlf
	sqlStr = sqlStr + " from (" + vbCrlf
	sqlStr = sqlStr + " 	select" + vbCrlf
	sqlStr = sqlStr + " 	sum(sellcash*itemno) as totsell " + vbCrlf
	sqlStr = sqlStr + " 	,sum(suplycash*itemno) as totsupp " + vbCrlf
	sqlStr = sqlStr + " 	,sum(shopbuyprice*itemno) as totshopbuy " + vbCrlf
	sqlStr = sqlStr + " 	from " + vbCrlf
	sqlStr = sqlStr + " 	[db_shop].[dbo].tbl_shop_ipchul_detail" + vbCrlf
	sqlStr = sqlStr + " 	where masteridx="  + CStr(moveidx) + vbCrlf
	sqlStr = sqlStr + " 	and deleteyn='N'" + vbCrlf
	sqlStr = sqlStr + " ) as T" + vbCrlf
	sqlStr = sqlStr + " where [db_shop].[dbo].tbl_shop_ipchul_master.idx=" + CStr(moveidx)

	'response.write & "<Br>"
	dbget.Execute(sqlStr)
	'-------------------------------도착매장

	'출발지 마스터 테이블에.. 도착지 idx 박아넣음
	sqlStr = " update [db_shop].[dbo].tbl_shop_ipchul_master set" + vbCrlf
	sqlStr = sqlStr & " ipchulmoveidx = "&moveidx&"" + vbCrlf
	sqlStr = sqlStr & " where idx = "&idx&"" + vbCrlf

	'response.write & "<Br>"
	dbget.Execute(sqlStr)

	'//기주문 업데이트
	PreOrderUpdateByBrand_off moveidx,chargeid,moveshopid

	response.write "<script type='text/javascript'>"
	response.write "	alert('OK');"
	response.write "	opener.location.reload();"
	response.write "	self.close();"
	response.write "</script>"
	dbget.close() : response.end

elseif mode="addipchullist" then

	if isarray(getupcheshopcontractinfo(shopid,chargeid)) then
		comm_cd = getupcheshopcontractinfo(shopid,chargeid)(7,0)
	else
		response.write "<script type='text/javascript'>"
		response.write "	alert('업체와 매장간 계약이 없습니다');"
		response.write "	location.replace('"& refer &"');"
		response.write "</script>"
		dbget.close() : response.end
	end if

	if comment <> "" then
		if checkNotValidHTML(comment) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
		response.write "	self.close();"
		response.write "</script>"
		dbget.close() : response.end
		end if
	end if

    ''isreq 입고요청. Flag , isbaljuExists 'Y'
	sqlStr = " insert into [db_shop].[dbo].tbl_shop_ipchul_master"
	sqlStr = sqlStr + " (chargeid,shopid,divcode,vatcode,scheduledate,statecd,songjangdiv,songjangno,reguserid,isbaljuExists,comment,comm_cd)"
	sqlStr = sqlStr + " values('" + chargeid + "'"
	sqlStr = sqlStr + " ,'" + shopid + "'"
	sqlStr = sqlStr + " ,'" + divcode + "'"
	sqlStr = sqlStr + " ,'" + vatcode + "'"
	sqlStr = sqlStr + " ,'" + scheduledt + "'"

	If waitflag = "on" Then
		sqlStr = sqlStr + " ,-5"
	Else
		if (isreq<>"") then
		    sqlStr = sqlStr + " ,-2"
		else
		    sqlStr = sqlStr + " ,0"
		end if
	End If

	sqlStr = sqlStr + " ,'" + songjangdiv + "'"
	sqlStr = sqlStr + " ,'" + songjangno + "'"
	sqlStr = sqlStr + " ,'" + session("ssBctId") + "'"

	if (isreq<>"") then
	    sqlStr = sqlStr + " ,'Y'"
	else
	    sqlStr = sqlStr + " ,'N'"
	end if

	sqlStr = sqlStr + " ,'" + html2db(comment) + "'"
	sqlStr = sqlStr + " ,'" + comm_cd + "'"
	sqlStr = sqlStr + " )"

	'response.write sqlStr &"<br>"
	dbget.Execute(sqlStr)

	sqlStr = " select ident_current('[db_shop].[dbo].tbl_shop_ipchul_master') as idx "
	rsget.Open sqlStr, dbget, 1
		idx = rsget("idx")
	rsget.close

	itemgubunarr = Left(itemgubunarr,Len(itemgubunarr)-1)
	itemarr = Left(itemarr,Len(itemarr)-1)
	itemoptionarr = Left(itemoptionarr,Len(itemoptionarr)-1)
	sellcasharr = Left(sellcasharr,Len(sellcasharr)-1)
	suplycasharr = Left(suplycasharr,Len(suplycasharr)-1)
	shopbuypricearr = Left(shopbuypricearr,Len(shopbuypricearr)-1)
	itemnoarr = Left(itemnoarr,Len(itemnoarr)-1)
	designerarr = Left(designerarr,Len(designerarr)-1)

	itemgubunarr = split(itemgubunarr,"|")
	itemarr = split(itemarr,"|")
	itemoptionarr = split(itemoptionarr,"|")
	sellcasharr = split(sellcasharr,"|")
	suplycasharr = split(suplycasharr,"|")
	shopbuypricearr = split(shopbuypricearr,"|")
	itemnoarr = split(itemnoarr,"|")
	designerarr = split(designerarr,"|")

	cnt = ubound(itemarr)

	for i=0 to cnt
		sqlStr = " insert into [db_shop].[dbo].tbl_shop_ipchul_detail" + vbCrlf
		sqlStr = sqlStr + " (masteridx,itemgubun,shopitemid,itemoption," + vbCrlf
		sqlStr = sqlStr + " designerid,sellcash,suplycash,shopbuyprice,itemno,reqno)"  + vbCrlf
		sqlStr = sqlStr + " values(" + CStr(idx)  + "," + vbCrlf
		sqlStr = sqlStr + "'" + requestCheckVar(itemgubunarr(i),2) + "'," + vbCrlf
		sqlStr = sqlStr + "" + requestCheckVar(itemarr(i),10) + "," + vbCrlf
		sqlStr = sqlStr + "'" + requestCheckVar(itemoptionarr(i),4) + "'," + vbCrlf
		sqlStr = sqlStr + "'" + requestCheckVar(designerarr(i),32) + "'," + vbCrlf
		sqlStr = sqlStr + "" + requestCheckVar(sellcasharr(i),20) + "," + vbCrlf
		sqlStr = sqlStr + "" + requestCheckVar(suplycasharr(i),20) + "," + vbCrlf
		sqlStr = sqlStr + "" + requestCheckVar(shopbuypricearr(i),20) + "," + vbCrlf
		sqlStr = sqlStr + "" + requestCheckVar(itemnoarr(i),10) + "," + vbCrlf
		if (isreq<>"") then
		    sqlStr = sqlStr + "" + requestCheckVar(itemnoarr(i),10) + vbCrlf
		else
		    sqlStr = sqlStr + "0" + vbCrlf
		end if
		sqlStr = sqlStr + "" + ")"

		dbget.Execute(sqlStr)
	next

	if IS_HIDE_BUYCASH = True and shopid <> "" then
		sqlStr = " IF EXISTS(select top 1 idx from [db_shop].[dbo].tbl_shop_ipchul_detail where masteridx = " + CStr(idx)  + " and suplycash < 0) "
		sqlStr = sqlStr + " BEGIN "
		sqlStr = sqlStr + " 	update d "
		''sqlStr = sqlStr + " 	set d.sellcash = T.sellcash, d.shopbuyprice = (case when T.suplycash < T.buycash then T.buycash else T.suplycash end), d.suplycash = T.buycash "
		sqlStr = sqlStr + " 	set d.suplycash = T.buycash "
		sqlStr = sqlStr + " 	FROM "
		sqlStr = sqlStr + " 		[db_shop].[dbo].tbl_shop_ipchul_detail d "
		sqlStr = sqlStr + " 		join ( "
		sqlStr = sqlStr + " 			select "
		sqlStr = sqlStr + " 				d.masteridx, d.itemgubun, d.shopitemid, d.itemoption "
		sqlStr = sqlStr + " 				, s.shopitemprice as sellcash "
		sqlStr = sqlStr + " 				, (case "
		sqlStr = sqlStr + " 						when s.shopbuyprice <> 0 then s.shopbuyprice "
		sqlStr = sqlStr + " 						when IsNull(m.defaultsuplymargin,0) = 0 and IsNull(m.defaultmargin,0) = 0 then Round(s.shopitemprice * (100.0 - (35 - 5))/100, 0) "
		sqlStr = sqlStr + " 						when IsNull(m.defaultsuplymargin,0) = 0 and IsNull(m.defaultmargin,0) <> 0 then Round(s.shopitemprice * (100.0 - (m.defaultmargin - 5))/100, 0) "
		sqlStr = sqlStr + " 						when IsNull(m.defaultsuplymargin,0) <> 0 then Round(s.shopitemprice * (100.0 - m.defaultsuplymargin)/100, 0) "
		sqlStr = sqlStr + " 						else s.shopitemprice end) as suplycash "
		sqlStr = sqlStr + " 				, (case "
		sqlStr = sqlStr + " 						when s.shopsuplycash <> 0 then s.shopsuplycash "
		sqlStr = sqlStr + " 						when IsNull(i.mwdiv, '') = 'M' and IsNull(i.buycash, 0) <> 0 and IsNull(m.comm_cd,'') <> 'B012' and IsNull(m.comm_cd,'') <> 'B022' then Round(IsNull(i.buycash,0),0) + Round(IsNull(o.optaddprice,0),0) "
		sqlStr = sqlStr + " 						when IsNull(m.defaultmargin,0) = 0 then Round(s.shopitemprice * (100.0 - 35)/100, 0) "
		sqlStr = sqlStr + " 						when IsNull(m.defaultmargin,0) <> 0 then Round(s.shopitemprice * (100.0 - IsNull(m.defaultmargin,0))/100, 0) "
		sqlStr = sqlStr + " 						else s.shopitemprice end) as buycash "
		sqlStr = sqlStr + " 			from "
		sqlStr = sqlStr + " 				[db_shop].[dbo].tbl_shop_ipchul_detail d "
		sqlStr = sqlStr + " 				join [db_shop].[dbo].tbl_shop_item s "
		sqlStr = sqlStr + " 				on "
		sqlStr = sqlStr + " 					1 = 1 "
		sqlStr = sqlStr + " 					and d.masteridx = " + CStr(idx)  + " "
		sqlStr = sqlStr + " 					and d.itemgubun = s.itemgubun "
		sqlStr = sqlStr + " 					and d.shopitemid = s.shopitemid "
		sqlStr = sqlStr + " 					and d.itemoption = s.itemoption "
		sqlStr = sqlStr + " 				left join [db_shop].[dbo].tbl_shop_designer m "
		sqlStr = sqlStr + " 				on	 "
		sqlStr = sqlStr + " 					1 = 1 "
		sqlStr = sqlStr + " 					and m.shopid = '" & shopid & "' "
		sqlStr = sqlStr + " 					and m.makerid = s.makerid "
		sqlStr = sqlStr + " 				left join [db_item].[dbo].tbl_item i "
		sqlStr = sqlStr + " 				on "
		sqlStr = sqlStr + " 					1 = 1 "
		sqlStr = sqlStr + " 					and s.itemgubun = '10' "
		sqlStr = sqlStr + " 					and s.shopitemid = i.itemid "
		sqlStr = sqlStr + " 				left join [db_item].[dbo].tbl_item_option o "
		sqlStr = sqlStr + " 				on "
		sqlStr = sqlStr + " 					1 = 1 "
		sqlStr = sqlStr + " 					and s.itemgubun='10' "
		sqlStr = sqlStr + " 					and s.shopitemid = o.itemid "
		sqlStr = sqlStr + " 					and s.itemoption=o.itemoption "
		sqlStr = sqlStr + " 		) T "
		sqlStr = sqlStr + " 		on "
		sqlStr = sqlStr + " 			1 = 1 "
		sqlStr = sqlStr + " 			and d.masteridx = T.masteridx "
		sqlStr = sqlStr + " 			and d.itemgubun = T.itemgubun "
		sqlStr = sqlStr + " 			and d.shopitemid = T.shopitemid "
		sqlStr = sqlStr + " 			and d.itemoption = T.itemoption "
		sqlStr = sqlStr + " 	WHERE "
		sqlStr = sqlStr + " 		d.suplycash < 0 "
		sqlStr = sqlStr + " END "
		rsget.Open sqlStr, dbget, 1
	end if

	sqlStr = " update [db_shop].[dbo].tbl_shop_ipchul_detail" + vbCrlf
	sqlStr = sqlStr + " set itemname=T.shopitemname" + vbCrlf
	sqlStr = sqlStr + " ,itemoptionname=T.shopitemoptionname" + vbCrlf
	sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item T" + vbCrlf
	sqlStr = sqlStr + " where [db_shop].[dbo].tbl_shop_ipchul_detail.masteridx=" + CStr(idx)
	sqlStr = sqlStr + " and [db_shop].[dbo].tbl_shop_ipchul_detail.itemgubun=T.itemgubun"
	sqlStr = sqlStr + " and [db_shop].[dbo].tbl_shop_ipchul_detail.shopitemid=T.shopitemid"
	sqlStr = sqlStr + " and [db_shop].[dbo].tbl_shop_ipchul_detail.itemoption=T.itemoption"

	dbget.Execute(sqlStr)

	'' Master Summary
	sqlStr = " update [db_shop].[dbo].tbl_shop_ipchul_master" + vbCrlf
	sqlStr = sqlStr + " set totalsellcash=IsNULL(T.totsell,0)" + vbCrlf
	sqlStr = sqlStr + " ,totalsuplycash=IsNULL(T.totsupp,0)" + vbCrlf
	sqlStr = sqlStr + " ,totalshopbuyprice=IsNULL(T.totshopbuy,0)" + vbCrlf
	sqlStr = sqlStr + " from (" + vbCrlf
	sqlStr = sqlStr + " select sum(sellcash*itemno) as totsell " + vbCrlf
	sqlStr = sqlStr + " ,sum(suplycash*itemno) as totsupp " + vbCrlf
	sqlStr = sqlStr + " ,sum(shopbuyprice*itemno) as totshopbuy " + vbCrlf
	sqlStr = sqlStr + " from " + vbCrlf
	sqlStr = sqlStr + " [db_shop].[dbo].tbl_shop_ipchul_detail" + vbCrlf
	sqlStr = sqlStr + " where masteridx="  + CStr(idx) + vbCrlf
	sqlStr = sqlStr + " and deleteyn='N'" + vbCrlf
	sqlStr = sqlStr + " ) as T" + vbCrlf
	sqlStr = sqlStr + " where [db_shop].[dbo].tbl_shop_ipchul_master.idx=" + CStr(idx)

	dbget.Execute(sqlStr)

	'//기주문 업데이트
	PreOrderUpdateByBrand_off idx,chargeid,shopid

	addshopid = Split(addshopid, ",")
	for i = 0 to UBOund(addshopid)
		if (Trim(addshopid(i)) <> "") then
			if isarray(getupcheshopcontractinfo(Trim(addshopid(i)),chargeid)) then
				comm_cd = getupcheshopcontractinfo(Trim(addshopid(i)),chargeid)(7,0)
			else
				response.write "<script type='text/javascript'>"
				response.write "	alert('업체와 매장간 계약이 없습니다');"
				''response.write "	location.replace('"& refer &"');"
				response.write "</script>"
				response.write "업체와 매장간 계약이 없습니다 - " & Trim(addshopid(i)) & " " & chargeid
				dbget.close() : response.end
			end if

			sqlStr = " exec [db_shop].[dbo].[usp_Ten_IpchulSheel_Cpoy] '" & CStr(idx) & "', '" & Trim(addshopid(i)) & "', '" & comm_cd & "' "
		    rsget.CursorLocation = adUseClient
		    rsget.Open sqlStr, dbget, adOpenForwardOnly
		    ''if Not rsget.Eof then
				newiid = rsget("masteridx")
				newtargetid = rsget("targetid")
				newbaljuid = rsget("baljuid")
		    ''end if
		    rsget.close

			if Not IsNull(newtargetid) then
				'//기주문 업데이트
				PreOrderUpdateByBrand_off newiid,chargeid,Trim(addshopid(i))
			end if
		end if
	next
else
	response.write mode
	dbget.close()	:	response.End

end if

if ((mode ="offitemreg") or (mode="arrins")) then
	if (InStr(refer,"&react=true")<1) then
		refer = refer + "&react=true"
	end if

elseif mode="addipchullist" then
	refer = "/common/offshop/shop_ipchullist.asp?menupos="&menupos&""
end if
%>

<script type='text/javascript'>
	alert('저장 되었습니다.');
	location.replace('<%= refer %>');
</script>

<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
