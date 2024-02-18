<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프라인 수기 매출
' History : 2012.08.07 한용민 생성
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/BarcodeFunction.asp"-->
<%
dim i , mode , bagidxarr , sqlStr ,menupos, shopid ,barcode ,sqlsearch ,shopregdate ,orderno ,posid ,result
dim adminuserid , masteridx ,cnt, nowdate, jungsandate
dim itemgubun, itemid, itemoption, itemprice, suplycash, buyprice, itemname, itemoptionname, makerid, extbarcode
dim itemgubunarr ,itemidarr ,itemoptionarr ,itemnamearr ,itemoptionnamearr ,sellcasharr ,suplycasharr
dim shopbuypricearr ,itemnoarr ,makeridarr ,extbarcodearr
dim imaechulgubun, tmpshopid
    mode = requestcheckvar(request("mode"),32)
    menupos = requestcheckvar(request("menupos"),10)
    shopregdate = requestcheckvar(request("shopregdate"),10)
	shopid = requestcheckvar(request("shopid"),32)
	barcode = requestcheckvar(request("barcode"),32)
	itemgubunarr = request("itemgubunarr")
	itemidarr = request("itemidarr")
	itemoptionarr = request("itemoptionarr")
	itemnamearr = request("itemnamearr")
	itemoptionnamearr = request("itemoptionnamearr")
	sellcasharr = request("sellcasharr")
	suplycasharr = request("suplycasharr")
	shopbuypricearr = request("shopbuypricearr")
	itemnoarr = request("itemnoarr")
	makeridarr = request("makeridarr")
	extbarcodearr = request("extbarcodearr")

adminuserid = session("ssBctId")
posid = 99
nowdate = now()
jungsandate = year(nowdate) & "-" & Format00(2,month(nowdate)) & "-" & "10"
'response.write mode

'//바코드 상품등록
if mode = "oneaddmanualItem" then

	if shopid = "" then
		response.write "<script type='text/javascript'>"
		response.write "	alert('매장이 없습니다.');"
		response.write "	self.close();"
		response.write "</script>"
		response.end	:	dbget.close()
	end if
	if barcode = "" then
		response.write "<script type='text/javascript'>"
		response.write "	alert('바코드를 입력 하세요.');"
		response.write "	self.close();"
		response.write "</script>"
		response.end	:	dbget.close()
	end if
	if len(barcode) < 11 then
		response.write "<script type='text/javascript'>"
		response.write "	alert('바코드의 길이가 짧습니다.\n물류코드나 범용바코드를 다시 확인후, 입력 하세요.');"
		response.write "	self.close();"
		response.write "</script>"
		response.end	:	dbget.close()
	end if

	if trim(barcode)<>"" then

		'//바코드가 있을경우, 범용바코드는 필수로 검색
		sqlStr = "select top 1"
		sqlStr = sqlStr + " itemgubun,shopitemid,itemoption"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item"
		sqlStr = sqlStr + " where extbarcode='" + trim(barcode) + "'"

		'response.write sqlStr & "<Br>"
		rsget.Open sqlStr,dbget,1
		if Not rsget.Eof then
			itemgubun = rsget("itemgubun")
			itemid = rsget("shopitemid")
			itemoption = rsget("itemoption")
		end if
		rsget.Close
	end if

	if itemid = "" then
		itemgubun 	= BF_GetItemGubun(barcode)
		itemid 		= BF_GetItemId(barcode)
		itemoption 	= BF_GetItemOption(barcode)
	end if

	sqlsearch = sqlsearch + " and s.itemgubun='"& itemgubun &"'"
	sqlsearch = sqlsearch + " and s.shopitemid="& itemid &""
	sqlsearch = sqlsearch + " and s.itemoption='"& itemoption &"'"

    sqlStr = " select top 1 s.itemgubun, s.shopitemid, s.itemoption, s.extbarcode, s.isusing as itemstatus"
    sqlStr = sqlStr + " , convert(varchar(32),s.regdate,20) as regdate"
	sqlStr = sqlStr + " ,(CASE"
	sqlStr = sqlStr + " 	when s.shopsuplycash = 0 and sd.comm_cd in ('B011','B012')"		'/매입가가 0 ,텐텐위탁, 업체위탁
	sqlStr = sqlStr + " 		then convert(int,s.shopitemprice*(100-IsNULL(sd.defaultmargin,100))/100)"
	sqlStr = sqlStr + " 	else s.shopsuplycash"
	sqlStr = sqlStr + "	end) as shopsuplycash"
	'sqlStr = sqlStr + " , s.shopsuplycash"
	sqlStr = sqlStr + " ,(CASE" & VbCRLF
	sqlStr = sqlStr + " 	when s.shopbuyprice = 0 and sd.comm_cd in ('B011','B012')"		'/매장출고가 0 ,텐텐위탁, 업체위탁
	sqlStr = sqlStr + " 		then convert(int,s.shopitemprice*(100-IsNULL(sd.defaultsuplymargin,100))/100)"
	sqlStr = sqlStr + "		else s.shopbuyprice"
	sqlStr = sqlStr + "	end) as shopbuyprice"
	'sqlStr = sqlStr + " , s.shopbuyprice"
    sqlStr = sqlStr + " , (CASE WHEN s.itemgubun='80' THEN 0 ELSE s.orgsellprice END) as orgsellprice"
    sqlStr = sqlStr + " , (CASE WHEN s.itemgubun='80' THEN 0 ELSE s.shopitemprice END) as shopitemprice"      ''판매가
    sqlStr = sqlStr + " , s.makerid ,s.extbarcode" '' 브랜드 ID
    sqlStr = sqlStr + " , s.shopitemname, s.shopitemoptionname"
    sqlStr = sqlStr + " , c.socname_kor"        '' 브랜드 명
	sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item s"
	sqlStr = sqlStr + " join db_shop.dbo.tbl_shop_designer sd" & VbCRLF
	sqlStr = sqlStr + " 	on sd.shopid='"&shopid&"' and s.makerid=sd.makerid" & VbCRLF
	sqlStr = sqlStr + " join [db_user].[dbo].tbl_user_c c"
	sqlStr = sqlStr + " 	on s.makerid=c.userid"
	sqlStr = sqlStr + " where 1=1 " & sqlsearch

	'response.write sqlStr & "<Br>"
	rsget.open sqlStr,dbget,1

    if Not rsget.Eof then
		itemgubun = rsget("itemgubun")
		itemid = rsget("shopitemid")
		itemoption = rsget("itemoption")
		itemprice = rsget("shopitemprice")
		suplycash = rsget("shopsuplycash")
		buyprice = rsget("shopbuyprice")
		itemname = rsget("shopitemname")
		itemoptionname = rsget("shopitemoptionname")
		makerid = rsget("makerid")
		extbarcode = rsget("extbarcode")
    end if

    rsget.close

	if itemid <> "" then
		response.write "<script type='text/javascript'>"
		response.write "	opener.ReActItems('"&itemgubun&"|','"&itemid&"|','"&itemoption&"|','"&itemprice&"|','"&suplycash&"|','"&buyprice&"|','1','"&itemname&"|','"&itemoptionname&"|','"&makerid&"|','"&extbarcode&"|');"
		response.write "	self.close();"
		response.write "</script>"
		response.end	:	dbget.close()
	else
		response.write "<script type='text/javascript'>"
		response.write "	alert('해당되는 상품이 없습니다.');"
		response.write "	self.close();"
		response.write "</script>"
		response.end	:	dbget.close()
	end if

'//매출전송
elseif mode = "addmanualItem" then

	if not(C_ADMIN_USER) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('권한이 없습니다');"
		response.write "	self.close();"
		response.write "</script>"
		response.end	:	dbget.close()
	end if
	if shopid = "" then
		response.write "<script type='text/javascript'>"
		response.write "	alert('매장이 없습니다.');"
		response.write "</script>"
		response.end	:	dbget.close()
	end if
	if shopregdate = "" then
		response.write "<script type='text/javascript'>"
		response.write "	alert('매출날짜가 지정되지 않았습니다.');"
		response.write "</script>"
		response.end	:	dbget.close()
	end if

	itemgubunarr = split(itemgubunarr,"|")
	itemidarr	= split(itemidarr,"|")
	itemoptionarr = split(itemoptionarr,"|")
	itemnamearr		= split(itemnamearr,"|")
	itemoptionnamearr = split(itemoptionnamearr,"|")
	sellcasharr = split(sellcasharr,"|")
	suplycasharr = split(suplycasharr,"|")
	shopbuypricearr = split(shopbuypricearr,"|")
	itemnoarr = split(itemnoarr,"|")
	makeridarr = split(makeridarr,"|")
	extbarcodearr = split(extbarcodearr,"|")

	'//두달 이전 매출 일경우
	if datediff("m", Left(shopregdate,10) , nowdate) >= 2 then
		response.write "<script type='text/javascript'>"
		response.write "	alert('두달 이전내역은 입력 하실수 없습니다.');"
		response.write "</script>"
		response.end	:	dbget.close()
	end if

	'//다음달 매출 입력시
	if datediff("m", Left(shopregdate,10) , nowdate) < 0 then
		response.write "<script type='text/javascript'>"
		response.write "	alert('다음달 매출은 입력이 불가능 합니다.');"
		response.write "</script>"
		response.end	:	dbget.close()
	end if

'	'//이전달 매출 입력시
'	if datediff("m", Left(shopregdate,10) , nowdate) = 1 then
'		if datediff("d", jungsandate , nowdate) > 0 then
'			response.write "<script type='text/javascript'>"
'			response.write "	alert('매출일이 정산이 마감이 된 날짜 입니다.');"
'			response.write "</script>"
'			if Not C_ADMIN_AUTH then
'				response.end	:	dbget.close()
'			else
'				response.write "<script type='text/javascript'>"
'				response.write "	alert('[관리자권한]\n\n강제등록합니다.');"
'				response.write "</script>"
'			end if
'		end if
'	end if

	cnt = UBound(itemidarr)

	for i=0 to cnt - 1

		'//판매가와 수량 둘다 마이너스 일경우..곱하면 플러스가 나옴.. 팅겨냄
		if left(trim(sellcasharr(i)),1)="-" and left(trim(itemnoarr(i)),1)="-" then
			response.write "<script type='text/javascript'>"
			response.write "	alert('판매가와 수량 둘다 마이너스 값이 될수 없습니다.\n마이너스 주문 입력시 수량만 마이너스로 입력해주세요');"
			response.write "</script>"
			response.end	:	dbget.close()
		end if
	next

	orderno = manualordernomake_off(shopid,posid)

    '/이미존재하는 주문번호인지 체크
    sqlStr = "select count(idx) as cnt"
	sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shopjumun_master"
	sqlStr = sqlStr + " where orderno='"&orderno&"'"

	'response.write sqlStr &"<br>"
	rsget.Open sqlStr,dbget,1

	if Not rsget.Eof then
	    if (rsget("cnt")>0) then result = "Y"
	end if

	rsget.close

	if result = "Y" then
		response.write "<script type='text/javascript'>"
		response.write "	alert('주문번호가 이미 존재 합니다. 관리자 문의요망.');"
		response.write "</script>"
		response.end	:	dbget.close()
	end if
	result = ""

    ''매출구분  /2013/12/17 추가
    imaechulgubun=""

    sqlStr = "select isNULL(tplcompanyid,'MANUAL') as maechulgubun"
    sqlStr = sqlStr&" from db_partner.dbo.tbl_partner"
    sqlStr = sqlStr&" where id='"&shopid&"'"
    rsget.Open sqlStr,dbget,1
	if Not rsget.Eof then
		imaechulgubun=rsget("maechulgubun")
	end if
	rsget.close

    if (imaechulgubun="") then
        imaechulgubun="MANUAL"
    end if

	'// 입력전 데이타 검증
	'// 1. 올바른 바코드인지
	'// 2. 계약이 있는지
	for i=0 to cnt - 1

		sqlStr = " 	select top 1 i.itemgubun, IsNull(s.shopid, '') as shopid " + vbcrlf
		sqlStr = sqlStr + " 	from db_shop.dbo.tbl_shop_item i" + vbcrlf
		sqlStr = sqlStr + " 	left join db_shop.dbo.tbl_shop_designer s" & VbCRLF
		sqlStr = sqlStr + " 		on s.shopid='"&shopid&"'"
		sqlStr = sqlStr + " 		and i.makerid=s.makerid" & VbCRLF
		sqlStr = sqlStr + " 	left join db_item.dbo.tbl_item ii" & VbCRLF
		sqlStr = sqlStr + " 		on i.shopitemid = ii.itemid" & VbCRLF
		sqlStr = sqlStr + " 		and i.itemgubun = '10'" & VbCRLF
		sqlStr = sqlStr + " 	where i.itemgubun = '"& requestCheckVar(trim(itemgubunarr(i)),2) &"'" + vbcrlf
		sqlStr = sqlStr + " 	and i.shopitemid = "& requestCheckVar(trim(itemidarr(i)),10) &"" + vbcrlf
		sqlStr = sqlStr + " 	and i.itemoption = '"& requestCheckVar(trim(itemoptionarr(i)),4) &"'" + vbcrlf
		tmpshopid = "XXXXXXXXX"
		rsget.Open sqlStr,dbget,1
		if Not rsget.Eof then
			tmpshopid = rsget("shopid")
		end if
		rsget.close

		if (tmpshopid = "XXXXXXXXX") then
			response.write "잘못된 상품코드 또는 오프라인 상품등록 이전 상품입니다. : " & itemgubunarr(i)
			dbget.close() : response.end
		elseif tmpshopid = "" then
			response.write "계약이 지정되지 않았습니다. : " & itemgubunarr(i)
			dbget.close() : response.end
		end if
	next

	'//마스터 테이블 등록
    sqlStr = "select * from [db_shop].[dbo].tbl_shopjumun_master where 1=0"
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew

	rsget("orderno")    = orderno
	rsget("shopid")     = shopid
	rsget("totalsum")   = 0
	rsget("realsum")    = 0
	rsget("jumundiv")   = "00"
	rsget("jumunmethod") = "01"
	rsget("shopregdate") = Left(shopregdate,10)
	rsget("cancelyn")   = "N"
	rsget("shopidx")    = "0"
	rsget("spendmile")  = "0"
	rsget("pointuserno") = ""
	rsget("gainmile") = "0"
	rsget("cashsum")    = 0
    rsget("cardsum")    = "0"
    rsget("casherid")   = adminuserid
    rsget("GiftCardPaySum") = "0"
    rsget("CardAppNo")      = ""
    rsget("CashReceiptNo")  = ""
    rsget("CashreceiptGubun") = ""
    rsget("CardInstallment")  = ""
	rsget("IXyyyymmdd") = Left(shopregdate,10)
	rsget("tableno")  = "0"
    rsget("TenGiftCardPaySum")  = "0"
	rsget("TenGiftCardMatchCode")  = ""
	rsget("refOrderNo")  = ""
	rsget("maechulgubun")  = imaechulgubun '"MANUAL"

	rsget.update
		masteridx = rsget("idx")
	rsget.close

	for i=0 to cnt - 1

		'//디테일 테이블 등록
        sqlStr = "insert into [db_shop].[dbo].tbl_shopjumun_detail" + vbcrlf
		sqlStr = sqlStr + " ( masteridx, orderno, itemgubun, itemid, itemoption" + vbcrlf
		sqlStr = sqlStr + " , itemno, itemname, itemoptionname, sellprice, realsellprice" + vbcrlf
		sqlStr = sqlStr + " , suplyprice" + vbcrlf
		sqlStr = sqlStr + " , shopbuyprice" + vbcrlf
		sqlStr = sqlStr + " , makerid, jungsanid, cancelyn" + vbcrlf
		sqlStr = sqlStr + " , shopidx, itempoint, discountKind, Iorgsellprice, Ishopitemprice" + vbcrlf
		sqlStr = sqlStr + " , jcomm_cd, addtaxcharge, vatinclude)" + vbcrlf
		sqlStr = sqlStr + " 	select" + vbcrlf
		sqlStr = sqlStr + " 	'"&masteridx&"','"&orderno&"',i.itemgubun ,i.shopitemid ,i.itemoption" + vbcrlf
		sqlStr = sqlStr + " 	,'"& requestCheckVar(trim(itemnoarr(i)),10) &"', i.shopitemname, i.shopitemoptionname,'"& requestCheckVar(trim(sellcasharr(i)),20) &"','"& requestCheckVar(trim(sellcasharr(i)),20) &"'" + vbcrlf
		sqlStr = sqlStr + " 	,(CASE" & VbCRLF
		sqlStr = sqlStr + " 		when isnull(ii.mwdiv,'')='M' and s.comm_cd not in ('B012')" & VbCRLF		'//온라인매입이고, 업체위탁이 아니면 온라인매입가로
		sqlStr = sqlStr + " 			THEN isnull(ii.buycash,0)" & VbCRLF
		'sqlStr = sqlStr + " 		when i.shopsuplycash = 0 and s.comm_cd in ('B011','B012','B013')" & VbCRLF		'/매입가가 0 ,텐텐위탁, 업체위탁 ,출고위탁
		sqlStr = sqlStr + " 		when i.shopsuplycash = 0" & VbCRLF		'매입가 다 무조건 꼿을것
		sqlStr = sqlStr + " 			then convert(int,i.shopitemprice*(100-IsNULL(s.defaultmargin,100))/100)" & VbCRLF
		sqlStr = sqlStr + " 		else i.shopsuplycash" & VbCRLF
		sqlStr = sqlStr + "			end) as shopsuplycash" & VbCRLF
		sqlStr = sqlStr + " 	,(CASE" & VbCRLF
		'sqlStr = sqlStr + " 		when i.shopbuyprice = 0 and s.comm_cd in ('B011','B012','B013')" & VbCRLF		'/매장출고가 0 ,텐텐위탁, 업체위탁 ,출고위탁
		sqlStr = sqlStr + " 		when i.shopbuyprice = 0" & VbCRLF		'매장출고가 다 무조건 꼿을것
		sqlStr = sqlStr + " 			then convert(int,i.shopitemprice*(100-IsNULL(s.defaultsuplymargin,100))/100)" & VbCRLF
		sqlStr = sqlStr + "			else i.shopbuyprice" & VbCRLF
		sqlStr = sqlStr + "			end) as shopbuyprice" & VbCRLF
		sqlStr = sqlStr + " 	, i.makerid, i.makerid, 'N'" + vbcrlf
		sqlStr = sqlStr + " 	,'0','0','0', i.orgsellprice, i.shopitemprice" + vbcrlf
		sqlStr = sqlStr + " 	, s.comm_cd, 0, i.vatinclude" + vbcrlf
		sqlStr = sqlStr + " 	from db_shop.dbo.tbl_shop_item i" + vbcrlf
		sqlStr = sqlStr + " 	join db_shop.dbo.tbl_shop_designer s" & VbCRLF
		sqlStr = sqlStr + " 		on s.shopid='"&shopid&"'"
		sqlStr = sqlStr + " 		and i.makerid=s.makerid" & VbCRLF
		sqlStr = sqlStr + " 	left join db_item.dbo.tbl_item ii" & VbCRLF
		sqlStr = sqlStr + " 		on i.shopitemid = ii.itemid" & VbCRLF
		sqlStr = sqlStr + " 		and i.itemgubun = '10'" & VbCRLF
		sqlStr = sqlStr + " 	where i.itemgubun = '"& requestCheckVar(trim(itemgubunarr(i)),2) &"'" + vbcrlf
		sqlStr = sqlStr + " 	and i.shopitemid = "& requestCheckVar(trim(itemidarr(i)),10) &"" + vbcrlf
		sqlStr = sqlStr + " 	and i.itemoption = '"& requestCheckVar(trim(itemoptionarr(i)),4) &"'" + vbcrlf

		'response.write sqlStr & "<Br>"
		dbget.Execute sqlStr
	next

	'//마스터 테이블 합산
	sqlStr = "update m" + vbcrlf
	sqlStr = sqlStr + " set m.totalsum = t.sellprice" + vbcrlf
	sqlStr = sqlStr + " ,m.realsum = t.realsellprice" + vbcrlf
	sqlStr = sqlStr + " ,m.cashsum = t.realsellprice" + vbcrlf
	sqlStr = sqlStr + " from db_shop.dbo.tbl_shopjumun_master m" + vbcrlf
	sqlStr = sqlStr + " join (" + vbcrlf
	sqlStr = sqlStr + " 	select" + vbcrlf
	sqlStr = sqlStr + " 	orderno ,sum((d.sellprice+addtaxcharge) * d.itemno) as sellprice" + vbcrlf
	sqlStr = sqlStr + " 	,sum((d.realsellprice+addtaxcharge) * d.itemno) as realsellprice" + vbcrlf
	sqlStr = sqlStr + " 	,sum((d.suplyprice+addtaxcharge) * d.itemno) as suplyprice" + vbcrlf
	sqlStr = sqlStr + " 	from [db_shop].[dbo].tbl_shopjumun_detail d" + vbcrlf
	sqlStr = sqlStr + " 	where d.cancelyn = 'N'" + vbcrlf
	sqlStr = sqlStr + " 	and d.orderno = '"&orderno&"'" + vbcrlf
	sqlStr = sqlStr + " 	group by orderno" + vbcrlf
	sqlStr = sqlStr + " ) as t" + vbcrlf
	sqlStr = sqlStr + " 	on m.orderno = t.orderno" + vbcrlf
	sqlStr = sqlStr + " 	and m.cancelyn = 'N'" + vbcrlf
	sqlStr = sqlStr + " where m.orderno = '"&orderno&"'"

	'response.write sqlStr & "<Br>"
	dbget.Execute sqlStr

	'// 중복입력 제거
	sqlStr = "[db_shop].[dbo].[usp_TEN_Shop_ManualOrder_DuppRemove] '" & orderno & "'"
	dbget.Execute sqlStr

	''재고 업데이트(No tran)
    sqlStr = "exec db_summary.dbo.sp_Ten_Shop_Stock_RegOrder '" & orderno & "'"

    'response.write sqlStr & "<Br>"
    dbget.Execute sqlStr

	response.write "<script type='text/javascript'>"
	response.write "	alert('OK');"
	response.write "	parent.self.close();"
	response.write "</script>"
	response.end	:	dbget.close()

else
	response.write "<script type='text/javascript'>"
	response.write "	alert('구분자가 없습니다.');"
	response.write "</script>"
	response.end	:	dbget.close()
end if
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/common/lib/poptail.asp"-->
