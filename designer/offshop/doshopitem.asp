<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<%
''업체 오프 상품 수정

dim mode,itemgubunarr,itemarr
dim itemoptionarr,itempricearr,isusingarr,extbarcodearr, itemsuplyarr
dim designer,itemgubun,itemname,sellcash,suplycash
dim detailidxarr, currjungsanidarr, shopitemnamearr
dim idx

dim shopbuyprice

mode = request("mode")
itemgubunarr = request("itemgubunarr")
itemarr = request("itemarr")
itemoptionarr = request("itemoptionarr")
itempricearr = request("itempricearr")
itemsuplyarr = request("itemsuplyarr")
isusingarr = request("isusingarr")
extbarcodearr = request("extbarcodearr")
detailidxarr  = request("detailidxarr")
currjungsanidarr = request("currjungsanidarr")
shopitemnamearr = (request("shopitemnamearr"))


designer = request("designer")
itemgubun = request("itemgubun")
itemname = request("itemname")
sellcash = request("sellcash")
suplycash = request("suplycash")
shopbuyprice = request("shopbuyprice")


dim sellcasharr,suplycasharr,itemnoarr,designerarr
dim discountsellpricearr, shopbuypricearr
sellcasharr = request("sellcasharr")
suplycasharr = request("suplycasharr")
itemnoarr  = request("itemnoarr")
designerarr = request("designerarr")
discountsellpricearr = request("discountsellpricearr")
shopbuypricearr = request("shopbuypricearr")

dim chargeid,shopid,divcode,vatcode
chargeid = request("chargeid")
shopid = request("shopid")
divcode = request("divcode")
vatcode = request("vatcode")

idx = request("idx")

dim cksel
cksel = request("cksel")

dim songjangdiv,songjangno
songjangdiv = request("songjangdiv")
songjangno = LeftB(Newhtml2db(request("songjangno")),16)


dim i,cnt,sqlStr
dim extbarcodeAlreadyExists
dim extbarcodeAlreadyExistsString
dim stockitemexists


dim refer
refer = request.ServerVariables("HTTP_REFERER")


if mode ="arrmodi" then
    '' 사용여부, 범용바코드만 수정가능.
	itemgubunarr = Left(itemgubunarr,Len(itemgubunarr)-1)
	itemarr = Left(itemarr,Len(itemarr)-1)
	itemoptionarr = Left(itemoptionarr,Len(itemoptionarr)-1)
	isusingarr = Left(isusingarr,Len(isusingarr)-1)

	itemgubunarr = split(itemgubunarr,"|")
	itemarr = split(itemarr,"|")
	itemoptionarr = split(itemoptionarr,"|")
	isusingarr = split(isusingarr,"|")
	extbarcodearr = split(extbarcodearr,"|")

	cnt = ubound(itemarr)

	for i=0 to cnt
		extbarcodeAlreadyExists = false

		if extbarcodearr(i)<>"" then
			sqlStr = " select top 1 * from [db_item].[dbo].tbl_item_option_stock" + VbCrlf
			sqlStr = sqlStr + " where barcode='" + trim(CStr(extbarcodearr(i))) + "'" + VbCrlf
			sqlStr = sqlStr + " and not ("
			sqlStr = sqlStr + " 	itemgubun='" + itemgubunarr(i) + "'" + VbCrlf
			sqlStr = sqlStr + " 	and itemid=" + CStr(itemarr(i)) + "" + VbCrlf
			sqlStr = sqlStr + " 	and itemoption='" + CStr(itemoptionarr(i)) + "'" + VbCrlf
			sqlStr = sqlStr + " ) "

			rsget.Open sqlStr,dbget,1
			if Not rsget.EOF then
				extbarcodeAlreadyExists = true
				extbarcodeAlreadyExistsString = extbarcodeAlreadyExistsString + extbarcodearr(i) + ","
			end if
			rsget.close
		end if

		if Not (extbarcodeAlreadyExists) then
			sqlStr = " update [db_shop].[dbo].tbl_shop_item"
			sqlStr = sqlStr + " set extbarcode='" + trim(CStr(extbarcodearr(i))) + "',"
			sqlStr = sqlStr + " isusing='" + CStr(isusingarr(i)) + "',"
			sqlStr = sqlStr + " updt=getdate()"

			sqlStr = sqlStr + " where itemgubun='" + itemgubunarr(i) + "'"
			sqlStr = sqlStr + " and shopitemid=" + CStr(itemarr(i)) + ""
			sqlStr = sqlStr + " and itemoption='" + CStr(itemoptionarr(i)) + "'"

			rsget.Open sqlStr,dbget,1


			''바코드 테이블 확인
			if trim(CStr(itemarr(i)))<>"" then
				sqlStr = " select top 1 * from [db_item].[dbo].tbl_item_option_stock" + VbCrlf
				sqlStr = sqlStr + " where itemgubun='" + itemgubunarr(i) + "'" + VbCrlf
				sqlStr = sqlStr + " and itemid=" + CStr(itemarr(i)) + "" + VbCrlf
				sqlStr = sqlStr + " and itemoption='" + CStr(itemoptionarr(i)) + "'" + VbCrlf
				rsget.Open sqlStr,dbget,1
				stockitemexists = (not rsget.Eof)
				rsget.close

				if (stockitemexists) then
					sqlStr = " update [db_item].[dbo].tbl_item_option_stock" + VbCrlf
					sqlStr = sqlStr + " set barcode='" + trim(CStr(extbarcodearr(i))) + "'" + VbCrlf
					sqlStr = sqlStr + " where itemgubun='" + itemgubunarr(i) + "'" + VbCrlf
					sqlStr = sqlStr + " and itemid=" + CStr(itemarr(i)) + "" + VbCrlf
					sqlStr = sqlStr + " and itemoption='" + CStr(itemoptionarr(i)) + "'" + VbCrlf

					rsget.Open sqlStr,dbget,1
				else
					sqlStr = " insert into [db_item].[dbo].tbl_item_option_stock" + VbCrlf
					sqlStr = sqlStr + " (itemgubun,itemid,itemoption,barcode)" + VbCrlf
					sqlStr = sqlStr + " values("
					sqlStr = sqlStr + " '" + itemgubunarr(i) + "'," + VbCrlf
					sqlStr = sqlStr + " " + CStr(itemarr(i)) + "," + VbCrlf
					sqlStr = sqlStr + " '" + itemoptionarr(i) + "'," + VbCrlf
					sqlStr = sqlStr + " '" + trim(extbarcodearr(i)) + "'" + VbCrlf
					sqlStr = sqlStr + " )" + VbCrlf
					rsget.Open sqlStr,dbget,1
				end if
			end if
		end if
	next

	if extbarcodeAlreadyExistsString<>"" then
		response.write "<script>alert('이미 존재하는 바코드 번호-" + extbarcodeAlreadyExistsString + "');</script>"

	end if

elseif mode ="arradd" then
	itemgubunarr = Left(itemgubunarr,Len(itemgubunarr)-1)
	itemarr = Left(itemarr,Len(itemarr)-1)
	itemoptionarr = Left(itemoptionarr,Len(itemoptionarr)-1)

	itemgubunarr = split(itemgubunarr,"|")
	itemarr = split(itemarr,"|")
	itemoptionarr = split(itemoptionarr,"|")

	cnt = ubound(itemarr)

	dim shopitemNotExists
	
	for i=0 to cnt
		shopitemNotExists = false

		sqlStr = " select count(s.shopitemid) as cnt" + VbCrlf
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item s " + VbCrlf
		sqlStr = sqlStr + " where shopitemid=" + itemarr(i) + VbCrlf
		sqlStr = sqlStr + " and itemoption='" + itemoptionarr(i) + "'" + VbCrlf
		sqlStr = sqlStr + " and itemgubun='" + itemgubunarr(i) + "'" + VbCrlf
		rsget.Open sqlStr,dbget,1
			shopitemNotExists = rsget("cnt")<1
		rsget.close

		if shopitemNotExists then
		    sqlStr = " insert into [db_shop].[dbo].tbl_shop_item" + VbCrlf
			sqlStr = sqlStr + " (itemgubun,shopitemid,itemoption," + VbCrlf
			sqlStr = sqlStr + " makerid,shopitemname,shopitemoptionname,shopitemprice,orgsellprice,shopsuplycash,shopbuyprice," + VbCrlf
			sqlStr = sqlStr + " extbarcode, vatinclude,catecdl,catecdm,catecdn)" + VbCrlf
			
			sqlStr = sqlStr + " select top 1 '10', i.itemid, '" + itemoptionarr(i) + "'"
			sqlStr = sqlStr + " , i.makerid, i.itemname, IsNull(v.optionname,'') as optname, i.sellcash, i.orgprice, 0, 0"
			sqlStr = sqlStr + " , IsNULL(s.barcode,'') as barcode ,i.vatinclude, i.itemserial_large, i.itemserial_mid, i.itemserial_small " + VbCrlf
			sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i " + VbCrlf
			sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option v on i.itemid=v.itemid and v.itemoption='" + itemoptionarr(i) + "'" + VbCrlf
			sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option_stock s on s.itemgubun='10' and i.itemid=s.itemid and s.itemoption='" + itemoptionarr(i) + "'" + VbCrlf
			sqlStr = sqlStr + " where i.itemid=" + itemarr(i) + VbCrlf

			dbget.Execute sqlStr
		end if

	next
elseif mode ="offitemreg" then
    '' MayBe Not Using
    response.write "<script>alert('사용 불가 메뉴입니다.- 관리자 문의요망')</script>"
    dbget.close()	:	response.End 
    
	'designer,itemname,sellcash
	dim nextitemid
	dim IsBrandShop

	IsBrandShop = false

	sqlStr = " select top 1 shopitemid"
	sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item"
	sqlStr = sqlStr + " where itemgubun='" + itemgubun + "'"
	sqlStr = sqlStr + " order by shopitemid desc"

	rsget.Open sqlStr,dbget,1
		if not rsget.Eof then
			nextitemid = rsget("shopitemid")+1
		else
			nextitemid = 1
		end if
	rsget.close


	sqlStr = " insert into [db_shop].[dbo].tbl_shop_item" + vbCrlf
	sqlStr = sqlStr + " (itemgubun,shopitemid,itemoption," + vbCrlf
	sqlStr = sqlStr + " makerid,shopitemname,shopitemoptionname,shopitemprice," + vbCrlf
	sqlStr = sqlStr + " shopsuplycash)" + vbCrlf
	sqlStr = sqlStr + " values(" + vbCrlf
	sqlStr = sqlStr + " '" + itemgubun + "'," + vbCrlf
	sqlStr = sqlStr + " " + CStr(nextitemid) + "," + vbCrlf
	sqlStr = sqlStr + " '0000'," + vbCrlf
	sqlStr = sqlStr + " '" + designer + "'," + vbCrlf
	sqlStr = sqlStr + " '" + Newhtml2DB(itemname) + "'," + vbCrlf
	sqlStr = sqlStr + " ''," + vbCrlf
	sqlStr = sqlStr + " " + CStr(sellcash) + "," + vbCrlf
	sqlStr = sqlStr + " " + CStr(suplycash) + "" + vbCrlf
	sqlStr = sqlStr + " )" + vbCrlf

	rsget.Open sqlStr,dbget,1

elseif mode="arrins" then

	sqlStr = " select top 1 statecd from [db_shop].[dbo].tbl_shop_ipchul_master"
	sqlStr = sqlStr + " where idx=" + CStr(idx)
	rsget.Open sqlStr, dbget, 1
	if not rsget.Eof then
		if rsget("statecd")<>"0" then
			response.write "<script>alert('현재 입고대기 상태가 아닙니다.');</script>"
			response.write "<script>location.replace('" + refer + "');</script>"
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
		sqlStr = sqlStr + "'" + itemgubunarr(i) + "'," + vbCrlf
		sqlStr = sqlStr + "" + itemarr(i) + "," + vbCrlf
		sqlStr = sqlStr + "'" + itemoptionarr(i) + "'," + vbCrlf
		sqlStr = sqlStr + "'" + designerarr(i) + "'," + vbCrlf
		sqlStr = sqlStr + "" + sellcasharr(i) + "," + vbCrlf
		sqlStr = sqlStr + "" + suplycasharr(i) + "," + vbCrlf
		sqlStr = sqlStr + "" + itemnoarr(i) + ")"

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
	''scheduledt = Cstr(dateserial(request("yyyy1"),request("mm1"),request("dd1")))
	scheduledt = request("scheduledt")

	sqlStr = " insert into [db_shop].[dbo].tbl_shop_ipchul_master"
	sqlStr = sqlStr + " (chargeid,shopid,divcode,vatcode,scheduledate,statecd,songjangdiv,songjangno)"
	sqlStr = sqlStr + " values('" + chargeid + "'"
	sqlStr = sqlStr + " ,'" + shopid + "'"
	sqlStr = sqlStr + " ,'" + divcode + "'"
	sqlStr = sqlStr + " ,'" + vatcode + "'"
	sqlStr = sqlStr + " ,'" + scheduledt + "'"
	sqlStr = sqlStr + " ,'0'"
	sqlStr = sqlStr + " ,'" + songjangdiv + "'"
	sqlStr = sqlStr + " ,'" + songjangno + "'"
	sqlStr = sqlStr + " )"

	dbget.Execute(sqlStr)


	sqlStr = " select ident_current('[db_shop].[dbo].tbl_shop_ipchul_master') as idx "
	rsget.Open sqlStr, dbget, 1
		idx = rsget("idx")
	rsget.close

	sqlStr = "update [db_shop].[dbo].tbl_shop_ipchul_master" + VbCrlf
	sqlStr = sqlStr + " set songjangname=IsNULL(T.divname,'')" + VbCrlf
	sqlStr = sqlStr + " from [db_order].[dbo].tbl_songjang_div T" + VbCrlf
	sqlStr = sqlStr + " where [db_shop].[dbo].tbl_shop_ipchul_master.idx=" + CStr(idx)
	sqlStr = sqlStr + " and [db_shop].[dbo].tbl_shop_ipchul_master.songjangdiv=T.divcd"

	dbget.Execute(sqlStr)


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
		sqlStr = sqlStr + "'" + itemgubunarr(i) + "'," + vbCrlf
		sqlStr = sqlStr + "" + itemarr(i) + "," + vbCrlf
		sqlStr = sqlStr + "'" + itemoptionarr(i) + "'," + vbCrlf
		sqlStr = sqlStr + "'" + designerarr(i) + "'," + vbCrlf
		sqlStr = sqlStr + "" + sellcasharr(i) + "," + vbCrlf
		sqlStr = sqlStr + "" + suplycasharr(i) + "," + vbCrlf
		sqlStr = sqlStr + "" + itemnoarr(i) + ")"

		dbget.Execute(sqlStr)
	next

	''상품명 옵션명
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
	sqlStr = sqlStr + " set totalsellcash=T.totsell," + vbCrlf
	sqlStr = sqlStr + " totalsuplycash=T.totsupp" + vbCrlf
	sqlStr = sqlStr + " from (select sum(sellcash*itemno) as totsell, " + vbCrlf
	sqlStr = sqlStr + " sum(suplycash*itemno) as totsupp from " + vbCrlf
	sqlStr = sqlStr + " [db_shop].[dbo].tbl_shop_ipchul_detail" + vbCrlf
	sqlStr = sqlStr + " where masteridx="  + CStr(idx) + vbCrlf
	sqlStr = sqlStr + " and deleteyn='N'" + vbCrlf
	sqlStr = sqlStr + " ) as T" + vbCrlf
	sqlStr = sqlStr + " where [db_shop].[dbo].tbl_shop_ipchul_master.idx=" + CStr(idx)

	dbget.Execute(sqlStr)
end if



if ((mode ="offitemreg") or (mode="arrins")) then
	if (InStr(refer,"&react=true")<1) then
		refer = refer + "&react=true"
	end if

elseif mode="addipchullist" then
	refer = "/designer/offshop/ipchullist.asp?menupos=196"
end if
%>

<script language="javascript">
alert('저장 되었습니다.');
location.replace('<%= refer %>');
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->