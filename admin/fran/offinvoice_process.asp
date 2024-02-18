<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<%

'==============================================================================
dim refer
refer = request.ServerVariables("HTTP_REFERER")

dim menupos
dim mode
dim masteridx, detailidx

dim shopid, workidx, invoiceno, invoicedate, delivermethod, exportmethod, jungsantype, priceunit, exchangerate, totalboxno, totalboxprice, totalgoodsprice, totalprice
dim exportdeclarefilename, exporteraddr, riskmesseraddr, notifyaddr, portname, destinationname, carriername, carrierdate, goodscomment1, goodscomment2
dim lccomment, lcbank, comment, reguserid

dim totalgoodspricecalc
dim totalboxpricecalc
dim totalpricecalc

dim productdetailmode, productdetailidx
dim productdetailcount, orderno, goodscomment
dim orderno_new, goodscomment_new

dim priceperbox
dim totalboxno_new, priceperbox_new, totalprice_new

dim jungsanidx

dim statecd

dim reportno, reportno2,reportno3
dim reportdate
dim reportpriceunit
dim reportexchangerate
dim reportforeigntotalprice
dim reporttotalprice

dim totalGoodsPriceWon, totalDeliverPriceWon, totalPriceWon, totalGoodsPriceForeign, totalDeliverPriceForeign, totalPriceForeign


'==============================================================================
menupos		= request("menupos")
mode		= request("mode")

masteridx	= request("masteridx")
detailidx	= request("detailidx")

shopid	= request("shopid")
workidx	= request("workidx")
invoiceno	= html2db(request("invoiceno"))
invoicedate	= html2db(request("invoicedate"))
delivermethod	= request("delivermethod")
exportmethod	= request("exportmethod")
jungsantype	= request("jungsantype")
priceunit	= request("priceunit")

exchangerate	= request("exchangerate")
totalboxno	= request("totalboxno")

'// 원화
totalboxprice	= Replace(request("totalboxprice"), ",", "")
totalgoodsprice	= Replace(request("totalgoodsprice"), ",", "")
totalprice	= Replace(request("totalprice"), ",", "")

'// 외화
totalgoodspricecalc	= Replace(request("totalgoodspricecalc"), ",", "")
totalboxpricecalc	= Replace(request("totalboxpricecalc"), ",", "")
totalpricecalc		= Replace(request("totalpricecalc"), ",", "")

totalGoodsPriceWon			= Replace(request("totalGoodsPriceWon"), ",", "")
totalDeliverPriceWon		= Replace(request("totalDeliverPriceWon"), ",", "")
totalPriceWon				= Replace(request("totalPriceWon"), ",", "")
totalGoodsPriceForeign		= Replace(request("totalGoodsPriceForeign"), ",", "")
totalDeliverPriceForeign	= Replace(request("totalDeliverPriceForeign"), ",", "")
totalPriceForeign			= Replace(request("totalPriceForeign"), ",", "")

'// 과거 데이타 호환
if (totalgoodsprice = "") and (totalGoodsPriceWon <> "") then
	totalgoodsprice = totalGoodsPriceWon
end if

if (totalboxprice = "") and (totalDeliverPriceWon <> "") then
	totalboxprice = totalDeliverPriceWon
end if

if (totalpricecalc = "") and (totalPriceForeign <> "") then
	totalpricecalc = totalPriceForeign
end if

totalprice = totalpricecalc

exportdeclarefilename	= html2db(request("exportdeclarefilename"))
exporteraddr	= html2db(request("exporteraddr"))
riskmesseraddr	= html2db(request("riskmesseraddr"))
notifyaddr	= html2db(request("notifyaddr"))
portname	= html2db(request("portname"))
destinationname	= html2db(request("destinationname"))
carriername	= html2db(request("carriername"))
carrierdate	= html2db(request("carrierdate"))
goodscomment1	= html2db(request("goodscomment1"))
goodscomment2	= html2db(request("goodscomment2"))

lccomment	= html2db(request("lccomment"))
lcbank	= html2db(request("lcbank"))
comment	= html2db(request("comment"))
reguserid	= session("ssBctid")

productdetailmode	= html2db(request("productdetailmode"))
productdetailidx	= html2db(request("productdetailidx"))
productdetailcount	= html2db(request("productdetailcount"))
orderno_new			= html2db(request("orderno_new"))
goodscomment_new	= html2db(request("goodscomment_new"))
totalboxno_new		= html2db(request("totalboxno_new"))
priceperbox_new		= html2db(request("priceperbox_new"))
totalprice_new		= html2db(request("totalprice_new"))

jungsanidx	= request("jungsanidx")

statecd	= request("statecd")

reportno				= request("reportno")
reportno2				= request("reportno2")
reportno3				= request("reportno3")

reportdate				= request("reportdate")
reportpriceunit			= request("reportpriceunit")
reportexchangerate		= request("reportexchangerate")
reportforeigntotalprice	= request("reportforeigntotalprice")
reporttotalprice		= request("reporttotalprice")

reportexchangerate		= Replace(request("reportexchangerate"), ",", "")
reportforeigntotalprice	= Replace(request("reportforeigntotalprice"), ",", "")
reporttotalprice		= Replace(request("reporttotalprice"), ",", "")


if (workidx = "") then
	workidx = 0
end if

if (jungsanidx = "") then
	jungsanidx = 0
end if

if (exchangerate = "") then
	exchangerate = 0
end if

if (totalboxno = "") or (CStr(totalboxno) = "0") then
	totalboxno = 1
end if

if (totalboxprice = "") then
	totalboxprice = 0
end if

if (totalgoodsprice = "") then
	totalgoodsprice = 0
end if

if (totalprice = "") then
	totalprice = 0
end if

if (reportexchangerate = "") then
	reportexchangerate = 0
end if

if (reportforeigntotalprice = "") then
	reportforeigntotalprice = 0
end if

if (reporttotalprice = "") then
	reporttotalprice = 0
end if

function insertDetailFromWork(byval masteridx, byval shopid, byval workidx)
	dim sqlStr

	sqlStr = " delete from [db_storage].[dbo].tbl_offline_invoice_detail "
	sqlStr = sqlStr + " where masteridx = " + CStr(masteridx) + " "
	dbget.Execute sqlStr

	if (workidx <> "") and (CStr(workidx) <> "0") then

		sqlStr = " insert into [db_storage].[dbo].tbl_offline_invoice_detail( "
		sqlStr = sqlStr + " 	masteridx "
		sqlStr = sqlStr + " 	, cartonboxno "
		sqlStr = sqlStr + " 	, goodscomment "
		sqlStr = sqlStr + " 	, nweight "
		sqlStr = sqlStr + " 	, gweight "
		sqlStr = sqlStr + " 	, emsPrice "
		sqlStr = sqlStr + " ) "
		sqlStr = sqlStr + " select "
		sqlStr = sqlStr + " 	" + CStr(masteridx) + " "
		sqlStr = sqlStr + " 	, cartoonboxno "
		sqlStr = sqlStr + " 	, 'Stationary & Gifts' "
		sqlStr = sqlStr + " 	, sum(innerboxweight) "
		sqlStr = sqlStr + " 	, max(cartoonboxweight) "
		sqlStr = sqlStr + " 	, db_storage.[dbo].[uf_getEmsPrice]('" + CStr(shopid) + "', max(cartoonboxweight)*1000) "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	db_storage.dbo.tbl_cartoonbox_detail "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + " 	and masteridx = " + CStr(workidx) + " "
		sqlStr = sqlStr + " 	and shopid = '" + CStr(shopid) + "' "
		sqlStr = sqlStr + " group by "
		sqlStr = sqlStr + " 	cartoonboxno "
		'response.write "aaaaaaaaaaaa" & sqlStr
		dbget.Execute sqlStr

		sqlStr = " update "
		sqlStr = sqlStr + " 	m "
		sqlStr = sqlStr + " set "
		sqlStr = sqlStr + " 	totalboxno = T.cnt "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	[db_storage].[dbo].tbl_offline_invoice_master m "
		sqlStr = sqlStr + " 	join ( "
		sqlStr = sqlStr + " 		select masteridx, count(idx) as cnt "
		sqlStr = sqlStr + " 		from "
		sqlStr = sqlStr + " 			[db_storage].[dbo].tbl_offline_invoice_detail "
		sqlStr = sqlStr + " 		where "
		sqlStr = sqlStr + " 			masteridx = " + CStr(masteridx) + " "
		sqlStr = sqlStr + " 		group by "
		sqlStr = sqlStr + " 			masteridx "
		sqlStr = sqlStr + " 	) T "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		m.idx = T.masteridx "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	m.idx = " + CStr(masteridx) + " "
		dbget.Execute sqlStr

	end if

end	function

function insertDefaultGoodsDescription(byval masteridx, byval totalGoodsPriceForeign, byval totalDeliverPriceForeign)
	dim sqlStr
	dim totalboxno

	sqlStr = " delete from [db_storage].[dbo].tbl_offline_invoice_product_detail "
	sqlStr = sqlStr + " where masteridx = " + CStr(masteridx) + " "
	dbget.Execute sqlStr

	'// =======================================================================
	'// 상품정보 자동입력
	sqlStr = " insert into [db_storage].[dbo].tbl_offline_invoice_product_detail( "
	sqlStr = sqlStr + " 	masteridx "
	sqlStr = sqlStr + " 	, orderno "
	sqlStr = sqlStr + " 	, goodscomment "
	sqlStr = sqlStr + " 	, totalboxno "
	sqlStr = sqlStr + " 	, priceperbox "
	sqlStr = sqlStr + " 	, totalprice "
	sqlStr = sqlStr + " ) "
	sqlStr = sqlStr + " values( "
	sqlStr = sqlStr + " 	" + CStr(masteridx) + " "
	sqlStr = sqlStr + " 	, 1"
	sqlStr = sqlStr + " 	, 'Stationary & Gifts' "
	sqlStr = sqlStr + " 	, " + CStr(1) + " "
	sqlStr = sqlStr + " 	, " + CStr(totalGoodsPriceForeign) + " "
	sqlStr = sqlStr + " 	, " + CStr(totalGoodsPriceForeign) + " "
	sqlStr = sqlStr + " ) "
	'response.write sqlStr
	dbget.Execute sqlStr

	'// =======================================================================
	'// 운임정보 자동입력
	sqlStr = " insert into [db_storage].[dbo].tbl_offline_invoice_product_detail( "
	sqlStr = sqlStr + " 	masteridx "
	sqlStr = sqlStr + " 	, orderno "
	sqlStr = sqlStr + " 	, goodscomment "
	sqlStr = sqlStr + " 	, totalprice "
	sqlStr = sqlStr + " ) "
	sqlStr = sqlStr + " values( "
	sqlStr = sqlStr + " 	" + CStr(masteridx) + " "
	sqlStr = sqlStr + " 	, 2"
	sqlStr = sqlStr + " 	, 'Freight charge' "
	sqlStr = sqlStr + " 	, " + CStr(totalDeliverPriceForeign) + " "
	sqlStr = sqlStr + " ) "
	dbget.Execute sqlStr

	totalboxno = 1
	sqlStr = " select totalboxno from [db_storage].[dbo].tbl_offline_invoice_master where idx = " + CStr(masteridx)
	rsget.Open sqlStr, dbget, 1
	If Not rsget.EOF Then
		totalboxno = rsget("totalboxno")
	end if

	sqlStr = " update [db_storage].[dbo].tbl_offline_invoice_product_detail "
	sqlStr = sqlStr + " set totalboxno = " + CStr(totalboxno) + ", priceperbox = (totalprice/" + CStr(totalboxno) + ") "
	sqlStr = sqlStr + " where masteridx = " + CStr(masteridx) + " and totalprice = " + CStr(totalGoodsPriceForeign) + " and totalboxno = 1 "
	dbget.Execute sqlStr

end	function

'==============================================================================
dim sqlStr,i, iid

if (mode = "newmaster") then

	sqlStr = " select * from [db_storage].[dbo].tbl_offline_invoice_master where 1=0"

	'response.write sqlStr &"<Br>"
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew
	rsget("shopid") = shopid
	rsget("workidx") = workidx
	rsget("invoiceno") = invoiceno
	rsget("invoicedate") = invoicedate
	rsget("delivermethod") = delivermethod
	rsget("exportmethod") = exportmethod
	rsget("jungsantype") = jungsantype
	rsget("priceunit") = priceunit
	rsget("exchangerate") = exchangerate
	rsget("totalboxno") = totalboxno
	rsget("totalboxprice") = totalboxprice
	rsget("totalgoodsprice") = totalgoodsprice
	rsget("totalprice") = totalprice
	rsget("exportdeclarefilename") = exportdeclarefilename
	rsget("exporteraddr") = exporteraddr
	rsget("riskmesseraddr") = riskmesseraddr
	rsget("notifyaddr") = notifyaddr
	rsget("portname") = portname
	rsget("destinationname") = destinationname
	rsget("carriername") = carriername
	rsget("carrierdate") = carrierdate
	rsget("goodscomment1") = goodscomment1
	rsget("goodscomment2") = goodscomment2
	rsget("lccomment") = lccomment
	rsget("lcbank") = lcbank
	rsget("comment") = comment
	rsget("statecd") = "1"		'// 작성중
	rsget("reguserid") = reguserid

	rsget("totalGoodsPriceWon") = totalGoodsPriceWon
	rsget("totalDeliverPriceWon") = totalDeliverPriceWon
	rsget("totalPriceWon") = totalPriceWon
	rsget("totalGoodsPriceForeign") = totalGoodsPriceForeign
	rsget("totalDeliverPriceForeign") = totalDeliverPriceForeign
	rsget("totalPriceForeign") = totalPriceForeign

	rsget.update
		iid = rsget("idx")
	rsget.close

	if (jungsanidx <> 0) then
		sqlStr = " update [db_shop].[dbo].tbl_fran_meachuljungsan_master "
		sqlStr = sqlStr + " set  "
		sqlStr = sqlStr + " 	invoiceidx = " + CStr(iid) + " "
		sqlStr = sqlStr + " 	, issuestatecd = '0' "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	idx = " + CStr(jungsanidx) + " "
		dbget.Execute sqlStr
	end if

	if (workidx <> 0) then
		Call insertDetailFromWork(iid, shopid, workidx)

		sqlStr = " update [db_storage].[dbo].tbl_cartoonbox_master "
		sqlStr = sqlStr + " set  "
		sqlStr = sqlStr + " 	invoiceidx = " + CStr(iid) + " "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	idx = " + CStr(workidx) + " "
		dbget.Execute sqlStr
	end if

	Call insertDefaultGoodsDescription(iid, totalGoodsPriceForeign, totalDeliverPriceForeign)

	refer = refer + "&idx=" + CStr(iid)

elseif (mode = "savemaster") then

	sqlStr = " update "
	sqlStr = sqlStr + " 	[db_storage].[dbo].tbl_offline_invoice_master "
	sqlStr = sqlStr + " set "
	sqlStr = sqlStr + " 	shopid = '" + CStr(shopid) + "' "
	sqlStr = sqlStr + " 	, workidx = " + CStr(workidx) + " "
	sqlStr = sqlStr + " 	, invoiceno = '" + CStr(invoiceno) + "' "
	sqlStr = sqlStr + " 	, invoicedate = '" + CStr(invoicedate) + "' "
	sqlStr = sqlStr + " 	, delivermethod = '" + CStr(delivermethod) + "' "
	sqlStr = sqlStr + " 	, exportmethod = '" + CStr(exportmethod) + "' "
	sqlStr = sqlStr + " 	, jungsantype = '" + CStr(jungsantype) + "' "
	sqlStr = sqlStr + " 	, priceunit = '" + CStr(priceunit) + "' "
	sqlStr = sqlStr + " 	, exchangerate = " + CStr(exchangerate) + " "
	sqlStr = sqlStr + " 	, totalboxno = " + CStr(totalboxno) + " "
	sqlStr = sqlStr + " 	, totalboxprice = " + CStr(totalboxprice) + " "
	sqlStr = sqlStr + " 	, totalgoodsprice = " + CStr(totalgoodsprice) + " "
	sqlStr = sqlStr + " 	, totalprice = " + CStr(totalprice) + " "
	sqlStr = sqlStr + " 	, exporteraddr = '" + CStr(exporteraddr) + "' "
	sqlStr = sqlStr + " 	, riskmesseraddr = '" + CStr(riskmesseraddr) + "' "
	sqlStr = sqlStr + " 	, notifyaddr = '" + CStr(notifyaddr) + "' "
	sqlStr = sqlStr + " 	, portname = '" + CStr(portname) + "' "
	sqlStr = sqlStr + " 	, destinationname = '" + CStr(destinationname) + "' "
	sqlStr = sqlStr + " 	, carriername = '" + CStr(carriername) + "' "
	sqlStr = sqlStr + " 	, carrierdate = '" + CStr(carrierdate) + "' "
	sqlStr = sqlStr + " 	, goodscomment1 = '" + CStr(goodscomment1) + "' "
	sqlStr = sqlStr + " 	, goodscomment2 = '" + CStr(goodscomment2) + "' "
	sqlStr = sqlStr + " 	, lccomment = '" + CStr(lccomment) + "' "
	sqlStr = sqlStr + " 	, lcbank = '" + CStr(lcbank) + "' "
	sqlStr = sqlStr + " 	, comment = '" + CStr(comment) + "' "
	sqlStr = sqlStr + " 	, reguserid = '" + CStr(reguserid) + "' "
	sqlStr = sqlStr + " 	, reportno = '" + CStr(reportno) + "' "
	sqlStr = sqlStr + " 	, reportno2 = '" + CStr(reportno2) + "' "
	sqlStr = sqlStr + " 	, reportno3 = '" + CStr(reportno3) + "' "
	sqlStr = sqlStr + " 	, reportdate = '" + CStr(reportdate) + "' "										'// 신고일자
	sqlStr = sqlStr + " 	, reportpriceunit = '" + CStr(reportpriceunit) + "' "
	sqlStr = sqlStr + " 	, reportexchangerate = '" + CStr(reportexchangerate) + "' "
	sqlStr = sqlStr + " 	, reportforeigntotalprice = '" + CStr(reportforeigntotalprice) + "' "
	sqlStr = sqlStr + " 	, reporttotalprice = '" + CStr(reporttotalprice) + "' "
	sqlStr = sqlStr + " 	, totalGoodsPriceWon = " + CStr(totalGoodsPriceWon) + " "
	sqlStr = sqlStr + " 	, totalDeliverPriceWon = " + CStr(totalDeliverPriceWon) + " "
	sqlStr = sqlStr + " 	, totalPriceWon = " + CStr(totalPriceWon) + " "
	sqlStr = sqlStr + " 	, totalGoodsPriceForeign = " + CStr(totalGoodsPriceForeign) + " "
	sqlStr = sqlStr + " 	, totalDeliverPriceForeign = " + CStr(totalDeliverPriceForeign) + " "
	sqlStr = sqlStr + " 	, totalPriceForeign = " + CStr(totalPriceForeign) + " "
	sqlStr = sqlStr + " 	, lastupdate = getdate() "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	idx = " + CStr(masteridx) + "	 "
	'response.write sqlStr &"<Br>"
	dbget.Execute sqlStr

	if (reportdate <> "") then
		sqlStr = " update "
		sqlStr = sqlStr + " [db_shop].[dbo].tbl_fran_meachuljungsan_master "
		sqlStr = sqlStr + " set taxdate = '" + CStr(reportdate) + "' "
		sqlStr = sqlStr + " where invoiceidx = " + CStr(masteridx) + " and papertype = 200 "
		dbget.Execute sqlStr
	end if

	if (productdetailmode = "deletedetailone") then
		sqlStr = " delete from [db_storage].[dbo].tbl_offline_invoice_product_detail "
		sqlStr = sqlStr + " where masteridx = " + CStr(masteridx) + " and idx = " + CStr(productdetailidx) + " "
		dbget.Execute sqlStr
	elseif (productdetailmode = "adddetailone") then
		if (totalboxno_new = "") then
			totalboxno_new = "NULL"
		end if
		if (priceperbox_new = "") then
			priceperbox_new = "NULL"
		end if
		if (totalprice_new = "") then
			totalprice_new = "NULL"
		end if

		sqlStr = " insert into [db_storage].[dbo].tbl_offline_invoice_product_detail( "
		sqlStr = sqlStr + " 	masteridx "
		sqlStr = sqlStr + " 	, orderno "
		sqlStr = sqlStr + " 	, goodscomment "
		sqlStr = sqlStr + " 	, totalboxno "
		sqlStr = sqlStr + " 	, priceperbox "
		sqlStr = sqlStr + " 	, totalprice "
		sqlStr = sqlStr + " ) "
		sqlStr = sqlStr + " values( "
		sqlStr = sqlStr + " 	" + CStr(masteridx) + " "
		sqlStr = sqlStr + " 	, " + CStr(orderno_new) + " "
		sqlStr = sqlStr + " 	, '" + CStr(goodscomment_new) + "' "
		sqlStr = sqlStr + " 	, " + CStr(totalboxno_new) + " "
		sqlStr = sqlStr + " 	, " + CStr(priceperbox_new) + " "
		sqlStr = sqlStr + " 	, " + CStr(totalprice_new) + " "
		sqlStr = sqlStr + " ) "
		'response.write "aaaaaaaaaaaa" & sqlStr
		dbget.Execute sqlStr
		'dbget.close
		'response.end
	elseif (productdetailmode = "modifydetailall") then
		for i = 0 to productdetailcount*1 - 1
			productdetailidx	= request("productdetailidx_" + CStr(i))
			orderno				= request("orderno_" + CStr(i))
			goodscomment		= html2db(request("goodscomment_" + CStr(i)))

			totalboxno			= html2db(request("totalboxno_" + CStr(i)))
			priceperbox			= html2db(request("priceperbox_" + CStr(i)))
			totalprice			= html2db(request("totalprice_" + CStr(i)))

			if (totalboxno = "") then
				totalboxno = "NULL"
			end if
			if (priceperbox = "") then
				priceperbox = "NULL"
			end if
			if (totalprice = "") then
				totalprice = "NULL"
			end if

			sqlStr = " update "
			sqlStr = sqlStr + " 	[db_storage].[dbo].tbl_offline_invoice_product_detail "
			sqlStr = sqlStr + " set "
			sqlStr = sqlStr + " 	orderno = " + CStr(orderno) + " "
			sqlStr = sqlStr + " 	, goodscomment = '" + CStr(goodscomment) + "' "

			sqlStr = sqlStr + " 	, totalboxno = " + CStr(totalboxno) + " "
			sqlStr = sqlStr + " 	, priceperbox = " + CStr(priceperbox) + " "
			sqlStr = sqlStr + " 	, totalprice = " + CStr(totalprice) + " "

			sqlStr = sqlStr + " where masteridx = " + CStr(masteridx) + " and idx = " + CStr(productdetailidx) + " "
			dbget.Execute sqlStr
		next
	end if

elseif (mode = "delmaster") then

	sqlStr = " delete from [db_storage].[dbo].tbl_offline_invoice_master "
	sqlStr = sqlStr + " where idx = " + CStr(masteridx) + " "
	dbget.Execute sqlStr

	sqlStr = " update [db_storage].[dbo].tbl_offline_invoice_detail "
	sqlStr = sqlStr + " set masteridx = null "
	sqlStr = sqlStr + " where masteridx = " + CStr(masteridx) + " "
	''dbget.Execute sqlStr

	sqlStr = " delete from [db_storage].[dbo].tbl_offline_invoice_product_detail "
	sqlStr = sqlStr + " where masteridx = " + CStr(masteridx)
	dbget.Execute sqlStr

	if (jungsanidx <> 0) then
		sqlStr = " update [db_shop].[dbo].tbl_fran_meachuljungsan_master "
		sqlStr = sqlStr + " set  "
		sqlStr = sqlStr + " 	invoiceidx = NULL "
		sqlStr = sqlStr + " 	, issuestatecd = NULL "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	idx = " + CStr(jungsanidx) + " "
		dbget.Execute sqlStr
	end if

	if (workidx <> 0) then
		sqlStr = " update [db_storage].[dbo].tbl_cartoonbox_master "
		sqlStr = sqlStr + " set  "
		sqlStr = sqlStr + " 	invoiceidx = NULL "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	idx = " + CStr(workidx) + " "
		dbget.Execute sqlStr
	end if

	refer = "/admin/fran/offinvoice_list.asp?menupos=" + CStr(menupos)

elseif (mode = "insertdetailfromwork") then

	Call insertDetailFromWork(masteridx, shopid, workidx)

elseif (mode = "insertdefaultdescription") then

	Call insertDefaultGoodsDescription(masteridx, totalGoodsPriceForeign, totalDeliverPriceForeign)

elseif (mode = "modifystate") then

	sqlStr = " update "
	sqlStr = sqlStr + " 	[db_storage].[dbo].tbl_offline_invoice_master "
	sqlStr = sqlStr + " set "
	sqlStr = sqlStr + " 	statecd = '" + CStr(statecd) + "' "
	sqlStr = sqlStr + " 	, reguserid = '" + CStr(reguserid) + "' "
	sqlStr = sqlStr + " 	, lastupdate = getdate() "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	idx = " + CStr(masteridx) + "	 "
	'response.write sqlStr &"<Br>"
	dbget.Execute sqlStr

	if (jungsanidx <> 0) then
		sqlStr = " update [db_shop].[dbo].tbl_fran_meachuljungsan_master "
		sqlStr = sqlStr + " set  "

		if (statecd = "7") then
			sqlStr = sqlStr + " 	issuestatecd = '9' "
		else
			sqlStr = sqlStr + " 	issuestatecd = '0' "
		end if

		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	idx = " + CStr(jungsanidx) + " "
		dbget.Execute sqlStr
	end if

end if

%>

<script language="javascript">
	alert('저장 되었습니다.');
	location.replace('<%= refer %>');
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->
