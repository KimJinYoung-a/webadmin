<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<%
session.codepage = 65001
response.Charset="UTF-8"
%>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : 텐바이텐 대량구매 주문서
' Hieditor : 2013.07.15 한용민 생성
' ///////////// 이 페이지 수정시 밑에 페이지도 똑같이 수정해야 한다. /////////////////////
'scm : http://webadmin.10x10.co.kr/common/popOrderSheet_foreign_excel.asp
'wholesale : http://wholesale.10x10.co.kr/mywholesale/order/popOrderSheet_foreign_excel.asp
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/BarcodeFunction.asp" -->
<!-- #include virtual="/lib/classes/stock/ipchulbarcodecls.asp"-->
<!-- #include virtual="/lib/classes/stock/cartoonboxcls.asp" -->
<!-- #include virtual="\lib\classes\stock\ordersheetcls.asp" -->
<%
dim chulgoyn, showdeleted, showmichulgo, michulgoreason, statecd, itemid, makerid, shopdiv
dim day5chulgo, shortchulgo, tempshort, danjong, etcshort, innerboxno, research, dateType
dim yyyy1,mm1 , dd1, yyyy2, mm2, dd2, fromDate, toDate
dim chkimg
dim masteridx, baljucode, shopid, i, jumunwait, isfixed, IsForeignOrder, IsForeign_confirmed, cartoonboxmasteridx
dim listgubun, oOneOrder, pagecode, sqlStr
	listgubun 		= requestCheckVar(request("listgubun"), 32)
	masteridx = getNumeric(requestCheckVar(request("masteridx"),10))
	shopid = requestCheckVar(request("shopid"),32)
	baljucode = requestCheckVar(request("baljucode"),32)
	cartoonboxmasteridx = getNumeric(requestCheckVar(request("cartoonboxmasteridx"),10))
	yyyy1 = requestCheckVar(request("yyyy1"),4)
	mm1 = requestCheckVar(request("mm1"),2)
	dd1 = requestCheckVar(request("dd1"),2)
	yyyy2 = requestCheckVar(request("yyyy2"),4)
	mm2 = requestCheckVar(request("mm2"),2)
	dd2 = requestCheckVar(request("dd2"),2)
	chulgoyn = requestCheckVar(request("chulgoyn"),1)
	showdeleted = requestCheckVar(request("showdel"),1)		'웹서버 웹나이트가 파라미터중 delete 문구가 있는 경우 막는다.
	showmichulgo = requestCheckVar(request("showmichulgo"),1)
	michulgoreason = requestCheckVar(request("michulgoreason"),32)
	innerboxno = requestCheckVar(request("innerboxno"),10)
	statecd = requestCheckVar(request("statecd"),10)
	itemid = requestCheckVar(request("itemid"),10)
	makerid = requestCheckVar(request("makerid"),32)
	shopdiv = requestCheckVar(request("shopdiv"),1)
	day5chulgo = requestCheckVar(request("day5chulgo"),1)
	shortchulgo = requestCheckVar(request("shortchulgo"),1)
	tempshort = requestCheckVar(request("tempshort"),10)
	danjong = requestCheckVar(request("danjong"),1)
	etcshort = requestCheckVar(request("etcshort"),1)
	research = requestCheckVar(request("research"),10)
	dateType = requestCheckVar(request("dateType"),1)
    chkimg = requestCheckVar(request("chkimg"),10)

Dim isImageVisible : isImageVisible = (chkimg="on")

if dateType="" then dateType="B"
if (research = "") then
	showdeleted = "N"
	michulgoreason = "all"
end if

michulgoreason = "|"
if (day5chulgo = "Y") then
	'5일내출고
	michulgoreason = michulgoreason + "5|"
end if
if (shortchulgo = "Y") then
	'재고부족
	michulgoreason = michulgoreason + "S|"
end if
if (tempshort = "Y") then
	'일시품절
	michulgoreason = michulgoreason + "T|"
end if
if (danjong = "Y") then
	'단종
	michulgoreason = michulgoreason + "D|"
end if
if (etcshort = "Y") then
	'기타
	michulgoreason = michulgoreason + "E|"
end if

'if (yyyy1="") then
'	yyyy1 = Cstr(Year(now()))
'	mm1 = Cstr(Month(now()))
'	dd1 = Cstr(day(now()))
'end if
'
'if (yyyy2="") then
'	yyyy2 = Cstr(Year(now()))
'	mm2 = Cstr(Month(now()))
'	dd2 = Cstr(day(now()))
'end if

if yyyy1<>"" and mm1<>"" and dd1<>"" then
	fromDate = DateSerial(yyyy1, mm1, dd1)
end if
if yyyy2<>"" and mm2<>"" and dd2<>"" then
	toDate = DateSerial(yyyy2, mm2, dd2+1)
end if

'/매장
if (C_IS_SHOP) then

	'//직영점일때
	if C_IS_OWN_SHOP then

		'/어드민권한 점장 미만
		'if getlevel_sn("",session("ssBctId")) > 6 then
			'shopid = C_STREETSHOPID		'"streetshop011"
		'end if
	else
		shopid = C_STREETSHOPID
	end if
else
	'/업체
	if (C_IS_Maker_Upche) then
		makerid = session("ssBctID")	'"7321"
	else
		if (Not C_ADMIN_USER) then
		    shopid = "X"                ''다른매장조회 막음.
		else
		end if
	end if
end if

'/패킹리스트
if cartoonboxmasteridx<>"" then
	listgubun = "PACKING"

	SET oOneOrder = new CCartoonBox
		oOneOrder.FRectMasterIdx=cartoonboxmasteridx
		oOneOrder.GetMasterOne

	if (oOneOrder.FresultCount<1) then
	    response.write "<script type='text/javascript'>alert('Invalid Order[1]');history.back;</script>"
	    session.codePage = 949
	    dbget.close() : response.end
	end if

	'pagecode = "_" & cartoonboxmasteridx
	pagecode = "_" & oOneOrder.FOneItem.Finvoiceidx
	if oOneOrder.FOneItem.Fdeliverdt<>"" and not(isnull(oOneOrder.FOneItem.Fdeliverdt)) then
		pagecode = pagecode & "_" & trim(left(oOneOrder.FOneItem.Fdeliverdt,10))
	end if

'/주문리스트
elseif masteridx<>"" or baljucode<>"" then
	listgubun = "JUMUN"

	if baljucode <> "" and masteridx="" then
		sqlStr = "select idx from db_storage.dbo.tbl_ordersheet_master where baljucode = '" & baljucode & "' "
		rsget.Open sqlStr,dbget,1
		if not rsget.EOF then
			masteridx = rsget("idx")
		end if
		rsget.Close
	end if

	SET oOneOrder = new CStorageMaster
		oOneOrder.frectsitename = "WSLWEB"
		oOneOrder.FRectOrderIDX=masteridx
		oOneOrder.FRectAuthMode = "none"
		oOneOrder.getShopOneOrderMaster

	if (oOneOrder.FresultCount<1) then
	    response.write "<script type='text/javascript'>alert('Invalid Order[2]');history.back;</script>"
	    session.codePage = 949
	    dbget.close() : response.end
	end if

	pagecode = "_" & baljucode
	if oOneOrder.FOneItem.Fscheduledate<>"" and not(isnull(oOneOrder.FOneItem.Fscheduledate)) then
		pagecode = pagecode & "_" & trim(left(oOneOrder.FOneItem.Fscheduledate,10))
	end if

'//구분자가 아예없는 전체 리스트
else
	listgubun = "NONE"

	pagecode = "_ITEMLIST"
end if

IsForeignOrder = false		'/업체접수주문
IsForeign_confirmed = false		'/업체접수주문 컨펌완료여부

dim oforeign_detail
set oforeign_detail = new CStorageDetail
	oforeign_detail.FPageSize = 5000
	oforeign_detail.FCurrPage = 1
	oforeign_detail.FRectbaljucode = baljucode
	oforeign_detail.FRectMasterIdx = masteridx
	oforeign_detail.FRectshopid = shopid
	oforeign_detail.FRectmakerid = makerid
	oforeign_detail.FRectItemid = itemid
	oforeign_detail.FRectstartdate = fromDate
	oforeign_detail.FRectenddate = toDate
	oforeign_detail.FRectinnerboxno = innerboxno
	oforeign_detail.FRectShopdiv = shopdiv
	oforeign_detail.FRectShowDeleted = "N"
	oforeign_detail.FRectMichulgoReason = michulgoreason
	oforeign_detail.FRectDateType = dateType
	oforeign_detail.FRectcartoonboxmasteridx = cartoonboxmasteridx

	if (statecd = "A") then
		oforeign_detail.FRectChulgoYN = "N"
	else
		oforeign_detail.FRectStatecd = statecd
	end if

	oforeign_detail.Getordersheet_foreign_detail

if listgubun = "JUMUN" then
	if oOneOrder.FOneItem.FStatecd=" " then
		jumunwait = true	'/주문서작성중
	end if
	isfixed = oOneOrder.FOneItem.IsFixed
	if oOneOrder.FOneItem.fforeign_statecd<>"" then
		IsForeignOrder=true

		if oOneOrder.FOneItem.fforeign_statecd>="3" then
		'if oOneOrder.FOneItem.fforeign_statecd="7" then
			IsForeign_confirmed = true
		end if
	else
		IsForeign_confirmed = true
	end if
end if

if (NOT isImageVisible) then
	Response.Expires=0
	response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "Content-Disposition", "attachment; filename=TEN_" & shopid & pagecode & ".xls"
	Response.CacheControl = "public"
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset="utf-8">
<title>텐바이텐주문서</title>
<style type='text/css'>
	.txt {mso-number-format:'\@'}
</style>
</head>
<body>
<table width="100%" align="center" border="1" cellpadding="3" cellspacing="0" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<% if isImageVisible then %>
		<td colspan="29">
	<% else %>
		<td colspan="35">
	<% end if %>

		검색결과 : <%= oforeign_detail.FTotalCount %>
		<%=CHKIIF(isImageVisible,"<br>엑셀에서 행높이 100 으로 설정한 후 카피 and 페이스트","")%>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
<% if NOT (isImageVisible) then %>
	<td>Order NO</td>
	<td>Release NO</td>
	<td>Order Date</td>
	<td>Packing Date</td>
	<td>Inner Box NO</td>
	<td>Carton Box NO</td>
	<td>Brand</td>
	<td>Item Code</td>
	<td>Barcode</td>
	<td>Item Name</td>
	<td>Option Name</td>
	<td>ORDER<br>QTY</td>
	<td>CONFIRM<br>QTY</td>

	<% 'if listgubun = "PACKING" or listgubun = "NONE" then %>
		<!--<td>CONFIRM<br>QTY</td>-->
	<% 'else %>
		<%
		'/주문서작성중이 아닌거
		'if not(jumunwait) then
		%>
			<!--<td>ORDER<br>QTY</td>-->
			<!--<td>CONFIRM<br>QTY</td>-->
		<% 'else %>
			<!--<td>ORDER<br>QTY</td>-->
		<% 'end if %>
	<% 'end if %>

	<td>Currencyunit</td>
	<td>Retail Price</td>
	<td>Wholesale Price</td>
	<td>Discount Rate<Br>(%)</td>
	<td>Total Price</td>
	<td>RRP</td>
	<td>exchangeRate</td>
	<td>multipleRate</td>
	<td></td>
	<td>Item Name<Br>[KR]</td>
	<td>Option Name<Br>[KR]</td>
	<td>Material<Br>[KR]</td>
	<td>Origin<Br>[KR]</td>
	<td></td>
	<td>Item Name<Br>[EN]</td>
	<td>Option Name<Br>[EN]</td>
	<td>Material<Br>[EN]</td>
	<td>Origin<Br>[EN]</td>
	<td>Item<Br>Weight(g)</td>
	<td>category1<Br>[CODE]</td>
	<td>category2<Br>[CODE]</td>
	<td>category3<Br>[CODE]</td>
	<!--<td>category1<Br>[KR]</td>
	<td>category2<Br>[KR]</td>
	<td>category3<Br>[KR]</td>
	<td>category2<Br>[EN]</td>
	<td>category3<Br>[EN]</td>-->
<% else %>
    <td>Order NO</td>
	<td>Release NO</td>
	<td>Inner Box NO</td>
	<td>Carton Box NO</td>
	<td>Item Code</td>
	<td>Item Name<Br>[EN]</td>
	<td>Option Name<Br>[EN]</td>
	<td>Material<Br>[EN]</td>
	<td>Origin<Br>[EN]</td>
	<td>Item<Br>Weight(g)</td>
	<td>Image</td>
<% end if %>
</tr>
<% if oforeign_detail.FresultCount > 0 then %>
<% for i=0 to oforeign_detail.FresultCount-1 %>
    <% if NOT (isImageVisible) then %>
	<tr bgcolor="#FFFFFF">
		<td><%= oforeign_detail.FItemList(i).fbaljucode %></td>
		<td><%= oforeign_detail.FItemList(i).falinkcode %></td>
		<td><%= oforeign_detail.FItemList(i).fregdate %></td>
		<td><%= oforeign_detail.FItemList(i).fbaljudate %></td>
		<td><%= oforeign_detail.FItemList(i).finnerboxno %></td>
		<td><%= oforeign_detail.FItemList(i).fcartoonboxno %></td>
		<td class='txt'><%= oforeign_detail.FItemList(i).fmakerid %></td>
		<td class='txt'>
			<%= BF_MakeTenBarcode(oforeign_detail.FItemList(i).fitemgubun, oforeign_detail.FItemList(i).Fitemid, oforeign_detail.FItemList(i).fitemoption) %>
		</td>
		<td class='txt'><%= oforeign_detail.FItemList(i).fextbarcode %></td>
		<td align="left"><%= oforeign_detail.FItemList(i).fitemname %></td>
		<td align="left"><%= oforeign_detail.FItemList(i).fitemoptionname %></td>
		<td>
			<%= oforeign_detail.FItemList(i).Fbaljuitemno %>
		</td>
		<td>
			<%=FormatNumber( getstateitemno(oforeign_detail.FItemList(i).Fstatecd, oforeign_detail.FItemList(i).Fforeign_statecd, oforeign_detail.FItemList(i).Fbaljuitemno, oforeign_detail.FItemList(i).Frealitemno) ,0)%>
		</td>

		<% 'if listgubun = "PACKING" or listgubun = "NONE" then %>
			<!--<td>-->
				<%'=FormatNumber( oforeign_detail.FItemList(i).Frealitemno ,0)%>
			<!--</td>-->
		<% 'else %>
			<%
			'/주문서작성중이 아닌거
			'if not(jumunwait) then
			%>
				<!--<td>-->
					<%'= oforeign_detail.FItemList(i).Fbaljuitemno %>
				<!--</td>
				<td>-->
					<%'=FormatNumber( getstateitemno(oOneOrder.FOneItem.Fstatecd, oOneOrder.FOneItem.Fforeign_statecd, oforeign_detail.FItemList(i).Fbaljuitemno, oforeign_detail.FItemList(i).Frealitemno) ,0)%>
				<!--</td>-->
			<% 'else %>
				<!--<td>-->
					<%'=FormatNumber( getstateitemno(oOneOrder.FOneItem.Fstatecd, oOneOrder.FOneItem.Fforeign_statecd, oforeign_detail.FItemList(i).Fbaljuitemno, oforeign_detail.FItemList(i).Frealitemno) ,0)%>
				<!--</td>-->
			<% 'end if %>
		<% 'end if %>

		<td><%= oforeign_detail.FItemList(i).fcurrencyunit %></td>
		<td><%= getdisp_price(oforeign_detail.FItemList(i).fsellcash,oforeign_detail.FItemList(i).fcurrencyunit) %></td>
		<td><%= getdisp_price(oforeign_detail.FItemList(i).fsuplycash,oforeign_detail.FItemList(i).fcurrencyunit) %></td>
		<td><%= oforeign_detail.FItemList(i).fdefaultsuplymargin %></td>
		<td><%= getdisp_price(oforeign_detail.FItemList(i).ftotalsuplycash,oforeign_detail.FItemList(i).fcurrencyunit) %></td>
		<td><%= FormatNumber(oforeign_detail.FItemList(i).flcprice, 2) %></td>
		<td><%= oforeign_detail.FItemList(i).fexchangeRate %></td>
		<td><%= oforeign_detail.FItemList(i).fmultipleRate %></td>
		<td></td>
		<td align="left"><%= oforeign_detail.FItemList(i).fitemname_10x10 %></td>
		<td align="left"><%= oforeign_detail.FItemList(i).foptionname_10x10 %></td>
		<td align="left"><%= oforeign_detail.FItemList(i).fitemsource_10x10 %></td>
		<td align="left"><%= oforeign_detail.FItemList(i).fsourcearea_10x10 %></td>
		<td></td>
		<td align="left"><%= oforeign_detail.FItemList(i).fitemname_en %></td>
		<td align="left"><%= oforeign_detail.FItemList(i).foptionname_en %></td>
		<td align="left"><%= oforeign_detail.FItemList(i).fitemsource_en %></td>
		<td align="left"><%= oforeign_detail.FItemList(i).fsourcearea_en %></td>
		<td><%= oforeign_detail.FItemList(i).fitemweight %></td>
		<td class='txt'><%= oforeign_detail.FItemList(i).fcatecdl %></td>
		<td class='txt'><%= oforeign_detail.FItemList(i).fcatecdm %></td>
		<td class='txt'><%= oforeign_detail.FItemList(i).fcatecdn %></td>
		<!--<td align="left"><%'= oforeign_detail.FItemList(i).fcatename1 %></td>
		<td align="left"><%'= oforeign_detail.FItemList(i).fcatename2 %></td>
		<td align="left"><%'= oforeign_detail.FItemList(i).fcatename3 %></td>
		<td align="left"><%'= oforeign_detail.FItemList(i).fcatename_eng2 %></td>
		<td align="left"><%'= oforeign_detail.FItemList(i).fcatename_eng3 %></td>-->
	</tr>
    <% else %>
    <tr height="100" bgcolor="#FFFFFF">
		<td><%= oforeign_detail.FItemList(i).fbaljucode %></td>
		<td><%= oforeign_detail.FItemList(i).falinkcode %></td>
		<td><%= oforeign_detail.FItemList(i).finnerboxno %></td>
		<td><%= oforeign_detail.FItemList(i).fcartoonboxno %></td>
		<td class='txt'>
			<%= oforeign_detail.FItemList(i).fitemgubun %>
			<%= CHKIIF(oforeign_detail.FItemList(i).Fitemid>=1000000,Format00(8,oforeign_detail.FItemList(i).Fitemid),Format00(6,oforeign_detail.FItemList(i).Fitemid)) %>
			<%= oforeign_detail.FItemList(i).fitemoption %>
		</td>
		<td align="left"><%= oforeign_detail.FItemList(i).fitemname_en %></td>
		<td align="left"><%= oforeign_detail.FItemList(i).foptionname_en %></td>
		<td align="left"><%= oforeign_detail.FItemList(i).fitemsource_en %></td>
		<td align="left"><%= oforeign_detail.FItemList(i).fsourcearea_en %></td>
		<td><%= oforeign_detail.FItemList(i).fitemweight %></td>
		<td height="100" width="100"><img src="<%= oforeign_detail.FItemList(i).Fmainimageurl %>" width="100" height="100"></td>
	</tr>
    <% end if %>
	
<% next %>
<% else %>
	<tr bgcolor="#FFFFFF">
	<% if isImageVisible then %>
		<td colspan="29" align="center">[검색결과가 없습니다.]</td>
	<% else %>
		<td colspan="35" align="center">[검색결과가 없습니다.]</td>
	<% end if %>
	</tr>
<% end if %>

</table>
</body>
</html>

<%
set oOneOrder = nothing
set oforeign_detail = nothing

session.codePage = 949
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
