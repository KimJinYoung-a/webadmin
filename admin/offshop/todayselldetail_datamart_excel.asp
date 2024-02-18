<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프라인 매출상품 상세 공용페이지 NO 페이징 버전 엑셀출력
' History : 2012.08.31 한용민 생성
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/BarcodeFunction.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopsellcls_datamart.asp"-->
<%
dim shopid , datefg , i ,makerid ,yyyy1,mm1,dd1,yyyy2,mm2,dd2, toDate,fromDate
dim totitemno ,totalsum ,totsuplysum ,totsellsum ,oldlist ,offgubun ,vOffCateCode ,offmduserid
dim vOffMDUserID ,vPurchaseType ,ordertype ,itemid ,itemname ,extbarcode ,reload, buyergubun
dim inc3pl, commCd, arrlist, menupos
	yyyy1 = requestCheckVar(request("yyyy1"),4)
	mm1 = requestCheckVar(request("mm1"),2)
	dd1 = requestCheckVar(request("dd1"),2)
	yyyy2 = requestCheckVar(request("yyyy2"),4)
	mm2 = requestCheckVar(request("mm2"),2)
	dd2 = requestCheckVar(request("dd2"),2)
	menupos = requestCheckVar(request("menupos"),10)
	shopid = requestCheckVar(request("shopid"),32)
	datefg = requestCheckVar(request("datefg"),32)
	makerid = requestCheckVar(request("makerid"),32)
	oldlist = requestCheckVar(request("oldlist"),10)
	offgubun = requestCheckVar(request("offgubun"),32)
	vOffCateCode = requestCheckVar(request("offcatecode"),32)
	vOffMDUserID = requestCheckVar(request("offmduserid"),32)
	vPurchaseType = requestCheckVar(request("purchasetype"),2)
	ordertype = requestCheckVar(request("ordertype"),32)
	itemid = requestCheckVar(request("itemid"),10)
	itemname = requestCheckVar(request("itemname"),124)
	extbarcode = requestCheckVar(request("extbarcode"),32)
	reload = requestCheckVar(request("reload"),2)
	buyergubun = requestCheckVar(request("buyergubun"),10)
    inc3pl = requestCheckVar(request("inc3pl"),32)
    commCd = requestCheckVar(request("commCd"),32)

if datefg = "" then datefg = "maechul"
if shopid<>"" then offgubun=""
if ordertype="" then ordertype="totalprice"
if reload <> "on" and offgubun = "" then offgubun = "95"

if (yyyy1="") then
	'fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now())), Cstr(day(now()))-7)
	fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now())), Cstr(day(now())))
else
	fromDate = DateSerial(yyyy1, mm1, dd1)
end if

if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

toDate = DateSerial(yyyy2, mm2, dd2+1)

yyyy1 = left(fromDate,4)
mm1 = Mid(fromDate,6,2)
dd1 = Mid(fromDate,9,2)

'C_IS_Maker_Upche = TRUE
'C_IS_SHOP = TRUE

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

''두타쪽 매출조회 권한
Dim isFixShopView
IF (session("ssBctID")="doota01") then
    shopid="streetshop014"
    C_IS_SHOP = TRUE
    isFixShopView = TRUE
ENd If

dim ooffsell
set ooffsell = new COffShopSellReport
	ooffsell.FPageSize = 10000
	ooffsell.FCurrPage = 1
	ooffsell.FRectOldData = oldlist
	ooffsell.FRectShopid = shopid
	ooffsell.FRectNormalOnly = "on"
	ooffsell.frectdatefg = datefg
    ooffsell.FRectTerms = ""
    ooffsell.FRectStartDay = fromDate
    ooffsell.FRectEndDay = toDate
    ooffsell.FRectDesigner = makerid
	ooffsell.FRectOffgubun = offgubun
	ooffsell.frectoffcatecode = vOffCateCode
	ooffsell.frectoffmduserid = vOffMDUserID
	ooffsell.FRectBrandPurchaseType = vPurchaseType
	ooffsell.FRectOrdertype = ordertype
	ooffsell.FRectitemid = itemid
	ooffsell.FRectitemname = itemname
	ooffsell.FRectextbarcode = extbarcode
	ooffsell.FRectbuyergubun = buyergubun
	ooffsell.FRectInc3pl = inc3pl
	ooffsell.FRectCommCD = commCd
    arrlist = ooffsell.GetDaylySellItemList_getrows

totitemno = 0
totalsum =0
totsuplysum = 0
totsellsum = 0

Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=TEN" & Left(CStr(now()),10) & ".xls"
Response.CacheControl = "public"
%>

<style type='text/css'>
	.txt {mso-number-format:'\@'}
</style>

<table width="100%" border="0" align="center" class="a" cellpadding="1" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF" height="25">
	<td colspan="20">
		검색결과 : <b><%=ooffsell.FresultCount%></b> ※최대 3,000건까지 노출 됩니다.(엑셀다운은 10,000건 제한)
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>바코드</td>
	<td>공용바코드</td>
	<td>브랜드</td>
	<td>상품명(옵션명)</td>
	<% if (NOT C_InspectorUser) then %>
	<td>판매액</td>
    <% end if %>
	<td>매출액</td>

	<% if not(C_IS_SHOP) then %>
		<td>매입액</td>
	<% end if %>

	<td>판매수량</td>
	<td>비고</td>
</tr>
<%
if isarray(arrlist) then

for i=0 to ubound(arrlist,2)

totitemno = totitemno + arrlist(3,i)
totalsum = totalsum + arrlist(0,i)
totsellsum = totsellsum + arrlist(2,i)
totsuplysum = totsuplysum + arrlist(1,i)
%>
<tr bgcolor="#FFFFFF" align="center">
	<td class='txt'>
		<%= arrlist(10,i) %><%= BF_GetFormattedItemId(arrlist(11,i)) %><%= arrlist(12,i) %>
	</td>
	<td class='txt'>
		<%= arrlist(15,i) %>
	</td>
	<td><%= arrlist(13,i) %></td>
	<td align="left">
		<%= arrlist(8,i) %>
		<% if arrlist(9,i) <> "" then %>
			(<%=arrlist(9,i)%>)
		<% end if %>
	</td>
	<% if (NOT C_InspectorUser) then %>
	<td align="right"><%= FormatNumber(arrlist(2,i),0) %></td>
    <% end if %>
	<td align="right"><%= FormatNumber(arrlist(0,i),0) %></td>
	
	<% if not(C_IS_SHOP) then %>
		<td align="right"><%= FormatNumber(arrlist(1,i),0) %></td>
	<% end if %>
	
	<td><%= arrlist(3,i) %></td>
	<td align="center"><%= arrlist(16,i) %></td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF" align="center">
	<td colspan=4><b>총계</b></td>
	<% if (NOT C_InspectorUser) then %>
	<td align="right"><%= FormatNumber(totsellsum,0) %></td>
    <% end if %>
	<td align="right"><%= FormatNumber(totalsum,0) %></td>

	<% if not(C_IS_SHOP) then %>
		<td align="right"><%= FormatNumber(totsuplysum,0) %></td>
	<% end if %>

	<td><%= FormatNumber(totitemno,0) %></td>
	<td></td>
</tr>
<% else %>
<tr  align="center" bgcolor="#FFFFFF">
	<td colspan="20">등록된 내용이 없습니다.</td>
</tr>
<% end if %>
</table>

<%
set ooffsell = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->