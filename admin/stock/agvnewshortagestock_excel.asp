<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/newshortagestockcls.asp"-->
<!-- #include virtual="/lib/BarcodeFunction.asp"-->
<%
const C_STOCK_DAY=7

''아래 두 페이지는 검색조건을 동일하게 맞춰야 한다.
''/admin/stock/newshortagestock.asp
''/admin/newstorage/popjumunitemNew.asp

dim page, mode, makerid, shopid,itemid, research
dim onlynotupchebeasong, onlyusingitem, onlyusingitemoption, onlynotdanjong, soldoutover7days, onlysell, onlynottempdanjong
dim onlynotmddanjong, includepreorder, skiplimitsoldout
dim onoffgubun, idx, shortagetype, onlystockminus
dim changemakerid
dim purchasetype
dim itemgubun, itemname
dim chkMinusStockGubun, minusStockGubun
dim mwdiv, excmkr, onlyOn, pagesize, onlyrealup, ordBy

shopid = requestCheckVar(("shopid"),32)
page = requestCheckVar(request("page"),32)
mode = requestCheckVar(request("mode"),32)
itemid = requestCheckVar(request("itemid"),32)
research = requestCheckVar(request("research"),32)
onlynotupchebeasong = requestCheckVar(request("onlynotupchebeasong"),32)
onlyusingitem = requestCheckVar(request("onlyusingitem"),32)
onlyusingitemoption = requestCheckVar(request("onlyusingitemoption"),32)
onlynotdanjong = requestCheckVar(request("onlynotdanjong"),32)
soldoutover7days = requestCheckVar(request("soldoutover7days"),32)
onoffgubun = requestCheckVar(request("onoffgubun"),32)
idx = requestCheckVar(request("idx"),32)
shortagetype = requestCheckVar(request("shortagetype"),32)
onlysell = requestCheckVar(request("onlysell"),32)
onlynottempdanjong = requestCheckVar(request("onlynottempdanjong"),32)
onlynotmddanjong = requestCheckVar(request("onlynotmddanjong"),32)
includepreorder = requestCheckVar(request("includepreorder"),32)
skiplimitsoldout = requestCheckVar(request("skiplimitsoldout"),32)
onlystockminus = requestCheckVar(request("onlystockminus"),32)
purchasetype = requestCheckVar(request("purchasetype"),32)
itemgubun = requestCheckVar(request("itemgubun"),32)
itemname = requestCheckVar(request("itemname"),128)
chkMinusStockGubun = requestCheckVar(request("chkMinusStockGubun"),32)
minusStockGubun = requestCheckVar(request("minusStockGubun"),32)
mwdiv = requestCheckVar(request("mwdiv"),32)
excmkr = requestCheckVar(request("excmkr"),32)
onlyOn = requestCheckVar(request("onlyOn"),32)
pagesize = requestCheckVar(request("pagesize"),32)
onlyrealup = requestCheckVar(request("onlyrealup"),32)
ordBy = requestCheckVar(request("ordBy"),32)

changemakerid = "Y"
if (changemakerid = "") then
	changemakerid = requestCheckVar(request("changemakerid"),32)
end if

makerid = request("makerid")
if (makerid = "") then
	makerid = requestCheckVar(request("suplyer"),32)
end if


if (research<>"on") then
	excmkr = "Y"
    chkMinusStockGubun="Y"
    minusStockGubun = "agv"
end if

if (research<>"on") and (onlynotupchebeasong = "") then
	onlynotupchebeasong = "on"
end if

if (research<>"on") and (onlyusingitem = "") then
	onlyusingitem = "on"
end if

if (research<>"on") and (onlyusingitemoption="") then
	onlyusingitemoption = "on"
end if

if (research<>"on") and (onlynotdanjong = "") then
	onlynotdanjong = "on"
end if

if (research<>"on") and (onoffgubun="") then
	onoffgubun = "online"
end if

if (research<>"on") and (itemgubun="") then
	itemgubun = "10"
end if

if (research<>"on") and (shortagetype="") then
	shortagetype = "7day"
end if

if (research<>"on") and (includepreorder="") then
	includepreorder = "on"
end if

if (pagesize="") then
	pagesize = 100
end if

if (research<>"on") and (onlyrealup="") then
	onlyrealup = "on"
end if



if page="" then page=1
if mode="" then mode="bybrand"
'상품코드 유효성 검사(2008.07.31;허진원)
if itemid<>"" then
	if Not(isNumeric(itemid)) then
		Response.Write "<script language=javascript>alert('[" & itemid & "]은(는) 유효한 상품코드가 아닙니다.');history.back();</script>"
		dbget.close()	:	response.End
	end if
end if

dim oshortagestock
set oshortagestock  = new CShortageStock
oshortagestock.FPageSize = 10000
oshortagestock.FCurrPage = 1

oshortagestock.FRectOnlySell			= onlysell
oshortagestock.FRectOnlyUsingItem		= onlyusingitem
oshortagestock.FRectOnlyUsingItemOption	= onlyusingitemoption
oshortagestock.FRectOnlyNotUpcheBeasong	= onlynotupchebeasong

oshortagestock.FRectOnlyNotDanjong		= onlynotdanjong
oshortagestock.FRectOnlyNotTempDanjong	= onlynottempdanjong
oshortagestock.FRectOnlyNotMDDanjong	= onlynotmddanjong
oshortagestock.FRectSkipLimitSoldOut	= skiplimitsoldout

oshortagestock.FRectPurchaseType		= purchasetype

oshortagestock.FRectMakerid				= makerid
oshortagestock.FRectItemId				= itemid
'oshortagestock.FRectItemOption			= makerid

oshortagestock.FRectItemGubun			= itemgubun

if (chkMinusStockGubun = "Y") then
	oshortagestock.FRectMinusStockGubun			= minusStockGubun
end if

if (itemname <> "") then
	if (makerid <> "") then
		oshortagestock.FRectItemName			= itemname
	else
		response.write "<script>alert('먼저 브랜드를 선택하세요.');</script>"
	end if
end if

oshortagestock.FRectMWDiv				= mwdiv
oshortagestock.FRectExcMkr				= excmkr
oshortagestock.FRectOnlyOn				= onlyOn
oshortagestock.FRectOnlyRealUp			= onlyrealup
oshortagestock.FRectOrderBy				= ordBy
oshortagestock.FRectAGVCheck			= "Y"
if (itemgubun = "10") then
	oshortagestock.GetShortageItemListOnline
else
	oshortagestock.GetShortageItemListOffline
end if



dim i, shopsuplycash, buycash
dim IsAvailDelete



'==============================================================================
dim yyyy1,yyyy2,mm1,mm2,dd1,dd2, nowdate, iStartDate, iEndDate

'재입고예정일
'오늘기준 +- 일주일은 검정색 표시 / 그 이외 회색표시
if (yyyy1="") then
    nowdate = Left(CStr(DateAdd("d",now(),-7)),10)
	yyyy1 = Left(nowdate,4)
	mm1   = Mid(nowdate,6,2)
	dd1   = Mid(nowdate,9,2)

    nowdate = Left(CStr(DateAdd("d",now(),+7)),10)
	yyyy2 = Left(nowdate,4)
	mm2   = Mid(nowdate,6,2)
	dd2   = Mid(nowdate,9,2)
end if

iStartDate  = Left(CStr(DateSerial(yyyy1,mm1,dd1)),10)
iEndDate    = Left(CStr(DateSerial(yyyy2,mm2,dd2)),10)

Response.Buffer = true    '버퍼사용여부
Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=TEN" & Left(CStr(now()),10) & "_AGVstock.xls"
Response.CacheControl = "public"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html;charset=euc-kr" />
<style type='text/css'>
	.txt {mso-number-format:'\@'}
</style>
</head>
<body>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#DDDDDD" border=1>
	<tr height="18" bgcolor="FFFFFF">
		<td colspan="19">
			검색결과 : <b><%= oshortagestock.FTotalCount %></b>
			&nbsp;
			최대 10000건 까지 검색 됩니다.
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF">
		<td>브랜드ID</td>
        <td>렉코드</td>
        <td>보조랙</td>
		<td>구분</td>
		<td>상품코드</td>
		<td>옵션코드</td>
		<td>상품명[옵션명]</td>
		<td>실사유효재고(V)</td>
		<td>재고파악재고</td>
		<td>예상재고</td>
        <td>AGV재고</td>
		<td>ON결제완료</td>
        <td>ON발주중</td>
		<td>OFF발주중</td>

		<td>총(<%= C_STOCK_DAY %>일)필요수량</td>
		<td>출고이전필요수량</td>
		<td>AGV부족수량</td>
		<td>ON(7일)판매</td>
		<td>OFF(7일)판매</td>
	</tr>
<% for i=0 to oshortagestock.FResultCount -1 %>
<%
    IsAvailDelete = (oshortagestock.FItemList(i).Ftotipgono=0) and (oshortagestock.FItemList(i).FtotSellNo=0) and (oshortagestock.FItemList(i).Fshortageno=0) and (oshortagestock.FItemList(i).Frealstock=0) and (oshortagestock.FItemList(i).Fpreorderno=0)
%>

	<% if oshortagestock.FItemList(i).IsInvalidOption then %>
	<tr align="center" bgcolor="#CCCCCC">
	<% else %>
	<tr align="center" bgcolor="#FFFFFF">
	<% end if %>
		<td><%= oshortagestock.FItemList(i).FMakerID %></td>
        <td><%= oshortagestock.FItemList(i).FrackcodeByOption %></td>
        <td><%= oshortagestock.FItemList(i).FsubRackcodeByOption %></td>
		<td><%= oshortagestock.FItemList(i).Fitemgubun %></td>
		<td><%= oshortagestock.FItemList(i).FItemID %></td>
		<td class="txt"><%= oshortagestock.FItemList(i).Fitemoption %></td>
		<td>
			<%= oshortagestock.FItemList(i).FItemName %>
			<% if oshortagestock.FItemList(i).FItemOption <> "0000" then %>
				<% if oshortagestock.FItemList(i).Foptionusing="Y" then %>
					[<%= oshortagestock.FItemList(i).FItemOptionName %>]
				<% else %>
					[<%= oshortagestock.FItemList(i).FItemOptionName %>]
				<% end if %>
			<% end if %>
		</td>
		<td><b><%= oshortagestock.FItemList(i).Frealstock %></b></td>
		<td><b><%= oshortagestock.FItemList(i).GetCheckStockNo %></b></td>
		<td><b><%= oshortagestock.FItemList(i).GetMaystock %></b></td>
        <td><b><%= oshortagestock.FItemList(i).FAGVStock %></b></td>

		<td><%= oshortagestock.FItemList(i).FIpkumdiv4 %></td>
        <td><%= oshortagestock.FItemList(i).FIpkumdiv5 %></td>
		<td><%= oshortagestock.FItemList(i).Foffconfirmno %></td>

		<td><b><%= oshortagestock.FItemList(i).Frequireno %></b></td>
		<td>
		    <!-- 출고이전 필요수량 -->
		    <%= (oshortagestock.FItemList(i).Fipkumdiv5 + oshortagestock.FItemList(i).Foffconfirmno+oshortagestock.FItemList(i).Fipkumdiv4 + oshortagestock.FItemList(i).Fipkumdiv2 + oshortagestock.FItemList(i).Foffjupno)*-1 %>
		</td>
		<td><b><%= oshortagestock.FItemList(i).GetAGVShortageNo %></b></td>
		<td><%= oshortagestock.FItemList(i).Fsell7days %></td>
		<td><%= oshortagestock.FItemList(i).Foffchulgo7days %></td>
	</tr>
<%
if i mod 1000 = 0 then
    Response.Flush		' 버퍼리플래쉬
end if
next
%>
</table>
</html>
<%
set oshortagestock = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
