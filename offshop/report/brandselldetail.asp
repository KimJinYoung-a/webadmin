<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/offshop/incSessionoffshop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/offshop/lib/offshopbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopsellcls.asp"-->
<%
dim shopid, designer, page

dim yyyy1,mm1,dd1,yyyy2,mm2,dd2
dim yyyymmdd1,yyymmdd2
dim fromDate,toDate
dim oldlist

shopid = request("shopid")
designer = request("designer")
page = request("page")
if page="" then page=1

yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")
yyyy2 = request("yyyy2")
mm2 = request("mm2")
dd2 = request("dd2")
oldlist = request("oldlist")

if (yyyy1="") then
	fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now())), Cstr(day(now()))-14)
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

dim ooffsell
set ooffsell = new COffShopSellReport
ooffsell.FRectShopid = shopid
ooffsell.FRectDesigner = designer
ooffsell.FRectNormalOnly = "on"
ooffsell.FRectStartDay = fromDate
ooffsell.FRectEndDay = toDate
ooffsell.FRectOldData = oldlist

ooffsell.GetDaylySellItemList

dim i, totalsum, totcnt

totalsum = 0
totcnt = 0

Dim CurrencyUnit, CurrencyChar, ExchangeRate
Dim FmNum, IsTaxAddCharge
Call fnGetOffCurrencyUnit(shopid,CurrencyUnit, CurrencyChar, ExchangeRate)
FmNum = CHKIIF(CurrencyUnit="WON" or CurrencyUnit="KRW",0,2)

IsTaxAddCharge = CHKIIF(CurrencyUnit<>"WON" and CurrencyUnit<>"KRW",true,false)
%>
<br>
<% if oldlist="on" then %>
<div class="a">검색기간 : <%= yyyy1 %>-<%= mm1 %>-<%= dd1 %> ~ <%= yyyy2 %>-<%= mm2 %>-<%= dd2 %> ( 3개월 이전 내역만 검색됩니다. )</a>
<% else %>
<div class="a">검색기간 : <%= yyyy1 %>-<%= mm1 %>-<%= dd1 %> ~ <%= yyyy2 %>-<%= mm2 %>-<%= dd2 %> ( 최근 3개월 내역만 검색됩니다. )</a>
<% end if %>
<table width="900" cellspacing="1" class="a" bgcolor=#3d3d3d>
<tr bgcolor="#DDDDFF">
	<td width="86">바코드</td>
	<td width="90">브랜드</td>
	<td width="100">상품명</td>
	<td width="100">옵션명</td>
	<td width="70">소비자가</td>
	<td width="70">판매가</td>
	<% if (IsTaxAddCharge) then %>
	<td width="70">Tax</td>
	<% end if %>
	<td width="60">갯수</td>
	<td width="80">합계</td>
</tr>
<% for i=0 to ooffsell.FresultCount-1 %>
<%
totcnt = totcnt + ooffsell.FItemList(i).FItemNo
totalsum = totalsum + ooffsell.FItemList(i).FSubTotal
%>
<tr bgcolor="#FFFFFF">
	<td><%= ooffsell.FItemList(i).GetBarCode %></td>
	<td><%= ooffsell.FItemList(i).FMakerID %></td>
	<td><%= ooffsell.FItemList(i).FItemName %></td>
	<td><%= ooffsell.FItemList(i).FItemOptionName %></td>
	<td align="right"><%= FormatNumber(ooffsell.FItemList(i).FSellPrice,FmNum) %></td>
	<td align="right"><%= FormatNumber(ooffsell.FItemList(i).FRealSellPrice,FmNum) %></td>
	<% if (IsTaxAddCharge) then %>
	<td align="right"><%= FormatNumber(ooffsell.FItemList(i).FAddTaxCharge,FmNum) %></td>
	<% end if %>
	<td align="center"><%= ooffsell.FItemList(i).FItemNo %></td>
	<td align="right"><%= FormatNumber(ooffsell.FItemList(i).FSubTotal,FmNum) %></td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF">
	<td colspan="7" ><b>총계</b></td>

	<td align="center"><%= totcnt %></td>
	<td align="right"><b><%= FormatNumber(totalsum,FmNum) %></b></td>
</tr>
</table>
<%
set ooffsell = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->