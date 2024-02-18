<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/offshop/incSessionoffshop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/offshop/lib/offshopbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopsellcls.asp"-->
<%
dim page,shopid
dim yyyy1,mm1,dd1,yyyy2,mm2,dd2
dim yyyymmdd1,yyymmdd2
dim fromDate,toDate
dim oldlist

shopid = session("ssBctID")
page = request("page")
if page="" then page=1


yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")
yyyy2 = request("yyyy2")
mm2 = request("mm2")
dd2 = request("dd2")
oldlist = request("oldlist")

if (shopid="doota01") then shopid="streetshop014"

if (yyyy1="") then
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


dim ooffsell
set ooffsell = new COffShopSellReport
ooffsell.FRectShopID = shopid
ooffsell.FRectNormalOnly = "on"
ooffsell.FRectStartDay = fromDate
ooffsell.FRectEndDay = toDate
ooffsell.FRectOnlyShop = "on"
ooffsell.FRectOldData = oldlist

ooffsell.GetBrandSellSumList

dim i, sum1,sum2

sum1 =0
sum2 =0


Dim CurrencyUnit, CurrencyChar, ExchangeRate
Dim FmNum
Call fnGetOffCurrencyUnit(shopid,CurrencyUnit, CurrencyChar, ExchangeRate)
FmNum = CHKIIF(CurrencyUnit="WON",0,2)

%>
<table width="800" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr>
		<td class="a" >
			<input type="checkbox" name="oldlist" <% if oldlist="on" then response.write "checked" %> >3개월이전내역
			&nbsp;
			매출일 : <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		</td>
		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>

<table width="800" cellspacing="1" class="a" bgcolor=#3d3d3d>
<tr bgcolor="#DDDDFF">
	<td width="100">브랜드</td>
	<td width="100">아이템건수</td>
	<td width="100">매출액</td>
	<td width="60">아이템목록</td>
<!--
	<td width="60">예상재고</td>
-->
</tr>
<% for i=0 to ooffsell.FresultCount-1 %>
<tr bgcolor="#FFFFFF">
	<% sum1 = sum1 + ooffsell.FItemList(i).FSum+ ooffsell.FItemList(i).FaddTaxChargeSum %>
	<td><%= ooffsell.FItemList(i).FMakerid %></td>
	<td align="center"><%= ooffsell.FItemList(i).FCount %></td>
	<td align="right"><%= FormatNumber(ooffsell.FItemList(i).FSum + ooffsell.FItemList(i).FaddTaxChargeSum,FmNum) %></td>
	<td align="center"><a href="brandselldetail.asp?menupos=<%= menupos %>&shopid=<%= shopid %>&designer=<%= ooffsell.FItemList(i).FMakerid %>&yyyy1=<%= yyyy1 %>&yyyy2=<%= yyyy2 %>&mm1=<%= mm1 %>&mm2=<%= mm2 %>&dd1=<%= dd1 %>&dd2=<%= dd2 %>&oldlist=<%= oldlist %>" target="_blank">보기</a></td>
<!--
	<td align="center"><a href="jaegolist.asp?menupos=204&shopid=<%= ooffsell.FItemList(i).FShopid %>&designer=<%= ooffsell.FItemList(i).FMakerid %>">보기</a></td>
-->
</tr>
<% next %>
<tr bgcolor="#FFFFFF">
	<td>합계</td>
	<td colspan="2" align="right"><%= FormatNumber(sum1,FmNum) %></td>
	<td></td>
</tr>
</table>
<%
set ooffsell = Nothing
%>
<!-- #include virtual="/offshop/lib/offshopbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->