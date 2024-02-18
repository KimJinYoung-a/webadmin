<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/offshop/incSessionoffshop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/offshop/lib/offshopbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopsellcls.asp"-->
<%
dim chargeid, shopid, terms , datefg 
	chargeid = request("chargeid")
	shopid = request("shopid")
	terms = request("terms")
	datefg = request("datefg")
	if datefg = "" then datefg = "maechul"

shopid = session("ssBctID") ''강제지정
if (shopid="doota01") then shopid="streetshop014"

dim ooffsell
set ooffsell = new COffShopSellReport
	ooffsell.FRectShopid = shopid
	ooffsell.FRectNormalOnly = "on"
	ooffsell.frectdatefg = datefg
    ooffsell.FRectTerms = ""
    ooffsell.FRectStartDay = terms
    ooffsell.FRectEndDay = CStr(dateAdd("d",1,terms))
    ooffsell.GetDaylySellItemList

dim i,totalsum
totalsum =0

Dim CurrencyUnit, CurrencyChar, ExchangeRate
Dim FmNum, IsTaxAddCharge
Call fnGetOffCurrencyUnit(shopid,CurrencyUnit, CurrencyChar, ExchangeRate)
FmNum = CHKIIF(CurrencyUnit="WON" or CurrencyUnit="KRW",0,2)

IsTaxAddCharge = CHKIIF(CurrencyUnit<>"WON" and CurrencyUnit<>"KRW",true,false)
%>
<table width="800" cellspacing="1" cellpadding="0" class="a" bgcolor=#3d3d3d>
<tr>
	<td width="100" bgcolor="#DDDDFF">기간</td>
	<td bgcolor="#FFFFFF"><%= terms %></td>
</tr>
<tr>
	<td width="100" bgcolor="#DDDDFF">샾 구분</td>
	<td bgcolor="#FFFFFF"><%= shopid %></td>
</tr>
</table>
<br>
<table width="800" cellspacing="1" class="a" bgcolor=#3d3d3d>
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
<% totalsum = totalsum + ooffsell.FItemList(i).FSubTotal %>
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
	<td><b>총계</b></td>
	<td colspan="8" align="right"><b><%= FormatNumber(totalsum,FmNum) %></b></td>
</tr>
</table>

<%
set ooffsell = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->