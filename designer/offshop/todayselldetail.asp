<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프라인
' History : 2010.03.26 한용민 추가
'####################################################
%>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopsellcls.asp"-->
<%
dim chargeid, shopid, terms , datefg , i,totalsum
	chargeid = session("ssBctID")
	shopid = request("shopid")
	terms = request("terms")
	datefg = request("datefg")
	if datefg = "" then datefg = "maechul"	

dim ooffsell
set ooffsell = new COffShopSellReport
	ooffsell.FRectShopid = shopid
	ooffsell.FRectNormalOnly = "on"
	ooffsell.frectdatefg = datefg	
    ooffsell.FRectTerms = ""
    ooffsell.FRectStartDay = terms
    ooffsell.FRectEndDay = CStr(dateAdd("d",1,terms))
	ooffsell.FRectDesigner = chargeid		
	ooffsell.GetDaylySellItemList

totalsum =0
%>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr>
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">기간</td>
	<td bgcolor="#FFFFFF"><%= terms %></td>
</tr>
<tr>
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">샾 구분</td>
	<td bgcolor="#FFFFFF"><%= shopid %></td>
</tr>
</table>

<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="86">바코드</td>
	<td width="90">범용바코드</td>
	<td width="90">브랜드</td>
	<td>상품명</td>
	<td>옵션명</td>
	<td width="70">소비자가</td>
	<td width="70">판매가</td>
	<td width="60">수량</td>
	<td width="80">판매가합계</td>
</tr>
<% for i=0 to ooffsell.FresultCount-1 %>
<% totalsum = totalsum + ooffsell.FItemList(i).FSubTotal %>
<tr align="center" bgcolor="#FFFFFF">
	<td><%= ooffsell.FItemList(i).GetBarCode %></td>
	<td><%= ooffsell.FItemList(i).fextbarcode %></td>
	<td><%= ooffsell.FItemList(i).FMakerID %></td>
	<td align="left"><%= ooffsell.FItemList(i).FItemName %></td>
	<td><%= ooffsell.FItemList(i).FItemOptionName %></td>
	<td align="right"><%= FormatNumber(ooffsell.FItemList(i).FSellPrice,0) %></td>
	<td align="right"><%= FormatNumber(ooffsell.FItemList(i).FRealSellPrice,0) %></td>
	<td><%= ooffsell.FItemList(i).FItemNo %></td>
	<td align="right"><%= FormatNumber(ooffsell.FItemList(i).FSubTotal,0) %></td>
</tr>
<% next %>
<tr align="center" bgcolor="#FFFFFF">
	<td><b>합계</b></td>
	<td colspan="9" align="right"><b><%= FormatNumber(totalsum,0) %></b></td>
</tr>
</table>

<%
set ooffsell = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->