<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프라인 주문서관리 통계상세
' History : 2010.06.11 한용민 생성
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/ordersheetcls.asp"-->
<%
dim idx, i,selltotal, suplytotal, buytotal ,selltotalfix, suplytotalfix, buytotalfix ,totalfixno
dim ojumundetail
	idx = request("idx")

if idx="" then idx=0

if right(trim(idx),1) = "," then idx = left(idx,len(idx)-1)

%>

<%
set ojumundetail= new COrderSheet
	ojumundetail.FRectIdx = idx
	ojumundetail.frectitemgubun = "10"
	ojumundetail.GetOrderSheetDetail_gubun

selltotal =0
suplytotal =0
buytotal =0
selltotalfix =0
suplytotalfix =0
buytotalfix =0
totalfixno = 0
%>
※온라인상품(10)
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td colspan="20" align="right">
		총건수: <%= ojumundetail.FResultCount %>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td>상품</td>
    <td>이미지</td>
	<td>바코드</td>
	<td>브랜드</td>
	<td>상품명<br>[옵션명]</td>
	<td>날짜</td>
	<td>주문시<br>판매가</td>
	<td>출고<br>공급가</td>
	<td>매입가</td>
	<td>주문수</td>
	<td>확정수</td>
	<td>주문<br>공급가</td>
	<td>주문<br>판매가</td>
	<td>확정<br>공급가</td>
	<td>확정<br>판매가</td>	
	<td>센터<br>매입<br>구분</td>
</tr>
<% 
if (ojumundetail.FResultCount>0) then
	
for i=0 to ojumundetail.FResultCount-1

selltotal  = selltotal + ojumundetail.FItemList(i).FSellcash * ojumundetail.FItemList(i).Fbaljuitemno
suplytotal = suplytotal + ojumundetail.FItemList(i).FSuplycash * ojumundetail.FItemList(i).Fbaljuitemno
buytotal   = buytotal + ojumundetail.FItemList(i).Fbuycash * ojumundetail.FItemList(i).Fbaljuitemno
selltotalfix  = selltotalfix + ojumundetail.FItemList(i).FSellcash * ojumundetail.FItemList(i).Frealitemno
suplytotalfix = suplytotalfix + ojumundetail.FItemList(i).FSuplycash * ojumundetail.FItemList(i).Frealitemno
buytotalfix   = buytotalfix + ojumundetail.FItemList(i).Fbuycash * ojumundetail.FItemList(i).Frealitemno
totalfixno = totalfixno + ojumundetail.FItemList(i).Frealitemno
%>
<tr align="center" bgcolor="#FFFFFF">
	<td>
		<%= ojumundetail.FItemList(i).FItemGubun %>-<%= ojumundetail.FItemList(i).FItemid %>-<%= ojumundetail.FItemList(i).FItemoption %>
	</td>
	<td><img src="<%= ojumundetail.FItemList(i).GetImageSmall %>" border="0" width="50" height="50" onError="this.src='http://image.10x10.co.kr/images/no_image.gif'"></td>
	<td><%= ojumundetail.FItemList(i).FItemGubun %><%= format00(6,ojumundetail.FItemList(i).FItemID) %><%= ojumundetail.FItemList(i).Fitemoption %></td>
	<td><%= ojumundetail.FItemList(i).Fmakerid %></td>
	<td>
		<%= ojumundetail.FItemList(i).Fitemname %>
		<% if ojumundetail.FItemList(i).Fitemoption <> "0000" then %>
			<br><%= ojumundetail.FItemList(i).Fitemoptionname %>
		<% end if %>
	</td>
	<td><%= ojumundetail.FItemList(i).fregdate %></td>
	<td>
		<% if   (ojumundetail.FItemList(i).FItemDefaultMwDiv<>"W") and (ojumundetail.FItemList(i).Fbuycash>ojumundetail.FItemList(i).Fsuplycash) then %>
		<b><font color=red><%= FormatNumber(ojumundetail.FItemList(i).Fsellcash,0) %></font></b>
		<% else %>
		<%= FormatNumber(ojumundetail.FItemList(i).Fsellcash,0) %>
		<% end if %>

		<% if (ojumundetail.FItemList(i).IsOnLineItem) and (ojumundetail.FItemList(i).Fsellcash<>ojumundetail.FItemList(i).Fonlinesellcash) then %>
		<br>
		<div ><font color=red>온:<%= FormatNumber(ojumundetail.FItemList(i).Fonlinesellcash,0) %></font></div>
		<% end if %>
	</td>
	<td>
		<%= ojumundetail.FItemList(i).Fsuplycash %>
		<% if (ojumundetail.FItemList(i).IsOnLineItem) and (ojumundetail.FItemList(i).GetOrgShopSuplycashbyMargine<>ojumundetail.FItemList(i).Fsuplycash) then %>
			<div ><font color=red><%= ojumundetail.FItemList(i).GetOrgShopSuplycashbyMargine %></font></div>
		<% end if %>
	</td>
	<td>
		<%= ojumundetail.FItemList(i).Fbuycash %>
		<% if (ojumundetail.FItemList(i).Fbuycash<>ojumundetail.FItemList(i).Fonlinebuycash) and ((ojumundetail.FItemList(i).FItemDefaultMwDiv="W") and (ojumundetail.FItemList(i).FoffChargeDiv="4")) then %>
		<div ><font color="red">온:<%= ojumundetail.FItemList(i).Fonlinebuycash %></font></div>
		<% end if %>
	</td>
	<td>
		<%= ojumundetail.FItemList(i).Fbaljuitemno %>
	</td>
	<td>
		<%= ojumundetail.FItemList(i).Frealitemno %>
	</td>
	<td align="center"><%= FormatNumber(ojumundetail.FItemList(i).FSellcash * ojumundetail.FItemList(i).Fbaljuitemno,0) %></td>
	<td align="center"><%= FormatNumber(ojumundetail.FItemList(i).FSuplycash * ojumundetail.FItemList(i).Fbaljuitemno,0) %></td>
	<td align="center"><%= FormatNumber(ojumundetail.FItemList(i).FSellcash * ojumundetail.FItemList(i).Frealitemno,0) %></td>
	<td align="center"><%= FormatNumber(ojumundetail.FItemList(i).FSuplycash * ojumundetail.FItemList(i).Frealitemno,0) %></td>
	<td align="center"><%= ojumundetail.FItemList(i).Fcentermwdiv %></td>	
</tr>
<% next %>
<tr bgcolor="#FFFFFF" align="center">
	<td colspan=10>총계</td>
	<td><%= totalfixno %></td>
	<td colspan=2></td>
	<td 
		<%= formatNumber(selltotal,0) %><br>
		<%= formatNumber(selltotalfix,0) %>
	</td>
	<td>
		<%= formatNumber(suplytotal,0) %><br>
		<%= formatNumber(suplytotalfix,0) %>
	</td>
	<td></td>
</tr>
<% else %>
<tr bgcolor="#FFFFFF" align="center">
	<td colspan=20>검색 결과가 없습니다</td>
</tr>
<% end if %>
</table>
<%
set ojumundetail = Nothing

set ojumundetail= new COrderSheet
	ojumundetail.FRectIdx = idx
	ojumundetail.frectitemgubun = "90"
	ojumundetail.GetOrderSheetDetail_gubun

selltotal =0
suplytotal =0
buytotal =0
selltotalfix =0
suplytotalfix =0
buytotalfix =0
totalfixno = 0
%>
<br>
※오프라인상품(90)
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td colspan="20" align="right">
		총건수: <%= ojumundetail.FResultCount %>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td>상품</td>
    <td>이미지</td>
	<td>바코드</td>
	<td>브랜드</td>
	<td>상품명<br>[옵션명]</td>
	<td>날짜</td>
	<td>주문시<br>판매가</td>
	<td>출고<br>공급가</td>
	<td>매입가</td>
	<td>주문수</td>
	<td>확정수</td>
	<td>주문<br>공급가</td>
	<td>주문<br>판매가</td>
	<td>확정<br>공급가</td>
	<td>확정<br>판매가</td>	
	<td>센터<br>매입<br>구분</td>
</tr>
<%
if ojumundetail.FResultCount > 0 then
	 
for i=0 to ojumundetail.FResultCount-1

selltotal  = selltotal + ojumundetail.FItemList(i).FSellcash * ojumundetail.FItemList(i).Fbaljuitemno
suplytotal = suplytotal + ojumundetail.FItemList(i).FSuplycash * ojumundetail.FItemList(i).Fbaljuitemno
buytotal   = buytotal + ojumundetail.FItemList(i).Fbuycash * ojumundetail.FItemList(i).Fbaljuitemno
selltotalfix  = selltotalfix + ojumundetail.FItemList(i).FSellcash * ojumundetail.FItemList(i).Frealitemno
suplytotalfix = suplytotalfix + ojumundetail.FItemList(i).FSuplycash * ojumundetail.FItemList(i).Frealitemno
buytotalfix   = buytotalfix + ojumundetail.FItemList(i).Fbuycash * ojumundetail.FItemList(i).Frealitemno
totalfixno = totalfixno + ojumundetail.FItemList(i).Frealitemno
%>
<tr align="center" bgcolor="#FFFFFF">
	<td>
		<%= ojumundetail.FItemList(i).FItemGubun %>-<%= ojumundetail.FItemList(i).FItemid %>-<%= ojumundetail.FItemList(i).FItemoption %>
	</td>
	<td><img src="<%= ojumundetail.FItemList(i).GetImageSmall %>" border="0" width="50" height="50" onError="this.src='http://image.10x10.co.kr/images/no_image.gif'"></td>
	<td><%= ojumundetail.FItemList(i).FItemGubun %><%= format00(6,ojumundetail.FItemList(i).FItemID) %><%= ojumundetail.FItemList(i).Fitemoption %></td>
	<td><%= ojumundetail.FItemList(i).Fmakerid %></td>
	<td>
		<%= ojumundetail.FItemList(i).Fitemname %>
		<% if ojumundetail.FItemList(i).Fitemoption <> "0000" then %>
			<br><%= ojumundetail.FItemList(i).Fitemoptionname %>
		<% end if %>
	</td>
	<td><%= ojumundetail.FItemList(i).fregdate %></td>
	<td>
		<% if   (ojumundetail.FItemList(i).FItemDefaultMwDiv<>"W") and (ojumundetail.FItemList(i).Fbuycash>ojumundetail.FItemList(i).Fsuplycash) then %>
		<b><font color=red><%= FormatNumber(ojumundetail.FItemList(i).Fsellcash,0) %></font></b>
		<% else %>
		<%= FormatNumber(ojumundetail.FItemList(i).Fsellcash,0) %>
		<% end if %>

		<% if (ojumundetail.FItemList(i).IsOnLineItem) and (ojumundetail.FItemList(i).Fsellcash<>ojumundetail.FItemList(i).Fonlinesellcash) then %>
		<br>
		<div ><font color=red>온:<%= FormatNumber(ojumundetail.FItemList(i).Fonlinesellcash,0) %></font></div>
		<% end if %>
	</td>
	<td>
		<%= ojumundetail.FItemList(i).Fsuplycash %>
		<% if (ojumundetail.FItemList(i).IsOnLineItem) and (ojumundetail.FItemList(i).GetOrgShopSuplycashbyMargine<>ojumundetail.FItemList(i).Fsuplycash) then %>
			<div ><font color=red><%= ojumundetail.FItemList(i).GetOrgShopSuplycashbyMargine %></font></div>
		<% end if %>
	</td>
	<td>
		<%= ojumundetail.FItemList(i).Fbuycash %>
		<% if (ojumundetail.FItemList(i).Fbuycash<>ojumundetail.FItemList(i).Fonlinebuycash) and ((ojumundetail.FItemList(i).FItemDefaultMwDiv="W") and (ojumundetail.FItemList(i).FoffChargeDiv="4")) then %>
		<div ><font color="red">온:<%= ojumundetail.FItemList(i).Fonlinebuycash %></font></div>
		<% end if %>
	</td>
	<td>
		<%= ojumundetail.FItemList(i).Fbaljuitemno %>
	</td>
	<td>
		<%= ojumundetail.FItemList(i).Frealitemno %>
	</td>
	<td align="center"><%= FormatNumber(ojumundetail.FItemList(i).FSellcash * ojumundetail.FItemList(i).Fbaljuitemno,0) %></td>
	<td align="center"><%= FormatNumber(ojumundetail.FItemList(i).FSuplycash * ojumundetail.FItemList(i).Fbaljuitemno,0) %></td>
	<td align="center"><%= FormatNumber(ojumundetail.FItemList(i).FSellcash * ojumundetail.FItemList(i).Frealitemno,0) %></td>
	<td align="center"><%= FormatNumber(ojumundetail.FItemList(i).FSuplycash * ojumundetail.FItemList(i).Frealitemno,0) %></td>
	<td align="center"><%= ojumundetail.FItemList(i).Fcentermwdiv %></td>	
</tr>
<% next %>

<tr bgcolor="#FFFFFF" align="center">
	<td colspan=10>총계</td>
	<td><%= totalfixno %></td>
	<td colspan=2></td>
	<td 
		<%= formatNumber(selltotal,0) %><br>
		<%= formatNumber(selltotalfix,0) %>
	</td>
	<td>
		<%= formatNumber(suplytotal,0) %><br>
		<%= formatNumber(suplytotalfix,0) %>
	</td>
	<td></td>
</tr>
<% else %>
<tr bgcolor="#FFFFFF" align="center">
	<td colspan=20>검색 결과가 없습니다</td>
</tr>
<% end if %>
</table>
<%
set ojumundetail = Nothing

set ojumundetail= new COrderSheet
	ojumundetail.FRectIdx = idx
	ojumundetail.frectitemgubun = "70"
	ojumundetail.GetOrderSheetDetail_gubun

selltotal =0
suplytotal =0
buytotal =0
selltotalfix =0
suplytotalfix =0
buytotalfix =0
totalfixno = 0
%>
<br>
※기타상품(70)
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td colspan="20" align="right">
		총건수: <%= ojumundetail.FResultCount %>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td>상품</td>
    <td>이미지</td>
	<td>바코드</td>
	<td>브랜드</td>
	<td>상품명<br>[옵션명]</td>
	<td>날짜</td>
	<td>주문시<br>판매가</td>
	<td>출고<br>공급가</td>
	<td>매입가</td>
	<td>주문수</td>
	<td>확정수</td>
	<td>주문<br>공급가</td>
	<td>주문<br>판매가</td>
	<td>확정<br>공급가</td>
	<td>확정<br>판매가</td>	
	<td>센터<br>매입<br>구분</td>
</tr>
<%
if ojumundetail.FResultCount > 0 then
	 
for i=0 to ojumundetail.FResultCount-1

selltotal  = selltotal + ojumundetail.FItemList(i).FSellcash * ojumundetail.FItemList(i).Fbaljuitemno
suplytotal = suplytotal + ojumundetail.FItemList(i).FSuplycash * ojumundetail.FItemList(i).Fbaljuitemno
buytotal   = buytotal + ojumundetail.FItemList(i).Fbuycash * ojumundetail.FItemList(i).Fbaljuitemno
selltotalfix  = selltotalfix + ojumundetail.FItemList(i).FSellcash * ojumundetail.FItemList(i).Frealitemno
suplytotalfix = suplytotalfix + ojumundetail.FItemList(i).FSuplycash * ojumundetail.FItemList(i).Frealitemno
buytotalfix   = buytotalfix + ojumundetail.FItemList(i).Fbuycash * ojumundetail.FItemList(i).Frealitemno
totalfixno = totalfixno + ojumundetail.FItemList(i).Frealitemno
%>
<tr align="center" bgcolor="#FFFFFF">
	<td>
		<%= ojumundetail.FItemList(i).FItemGubun %>-<%= ojumundetail.FItemList(i).FItemid %>-<%= ojumundetail.FItemList(i).FItemoption %>
	</td>
	<td><img src="<%= ojumundetail.FItemList(i).GetImageSmall %>" border="0" width="50" height="50" onError="this.src='http://image.10x10.co.kr/images/no_image.gif'"></td>
	<td><%= ojumundetail.FItemList(i).FItemGubun %><%= format00(6,ojumundetail.FItemList(i).FItemID) %><%= ojumundetail.FItemList(i).Fitemoption %></td>
	<td><%= ojumundetail.FItemList(i).Fmakerid %></td>
	<td>
		<%= ojumundetail.FItemList(i).Fitemname %>
		<% if ojumundetail.FItemList(i).Fitemoption <> "0000" then %>
			<br><%= ojumundetail.FItemList(i).Fitemoptionname %>
		<% end if %>
	</td>
	<td><%= ojumundetail.FItemList(i).fregdate %></td>
	<td>
		<% if   (ojumundetail.FItemList(i).FItemDefaultMwDiv<>"W") and (ojumundetail.FItemList(i).Fbuycash>ojumundetail.FItemList(i).Fsuplycash) then %>
		<b><font color=red><%= FormatNumber(ojumundetail.FItemList(i).Fsellcash,0) %></font></b>
		<% else %>
		<%= FormatNumber(ojumundetail.FItemList(i).Fsellcash,0) %>
		<% end if %>

		<% if (ojumundetail.FItemList(i).IsOnLineItem) and (ojumundetail.FItemList(i).Fsellcash<>ojumundetail.FItemList(i).Fonlinesellcash) then %>
		<br>
		<div ><font color=red>온:<%= FormatNumber(ojumundetail.FItemList(i).Fonlinesellcash,0) %></font></div>
		<% end if %>
	</td>
	<td>
		<%= ojumundetail.FItemList(i).Fsuplycash %>
		<% if (ojumundetail.FItemList(i).IsOnLineItem) and (ojumundetail.FItemList(i).GetOrgShopSuplycashbyMargine<>ojumundetail.FItemList(i).Fsuplycash) then %>
			<div ><font color=red><%= ojumundetail.FItemList(i).GetOrgShopSuplycashbyMargine %></font></div>
		<% end if %>
	</td>
	<td>
		<%= ojumundetail.FItemList(i).Fbuycash %>
		<% if (ojumundetail.FItemList(i).Fbuycash<>ojumundetail.FItemList(i).Fonlinebuycash) and ((ojumundetail.FItemList(i).FItemDefaultMwDiv="W") and (ojumundetail.FItemList(i).FoffChargeDiv="4")) then %>
		<div ><font color="red">온:<%= ojumundetail.FItemList(i).Fonlinebuycash %></font></div>
		<% end if %>
	</td>
	<td>
		<%= ojumundetail.FItemList(i).Fbaljuitemno %>
	</td>
	<td>
		<%= ojumundetail.FItemList(i).Frealitemno %>
	</td>
	<td align="center"><%= FormatNumber(ojumundetail.FItemList(i).FSellcash * ojumundetail.FItemList(i).Fbaljuitemno,0) %></td>
	<td align="center"><%= FormatNumber(ojumundetail.FItemList(i).FSuplycash * ojumundetail.FItemList(i).Fbaljuitemno,0) %></td>
	<td align="center"><%= FormatNumber(ojumundetail.FItemList(i).FSellcash * ojumundetail.FItemList(i).Frealitemno,0) %></td>
	<td align="center"><%= FormatNumber(ojumundetail.FItemList(i).FSuplycash * ojumundetail.FItemList(i).Frealitemno,0) %></td>
	<td align="center"><%= ojumundetail.FItemList(i).Fcentermwdiv %></td>	
</tr>
<% next %>

<tr bgcolor="#FFFFFF" align="center">
	<td colspan=10>총계</td>
	<td><%= totalfixno %></td>
	<td colspan=2></td>
	<td 
		<%= formatNumber(selltotal,0) %><br>
		<%= formatNumber(selltotalfix,0) %>
	</td>
	<td>
		<%= formatNumber(suplytotal,0) %><br>
		<%= formatNumber(suplytotalfix,0) %>
	</td>
	<td></td>
</tr>
<% else %>
<tr bgcolor="#FFFFFF" align="center">
	<td colspan=20>검색 결과가 없습니다</td>
</tr>
<% end if %>
</table>

<%
set ojumundetail = Nothing
%>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->