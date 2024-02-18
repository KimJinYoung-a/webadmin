<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프라인 샵별패킹리스트(박스별)
' History : 이상구 생성
'			2017.04.11 한용민 수정(보안관련처리)
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/stock/offinvoicecls.asp"-->
<%

menupos = requestCheckVar(request("menupos"),10)

dim page, shopid, statecd
dim research, i

page = requestCheckVar(request("page"),10)
shopid = requestCheckVar(request("shopid"),32)
research = requestCheckVar(request("research"),2)

if (C_IS_SHOP) then
    shopid = C_STREETSHOPID
    statecd = "7"
end if

if (page = "") then
	page = 1
end if

dim ocoffinvoice
set ocoffinvoice = new COffInvoice
	ocoffinvoice.FRectShopid = shopid
	ocoffinvoice.FRectStateCD = statecd
	ocoffinvoice.FCurrPage = page
	ocoffinvoice.Fpagesize = 25
	ocoffinvoice.GetMasterList

%>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		ShopID :
		<% if (C_IS_SHOP) then %>
			<%= shopid %>
		<% else %>
			<% 'drawSelectBoxOffShop "shopid",shopid %>
			<% Call NewDrawSelectBoxDesignerwithNameAndUserDIV("shopid",shopid, "21") %>
		<% end if %>
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->

<br>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		검색결과 : <b><%= ocoffinvoice.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %> / <%= ocoffinvoice.FTotalpage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="40">IDX</td>
	<td>샵아이디</td>
	<td>인보이스<br>NO</td>
	<td>운송<br>방법</td>
	<td>운임<br>부담</td>
	<td>정산<br>시기</td>
	<td>박스<br>수량</td>
	<td>총상품금액<br>(원)</td>
	<td>총운임<br>(원)</td>
	<td>작성화폐</td>
	<td>수출환율</td>
	<td>총상품금액<br>(외환)</td>
	<td>총운임<br>(외환)</td>
	<td width="80">등록일</td>
	<td>비고</td>
</tr>
<% if ocoffinvoice.FResultCount >0 then %>
<% for i=0 to ocoffinvoice.FResultcount-1 %>
<tr bgcolor="#FFFFFF">
	<td align="center"><%= ocoffinvoice.FItemList(i).Fidx %></td>
	<td align="center"><a href="shop_offinvoiceview.asp?menupos=<%= menupos %>&idx=<%= ocoffinvoice.FItemList(i).Fidx %>"><%= ocoffinvoice.FItemList(i).Fshopid %><br><%= ocoffinvoice.FItemList(i).Fshopname %></a></td>
	<td align="center"><%= ocoffinvoice.FItemList(i).Finvoiceno %></td>
	<td align="center"><%= ocoffinvoice.FItemList(i).GetDeliverMethodName %></td>
	<td align="center"><%= ocoffinvoice.FItemList(i).GetExportMethodName %></td>
	<td align="center"><%= ocoffinvoice.FItemList(i).GetJungsanTypeName %></td>
	<td align="center"><%= ocoffinvoice.FItemList(i).Ftotalboxno %></td>
	<td align="right">
		<%= FormatNumber(ocoffinvoice.FItemList(i).Ftotalgoodsprice, 0) %>&nbsp;
	</td>
	<td align="right">
		<%= FormatNumber(ocoffinvoice.FItemList(i).Ftotalboxprice, 0) %>&nbsp;
	</td>
	<td align="center"><%= ocoffinvoice.FItemList(i).Fpriceunit %></td>
	<td align="center"><%= FormatNumber(ocoffinvoice.FItemList(i).Fexchangerate, 0) %> 원</td>
	<td align="right">
		<% if (ocoffinvoice.FItemList(i).Fexchangerate <> "") and (Not IsNull(ocoffinvoice.FItemList(i).Fexchangerate)) and (ocoffinvoice.FItemList(i).Fexchangerate <> "0") then %>
			<%= FormatNumber(Round((ocoffinvoice.FItemList(i).Ftotalgoodsprice/ocoffinvoice.FItemList(i).Fexchangerate), 2), 2) %>&nbsp;
		<% end if %>
	</td>
	<td align="right">
		<% if (ocoffinvoice.FItemList(i).Fexchangerate <> "") and (Not IsNull(ocoffinvoice.FItemList(i).Fexchangerate)) and (ocoffinvoice.FItemList(i).Fexchangerate <> "0") then %>
			<%= FormatNumber(Round((ocoffinvoice.FItemList(i).Ftotalboxprice/ocoffinvoice.FItemList(i).Fexchangerate), 2), 2) %>&nbsp;
		<% end if %>
	</td>
	<td align="center"><%= Left(ocoffinvoice.FItemList(i).Fregdate, 10) %></td>
	<td align="center">
	</td>
</tr>
<% next %>
<% else %>
<tr bgcolor="#FFFFFF">
	<td colspan=15 align=center>[ 검색결과가 없습니다. ]</td>
</tr>
<% end if %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
		<%
		dim strparam
		strparam = "&shopid=" + CStr(shopid)

		strparam = strparam + "&menupos=" + CStr(menupos)

		%>
		<% if ocoffinvoice.HasPreScroll then %>
			<a href="?page=<%= ocoffinvoice.StartScrollPage-1 %>&research=on<%= strparam %>">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + ocoffinvoice.StartScrollPage to ocoffinvoice.FScrollCount + ocoffinvoice.StartScrollPage - 1 %>
			<% if i>ocoffinvoice.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="?page=<%= i %>&research=on<%= strparam %>">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if ocoffinvoice.HasNextScroll then %>
			<a href="?page=<%= i %>&research=on<%= strparam %>">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>
</table>

<%
set ocoffinvoice = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
