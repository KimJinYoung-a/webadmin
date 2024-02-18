<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프라인 기프트카드 프로모션 관리
' History : 2018.01.15 이상구 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopcardcls.asp"-->
<%

menupos = request("menupos")



dim page, shopid
dim research, i

page = request("page")
shopid = request("shopid")
research = request("research")

if (page = "") then
	page = 1
end if


'================================================================================
dim oOffShopCardPromotion

set oOffShopCardPromotion = new COffShopCardPromotion

oOffShopCardPromotion.FRectShopid = shopid

oOffShopCardPromotion.FCurrPage = page
oOffShopCardPromotion.Fpagesize = 25

oOffShopCardPromotion.COffShopCardPromotionList

%>
<script>

function fnGoto(page) {
	var frm = document.frm;
	frm.page = page;
	frm.submit();
}

function popCardPromotionModi(idx) {
	var popwin = window.open('pop_card_promotion_modi.asp?idx=' + idx,'popCardPromotionModi','width=400, height=300, scrollbars=yes, resizable=yes');
	popwin.focus();
}

</script>
<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="page" value="1">
	<tr align="center" bgcolor="#FFFFFF" >
		<td width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			ShopID : 
			<% drawSelectBoxOffShop "shopid",shopid %>
			<% Call NewDrawSelectBoxDesignerwithNameAndUserDIV("shopid",shopid, "21") %>
		</td>
		<td width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->

<p />

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="right">
		<input type="button" value="프로모션 등록" onclick="popCardPromotionModi(-1);" class="button">
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<p />

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="18">
			검색결과 : <b><%= oOffShopCardPromotion.FTotalCount %></b>
			&nbsp;
			페이지 : <b><%= page %> / <%= oOffShopCardPromotion.FTotalpage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="40">IDX</td>
		<td>샵아이디</td>
		<td>기프트카드금액</td>
		<td>시작일</td>
		<td>종료일</td>
		<td>지급기준</td>
		<td>지급혜택</td>
		<td>비고</td>
	</tr>
	<% if oOffShopCardPromotion.FResultCount > 0 then %>
	<% for i=0 to oOffShopCardPromotion.FResultcount-1 %>
	<tr bgcolor="#FFFFFF">
		<td align="center"><%= oOffShopCardPromotion.FItemList(i).Fidx %></td>
		<td align="center"><a href="javascript:popCardPromotionModi(<%= oOffShopCardPromotion.FItemList(i).Fidx %>)"><%= oOffShopCardPromotion.FItemList(i).Fshopid %></a></td>
		<td align="center"><%= FormatNumber(oOffShopCardPromotion.FItemList(i).FcardPrice, 0) %></td>
		<td align="center"><%= oOffShopCardPromotion.FItemList(i).FstartDate %></td>
		<td align="center"><%= oOffShopCardPromotion.FItemList(i).FendDate %></td>
		<td align="center"><%= oOffShopCardPromotion.FItemList(i).getRateGubunName %></td>
		<td align="center"><%= oOffShopCardPromotion.FItemList(i).FrateAmmount %></td>
		<td></td>
	</tr>
	<% next %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="18" align="center">
			<%
			dim strparam
			strparam = "&shopid=" + CStr(shopid)

			strparam = strparam + "&menupos=" + CStr(menupos)

			%>
			<% if oOffShopCardPromotion.HasPreScroll then %>
				<a href="javascript:fnGoto(<%= oOffShopCardPromotion.StartScrollPage-1 %>)">[pre]</a>
			<% else %>
				[pre]
			<% end if %>

			<% for i=0 + oOffShopCardPromotion.StartScrollPage to oOffShopCardPromotion.FScrollCount + oOffShopCardPromotion.StartScrollPage - 1 %>
				<% if i>oOffShopCardPromotion.FTotalpage then Exit for %>
				<% if CStr(page)=CStr(i) then %>
				<font color="red">[<%= i %>]</font>
				<% else %>
				<a href="javascript:fnGoto(<%= i %>)">[i]</a>
				<% end if %>
			<% next %>

			<% if oOffShopCardPromotion.HasNextScroll then %>
				<a href="javascript:fnGoto(<%= i %>)">[next]</a>
			<% else %>
				[next]
			<% end if %>
		</td>
	</tr>
	<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="8" align=center>[ 검색결과가 없습니다. ]</td>
	</tr>
	<% end if %>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
