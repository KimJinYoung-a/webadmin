<%
	'// 현재 브랜드의 동일 카테고리 내의 다른 베스트셀러 (초반은 대분류만, 상품이 많아지면 중분류까지 지정)
	dim oRTItem
	set oRTItem = new DIYItemPrdCls
	oRTItem.FRectItemid			= itemid		'현재 상품 제외
	oRTItem.FRectMakerid		= oItem.Prd.FMakerid
	oRTItem.FRectCateCode		= vDisp
	oRTItem.FPageSize			= 6
	oRTItem.FSellScope			= "Y"			'판매상품만
	oRTItem.getDIYBESTItemList

	if oRTItem.FResultCount >0 then
%>
<div class="thumbList">
	<div class="swiper-container">
		<div class="swiper-wrapper">
			<% for lp=0 to oRTItem.FResultCount-1 %>
			<div class="swiper-slide">
				<div class="thumbImg"><a href="/diyshop/shop_prd.asp?itemid=<%= oRTItem.FItemList(lp).FItemID %>"/><img src="<%=oRTItem.FItemList(lp).FImageList120%>" alt="<%=oRTItem.FItemList(lp).FItemName%>" /></a></div>
				<p class="title"><%=oRTItem.FItemList(lp).FItemName%></p>
				<% if oRTItem.FItemList(lp).IsSaleItem or oRTItem.FItemList(lp).isCouponItem Then %>
					<% IF oRTItem.FItemList(lp).IsSaleItem then %>
					<p class="price"><% = FormatNumber(oRTItem.FItemList(lp).getRealPrice,0) %>원 <span>[<% = oRTItem.FItemList(lp).getSalePro %>]</span></p>
					<% End If %>
					<% IF oRTItem.FItemList(lp).IsCouponItem then %>
					<p class="price"><% = FormatNumber(oRTItem.FItemList(lp).GetCouponAssignPrice,0) %>원 <span>[<% = oRTItem.FItemList(lp).GetCouponDiscountStr %>]</span></p>
					<% End If %>
				<% Else %>
					<p class="price"><% = FormatNumber(oRTItem.FItemList(lp).getRealPrice,0) %>원</p>
				<% End if %>
			</div>
			<% next %>
		</div>
	</div>
</div>
<%
	end if
	set oRTItem = Nothing
%>