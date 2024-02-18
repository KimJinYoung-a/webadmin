<%
	'// 현재 브랜드를 제외한 같은 카테고리 내의 다른 베스트셀러 (초반은 대분류만, 상품이 많아지면 중분류까지 지정)
	dim oMHItem
	set oMHItem = new SearchItemCls
	oMHItem.FRectSortMethod		= "be"			'베스트셀러
	oMHItem.FRectSearchItemDiv	= ""
	oMHItem.FRectNotMakerid		= oItem.Prd.FMakerid '현재 브랜드 제외
	oMHItem.FRectCateCode		= vDisp
	oMHItem.FCurrPage			= 1
	oMHItem.FPageSize			= 6
	oMHItem.FScrollCount		= 1
	oMHItem.FSellScope			= "Y"			'판매상품만
	oMHItem.getSearchList

	if oMHItem.FResultCount >0 then
%>
<div class="box1 fingersChoice">
	<h2>핑거스 초이스</h2>
	<div class="thumbList">
		<div class="swiper-container">
			<div class="swiper-wrapper">

				<% for lp=0 to oMHItem.FResultCount-1 %>
				<div class="swiper-slide">
					<div class="thumbImg"><a href="/diyshop/shop_prd.asp?itemid=<%= oMHItem.FItemList(lp).FItemID %>&disp=<%=vDisp%>"><img src="<%=oMHItem.FItemList(lp).FImageList120%>" alt="<%=oMHItem.FItemList(lp).FItemName%>" /></a></div>
					<p class="title"><%=oMHItem.FItemList(lp).FItemName%></p>
					<% if oMHItem.FItemList(lp).IsSaleItem or oMHItem.FItemList(lp).isCouponItem Then %>
						<% IF oMHItem.FItemList(lp).IsSaleItem then %>
						<p class="price"><% = FormatNumber(oMHItem.FItemList(lp).getRealPrice,0) %>원 <span>[<% = oMHItem.FItemList(lp).getSalePro %>]</span></p>
						<% End IF %>
						<% IF oMHItem.FItemList(lp).IsCouponItem then %>
						<p class="price"><% = FormatNumber(oMHItem.FItemList(lp).GetCouponAssignPrice,0) %>원 <span>[<% = oMHItem.FItemList(lp).GetCouponDiscountStr %>]</span></p>
						<% End IF %>
					<% Else %>
						<p class="price"><% = FormatNumber(oMHItem.FItemList(lp).getRealPrice,0) %>원</p>
					<% End if %>
				</div>
				<% Next %>
			</div>
		</div>
	</div>
	<!--<a href="" class="btnMoreView"><span>작품 전체보기</span></a> -->
</div>
<%
	end if
	set oMHItem = Nothing
%>