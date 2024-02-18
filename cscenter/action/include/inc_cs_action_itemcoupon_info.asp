<%
'###########################################################
' Description : 상품쿠폰관련 정보
' Hieditor : 2023.10.16 한용민 생성
'###########################################################
%>
<%
' 주문취소 or (반품접수(업체배송) or 회수신청(텐바이텐배송))
if (IsCSCancelProcess(divcd) or IsCSReturnProcess(divcd)) then
%>
	<% if oCsItemCoupon.FResultCount>0 then %>
		<tr bgcolor="FFFFFF" align="center">
			<td>쿠폰명</td>
			<td width="60">할인값</td>
			<td width="150">유효기간</td>
			<td width="80">상태</td>
		</tr>
		<% for i = 0 to oCsItemCoupon.FResultCount-1 %>
			<tr bgcolor="FFFFFF" align="center">
				<td align="left" >
					<%= oCsItemCoupon.FItemList(i).fitemcouponname %>
					<br>쿠폰코드:<%= oCsItemCoupon.FItemList(i).fitemcouponidx %>
				</td>
				<td >
					<%= oCsItemCoupon.FItemList(i).GetDiscountStr %>
				</td>
				<td >
					<%= ChkIIF(Right(oCsItemCoupon.FItemList(i).Fitemcouponstartdate,8)="00:00:00",Left(oCsItemCoupon.FItemList(i).Fitemcouponstartdate,10),oCsItemCoupon.FItemList(i).Fitemcouponstartdate) %>
					~
					<%= ChkIIF(Right(oCsItemCoupon.FItemList(i).Fitemcouponexpiredate,8)="23:59:59",Left(oCsItemCoupon.FItemList(i).Fitemcouponexpiredate,10),oCsItemCoupon.FItemList(i).Fitemcouponexpiredate) %>
				</td>
				<td >
					<%= oCsItemCoupon.FItemList(i).GetOpenStateName %>

					<% if (oCsItemCoupon.FItemList(i).forderserial="" or isnull(oCsItemCoupon.FItemList(i).forderserial)) and oCsItemCoupon.FItemList(i).fusedyn<>"Y" then %>
						<br>쿠폰미사용
					<% else %>
						<br>쿠폰사용
					<% end if %>

					<% if not(oCsItemCoupon.FItemList(i).IsItemCouponCopyValid) then %>
						<br><font color="red">재발급불가</font>
					<% else %>
						<br>재발급가능
					<% end if %>
				</td>
			</tr>
		<% next %>
	<% end if %>
<% end if %>