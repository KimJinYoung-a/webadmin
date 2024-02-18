<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/datamart/brandStaticCls.asp"-->
<%

Dim makerID

makerID = req("makerID", "")

dim oCBrandServiceByMeachul, oCBrandServiceByAction, oCBrandServiceByDelivery, oCBrandServiceByClaim

set oCBrandServiceByMeachul = new CBrandService
if (makerID <> "") then
	oCBrandServiceByMeachul.GetBrandServiceByMeachulOne(makerid)
end if

set oCBrandServiceByAction = new CBrandService
if (makerID <> "") then
	oCBrandServiceByAction.GetBrandServiceByActionOne(makerid)
end if

set oCBrandServiceByDelivery = new CBrandService
if (makerID <> "") then
	oCBrandServiceByDelivery.GetBrandServiceByDeliveryOne(makerid)
end if

set oCBrandServiceByClaim = new CBrandService
if (makerID <> "") then
	oCBrandServiceByClaim.GetBrandServiceByClaimOne(makerid)
end if

dim val

function dispUpDnRate(currPrc, prevPrc, currDt, prevDt)
	dim val
	if (currPrc = 0 or prevPrc = 0) then
		dispUpDnRate = "-"
	elseif (1.0 * (currPrc * prevDt) / (prevPrc * currDt) * 100) > 500 then
		dispUpDnRate = "500%+"
	else
		val = (1.0 * (currPrc * prevDt) / (prevPrc * currDt) * 100)
		if (val > 100) then
			val = "<font color='red'>" & FormatNumber(val, 2) & "%" & "</font>"
		elseif (val < 100) then
			val = "<font color='blue'>" & FormatNumber(val, 2) & "%" & "</font>"
		else
			val = FormatNumber(val, 2) & "%"
		end if

		dispUpDnRate = val
	end if
end function

%>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="80" bgcolor="<%= adminColor("gray") %>">검색조건</td>
		<td align="left">
			&nbsp;
			브랜드ID :
			<input type="text" class="text" name="makerID" value="<%=makerID%>">
		</td>

		<td rowspan="2" width="80" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->

<p />

<!-- 리스트 시작 -->
[매출지수]
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
 	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td width="80" rowspan="2">기준일</td>
		<td width="200" rowspan="2">브랜드</td>
		<td width="320" colspan="4">1일 판매내역</td>
		<td width="320" colspan="4">7일 판매내역</td>
		<td width="320" colspan="4">30일 판매내역</td>
		<td width="320" colspan="4">90일 판매내역</td>
		<td width="240" colspan="3">360일 판매내역</td>
		<td rowspan="2">비고</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="80">판매수량</td>
		<td width="80">판매금액</td>
		<td width="80">주문건수</td>
		<td width="80">등락율<br />(금액)</td>
		<td width="80">판매수량</td>
		<td width="80">판매금액</td>
		<td width="80">주문건수</td>
		<td width="80">등락율<br />(금액)</td>
		<td width="80">판매수량</td>
		<td width="80">판매금액</td>
		<td width="80">주문건수</td>
		<td width="80">등락율<br />(금액)</td>
		<td width="80">판매수량</td>
		<td width="80">판매금액</td>
		<td width="80">주문건수</td>
		<td width="80">등락율<br />(금액)</td>
		<td width="80">판매수량</td>
		<td width="80">판매금액</td>
		<td width="80">주문건수</td>
	</tr>
	<% if (oCBrandServiceByMeachul.FresultCount > 0) then %>
	<tr align="center" bgcolor="#FFFFFF">
		<td><%= oCBrandServiceByMeachul.FOneItem.Fyyyymmdd %></td>
		<td><%= oCBrandServiceByMeachul.FOneItem.Fmakerid %></td>
		<td><%= FormatNumber(oCBrandServiceByMeachul.FOneItem.FoneDaySellItemCnt,0) %></td>
		<td><%= FormatNumber(oCBrandServiceByMeachul.FOneItem.FoneDaySelltotalPrice,0) %></td>
		<td><%= FormatNumber(oCBrandServiceByMeachul.FOneItem.FoneDaySellOrderCnt,0) %></td>
		<td><%= dispUpDnRate(oCBrandServiceByMeachul.FOneItem.FoneDaySelltotalPrice, oCBrandServiceByMeachul.FOneItem.FoneWeekSelltotalPrice, 1, 7) %></td>
		<td><%= FormatNumber(oCBrandServiceByMeachul.FOneItem.FoneWeekSellItemCnt,0) %></td>
		<td><%= FormatNumber(oCBrandServiceByMeachul.FOneItem.FoneWeekSelltotalPrice,0) %></td>
		<td><%= FormatNumber(oCBrandServiceByMeachul.FOneItem.FoneWeekSellOrderCnt,0) %></td>
		<td><%= dispUpDnRate(oCBrandServiceByMeachul.FOneItem.FoneWeekSelltotalPrice, oCBrandServiceByMeachul.FOneItem.FoneMonthSelltotalPrice, 7, 30) %></td>
		<td><%= FormatNumber(oCBrandServiceByMeachul.FOneItem.FoneMonthSellItemCnt,0) %></td>
		<td><%= FormatNumber(oCBrandServiceByMeachul.FOneItem.FoneMonthSelltotalPrice,0) %></td>
		<td><%= FormatNumber(oCBrandServiceByMeachul.FOneItem.FoneMonthSellOrderCnt,0) %></td>
		<td><%= dispUpDnRate(oCBrandServiceByMeachul.FOneItem.FoneMonthSelltotalPrice, oCBrandServiceByMeachul.FOneItem.FthreeMonthSelltotalPrice, 30, 90) %></td>
		<td><%= FormatNumber(oCBrandServiceByMeachul.FOneItem.FthreeMonthSellItemCnt,0) %></td>
		<td><%= FormatNumber(oCBrandServiceByMeachul.FOneItem.FthreeMonthSelltotalPrice,0) %></td>
		<td><%= FormatNumber(oCBrandServiceByMeachul.FOneItem.FthreeMonthSellOrderCnt,0) %></td>
		<td><%= dispUpDnRate(oCBrandServiceByMeachul.FOneItem.FthreeMonthSelltotalPrice, oCBrandServiceByMeachul.FOneItem.FoneYearSelltotalPrice, 90, 360) %></td>
		<td><%= FormatNumber(oCBrandServiceByMeachul.FOneItem.FoneYearSellItemCnt,0) %></td>
		<td><%= FormatNumber(oCBrandServiceByMeachul.FOneItem.FoneYearSelltotalPrice,0) %></td>
		<td><%= FormatNumber(oCBrandServiceByMeachul.FOneItem.FoneYearSellOrderCnt,0) %></td>
		<td></td>
	</tr>
	<% end if %>
</table>

<p />

[활동지수]
<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
 	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td width="80">년월</td>
		<td width="200">브랜드</td>
		<td width="80">이벤트<br />등록건수</td>
		<td width="80">신상품<br />등록건수</td>
		<td width="80">상품후기<br />등록건수</td>
		<td width="80">상품후기<br />평점</td>
		<td width="80">상품위시<br />등록건수</td>
		<td width="80">브랜드찜<br />등록건수</td>
		<td width="80">상품문의<br />등록건수</td>
		<td width="80">평균답변<br />등록일수</td>
		<td>비고</td>
	</tr>
	<% if (oCBrandServiceByAction.FresultCount > 0) then %>
	<tr align="center" bgcolor="#FFFFFF">
		<td><%= oCBrandServiceByAction.FOneItem.Fyyyymm %></td>
		<td><%= oCBrandServiceByAction.FOneItem.Fmakerid %></td>
		<td><%= FormatNumber(oCBrandServiceByAction.FOneItem.FeventRegCnt,0) %></td>
		<td><%= FormatNumber(oCBrandServiceByAction.FOneItem.FnewItemRegCnt,0) %></td>
		<td><%= FormatNumber(oCBrandServiceByAction.FOneItem.FitemReviewCnt,0) %></td>
		<td>
			<%
			if (oCBrandServiceByAction.FOneItem.FitemReviewCnt > 0) then
				response.write FormatNumber(oCBrandServiceByAction.FOneItem.FitemReviewPointSUM/oCBrandServiceByAction.FOneItem.FitemReviewCnt,2)
			else
				response.write "-"
			end if
			%>
		</td>
		<td><%= FormatNumber(oCBrandServiceByAction.FOneItem.FitemWishCnt,0) %></td>
		<td><%= FormatNumber(oCBrandServiceByAction.FOneItem.FbrandZzimCnt,0) %></td>
		<td><%= FormatNumber(oCBrandServiceByAction.FOneItem.FitemQnaRegCnt,0) %></td>
		<td>
			<%
			if (oCBrandServiceByAction.FOneItem.FitemQnaRegCnt > 0) then
				response.write FormatNumber(oCBrandServiceByAction.FOneItem.FitemQnaAnsDaySUM/oCBrandServiceByAction.FOneItem.FitemQnaRegCnt,2)
			else
				response.write "-"
			end if
			%>
		</td>
		<td></td>
	<% end if %>
</table>

<p />

[배송지수]
<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
 	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td rowspan="2" width="80">
			년월
		</td>
		<td rowspan="2" width="200">브랜드</td>
		<td width="80" rowspan="2">총발주건수<br>(업체배송)</td>
        <td colspan="6">고객불만 취소(반품)건수</td>
        <td colspan="4" width="80">평균배송소요일</td>
		<td rowspan="2" width="80"><b>서비스지수</b></td>
		<td rowspan="2">비고</td>
	</tr>
 	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
        <td width="80">품절</td>
		<td width="80">배송지연</td>
		<td width="80">상품불량</td>
		<td width="80">오배송</td>
		<td width="80">합계</td>
		<td width="80"><b>비율<br />(발주대비)</b></td>
		<td width="80">출고건수</td>
		<td width="80">배송일기준</td>
		<td width="80">송장조회기준</td>
		<td width="80">허위송장건수</td>
	</tr>
	<% if (oCBrandServiceByDelivery.FresultCount > 0) then %>
	<tr align="center" bgcolor="#FFFFFF">
		<td><%= oCBrandServiceByDelivery.FOneItem.Fyyyymm %></td>
		<td><%= oCBrandServiceByDelivery.FOneItem.Fmakerid %></td>
		<td><%= oCBrandServiceByDelivery.FOneItem.FbaljuCnt %></td>
		<td><%= oCBrandServiceByDelivery.FOneItem.FstockoutCnt %></td>
		<td><%= oCBrandServiceByDelivery.FOneItem.FdelayCnt %></td>
		<td><%= oCBrandServiceByDelivery.FOneItem.FbaditemCnt %></td>
		<td><%= oCBrandServiceByDelivery.FOneItem.FerrdeliveryCnt %></td>
		<td><%= oCBrandServiceByDelivery.FOneItem.GetSUM %></td>
		<td>
			<%
			if oCBrandServiceByDelivery.FOneItem.FbaljuCnt > 0 then
				val = Round((1.0 * oCBrandServiceByDelivery.FOneItem.GetSUM / oCBrandServiceByDelivery.FOneItem.FbaljuCnt * 100), 1)
				if (val >= 5) then
					response.write "<font color='red'><b>" & val & "%</b></font>"
				else
					response.write val & "%"
				end if
			else
				response.write "-"
			end if
			%>
		</td>
		<td><%= oCBrandServiceByDelivery.FOneItem.FchulgoCnt %></td>
		<% if oCBrandServiceByDelivery.FOneItem.FchulgoCnt > 0 then %>
		<td><%= Round(1.0*oCBrandServiceByDelivery.FOneItem.FchulgoNDaySum/oCBrandServiceByDelivery.FOneItem.FchulgoCnt,1) %></td>
		<td><%= Round(1.0*(oCBrandServiceByDelivery.FOneItem.FchulgoNDaySum+oCBrandServiceByDelivery.FOneItem.FrealOverNDaySum)/oCBrandServiceByDelivery.FOneItem.FchulgoCnt,1) %></td>
		<td>
			<% if (oCBrandServiceByDelivery.FOneItem.FfalsehoodSongjangCnt > 0) then %>
			<font color="red"><b><%= oCBrandServiceByDelivery.FOneItem.FfalsehoodSongjangCnt %></b></font>
			<% else %>
			<%= oCBrandServiceByDelivery.FOneItem.FfalsehoodSongjangCnt %>
			<% end if %>
		</td>
		<% else %>
		<td>-</td>
		<td>-</td>
		<td>-</td>
		<% end if %>
		<td></td>
		<td></td>
	</tr>
	<% end if %>
</table>

<p />

[클래임비용]
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
 	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td rowspan="2" width="80">
			년월
		</td>
		<td rowspan="2" width="200">브랜드</td>
		<td width="80" rowspan="2">총건수<br>(업체배송)</td>
		<td width="80" rowspan="2">총비용<br>(업체배송)</td>
        <td colspan="6">클래임 건수</td>
		<td colspan="6">클래임 비용</td>
		<td rowspan="2">비고</td>
	</tr>
 	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="80">배송지연</td>
        <td width="80">품절</td>
		<td width="80">오배송</td>
		<td width="80">상품등록오류</td>
		<td width="80">업체대응불량</td>
		<td width="80">기타업체과실</td>
		<td width="80">배송지연</td>
        <td width="80">품절</td>
		<td width="80">오배송</td>
		<td width="80">상품등록오류</td>
		<td width="80">업체대응불량</td>
		<td width="80">기타업체과실</td>
	</tr>
	<% if (oCBrandServiceByClaim.FresultCount > 0) then %>
	<tr align="center" bgcolor="#FFFFFF">
		<td><%= oCBrandServiceByClaim.FOneItem.Fyyyymm %></td>
		<td><%= oCBrandServiceByClaim.FOneItem.Fmakerid %></td>
		<td><%= oCBrandServiceByClaim.FOneItem.FtotCnt %></td>
		<td><%= FormatNumber(oCBrandServiceByClaim.FOneItem.FtotSum,0) %></td>
		<td><%= oCBrandServiceByClaim.FOneItem.FdelayCnt %></td>
		<td><%= oCBrandServiceByClaim.FOneItem.FstockoutCnt %></td>
		<td><%= oCBrandServiceByClaim.FOneItem.FerrdeliveryCnt %></td>
		<td><%= oCBrandServiceByClaim.FOneItem.FitemregerrCnt %></td>
		<td><%= oCBrandServiceByClaim.FOneItem.FupcheerrCnt %></td>
		<td><%= oCBrandServiceByClaim.FOneItem.FetcupcheerrCnt %></td>
		<td><%= FormatNumber(oCBrandServiceByClaim.FOneItem.FdelaySum,0) %></td>
		<td><%= FormatNumber(oCBrandServiceByClaim.FOneItem.FstockoutSum,0) %></td>
		<td><%= FormatNumber(oCBrandServiceByClaim.FOneItem.FerrdeliverySum,0) %></td>
		<td><%= FormatNumber(oCBrandServiceByClaim.FOneItem.FitemregerrSum,0) %></td>
		<td><%= FormatNumber(oCBrandServiceByClaim.FOneItem.FupcheerrSum,0) %></td>
		<td><%= FormatNumber(oCBrandServiceByClaim.FOneItem.FetcupcheerrSum,0) %></td>
		<td></td>
	</tr>
	<% end if %>
</table>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
