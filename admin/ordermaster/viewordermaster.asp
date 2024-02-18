<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->

<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/jumuncls.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%
'' #include virtual="/lib/checkAllowIPWithLog.asp" 주석처리 ithinkso
dim ojumun
set ojumun = new CJumunMaster

ojumun.FRectOrderSerial = requestCheckvar(request("orderserial"),20)
if (ojumun.FRectOrderSerial<>"") then
    ojumun.SearchJumunList
end if

if (ojumun.FResultCount<1) then
    dbget.close() : response.end
end if

dim ix
%>
<script language='javascript'>
function popOrderDetailEdit(idx){
	var popwin = window.open('/common/orderdetailedit.asp?idx=' + idx,'orderdetailedit','width=600,height=480,scrollbars=yes,resizable=yes');
	popwin.focus();
}
</script>
<table width="100%" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
<tr bgcolor="#FFFFFF">
  <td bgcolor="#22AAAA" width="100">주문번호</td>
  <td bgcolor="#DDDDDD" width="200"><%= ojumun.FMasterItemList(0).FOrderSerial %></td>
  <td bgcolor="#22AAAA" width="100">사이트</td>
  <td bgcolor="#DDDDDD" width="200"><%= ojumun.FMasterItemList(0).FSitename %></td>
</tr>
<tr bgcolor="#FFFFFF">
  <td bgcolor="#22AAAA" width="100">결제방식</td>
  <td bgcolor="#DDDDDD" width="200"><%= ojumun.FMasterItemList(0).JumunMethodName %></td>
  <td bgcolor="#22AAAA" width="100">주문상태</td>
  <td bgcolor="#DDDDDD" width="200"><font color="<%= ojumun.FMasterItemList(0).IpkumDivColor %>"><%= ojumun.FMasterItemList(0).IpkumDivName %></font></td>
</tr>
<tr bgcolor="#FFFFFF">
  <td bgcolor="#22AAAA" width="100">할인율</td>
  <td bgcolor="#DDDDDD" width="200"><%= ojumun.FMasterItemList(0).FDiscountRate %></td>
  <td bgcolor="#22AAAA" width="100">취소여부</td>
  <td bgcolor="#DDDDDD" width="200"><font color="<%= ojumun.FMasterItemList(0).CancelYnColor %>"><%= ojumun.FMasterItemList(0).CancelYnName %></font></td>
</tr>
<tr bgcolor="#FFFFFF">
  <td bgcolor="#22AAAA" width="100">결제금액</td>
  <td bgcolor="#DDDDDD" width="200"><%= FormatNumber(ojumun.FMasterItemList(0).FSubTotalPrice,0) %></td>
  <td bgcolor="#22AAAA" width="100">주문금액</td>
  <td bgcolor="#DDDDDD" width="200"><%= FormatNumber(ojumun.FMasterItemList(0).FTotalSum,0) %></td>
</tr>
<tr bgcolor="#FFFFFF">
  <td bgcolor="#22AAAA" width="100">주문일</td>
  <td bgcolor="#DDDDDD" width="200"><%= ojumun.FMasterItemList(0).FRegDate %></td>
  <td bgcolor="#22AAAA" width="100">입금일</td>
  <td bgcolor="#DDDDDD" width="200"><%= ojumun.FMasterItemList(0).FIpkumDate %></td>
</tr>
<tr bgcolor="#FFFFFF">
  <td bgcolor="#22AAAA" width="100">구매자</td>
  <td bgcolor="#DDDDDD" width="200"><%= ojumun.FMasterItemList(0).FBuyName %></td>
  <td bgcolor="#22AAAA" width="100">수령인</td>
  <td bgcolor="#DDDDDD" width="200"><%= ojumun.FMasterItemList(0).FReqName %></td>
</tr>
<% If ojumun.FMasterItemList(0).FSitename = "cnglob10x10" Then %>
<tr bgcolor="#FFFFFF">
  <td bgcolor="#22AAAA" width="100">제휴 주문번호</td>
  <td bgcolor="#DDDDDD" colspan="3"><%= ojumun.FMasterItemList(0).Fauthcode %></td>
</tr>
<% End If %>

<% if (FALSE) then %> <% ''2015/09/22 주석처리 ithinso 에서 이페이지 조회 개인정보 뺌 %>
    <tr bgcolor="#FFFFFF">
  <td bgcolor="#22AAAA" width="100">구매자ID</td>
  <td bgcolor="#DDDDDD" width="200"><%= ojumun.FMasterItemList(0).FUserID %></td>
  <td bgcolor="#22AAAA" width="100">구매자</td>
  <td bgcolor="#DDDDDD" width="200"><%= ojumun.FMasterItemList(0).FBuyName %></td>
</tr>
<tr bgcolor="#FFFFFF">
  <td bgcolor="#22AAAA" width="100">구매자전화</td>
  <td bgcolor="#DDDDDD" width="200"><%= ojumun.FMasterItemList(0).FBuyPhone %></td>
  <td bgcolor="#22AAAA" width="100">구매자핸드폰</td>
  <td bgcolor="#DDDDDD" width="200"><%= ojumun.FMasterItemList(0).FBuyHp %></td>
</tr>
<tr bgcolor="#FFFFFF">
  <td bgcolor="#22AAAA" width="100">구매자이메일</td>
  <td bgcolor="#DDDDDD" width="200"><a href="mailto:<%= ojumun.FMasterItemList(0).FBuyEmail %>" class="zzz"><%= ojumun.FMasterItemList(0).FBuyEmail %></a></td>
  <td bgcolor="#22AAAA" width="100">입금자</td>
  <td bgcolor="#DDDDDD" width="200"><%= ojumun.FMasterItemList(0).FAccountName %></td>
</tr>
<tr bgcolor="#FFFFFF">
  <td bgcolor="#22AAAA" width="100">수령인</td>
  <td bgcolor="#DDDDDD" width="200"><%= ojumun.FMasterItemList(0).FReqName %></td>
  <td bgcolor="#22AAAA" width="100"></td>
  <td bgcolor="#DDDDDD" width="200"></td>
</tr>
<tr bgcolor="#FFFFFF">
  <td bgcolor="#22AAAA" width="100">수령인전화</td>
  <td bgcolor="#DDDDDD" width="200"><%= ojumun.FMasterItemList(0).FReqPhone %></td>
  <td bgcolor="#22AAAA" width="100">수령인핸드폰</td>
  <td bgcolor="#DDDDDD" width="200"><%= ojumun.FMasterItemList(0).FReqHp %></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#22AAAA" width="100">수령인주소</td>
	<td bgcolor="#DDDDDD" colspan="3">
		ojumun.FMasterItemList(0).FReqZipCode
		<br>
		<%= ojumun.FMasterItemList(0).FReqZipAddr %>
		&nbsp;<%= ojumun.FMasterItemList(0).FReqAddress %>
	</td>
</tr>

<tr bgcolor="#FFFFFF">
  <td bgcolor="#22AAAA" width="100">기타사항</td>
  <td bgcolor="#DDDDDD" colspan="3">
  <%= ojumun.FMasterItemList(0).FComment %>
  </td>
</tr>
<tr bgcolor="#FFFFFF">
  <td bgcolor="#22AAAA" width="100">사용마일리지</td>
  <td bgcolor="#DDDDDD" width="200"><%= ojumun.FMasterItemList(0).FMileTotalPrice %></td>
  <td bgcolor="#22AAAA" width="100">송장번호</td>
  <td bgcolor="#DDDDDD" width="200"></td>
</tr>
<tr bgcolor="#FFFFFF">
  <td bgcolor="#22AAAA" width="100">카드승인번호</td>
  <td bgcolor="#DDDDDD" width="200"><%= ojumun.FMasterItemList(0).FAuthcode %></td>
  <td bgcolor="#22AAAA" width="100">카드결과</td>
  <td bgcolor="#DDDDDD" width="200"><%= ojumun.FMasterItemList(0).FResultmsg %></td>
</tr>
<tr bgcolor="#FFFFFF">
  <td bgcolor="#22AAAA" width="100">Inicis-ID</td>
  <td bgcolor="#DDDDDD" width="200"><%= ojumun.FMasterItemList(0).FPaygatetID %></td>
  <td bgcolor="#22AAAA" width="100">정산유무</td>
  <td bgcolor="#DDDDDD" width="200"><%= ojumun.FMasterItemList(0).Fjungsanflag %></td>
</tr>
<% end if %>
</table>
<%
ojumun.SearchJumunDetail ojumun.FRectOrderSerial
%>
<br><br>
<table width="100%" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
<!--
<tr bgcolor="#FFFFFF">
	<td width="100">배송옵션</td>
	<td width="200"><%= ojumun.FJumunDetail.BeasongOptionStr %></td>
</tr>
-->
<tr bgcolor="#FFFFFF">
	<td>배송비</td>
	<td align="right"><%= FormatNumber(ojumun.FJumunDetail.BeasongPay,0) %></td>
</tr>
</table>

<table width="100%" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
<tr bgcolor="#FFFFFF">
	<td width="50" align="center">상품ID</td>
	<td width="50" align="center">이미지</td>
	<td width="100" align="center">상품명</td>
	<td width="50" align="center">수량</td>
	<td width="70" align="center">옵션Code</td>
	<td width="100" align="center">옵션명</td>
<% If (session("ssBctDiv") <= 9) Then %>
	<td width="100" align="center">매입가</td>
<% End If %>
	<td width="70" align="center">Price</td>
	<td width="70" align="center">취소삭제</td>
	<td width="70" align="center">상태</td>
	<td width="70" align="center">송장구분</td>
	<td width="70" align="center">송장번호</td>
	<td width="70" align="center">배송일</td>
	<td width="70" align="center">업체배송</td>
	<td width="70" align="center">세일상품</td>
	<% if C_ADMIN_AUTH then %>
	<td width="40" align="center">수정</td>
	<% end if %>
</tr>
<% for ix=0 to ojumun.FJumunDetail.FDetailCount-1 %>
<% if ojumun.FJumunDetail.FJumunDetailList(ix).Fitemid <>0 then %>
<tr bgcolor="#FFFFFF">
	<td align="center"><%= ojumun.FJumunDetail.FJumunDetailList(ix).Fitemid %></td>
	<td align="center"><a href="http://www.10x10.co.kr/street/designershop.asp?itemid=<%= ojumun.FJumunDetail.FJumunDetailList(ix).Fitemid %>" target="_blank"><img src="<%= ojumun.FJumunDetail.FJumunDetailList(ix).FImageSmall %>" border="0"></a></td>
	<td align="center"><%= ojumun.FJumunDetail.FJumunDetailList(ix).FItemName %></td>
	<td align="center"><%= ojumun.FJumunDetail.FJumunDetailList(ix).FItemNo %></td>
	<td align="center"><%= ojumun.FJumunDetail.FJumunDetailList(ix).FItemOption %></td>
	<td align="center"><%= ojumun.FJumunDetail.FJumunDetailList(ix).FItemoptionName %></td>
<% If (session("ssBctDiv") <= 9) Then %>
	<td align="right"><%= FormatNumber(ojumun.FJumunDetail.FJumunDetailList(ix).Fbuycash,0) %></td>
<% End If %>
	<td align="right"><%= FormatNumber(ojumun.FJumunDetail.FJumunDetailList(ix).Fitemcost,0) %></td>
	<td align="center"><%= ojumun.FJumunDetail.FJumunDetailList(ix).CancelStateStr %></td>
	<td align="center"><%= ojumun.FJumunDetail.FJumunDetailList(ix).Fcurrstate %></td>
	<td align="center"><%= DeliverDivCd2Nm(ojumun.FJumunDetail.FJumunDetailList(ix).Fsongjangdiv) %></td>
	<td align="center"><%= ojumun.FJumunDetail.FJumunDetailList(ix).Fsongjangno %></td>
	<td align="center"><%= ojumun.FJumunDetail.FJumunDetailList(ix).Fbeasongdate %></td>
	<td align="center"><%= ojumun.FJumunDetail.FJumunDetailList(ix).Fisupchebeasong %></td>
	<td align="center"><%= ojumun.FJumunDetail.FJumunDetailList(ix).Fissailitem %></td>
	<% if C_ADMIN_AUTH then %>
	<td align="center"><input type="button" value="수정" onclick="popOrderDetailEdit('<%= ojumun.FJumunDetail.FJumunDetailList(ix).Fdetailidx %>');">
	<% end if %>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan="16"><%= ojumun.FJumunDetail.FJumunDetailList(ix).Frequiredetail %></td>
</tr>
<% end if %>
<% next %>
</table>
<%
set ojumun = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
