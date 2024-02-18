<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/company/incSessionCompany.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/company/jikbang/jikbangCls.asp" -->

<%
''not using
response.end

Dim ojumun
Set ojumun = new CJumunMaster
ojumun.FRectOrderSerial = requestCheckvar(request("orderserial"),16)

if (ojumun.FRectOrderSerial<>"") then
    ojumun.GetOnejikbangJumunMaster
end if

Dim ix

if (ojumun.FResultCount<1) then
    response.write "해당 내역이 없습니다."
    dbget.close
    response.end
end if

ojumun.SearchJumunDetail ojumun.FRectOrderSerial

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<script language="JavaScript" src="/js/xl.js"></script>
<script language="JavaScript" src="/js/common.js"></script>
<script language="JavaScript" src="/js/report.js"></script>
<link rel="stylesheet" href="/css/scm.css" type="text/css">
</head>
<body>

<table width="70%" border="0" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
<tr height="25" bgcolor="#FFFFFF">
	<td width="220" bgcolor="#FFD8D8" width="100">주문번호</td>
	<td ><%= ojumun.FOneItem.FOrderSerial %></td>
	<td width="220" bgcolor="#FFD8D8" width="100">매출코드</td>
	<td ><%= ojumun.FOneItem.getRdSiteName %><%=CHKIIF(ojumun.FOneItem.isMobileOrder,"(모바일)","") %></td>
</tr>

<tr height="25" bgcolor="#FFFFFF">
	<td bgcolor="#FFD8D8" width="100">주문금액</td>
	<td ><%= FormatNumber(ojumun.FOneItem.FTotalSum,0) %></td>
	<td bgcolor="#FFD8D8" width="100"></td>
	<td ></td>
</tr>
<tr height="25" bgcolor="#FFFFFF">
    <td bgcolor="#FFD8D8" width="100">할인금액</td>
	<td ><%= FormatNumber(ojumun.FOneItem.getEnuiSum,0) %></td>
	<td bgcolor="#FFD8D8" width="100">배송비</td>
	<td ><%= FormatNumber(ojumun.FOneItem.getDlvPaySum,0) %></td>
</tr>
<tr height="25" bgcolor="#FFFFFF">
	<td bgcolor="#FFD8D8" width="100">정산대상액 (VAT포함)</td>
	<td ><%= FormatNumber(ojumun.FOneItem.FreducedpriceSum,0) %></td>
	<td bgcolor="#FFD8D8" width="100">정산대상액 (VAT제외)</td>
	<td ><%= FormatNumber(ojumun.FOneItem.getJungsanTargetNoVatSum,0) %></td>
</tr>
<tr height="25" bgcolor="#FFFFFF">
	<td bgcolor="#FFD8D8" width="100">주문일</td>
	<td ><%= ojumun.FOneItem.FRegDate %></td>
	<td bgcolor="#FFD8D8" width="100">결제일</td>
	<td ><%= ojumun.FOneItem.FIpkumDate %></td>
</tr>
<tr height="25" bgcolor="#FFFFFF">
	<td bgcolor="#FFD8D8" width="100">취소일</td>
	<td >
	<% if ojumun.FOneItem.FCancelYn<>"N" then %>
	<%= ojumun.FOneItem.FCanceldate %>
	<% end if %>
	</td>
	<td bgcolor="#FFD8D8" width="100">정산일</td>
	<td ><%= ojumun.FOneItem.getJungsanFixdate %></td>
</tr>
<tr height="25" bgcolor="#FFFFFF">
	<td bgcolor="#FFD8D8" width="100">주문상태</td>
	<td >
	<% if ojumun.FOneItem.FCancelYn<>"N" then %>
	<%= ojumun.FOneItem.IpkumDivName %> 상태중 취소
	<% else %>
	<%= ojumun.FOneItem.IpkumDivName %>
	<% end if %>
	</td>
	<td bgcolor="#FFD8D8" width="100"></td>
	<td ></td>
</tr>
</table>


<p>
<table width="70%" border="0" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
<tr bgcolor="#FFD8D8" height="25">
	<td width="50" align="center">상품코드</td>
	<td width="50" align="center">이미지</td>
	<td width="100" align="center">상품명</td>
	<td width="50" align="center">수량</td>
	<td width="100" align="center">옵션명</td>
	<td width="70" align="center">판매가</td>
	<td width="70" align="center">%쿠폰할인가</td>
	<td width="70" align="center">과세구분</td>
	<td width="70" align="center">상태</td>
</tr>
<%
	For ix=0 to ojumun.FJumunDetail.FDetailCount - 1
%>
<tr bgcolor="#FFFFFF" >
    <% if ojumun.FJumunDetail.FJumunDetailList(ix).Fitemid=0 then %>
    <td align="center">-</td>
	<td align="center">-</td>
	<td align="center"><%= ojumun.FJumunDetail.FJumunDetailList(ix).FItemName %>&nbsp;</td>
	<td align="center"><%= ojumun.FJumunDetail.FJumunDetailList(ix).FItemNo %>&nbsp;</td>
	<td align="center"><%= ojumun.FJumunDetail.FJumunDetailList(ix).FItemoptionName %>&nbsp;</td>
	<td align="right"><%= FormatNumber(ojumun.FJumunDetail.FJumunDetailList(ix).Fitemcost,0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(ojumun.FJumunDetail.FJumunDetailList(ix).FreducedPrice,0) %>&nbsp;</td>
	<td align="center">&nbsp;</td>
	<td align="center">
	<% if ojumun.FJumunDetail.FJumunDetailList(ix).FCancelyn="Y" then %>
	취소
	<% end if %>
	</td>
    <% else %>
	<td align="center"><%= ojumun.FJumunDetail.FJumunDetailList(ix).Fitemid %>&nbsp;</td>
	<td align="center"><img src="<%= ojumun.FJumunDetail.FJumunDetailList(ix).FImageSmall %>" border="0"></td>
	<td align="center"><%= ojumun.FJumunDetail.FJumunDetailList(ix).FItemName %>&nbsp;</td>
	<td align="center"><%= ojumun.FJumunDetail.FJumunDetailList(ix).FItemNo %>&nbsp;</td>
	<td align="center"><%= ojumun.FJumunDetail.FJumunDetailList(ix).FItemoptionName %>&nbsp;</td>
	<td align="right"><%= FormatNumber(ojumun.FJumunDetail.FJumunDetailList(ix).Fitemcost,0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(ojumun.FJumunDetail.FJumunDetailList(ix).FreducedPrice,0) %>&nbsp;</td>
	<td align="center"><%= ChkIIF(ojumun.FJumunDetail.FJumunDetailList(ix).Fvatinclude="Y","과세","비과세") %>&nbsp;</td>
	<td align="center">
	<% if ojumun.FJumunDetail.FJumunDetailList(ix).FCancelyn="Y" then %>
	취소
	<% end if %>
	</td>
	<% end if %>
</tr>
<%
	Next
%>
</table>
<% Set ojumun = Nothing %>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->