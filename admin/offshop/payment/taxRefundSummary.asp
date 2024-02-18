<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프라인 taxRefund 관리
' History : 2014.01.17 서동석
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/payment/taxRefundMngCls.asp"-->
<%
dim page,shopid,yyyy1,mm1, onlythatdate

shopid = requestCheckvar(request("shopid"),32)
page = requestCheckvar(request("page"),10)
if page="" then page=1
yyyy1 = requestCheckvar(request("yyyy1"),4)
mm1 = requestCheckvar(request("mm1"),2)
onlythatdate = requestCheckvar(request("onlythatdate"),10)


dim oTaxRefund
set oTaxRefund = new CTaxRefund
%>
<!-- 표 상단바 시작-->
<table width="100%" align="center" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>" class="A">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>" rowspan="2">검색<br>조건</td>
	<td align="left">
		<table border="0" width="100%" cellpadding="3" cellspacing="0" class="a">
		<tr>
			<td>
				* 정산월 :
				&nbsp;&nbsp;
                <input type="checkbox" name="onlythatdate" <%=CHKIIF(onlythatdate="on","checked","") %> >해당월만
                &nbsp;&nbsp;
                &nbsp;&nbsp;
				<%
				'직영/가맹점
				if (C_IS_SHOP) then
				%>
					<% if not C_IS_OWN_SHOP and shopid <> "" then %>
						* 매장 : <%=shopid%><input type="hidden" name="shopid" value="<%= shopid %>">
					<% else %>
						* 매장 : <% drawSelectBoxOffShopAll "shopid",shopid %>
					<% end if %>
				<% else %>
					* 매장 : <% drawSelectBoxOffShopAll "shopid",shopid %>
				<% end if %>
			</td>
		</tr>
	    </table>
    </td>
	<td  width="50" bgcolor="<%= adminColor("gray") %>" rowspan="2">
		<input type="button" class="button_s" value="검색" onClick="frm.submit();">
	</td>
</tr>
<tr bgcolor="#FFFFFF" >
    <td>
    * 검색구분 :

    </td>
</tr>

</form>
</table>

<!-- 표 상단바 끝-->
<Br>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		검색결과 : <b><%= oTaxRefund.FTotalCount %></b>
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
	<td>정산월</td>
	<td>구매월</td>
	<td>매장</td>
	<td>카드</td>
	<td>현금</td>
	<td>마일리지</td>
	<td>상품권</td>
	<td>기프트카드</td>
	<td>합계</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
	<td >합계</td>
	<td align="right"></td>
	<td align="right"></td>
	<td align="right"></td>
	<td align="right"></td>
	<td align="right"></td>
	<td align="right"></td>
	<td align="right"></td>
	<td align="right"></td>
</tr>

<tr align="center" bgcolor="#FFFFFF">
	<td colspan="15">검색 결과가 없습니다.</td>
</tr>

</table>
<%
set oTaxRefund=Nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->