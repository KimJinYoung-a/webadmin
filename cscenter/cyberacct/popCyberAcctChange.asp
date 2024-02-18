<% option Explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/order/new_ordercls.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->

<%
dim orderserial : orderserial = requestCheckVar(request("orderserial"),11)

'==============================================================================
dim ojumun
set ojumun = new COrderMaster

if (orderserial <> "") then
    ojumun.FRectOrderSerial = orderserial
    ojumun.QuickSearchOrderMaster
end if


if (ojumun.FResultCount<1) and (Len(orderserial)=11) and (IsNumeric(orderserial)) then
    ojumun.FRectOldOrder = "on"
    ojumun.QuickSearchOrderMaster
end if
'==============================================================================

dim CyberAcctEditEnable , CyberAcct_subtotalPrice, CyberAcct_CLOSEDate
CyberAcctEditEnable = false

dim sqlStr
sqlStr = "select top 1 subtotalPrice, convert(varchar(19),CLOSEDate,20) as CLOSEDate from db_order.dbo.tbl_order_CyberAccountLog"
sqlStr = sqlStr & " where orderserial='" & orderserial & "'"
sqlStr = sqlStr & " order by differencekey desc"

rsget.Open sqlStr,dbget,1
if (Not rsget.Eof) then
    CyberAcctEditEnable = true
    CyberAcct_subtotalPrice = rsget("subtotalPrice")
    CyberAcct_CLOSEDate = rsget("CLOSEDate")
end if
rsget.Close

if (ojumun.FOneItem.FIpkumdiv<>"2") then
    CyberAcctEditEnable = false
end if
%>
<script language='javascript'>
function editCyberAcct(frm){
    if (confirm('수정 하시겠습니까?')){
        frm.submit();
    }
}
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <form name="frmAct" method="post" action="popCyberAcctChange_process.asp">
    <input type="hidden" name="orderserial" value="<%=ojumun.FOneItem.Forderserial%>">
	<tr align="center">
		<td width="70" bgcolor="<%= adminColor("tabletop") %>">주문번호</td>
		<td align="left" bgcolor="#FFFFFF" colspan="2"><%=ojumun.FOneItem.Forderserial%></td>
	</tr>
	<tr align="center">
		<td width="70" bgcolor="<%= adminColor("tabletop") %>">결제방법</td>
		<td align="left" bgcolor="#FFFFFF" colspan="2"><%= ojumun.FOneItem.JumunMethodName %></td>
	</tr>
	<tr align="center">
		<td width="70" bgcolor="<%= adminColor("tabletop") %>">주문상태</td>
		<td align="left" bgcolor="#FFFFFF" colspan="2"><font color="<%= ojumun.FOneItem.IpkumDivColor %>"><%= ojumun.FOneItem.IpkumDivName %></font>
		<% if ojumun.FOneItem.FCancelYn<>"N" then %>
        <font color="<%= ojumun.FOneItem.CancelYnColor %>"><%= ojumun.FOneItem.CancelYnName %></font>
        <% end if %>
		</td>
	</tr>
	<tr align="center">
		<td width="70" bgcolor="<%= adminColor("tabletop") %>">결제계좌</td>
		<td align="left" bgcolor="#FFFFFF" ><%= ojumun.FOneItem.FAccountNo %>
		</td>
		<td align="left" bgcolor="#FFFFFF">
		<% if ojumun.FOneItem.IsDacomCyberAccountPay then %>
		(데이콤 가상계좌)
		<% end if %>
		</td>
	</tr>
	<tr align="center">
		<td width="70" bgcolor="<%= adminColor("tabletop") %>">금액</td>
		<td align="left" bgcolor="#FFFFFF"><%= ojumun.FOneItem.FSubtotalPrice %>
		<% if ojumun.FOneItem.FSubtotalPrice<>CyberAcct_subtotalPrice then %>
		(<font color=red><%= CyberAcct_subtotalPrice %></font>)
		<% end if %>
		</td>
		<td align="left" bgcolor="#FFFFFF">
		<input type="text" class="text" name="subtotalprice" value="<%= ojumun.FOneItem.TotalMajorPaymentPrice %>" size="10" maxlength="10" readOnly>
		</td>
	</tr>
	<tr align="center">
		<td width="70" bgcolor="<%= adminColor("tabletop") %>">입금마감일</td>
		<td align="left" bgcolor="#FFFFFF">
		<%= CyberAcct_CLOSEDate %>
		
		</td>
		<td align="left" bgcolor="#FFFFFF">
		<input type="text" class="text" name="CLOSEDATE" value="<%= Left(CyberAcct_CLOSEDate,10) %>" size="10" maxlength="10" readOnly>
		<input type="text" class="text" name="buf1" value="23:59:59" size="8" maxlength="8" readOnly>
		<a href="javascript:calendarOpen(frmAct.CLOSEDATE);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>
		</td>
	</tr>
	</form>
	<tr>
	    <td colspan="3" bgcolor="#FFFFFF" align="center">
	    <% if (CyberAcctEditEnable) then %>
	    <input type="button" value="수정" onClick="editCyberAcct(frmAct);">
	    <% else %>
	        수정 하실 수 없습니다.
	    <% end if %>
	    </td>
	</tr>
</table>
<%
set ojumun = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/admin/lib/poptail.asp"-->