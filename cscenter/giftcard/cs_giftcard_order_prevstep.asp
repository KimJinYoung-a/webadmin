<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/cscenter/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_giftcard_ordercls.asp" -->
<%

dim giftorderserial
giftorderserial = RequestCheckVar(request("giftorderserial"),11)

'==============================================================================
dim oGiftOrder

set oGiftOrder = new cGiftCardOrder

if (giftorderserial <> "") then
	oGiftOrder.FRectGiftOrderSerial = giftorderserial

	oGiftOrder.getCSGiftcardOrderDetail
end if


dim IsAdminLogin
IsAdminLogin = C_ADMIN_AUTH



'==============================================================================
dim nextjumundiv, nextipkumdiv, prevjumundiv, previpkumdiv

if (oGiftOrder.FOneItem.Fjumundiv = "7") then
	'전송완료전환
	prevjumundiv = "5"
end if



%>
<script language='javascript'>
function SubmitForm() {
	<% if (IsAdminLogin <> True) then %>
		alert("시스템팀 전용메뉴입니다.");
		return;
	<% end if %>

	alert("작업중");
	return;

    if (frm.jumundiv.value != '7') {
        alert('이전단계로 전환할 수 없습니다.');
        return;
    }

    if (frm.cancelyn.value!='N'){
        alert('취소된 내역은 이전 상태로 전환이 불가능 합니다.');
        return;
    }

    if (confirm("이전 상태로 전환 하시겠습니까?") == true) {
        frm.submit();
    }
}
</script>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" onsubmit="return false;" action="cs_giftcard_order_info_edit_process.asp">
<input type="hidden" name="mode" value="jumundivprevstep">
<input type="hidden" name="giftorderserial" value="<%= oGiftOrder.FOneItem.FgiftOrderSerial %>">
<input type="hidden" name="jumundiv" value="<%= oGiftOrder.FOneItem.Fjumundiv %>">
<input type="hidden" name="ipkumdiv" value="<%= oGiftOrder.FOneItem.Fipkumdiv %>">
<input type="hidden" name="userid" value="<%= oGiftOrder.FOneItem.Fuserid %>">
<input type="hidden" name="cancelyn" value="<%= oGiftOrder.FOneItem.Fcancelyn %>">
    <tr height="25" bgcolor="<%= adminColor("topbar") %>">
	    <td colspan="2">
	        <table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
	    		<tr>
	    			<td width="160">
	    				<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>이전 상태 진행</b>
				    </td>
				    <td align="right">
				    	<input type="button" value="이전 상태 전환" class="csbutton" onclick="javascript:SubmitForm();">
				    </td>
				</tr>
			</table>
	    </td>
	</tr>
    <tr height="25" align="center" bgcolor="<%= adminColor("topbar") %>">
        <td width="150" >현재상태</td>
        <td width="150" >다음상태</td>
    </tr>
    <tr height="25" align="center" bgcolor="#FFFFFF">
        <td width="150"><font color="<%= oGiftOrder.FOneItem.IpkumDivColor %>"><%= oGiftOrder.FOneItem.GetJumunDivName %></font></td>
        <%
        nextjumundiv = oGiftOrder.FOneItem.Fjumundiv
        nextipkumdiv = oGiftOrder.FOneItem.Fipkumdiv
        oGiftOrder.FOneItem.Fjumundiv = prevjumundiv
        oGiftOrder.FOneItem.Fipkumdiv = previpkumdiv
        %>
        <td width="150"><font color="<%= oGiftOrder.FOneItem.IpkumDivColor %>"><%= oGiftOrder.FOneItem.GetJumunDivName %></font></td>
        <%
        oGiftOrder.FOneItem.Fjumundiv = nextjumundiv
        oGiftOrder.FOneItem.Fipkumdiv = nextipkumdiv
        %>
    </tr>
    <tr height="25" align="center" bgcolor="#FFFFFF">
        <td colspan="2"><%= oGiftOrder.FOneItem.GetAccountdivName %> <font color="<%= oGiftOrder.FOneItem.CancelYnColor %>"><%= oGiftOrder.FOneItem.CancelYnName %></font></td>
    </tr>
</table>
<%
set oGiftOrder = Nothing
%>
<!-- #include virtual="/cscenter/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->