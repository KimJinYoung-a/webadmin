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

dim iscreatenewcode
iscreatenewcode = RequestCheckVar(request("iscreatenewcode"),11)

'==============================================================================
dim oGiftOrder

set oGiftOrder = new cGiftCardOrder

if (giftorderserial <> "") then
	oGiftOrder.FRectGiftOrderSerial = giftorderserial

	oGiftOrder.getCSGiftcardOrderDetail
end if



dim title
if (iscreatenewcode = "Y") then
	title = "신규인증코드 전송"
else
	title = "기존인증코드 재전송"
end if

%>
<script language='javascript'>
function SubmitForm() {
    if (frm.jumundiv.value != '5') {
        alert('전송완료된 내역만 재전송이 가능합니다.');
        return;
    }

    if (frm.cancelyn.value!='N'){
        alert('취소된 내역은 다음 상태로 진행이 불가능 합니다.');
        return;
    }

    if (confirm("전송 하시겠습니까?") == true) {
        frm.submit();
    }
}
</script>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" onsubmit="return false;" action="cs_giftcard_order_info_edit_process.asp">
<input type="hidden" name="mode" value="resendcardcode">
<input type="hidden" name="iscreatenewcode" value="<%= iscreatenewcode %>">
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
	    				<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b><%= title %></b>
				    </td>
				    <td align="right">
				    	<input type="button" value="전송하기" class="csbutton" onclick="javascript:SubmitForm();">
				    </td>
				</tr>
			</table>
	    </td>
	</tr>
    <tr height="25" align="center" bgcolor="<%= adminColor("topbar") %>">
        <td width="100" >보내는분HP</td>
        <td align="left">
        	<%= oGiftOrder.FOneItem.Fsendhp %>
        </td>
    </tr>
    <tr height="25" align="center" bgcolor="<%= adminColor("topbar") %>">
        <td>받는분HP</td>
        <td align="left">
        	<%= oGiftOrder.FOneItem.Freqhp %>
        </td>
    </tr>
    <tr height="25" align="center" bgcolor="<%= adminColor("topbar") %>">
        <td>이메일전송여부</td>
        <td align="left">
	    	<% if (oGiftOrder.FOneItem.FsendDiv = "E") then %>
	    		동시전송
	    	<% else %>
	    		발송안함
	    	<% end if %>
        </td>
    </tr>
    <tr height="25" align="center" bgcolor="<%= adminColor("topbar") %>">
        <td>받는분Email</td>
        <td align="left">
        	<%= oGiftOrder.FOneItem.FMMSTitle %>
        </td>
    </tr>
    <tr height="25" align="center" bgcolor="<%= adminColor("topbar") %>">
        <td>받는분Email</td>
        <td align="left">
        	<%= nl2br(oGiftOrder.FOneItem.FMMSContent) %>
        </td>
    </tr>
    <tr height="25" align="center" bgcolor="#FFFFFF">
        <td colspan="2">
        	<%= oGiftOrder.FOneItem.GetAccountdivName %>
        	<font color="<%= oGiftOrder.FOneItem.CancelYnColor %>"><%= oGiftOrder.FOneItem.CancelYnName %></font>
        	<font color="<%= oGiftOrder.FOneItem.IpkumDivColor %>"><%= oGiftOrder.FOneItem.GetJumunDivName %></font>
        </td>
    </tr>
</table>
<%
set oGiftOrder = Nothing
%>
<!-- #include virtual="/cscenter/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->