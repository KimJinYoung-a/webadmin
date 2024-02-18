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

dim ix

%>
<script language="javascript" SRC="/js/confirm.js"></script>
<script>
function SubmitForm() {
	if (validate(frm)==false) {
		return ;
	}

    if (confirm("저장하시겠습니까?") == true) {
        frm.submit();
    }
}

document.title = "구매자정보";
</script>


<!-- 구매자정보 -->
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <form name="frm" onsubmit="return false;" action="cs_giftcard_order_info_edit_process.asp">
    <input type="hidden" name="mode" value="modifybuyerinfo">
	<tr height="25" bgcolor="<%= adminColor("topbar") %>">
	    <td colspan="2">
	        <table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
	    		<tr>
	    			<td width="100">
	    				<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>구매자 정보</b>
				    </td>
				    <td align="right">
				        <input type="button" value="저장하기" class="csbutton" onClick="SubmitForm();">
				    </td>
				</tr>
			</table>
	    </td>
	</tr>
	<tr height="25">
	    <td bgcolor="<%= adminColor("topbar") %>">구매자ID</td>
	    <td bgcolor="#FFFFFF"><input type="text" class="text_ro" name="userid" id="[off,off,off,off][구매자ID]" value="<%= oGiftOrder.FOneItem.FUserID %>" readonly></td>
	</tr>
	<tr height="25">
	    <td bgcolor="<%= adminColor("topbar") %>">주문번호</td>
	    <td bgcolor="#FFFFFF"><input type="text" class="text_ro" name="giftorderserial" id="[off,off,off,off][주문번호]" value="<%= oGiftOrder.FOneItem.FgiftOrderSerial %>" readonly></td>
	</tr>
	<tr height="25">
	    <td bgcolor="<%= adminColor("topbar") %>">구매자명</td>
	    <td bgcolor="#FFFFFF">
	        <input type="text" class="text" name="buyname" id="[on,off,1,16][구매자명]" value="<%= oGiftOrder.FOneItem.FBuyName %>" size="8" >
	        <font color="<%= oGiftOrder.FOneItem.GetUserLevelColor %>"><%= oGiftOrder.FOneItem.GetUserLevelName %></a></font>
	    </td>
	</tr>
	<tr height="25">
	    <td bgcolor="<%= adminColor("topbar") %>">전화번호</td>
	    <td bgcolor="#FFFFFF"><input type="text" class="text" name="buyphone" id="[on,off,1,16][구매자전화번호]" value="<%= oGiftOrder.FOneItem.FBuyPhone %>" ></td>
	</tr>
	<tr height="25">
	    <td bgcolor="<%= adminColor("topbar") %>">핸드폰</td>
	    <td bgcolor="#FFFFFF"><input type="text" class="text" name="buyhp" id="[on,off,1,16][구매자핸드폰]" value="<%= oGiftOrder.FOneItem.FBuyHp %>" ></td>
	</tr>
	<tr height="25">
	    <td bgcolor="<%= adminColor("topbar") %>">이메일</td>
	    <td bgcolor="#FFFFFF"><input type="text" class="text" name="buyemail" id="[on,off,1,128][이메일]" value="<%= oGiftOrder.FOneItem.FBuyEmail %>" ></td>
	</tr>
	<tr height="25">
	    <td bgcolor="<%= adminColor("topbar") %>">결재방법</td>
	    <td bgcolor="#FFFFFF"><%= oGiftOrder.FOneItem.GetAccountdivName %> / <font color="<%= oGiftOrder.FOneItem.IpkumDivColor %>"><%= oGiftOrder.FOneItem.GetIpkumDivName %></font></td>
	</tr>
	<tr height="25">
	    <td bgcolor="<%= adminColor("topbar") %>">입금자명</td>
	    <td bgcolor="#FFFFFF"><input type="text" class="text" name="accountname" id="[on,off,1,16][입금자명]" value="<%= oGiftOrder.FOneItem.Fipkumname %>" ></td>
	</tr>
	</form>
</table>
<!-- 구매자정보 -->



<%
set oGiftOrder = Nothing
%>
<!-- #include virtual="/cscenter/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->