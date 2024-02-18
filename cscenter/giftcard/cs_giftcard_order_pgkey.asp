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
// window.resizeTo(600,300);
function SubmitForm() {
	if (validate(frm)==false) {
		return ;
	}

    if (confirm("저장하시겠습니까?") == true) {
        frm.submit();
    }
}

document.title = "PG사 ID";
</script>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <form name="frm" onsubmit="return false;" action="cs_giftcard_order_info_edit_process.asp">
		<input type="hidden" name="mode" value="modipgkey">
		<input type="hidden" name="giftorderserial" value="<%= oGiftOrder.FOneItem.FgiftOrderSerial %>">
		<tr height="25" bgcolor="<%= adminColor("topbar") %>">
			<td colspan="2">
				<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
	    			<tr>
	    				<td width="100">
	    					<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>PG사 ID</b>
						</td>
						<td align="right">
				    		<input type="button" value="수정하기" class="csbutton" onclick="javascript:SubmitForm();">
						</td>
					</tr>
				</table>
			</td>
		</tr>
		<tr height="25" bgcolor="#FFFFFF">
			<td bgcolor="<%= adminColor("topbar") %>">PG사 ID</td>
			<td><input type="text" class="text" name="paydateid" size="50" id="[on,off,1,128][PG사ID]" value="<%= oGiftOrder.FOneItem.Fpaydateid %>"></td>
		</tr>
	</form>
</table>
<%
set oGiftOrder = Nothing
%>
<!-- #include virtual="/cscenter/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
