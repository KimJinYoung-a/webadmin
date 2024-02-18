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

function PopSearchZipcode(frmname) {
	var popwin = window.open("/lib/searchzip3.asp?target=" + frmname,"PopSearchZipcode","width=460 height=240 scrollbars=yes resizable=yes");
	popwin.focus();
}

function CopyZip(frmname, post1, post2, addr, dong) {
    eval(frmname + ".reqzipcode").value = post1 + "-" + post2;

    eval(frmname + ".reqzipaddr").value = addr;
    eval(frmname + ".reqaddress").value = dong;
}

document.title = "수령자 정보";
</script>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <form name="frm" onsubmit="return false;" action="cs_giftcard_order_info_edit_process.asp">
    <input type="hidden" name="mode" value="modifyreceiverinfo">
    <input type="hidden" name="giftorderserial" value="<%= oGiftOrder.FOneItem.FgiftOrderSerial %>">
	<tr height="25" bgcolor="<%= adminColor("topbar") %>">
	    <td colspan="2">
	        <table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
	    		<tr>
	    			<td width="100">
	    				<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>수령자 정보</b>
				    </td>
				    <td align="right">
				    	<input type="button" value="저장하기" class="csbutton" onclick="javascript:SubmitForm();">
				    </td>
				</tr>
			</table>
	    </td>
	</tr>
	<tr height="25" bgcolor="#FFFFFF">
	    <td bgcolor="<%= adminColor("topbar") %>">핸드폰</td>
	    <td><input type="text" class="text" name="reqhp" id="[on,off,1,16][핸드폰]" value="<%= oGiftOrder.FOneItem.FReqHp %>"></td>
	</tr>
	<tr height="25" bgcolor="#FFFFFF">
	    <td bgcolor="<%= adminColor("topbar") %>">이메일</td>
	    <td><input type="text" class="text" name="reqemail" id="[on,off,1,128][이메일]" value="<%= oGiftOrder.FOneItem.FReqEmail %>"></td>
	</tr>
</table>
<%
set oGiftOrder = Nothing
%>
<!-- #include virtual="/cscenter/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->