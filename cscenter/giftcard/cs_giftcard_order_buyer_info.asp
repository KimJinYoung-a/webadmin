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

    if (confirm("�����Ͻðڽ��ϱ�?") == true) {
        frm.submit();
    }
}

document.title = "����������";
</script>


<!-- ���������� -->
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <form name="frm" onsubmit="return false;" action="cs_giftcard_order_info_edit_process.asp">
    <input type="hidden" name="mode" value="modifybuyerinfo">
	<tr height="25" bgcolor="<%= adminColor("topbar") %>">
	    <td colspan="2">
	        <table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
	    		<tr>
	    			<td width="100">
	    				<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>������ ����</b>
				    </td>
				    <td align="right">
				        <input type="button" value="�����ϱ�" class="csbutton" onClick="SubmitForm();">
				    </td>
				</tr>
			</table>
	    </td>
	</tr>
	<tr height="25">
	    <td bgcolor="<%= adminColor("topbar") %>">������ID</td>
	    <td bgcolor="#FFFFFF"><input type="text" class="text_ro" name="userid" id="[off,off,off,off][������ID]" value="<%= oGiftOrder.FOneItem.FUserID %>" readonly></td>
	</tr>
	<tr height="25">
	    <td bgcolor="<%= adminColor("topbar") %>">�ֹ���ȣ</td>
	    <td bgcolor="#FFFFFF"><input type="text" class="text_ro" name="giftorderserial" id="[off,off,off,off][�ֹ���ȣ]" value="<%= oGiftOrder.FOneItem.FgiftOrderSerial %>" readonly></td>
	</tr>
	<tr height="25">
	    <td bgcolor="<%= adminColor("topbar") %>">�����ڸ�</td>
	    <td bgcolor="#FFFFFF">
	        <input type="text" class="text" name="buyname" id="[on,off,1,16][�����ڸ�]" value="<%= oGiftOrder.FOneItem.FBuyName %>" size="8" >
	        <font color="<%= oGiftOrder.FOneItem.GetUserLevelColor %>"><%= oGiftOrder.FOneItem.GetUserLevelName %></a></font>
	    </td>
	</tr>
	<tr height="25">
	    <td bgcolor="<%= adminColor("topbar") %>">��ȭ��ȣ</td>
	    <td bgcolor="#FFFFFF"><input type="text" class="text" name="buyphone" id="[on,off,1,16][��������ȭ��ȣ]" value="<%= oGiftOrder.FOneItem.FBuyPhone %>" ></td>
	</tr>
	<tr height="25">
	    <td bgcolor="<%= adminColor("topbar") %>">�ڵ���</td>
	    <td bgcolor="#FFFFFF"><input type="text" class="text" name="buyhp" id="[on,off,1,16][�������ڵ���]" value="<%= oGiftOrder.FOneItem.FBuyHp %>" ></td>
	</tr>
	<tr height="25">
	    <td bgcolor="<%= adminColor("topbar") %>">�̸���</td>
	    <td bgcolor="#FFFFFF"><input type="text" class="text" name="buyemail" id="[on,off,1,128][�̸���]" value="<%= oGiftOrder.FOneItem.FBuyEmail %>" ></td>
	</tr>
	<tr height="25">
	    <td bgcolor="<%= adminColor("topbar") %>">������</td>
	    <td bgcolor="#FFFFFF"><%= oGiftOrder.FOneItem.GetAccountdivName %> / <font color="<%= oGiftOrder.FOneItem.IpkumDivColor %>"><%= oGiftOrder.FOneItem.GetIpkumDivName %></font></td>
	</tr>
	<tr height="25">
	    <td bgcolor="<%= adminColor("topbar") %>">�Ա��ڸ�</td>
	    <td bgcolor="#FFFFFF"><input type="text" class="text" name="accountname" id="[on,off,1,16][�Ա��ڸ�]" value="<%= oGiftOrder.FOneItem.Fipkumname %>" ></td>
	</tr>
	</form>
</table>
<!-- ���������� -->



<%
set oGiftOrder = Nothing
%>
<!-- #include virtual="/cscenter/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->