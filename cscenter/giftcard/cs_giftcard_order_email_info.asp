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

dim ix, i
dim tmpvalue, tmpselected

%>
<script language="javascript" SRC="/js/confirm.js"></script>
<script>
function SubmitForm() {
	if (frm.sendDiv[1].checked != true) {
		if (validate(frm)==false) {
			return ;
		}
	}

    if (confirm("�����Ͻðڽ��ϱ�?") == true) {
        frm.submit();
    }
}

document.title = "�����̸��� ����";
</script>


<!-- ���������� -->
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <form name="frm" onsubmit="return false;" action="cs_giftcard_order_info_edit_process.asp">
    <input type="hidden" name="mode" value="modifyemailinfo">
    <input type="hidden" name="giftorderserial" value="<%= oGiftOrder.FOneItem.FgiftOrderSerial %>">
	<tr height="25" bgcolor="<%= adminColor("topbar") %>">
	    <td colspan="2">
	        <table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
	    		<tr>
	    			<td width="200">
	    				<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>�����̸��� ����</b>
				    </td>
				    <td align="right">
				        <input type="button" value="�����ϱ�" class="csbutton" onClick="SubmitForm();">
				    </td>
				</tr>
			</table>
	    </td>
	</tr>
	<tr height="25">
	    <td bgcolor="<%= adminColor("topbar") %>">���ۿ���</td>
	    <td bgcolor="#FFFFFF">
	    	<input type="radio" name="sendDiv" value="E" <% if (oGiftOrder.FOneItem.FsendDiv = "E") then %>checked<% end if %>> ��������
	    	<input type="radio" name="sendDiv" value="S" <% if (oGiftOrder.FOneItem.FsendDiv <> "E") then %>checked<% end if %>> �߼۾���
	    </td>
	</tr>
	<tr height="25">
	    <td bgcolor="<%= adminColor("topbar") %>">�����º�Email</td>
	    <td bgcolor="#FFFFFF"><input type="text" class="text" name="sendemail" id="[on,off,1,100][�����º�Email]" value="<%= oGiftOrder.FOneItem.Fsendemail %>"></td>
	</tr>
	<tr height="25">
	    <td bgcolor="<%= adminColor("topbar") %>">�޴º�Email</td>
	    <td bgcolor="#FFFFFF"><input type="text" class="text" name="reqEmail" id="[on,off,1,100][�޴º�Email]" value="<%= oGiftOrder.FOneItem.FreqEmail %>"></td>
	</tr>
	<tr height="25">
	    <td bgcolor="<%= adminColor("topbar") %>">Email����</td>
	    <td bgcolor="#FFFFFF">
	    	<input type="text" class="text" name="emailTitle" id="[on,off,1,64][Email����]" size="50" value="<%= oGiftOrder.FOneItem.FemailTitle %>">
	    </td>
	</tr>
	<tr height="25">
	    <td bgcolor="<%= adminColor("topbar") %>">Email����</td>
	    <td bgcolor="#FFFFFF">
	    	<textarea name="emailContent" cols=45 rows=9 id="[on,off,1,3000][Email����]"><%= oGiftOrder.FOneItem.FemailContent %></textarea>
	    </td>
	</tr>
	</form>
</table>
<!-- ���������� -->



<%
set oGiftOrder = Nothing
%>
<!-- #include virtual="/cscenter/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->