<%@ language=vbscript %>
<% option explicit %>
<%
session.codePage = 949
Response.CharSet = "EUC-KR"
%>
<%
'###########################################################
' Description : cs����
' History : 2009.04.17 �̻� ����
'			2016.06.30 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/cscenter/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/order/new_ordercls.asp"-->

<%
dim ojumun, orderserial, AlertMsg, IsOldOrder, ix
	orderserial = requestCheckVar(request("orderserial"),11)

set ojumun = new COrderMaster
	ojumun.FRectOrderSerial = orderserial
	ojumun.QuickSearchOrderMaster

	if (ojumun.FResultCount<1) and (Len(orderserial)=11) and (IsNumeric(orderserial)) then
	    ojumun.FRectOldOrder = "on"
	    ojumun.QuickSearchOrderMaster

	    if (ojumun.FResultCount>0) then
	        IsOldOrder = true
	        AlertMsg = "6���� ���� �ֹ��Դϴ�."
	    end if

	end if

%>
<script language="javascript" SRC="/js/confirm.js"></script>
<script type="text/javascript">

window.resizeTo(800,600);

function SubmitForm() {
	if (validate(frm)==false) {
		return ;
	}

    if (document.frm.reqphone.value == '--') {
        document.frm.reqphone.value = '000-000-0000';
    }

    if (document.frm.reqhp.value == '--') {
        document.frm.reqhp.value = '000-000-0000';
    }

    if (confirm("�����Ͻðڽ��ϱ�?") == true) {
        frm.submit();
    }
}

function CopyZip(frmname, post1, post2, addr, dong) {
    eval(frmname + ".reqzipcode").value = post1 + "-" + post2;

    eval(frmname + ".reqzipaddr").value = addr;
    eval(frmname + ".reqaddress").value = dong;
}

document.title = "��� ����";

</script>

<form name="frm" onsubmit="return false;" action="order_info_edit_process.asp" style="margin:0px;">
<input type="hidden" name="mode" value="modifyreceiverinfo">
<input type="hidden" name="orderserial" value="<%= ojumun.FOneItem.FOrderSerial %>">
<input type="hidden" name="acctdiv" value="<%=ojumun.FOneItem.FAccountDiv%>">
<input type="hidden" name="paygatetid" value="<%=ojumun.FOneItem.Fpaygatetid%>">
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="<%= adminColor("topbar") %>">
    <td colspan="2">
        <table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
    		<tr>
    			<td width="100">
    				<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>��� ����</b>
			    </td>
			    <td align="right">
			    	<input type="button" value="�����ϱ�" class="csbutton" onclick="javascript:SubmitForm();" >
			    </td>
			</tr>
		</table>
    </td>
</tr>
<tr height="25" bgcolor="#FFFFFF">
    <td bgcolor="<%= adminColor("topbar") %>">�����θ�</td>
    <td><input type="text" class="text" name="reqname" id="[on,off,1,32][�����θ�]" value="<%= ojumun.FOneItem.FReqName %>"></td>
</tr>
<tr height="25" bgcolor="#FFFFFF">
    <td bgcolor="<%= adminColor("topbar") %>">��ȭ��ȣ</td>
    <td><input type="text" class="text" name="reqphone" id="[on,off,1,24][��ȭ��ȣ]" value="<%= ojumun.FOneItem.FReqPhone %>"></td>
</tr>
<tr height="25" bgcolor="#FFFFFF">
    <td bgcolor="<%= adminColor("topbar") %>">�ڵ���</td>
    <td><input type="text" class="text" name="reqhp" id="[on,off,1,16][�ڵ���]" value="<%= ojumun.FOneItem.FReqHp %>"></td>
</tr>
<tr height="25" bgcolor="#FFFFFF">
    <td rowspan="3" valign="top" bgcolor="<%= adminColor("topbar") %>">�����ּ�</td>
    <td>
        <input type="text" class="text" name="reqzipcode" value="<%= ojumun.FOneItem.FReqZipCode %>" size="7" readonly><!-- id="[on,off,7,7][�����ȣ]" -->
        <input type="button" class="button" value="�˻�" onClick="FnFindZipNew('frm','A')">
        <input type="button" class="button" value="�˻�(��)" onClick="TnFindZipNew('frm','A')">
    </td>
</tr>
<tr height="25" bgcolor="#FFFFFF">
    <td ><input type="text" class="text" name="reqzipaddr" id="[on,off,1,64][�ּ�]" size="35" value="<%= ojumun.FOneItem.FReqZipAddr %>"></td>
</tr>
<tr height="25" bgcolor="#FFFFFF">
    <td>
        <input type="text" class="text" name="reqaddress" id="[on,off,1,200][�ּ�]" size="35" value="<%= ojumun.FOneItem.FReqAddress %>">
    </td>
</tr>
<tr height="25" bgcolor="#FFFFFF">
    <td bgcolor="<%= adminColor("topbar") %>">��Ÿ����</td>
    <td>
        <textarea class="textarea" rows="3" cols="35" name="comment" id="[off,off,off,off][��Ÿ����]"><%= ojumun.FOneItem.FComment %></textarea>
	</td>
</tr>
</table>
</form>

<script type="text/javascript">
    <% if (AlertMsg<>"") then %>
        alert('<%= AlertMsg %>');
    <% end if %>
</script>

<%
set ojumun = Nothing
%>
<!-- #include virtual="/cscenter/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
