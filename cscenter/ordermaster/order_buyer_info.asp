<%@ language=vbscript %>
<% option explicit %>
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
dim ojumun, orderserial, ix
dim AlertMsg, IsOldOrder
	orderserial= requestCheckVar(request("orderserial"),11)

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
<form name="frm" onsubmit="return false;" action="order_info_edit_process.asp">
<input type="hidden" name="mode" value="modifybuyerinfo">
<tr height="25" bgcolor="<%= adminColor("topbar") %>">
    <td colspan="2">
        <table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
    		<tr>
    			<td width="100">
    				<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>������ ����</b>
			    </td>    				    
			    <td align="right">
			        <input type="button" value="�����ϱ�" class="csbutton" onClick="SubmitForm();" <%= chkIIF(IsOldOrder,"disabled","") %>>
			    </td>
			</tr>
		</table>
    </td>
</tr>
<tr height="25">
    <td bgcolor="<%= adminColor("topbar") %>">������ID....</td>
    <td bgcolor="#FFFFFF">
    	<input type="text" class="text_ro" name="userid" id="[off,off,off,off][������ID]" value="<%= ojumun.FOneItem.FUserID %>" readonly>
    </td>
</tr>
<tr height="25">
    <td bgcolor="<%= adminColor("topbar") %>">�ֹ���ȣ</td>
    <td bgcolor="#FFFFFF">
    	<input type="text" class="text_ro" name="orderserial" id="[off,off,off,off][�ֹ���ȣ]" value="<%= ojumun.FOneItem.FOrderSerial %>" readonly>
    </td>
</tr>
<tr height="25">
    <td bgcolor="<%= adminColor("topbar") %>">�����ڸ�</td>
    <td bgcolor="#FFFFFF">
        <input type="text" class="text" name="buyname" id="[on,off,1,32][�����ڸ�]" value="<%= ojumun.FOneItem.FBuyName %>" size="8" >
        <font color="<%= getUserLevelColorByDate(ojumun.FOneItem.fUserLevel, left(ojumun.FOneItem.Fregdate,10)) %>">
        <%= getUserLevelStrByDate(ojumun.FOneItem.fUserLevel, left(ojumun.FOneItem.Fregdate,10)) %></a></font>
    </td>
</tr>
<tr height="25">
    <td bgcolor="<%= adminColor("topbar") %>">��ȭ��ȣ</td>
    <td bgcolor="#FFFFFF"><input type="text" class="text" name="buyphone" id="[on,off,1,16][��������ȭ��ȣ]" value="<%= ojumun.FOneItem.FBuyPhone %>" ></td>
</tr>
<tr height="25">
    <td bgcolor="<%= adminColor("topbar") %>">�ڵ���</td>
    <td bgcolor="#FFFFFF"><input type="text" class="text" name="buyhp" id="[on,off,1,16][�������ڵ���]" value="<%= ojumun.FOneItem.FBuyHp %>" ></td>
</tr>
<tr height="25">
    <td bgcolor="<%= adminColor("topbar") %>">�̸���</td>
    <td bgcolor="#FFFFFF"><input type="text" class="text" name="buyemail" id="[on,off,1,128][�̸���]" value="<%= ojumun.FOneItem.FBuyEmail %>" ></td>
</tr>
<tr height="25">
    <td bgcolor="<%= adminColor("topbar") %>">������</td>
    <td bgcolor="#FFFFFF">
    	<%= ojumun.FOneItem.JumunMethodName %> / <font color="<%= ojumun.FOneItem.IpkumDivColor %>"><%= ojumun.FOneItem.IpkumDivName %></font>
    </td>
</tr>
<tr height="25">
    <td bgcolor="<%= adminColor("topbar") %>">�Ա��ڸ�</td>
    <td bgcolor="#FFFFFF"><input type="text" class="text" name="accountname" id="[on,off,1,16][�Ա��ڸ�]" value="<%= ojumun.FOneItem.FAccountName %>" ></td>
</tr>
</form>
</table>
<!-- ���������� -->

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