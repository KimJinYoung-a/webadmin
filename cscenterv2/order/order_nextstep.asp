<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/cscenterv2/lib/incSessionAdminCS.asp" -->
<!-- #include virtual="/cscenterv2/lib/db/dbopen.asp" -->
<!-- #include virtual="/cscenterv2/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/cscenterv2/lib/classes/order/ordercls.asp"-->
<%
dim ojumun
set ojumun = new COrderMaster
ojumun.FRectOrderSerial = requestCheckvar(request("orderserial"),16)
ojumun.QuickSearchOrderMaster

dim ix, preipkumdiv

%>
<script language='javascript'>
function SubmitForm() {
    //if ((frm.ipkumdiv.value!='2')&&(frm.ipkumdiv.value!='5')&&(frm.ipkumdiv.value!='6')){
    if ((frm.ipkumdiv.value!='2')){
        alert('�ֹ� ���������� ���� ���·� ������ ���� �մϴ�.');
        return;
    }

    if (frm.cancelyn.value!='N'){
        alert('��ҵ� ������ ���� ���·� ������ �Ұ��� �մϴ�.');
        return;
    }

    if (confirm("���� ���·� ���� �Ͻðڽ��ϱ�?") == true) {
        frm.submit();
    }
}
</script>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" onsubmit="return false;" action="order_info_edit_process.asp">
<input type="hidden" name="mode" value="ipkumdivnextstep">
<input type="hidden" name="orderserial" value="<%= ojumun.FOneItem.Forderserial %>">
<input type="hidden" name="ipkumdiv" value="<%= ojumun.FOneItem.Fipkumdiv %>">
<input type="hidden" name="userid" value="<%= ojumun.FOneItem.Fuserid %>">
<input type="hidden" name="cancelyn" value="<%= ojumun.FOneItem.Fcancelyn %>">
    <tr height="25" bgcolor="<%= adminColor("topbar") %>">
	    <td colspan="2">
	        <table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
	    		<tr>
	    			<td width="160">
	    				<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>���� ���� ����</b>
				    </td>
				    <td align="right">
				    	<input type="button" value="���� ���� ����" class="csbutton" onclick="javascript:SubmitForm();">
				    </td>
				</tr>
			</table>
	    </td>
	</tr>
    <tr height="25" align="center" bgcolor="<%= adminColor("topbar") %>">
        <td width="150" >�������</td>
        <td width="150" >��������</td>
    </tr>
    <tr height="25" align="center" bgcolor="#FFFFFF">
        <td width="150"><font color="<%= ojumun.FOneItem.IpkumDivColor %>"><%= ojumun.FOneItem.IpkumDivName %></font></td>
        <%
        preipkumdiv = ojumun.FOneItem.Fipkumdiv
        if (preipkumdiv="2") then
            ojumun.FOneItem.Fipkumdiv = "4"
        elseif (preipkumdiv="4") then
            ojumun.FOneItem.Fipkumdiv = "5"
        elseif (preipkumdiv="5") then
            ojumun.FOneItem.Fipkumdiv = "6"
        elseif (preipkumdiv="6") then
            ojumun.FOneItem.Fipkumdiv = "7"
        end if
        %>
        <td width="150"><font color="<%= ojumun.FOneItem.IpkumDivColor %>"><%= ojumun.FOneItem.IpkumDivName %></font></td>
        <%
        ojumun.FOneItem.Fipkumdiv = preipkumdiv
        %>
    </tr>
    <tr height="25" align="center" bgcolor="#FFFFFF">
        <td colspan="2"><%= ojumun.FOneItem.JumunMethodName %> <font color="<%= ojumun.FOneItem.CancelYnColor %>"><%= ojumun.FOneItem.CancelYnName %></font></td>
    </tr>
    <% if (ojumun.FOneItem.Fipkumdiv="2") and (ojumun.FOneItem.Fcancelyn="N") then %>
    <tr height="25" align="center" bgcolor="#FFFFFF">
        <td><input type="checkbox" name="emailok" checked >�Ա�Ȯ�θ��Ϲ߼�</td>
        <td ><%= ojumun.FOneItem.Fbuyemail %>
    </tr>
    <tr height="25" align="center" bgcolor="#FFFFFF">
        <td><input type="checkbox" name="smsok" checked >�Ա�Ȯ��SMS�߼�</td>
        <td><%= ojumun.FOneItem.Fbuyhp %></td>
    </tr>
    <% end if %>
</table>
<%
set ojumun = Nothing
%>
<!-- #include virtual="/cscenter/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->