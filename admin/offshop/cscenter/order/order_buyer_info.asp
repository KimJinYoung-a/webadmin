<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �������� ������
' Hieditor : 2011.03.07 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/offshop/cscenter/popheader_cs_off.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/order/order_cls.asp"-->
<!-- #include virtual="/admin/offshop/cscenter/cscenter_Function_off.asp"-->
<%
dim ojumun, masteridx ,ix
	masteridx= requestCheckVar(request("masteridx"),16)

set ojumun = new COrder
    ojumun.FRectmasteridx = masteridx
    ojumun.fQuickSearchOrderMaster
%>

<script language="javascript" SRC="/js/confirm.js"></script>
<script language="javascript">

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
<form name="frm" onsubmit="return false;" action="/admin/offshop/cscenter/order/order_process.asp">
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
    <td bgcolor="<%= adminColor("topbar") %>">�ϷĹ�ȣ</td>
    <td bgcolor="#FFFFFF"><input type="text" class="text_ro" name="masteridx" id="[off,off,off,off][�ϷĹ�ȣ]" value="<%= ojumun.FOneItem.fmasteridx %>" readonly></td>
</tr>	
<tr height="25">
    <td bgcolor="<%= adminColor("topbar") %>">�ֹ���ȣ</td>
    <td bgcolor="#FFFFFF"><input type="text" class="text_ro" name="orderno" id="[off,off,off,off][�ֹ���ȣ]" value="<%= ojumun.FOneItem.forderno %>" readonly></td>
</tr>
<tr height="25">
    <td bgcolor="<%= adminColor("topbar") %>">�����ڸ�</td>
    <td bgcolor="#FFFFFF">
        <input type="text" class="text" name="buyname" id="[on,off,1,32][�����ڸ�]" value="<%= ojumun.FOneItem.FBuyName %>" size="8">        
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
</form>
</table>
<!-- ���������� -->

<%
set ojumun = Nothing
%>
<!-- #include virtual="/admin/offshop/cscenter/poptail_cs_off.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->