<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �������� ������
' Hieditor : 2011.03.08 �ѿ�� ����
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
dim ojumun, masteridx, AlertMsg, IsOldOrder , ix
	masteridx = requestCheckVar(request("masteridx"),10)

set ojumun = new COrder
	ojumun.FRectmasteridx = masteridx
	ojumun.fQuickSearchOrderMaster
%>

<script language="javascript" SRC="/js/confirm.js"></script>
<script type="text/javascript">

window.resizeTo(600,400);

function SubmitForm() {
	if (validate(frm)==false) {
		return ;
	}

    if (confirm("�����Ͻðڽ��ϱ�?") == true) {
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

document.title = "��� ����";

</script>

<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" onsubmit="return false;" action="/admin/offshop/cscenter/order/order_process.asp">
<input type="hidden" name="mode" value="modifyreceiverinfo">
<input type="hidden" name="masteridx" value="<%= ojumun.FOneItem.Fmasteridx %>">
<input type="hidden" name="orderno" value="<%= ojumun.FOneItem.forderno %>">
<tr height="25" bgcolor="<%= adminColor("topbar") %>">
    <td colspan="2">
        <table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
    		<tr>
    			<td width="100">
    				<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>��� ����</b>
			    </td>    				    
			    <td align="right">
			    	<input type="button" value="�����ϱ�" class="csbutton" onclick="javascript:SubmitForm();" <%= chkIIF(IsOldOrder,"disabled","") %>>
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
    <td><input type="text" class="text" name="reqphone" id="[on,off,1,16][��ȭ��ȣ]" value="<%= ojumun.FOneItem.FReqPhone %>"></td>
</tr>
<tr height="25" bgcolor="#FFFFFF">
    <td bgcolor="<%= adminColor("topbar") %>">�ڵ���</td>
    <td><input type="text" class="text" name="reqhp" id="[on,off,1,16][�ڵ���]" value="<%= ojumun.FOneItem.FReqHp %>"></td>
</tr>
<tr height="25" bgcolor="#FFFFFF">
    <td rowspan="3" valign="top" bgcolor="<%= adminColor("topbar") %>">�����ּ�</td>
    <td>
        <input type="text" class="text" name="reqzipcode" value="<%= ojumun.FOneItem.FReqZipCode %>" size="7"  readonly><!-- id="[on,off,7,7][�����ȣ]" -->
        <input type="button" class="button" value="�˻�" onClick="FnFindZipNew('frm','A')">
        <input type="button" class="button" value="�˻�(��)" onClick="TnFindZipNew('frm','A')">
        <% '<input type="button" class="button" value="�˻�(��)" onClick="PopSearchZipcode('frm')"> %>
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

<script language='javascript'>
    <% if (AlertMsg<>"") then %>
        alert('<%= AlertMsg %>');
    <% end if %>
</script>  

<%
set ojumun = Nothing
%>
<!-- #include virtual="/cscenter/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->