<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db_TPLOpen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/cscenter/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/3pl/userCls.asp" -->
<!-- #include virtual="/lib/classes/3pl/common.asp" -->
<%

dim userid

userid = requestCheckVar(request("userid"),32)


dim oCTPLUser
set oCTPLUser = New CTPLUser
	oCTPLUser.FRectUserID				= userid

oCTPLUser.GetTPLUserOne

if (userid = "") then
	oCTPLUser.FOneItem.Fuseyn = "Y"
	oCTPLUser.FOneItem.Fregdate = Now()
	oCTPLUser.FOneItem.Flastupdt = Now()
end if

%>
<script language="javascript" SRC="/js/confirm.js"></script>
<script type="text/javascript">

function SubmitForm() {
	var frm = document.frm;

	if (validate(frm)==false) {
		return;
	}

	if (frm.useyn.value == '') {
		alert('��뿩�θ� �����ϼ���.');
		return;
	}

    if (confirm("�����Ͻðڽ��ϱ�?") == true) {
        frm.submit();
    }
}

</script>

<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" onsubmit="return false;" action="user_process.asp">
<input type="hidden" name="mode" value="<%= CHKIIF(userid<>"", "modi", "ins") %>">

<tr height="25" bgcolor="<%= adminColor("topbar") %>">
    <td colspan="2">
        <table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
    		<tr>
    			<td width="300">
    				<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>����� ����</b>
			    </td>
			    <td align="right">
			    	<input type="button" value="�����ϱ�" class="csbutton" onclick="javascript:SubmitForm();">
			    </td>
			</tr>
		</table>
    </td>
</tr>
<tr height="25" bgcolor="#FFFFFF">
    <td bgcolor="<%= adminColor("topbar") %>">����</td>
    <td>
		<% Call SelectBoxCompanyID("companyid", oCTPLUser.FOneItem.Fcompanyid, CHKIIF(userid<>"", "", "Y")) %>
	</td>
</tr>
<tr height="25" bgcolor="#FFFFFF">
    <td bgcolor="<%= adminColor("topbar") %>">�α��ξ��̵�</td>
    <td>
		<% if (userid = "") then %>
		<input type="text" class="text" name="userid" id="[on,off,4,32][���̵�]" value="<%= oCTPLUser.FOneItem.Fuserid %>">
		<% else %>
		<%= oCTPLUser.FOneItem.Fuserid %>
		<input type="hidden" name="userid" value="<%= userid %>">
		<% end if %>
	</td>
</tr>
<tr height="25" bgcolor="#FFFFFF">
    <td bgcolor="<%= adminColor("topbar") %>">����ڸ�</td>
    <td>
		<input type="text" class="text" name="username" id="[on,off,1,16][����ڸ�]" value="<%= oCTPLUser.FOneItem.Fusername %>">
	</td>
</tr>
<tr height="25" bgcolor="#FFFFFF">
    <td bgcolor="<%= adminColor("topbar") %>">��뿩��</td>
    <td>
		<% Call drawSelectBoxUsingYN("useyn", oCTPLUser.FOneItem.Fuseyn) %>
	</td>
</tr>
<tr height="25" bgcolor="#FFFFFF">
    <td bgcolor="<%= adminColor("topbar") %>">�����</td>
    <td>
		<%= oCTPLUser.FOneItem.Fregdate %>
	</td>
</tr>
<tr height="25" bgcolor="#FFFFFF">
    <td bgcolor="<%= adminColor("topbar") %>">��������</td>
    <td>
		<%= oCTPLUser.FOneItem.Flastupdt %>
	</td>
</tr>
</table>

<%
set oCTPLUser = Nothing
%>
<!-- #include virtual="/cscenter/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/db_TPLClose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
