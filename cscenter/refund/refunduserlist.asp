<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : [CS]���� ȯ�ұ��� ���� 
' History : �̻� ����
'			2020.04.08 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/checkAllowIPWithLog.asp" -->
<!-- #include virtual="/lib/classes/cscenter/cs_refundusercls.asp" -->
<%
dim userid, i

'==============================================================================
dim occscenterrefunduser
set occscenterrefunduser = new CCSCenterRefundUser

occscenterrefunduser.FPageSize = 50
occscenterrefunduser.FCurrPage = 1

occscenterrefunduser.GetCSCenterRefundUserList


if not(C_CSPowerUser or C_ADMIN_AUTH) then
	response.write "������ ��Ʈ�� �̻� ������ �ʿ��� �Ŵ� �Դϴ�."
	response.end
end if

dim IsSystemPsn	: IsSystemPsn = False
if (session("ssAdminPsn") = 7) then
	IsSystemPsn = True
end if

%>
<script language='javascript'>

function ModifyRefundUserInfo(frm)
{
	if ((frm.userid.value == "") && (frm.useyn.value == "Y")) {
		alert("���̵� �����ϼ���.\n\n�Ǵ� ���������� �����ϼ���.");
		return;
	}

	if ((frm.defaultCSRefundLimit.value == "") || (frm.defaultCSRefundLimit.value*0 != 0)) {
		alert("�߸��� ȯ�ұݾ��Դϴ�.");
		return;
	}

	if (confirm("�����Ͻðڽ��ϱ�?") == true) {
		frm.submit();
	}
}

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			<!--
			���̵� : <input type="text" class="text" name="userid" value="<%= userid %>">
			-->
		</td>

		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
          	<input type="button" class="button_s" value="�˻�" onclick="document.frm.submit()">
		</td>
	</tr>
	</form>
</table>
<br>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
        <td width="150">���̵�</td>
        <td width="120">�⺻ȯ�ұ���</td>
        <td width="90">���</td>
        <td width="180">������</td>
        <td>���</td>
    </tr>
<% if occscenterrefunduser.FTotalCount > 0 then %>
	<% for i = 0 to (occscenterrefunduser.FResultCount - 1) %>
    	<% if (occscenterrefunduser.FItemList(i).Fuseyn = "N") then %>
    <tr align="center" bgcolor="#DDDDDD" height="25">
    	<% else %>
    <tr align="center" bgcolor="#FFFFFF" height="25">
    	<% end if %>
    	<form name="frm<%= i %>" method="post" action="refunduser_process.asp">
    	<input type="hidden" class="text" name="menupos" value="<%= menupos %>">
    	<input type="hidden" name="mode" value="modify">
    	<input type="hidden" class="text" name="idx" value="<%= occscenterrefunduser.FItemList(i).Fidx %>">
        <td>
			<input type="text" class="text" name="userid" value="<%= occscenterrefunduser.FItemList(i).Fuserid %>" size="16">
		</td>
        <td>
			<input type="text" class="text" name="defaultCSRefundLimit" value="<%= occscenterrefunduser.FItemList(i).FdefaultCSRefundLimit %>" size="10">
        </td>
        <td>
			<select name="useyn" class="select">
				<option value="Y" <% if (occscenterrefunduser.FItemList(i).Fuseyn = "Y") then %>selected<% end if %>>�����
				<option value="N" <% if (occscenterrefunduser.FItemList(i).Fuseyn = "N") then %>selected<% end if %>>������
			</select>
        </td>
        <td><%= occscenterrefunduser.FItemList(i).Flastupdate %></td>
        <td align="left">
        	&nbsp;
        	<input type="button" class="button" value="����" onClick="ModifyRefundUserInfo(frm<%= i %>)">
        </td>
        </form>
    </tr>
	<% next %>
<% else %>
    <tr bgcolor="#FFFFFF" align="center">
        <td height="25" colspan="13">�˻������ �����ϴ�.</td>
    </tr>
<% end if %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="13" align="center">
		</td>
	</tr>
</table>

<%
set occscenterrefunduser = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
