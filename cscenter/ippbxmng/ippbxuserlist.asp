<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ������ IPP ����ڿ���
' Hieditor : 2015.05.27 �̻� ����
'			 2021.04.09 �ѿ�� ����(�ƿ��ҽ� ��Ź��ü ���� �߰�)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/checkAllowIPWithLog.asp" -->
<!-- #include virtual="/lib/classes/cscenter/cs_ippbxusercls.asp" -->
<%
dim userid, i, occscenterippbxuser

set occscenterippbxuser = new CCSCenterIppbxUser
	occscenterippbxuser.FPageSize = 50
	occscenterippbxuser.FCurrPage = 1
	occscenterippbxuser.GetCSCenterIppbxUserList

%>
<script type='text/javascript'>

function ModifyIppbxInfo(frm){
	if (confirm("�����Ͻðڽ��ϱ�?") == true) {
		frm.submit();
	}
}

</script>

<!-- �˻� ���� -->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		���̵� : <input type="text" class="text" name="userid" value="<%= userid %>">
	</td>

	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onclick="document.frm.submit()">
	</td>
</tr>
</table>
</form>

<br>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="10">
		�˻���� : <b><%= occscenterippbxuser.FTotalCount %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
	<td >������ȣ</td>
	<td >���ξ��̵�</td>
	<td >��뿩��</td>
	<td >������</td>
	<td>���</td>
</tr>
<% if occscenterippbxuser.FTotalCount > 0 then %>
	<% for i = 0 to (occscenterippbxuser.FResultCount - 1) %>
    <tr align="center" bgcolor="#FFFFFF" height="25">
    	<form name="frm<%= i %>" method="post" action="/cscenter/ippbxmng/ippbxuser_process.asp" style="margin:0px;">
    	<input type="hidden" class="text" name="menupos" value="<%= menupos %>">
    	<input type="hidden" class="text" name="localcallno" value="<%= occscenterippbxuser.FItemList(i).Flocalcallno %>">
        <td><%= occscenterippbxuser.FItemList(i).Flocalcallno %></td>
        <td><input type="text" class="text" name="userid" value="<%= occscenterippbxuser.FItemList(i).Fuserid %>"></td>
        <td>
			<select name="useyn" class="select">
				<option value="Y" <% if (occscenterippbxuser.FItemList(i).Fuseyn = "Y") then %>selected<% end if %>>�����
				<option value="N" <% if (occscenterippbxuser.FItemList(i).Fuseyn = "N") then %>selected<% end if %>>������
			</select>
        </td>
        <td><%= occscenterippbxuser.FItemList(i).Flastupdate %></td>
        <td>
			<%
			' ��Ź��ü ���� ����� ��Ź��ü �Ϲ������� �������� ����.
			if C_CSUser or C_ADMIN_AUTH then
			%>
				<% if C_CSOutsourcingUser then %>
					<% if C_CSOutsourcingPowerUser then %>
						<input type="button" class="button" value="����" onClick="ModifyIppbxInfo(frm<%= i %>)">
					<% end if %>
				<% else %>	
					<input type="button" class="button" value="����" onClick="ModifyIppbxInfo(frm<%= i %>)">
				<% end if %>
			<% end if %>
        </td>
        </form>
    </tr>
	<% next %>
<% else %>
    <tr bgcolor="#FFFFFF" align="center">
        <td height="25" colspan="10">�˻������ �����ϴ�.</td>
    </tr>
<% end if %>
</table>

<%
set occscenterippbxuser = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->