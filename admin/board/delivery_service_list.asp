<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  �ù��ü����
' History : 2007.10.29 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->

<Script language="javascript">
function chkEditFrm(frm){
	if (frm.divname.value==''){
		alert('��۾�ü�� �Է��� �ּ���');
		frm.divname.focus();

	}
	frm.submit();

}
function chkAddFrm(frm){
	if (frm.divcd.value==''){

		alert('��ȣ�� �Է��� �ּ���');
		frm.divcd.focus();
		return false;

	}

	if (eval('document.editFrm_' + frm.divcd.value)!=null) {
		alert('�ߺ��� ��ȣ�� ����Ҽ� �����ϴ�');
		return false;
	}


	if (frm.divname.value==''){
		alert('��۾�ü�� �Է��� �ּ���');
		frm.divname.focus();
		return false;
	}

	frm.submit();
}
</script>

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="90" align="center">��ȣ</td>
		<td width="150" align="center">��۾�ü</td>
		<td width="150" align="center">��ǥ��ȭ��ȣ</td>
		<td align="center">�����ȸ URL</td>
		<td align="center">��ǰ���� URL</td>
		<td width="80" align="center">�������</td>
		<td width="80" align="center">10X10���</td>
		<td width="50" align="center">����</td>
	</tr>

<%
dim sql
sql = " SELECT divcd,divname,findurl, returnURL, isUsing, isTenUsing,tel " &_
			" FROM db_order.[dbo].tbl_songjang_div " &_
			" ORDER BY isTenUsing desc ,divcd "

rsget.open sql,dbget,1

if not (rsget.eof or rsget.bof) then
	do until rsget.eof

	dim defBgColor
	if rsget("isUsing") = "Y" then
		if rsget("isTenUsing") ="Y" then
		defBgColor 	=	"#FFFFFF"
		else
		defBgColor	=	"#FFFFFF"
		end if

	else
		defBgColor="#CCCCCC"
	end if
	%>
	<tr align="center" bgcolor="<%= defBgColor %>">
	<form name="editFrm_<%= rsget("divcd") %>" method="post" target="subFrame" action="delivery_service_process.asp">
	<input type="hidden" name="mode" value="edit" />
	<input type="hidden" name="divcd" value="<%= rsget("divcd") %>" />
		<td><%= rsget("divcd") %></td>
		<td><input type="text" class="text" name="divname" size="15" value="<%= db2html(rsget("divname")) %>"></td>
		<td><input type="text" class="text" name="tel" size="15" value="<%= db2html(rsget("tel")) %>">
		</td>
		<td align="left"><input type="text" class="text" name="findurl" size="70" value="<%= db2html(rsget("findurl")) %>"></td>
		<td align="left"><input type="text" class="text" name="returnURL" size="70" value="<%= db2html(rsget("returnURL")) %>"></td>
		<td>
			<select class="select" name="isusing">
				<option value="Y" <% if rsget("isUsing") = "Y" then response.write "selected" %>>����� </option>
				<option value="N" <% if rsget("isUsing") = "N" then response.write "selected" %>>������</option>
			</select>
		</td>
		<td>
			<select class="select" name="isTenUsing">
				<option value="Y" <% if rsget("isTenUsing") = "Y" then response.write "selected" %>>����� </option>
				<option value="N" <% if rsget("isTenUsing") = "N" then response.write "selected" %>>������</option>
			</select>
		</td>
		<td align="center">
			<input type="button" class="button" value="����" onclick="chkEditFrm(this.form);">
			</td>
	</form>
	</tr>

<%
	rsget.movenext
	loop
end if

rsget.close
%>

</table>
<br />
<!-- �ű��Է� ���̺� -->
<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			<b>�ű� �Է�</b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="60" align="center">��ȣ</td>
		<td width="150" align="center">��۾�ü</td>
		<td width="150" align="center">��ǥ��ȭ��ȣ</td>
		<td align="center">�����ȸ URL</td>
		<td align="center">��ǰ���� URL</td>
		<td width="50" align="center"></td>
	</tr>
	<form name="addFrm" method="post" target="subFrame" action="delivery_service_process.asp">
	<input type="hidden" name="mode" value="add" />
	<tr bgcolor="#FFFFFF">
		<td align="center"><input type="text" name="divcd" value="" size="4" style="border:1px solid #CCCCCC;" /></td>
		<td align="center"><input type="text" name="divname" size="15" value="" style="border:1px solid #CCCCCC;" /></td>
		<td align="center"><input type="text" name="tel" size="15" value="" style="border:1px solid #CCCCCC;" /></td>
		<td align="left"><input type="text" name="findurl" size="70" value="" style="border:1px solid #CCCCCC;" /></td>
		<td align="left"><input type="text" name="returnURL" size="70" value="" style="border:1px solid #CCCCCC;" /></td>
		<td align="center"><input type="button" class="button" value="�Է�" onclick="chkAddFrm(this.form);"></td>
	</tr>
	</form>
</table>
<iframe src="" name="subFrame" frameborder="0" width="0" height="0"></iframe>
<br /><br /><br /><br /><br /><br />

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
