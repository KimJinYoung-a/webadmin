<%@ language=vbscript %>
<% option explicit %>
<%
'############### 2008�� 11�� 4�� �ѿ�� ����
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/diary2009/classes/DiaryCls.asp"-->


�� �ش��̺�Ʈ�� ��뿩��Y , �̺�Ʈ������ �ϰ�츸 ����˴ϴ�. ������ ����� ����ϼŵ� �ڵ����� ������� �ʽ��ϴ�.
<table width="100%" border="0" align="center" class="a" cellpadding="0" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" action="/admin/diary2009/event_process.asp" mothod="get">
	<input type="hidden" name="mode" value="new">
	<input type="hidden" name="event_type" value="diary">
	<tr bgcolor="FFFFFF" align="center">
		<td>
			��ũŸ��
		</td>
		<td>
			<select name="event_link">
				<option value="event" checked>event</option>
				<option value="item">item</option>
			</select><br><br>
			<font color="green"><b>item�� ���ý�</b></font> �ϴ�<font color="green"><b>��ǰ�ڵ�</b></font> ��ǰ�������� �̵�<br>
			<font color="red"><b>event�� ���ý�</b></font> �ϴ�<font color="red"><b>�̺�Ʈ�ڵ�</b></font> �̺�Ʈ�������� �̵�
		</td>
	</tr>
	<tr bgcolor="FFFFFF" align="center">
		<td>�̺�Ʈ�ڵ�</td>
		<td><input type="text" name="evt_code" value=""></td>
	</tr>
	<tr bgcolor="FFFFFF" align="center">
		<td>��ǰ�ڵ�</td>
		<td><input type="text" name="itemid" value=""></td>
	</tr>
	<tr bgcolor="FFFFFF" align="center">
		<td>����</td>
		<td><% SelectList "cate","" %></td>
	</tr>
	<tr bgcolor="FFFFFF" align="center">
		<td>�������</td>
		<td><input type="text" name="idx_order" value="">�⺻��0</td>
	</tr>
	<tr bgcolor="FFFFFF" align="center">
		<td>��뿩��</td>
		<td>
			<select name="isusing">
				<option value="N" checked>N</option>
				<option value="Y">Y</option>
			</select>
		</td>
	</tr>
	<tr bgcolor="FFFFFF" align="left">
		<td colspan=3><input type="button" class="button" value="����" onclick="javascript:frm.submit();"></td>
	</tr>
</form>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->


