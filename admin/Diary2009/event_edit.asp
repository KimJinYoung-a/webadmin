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
<%
dim idx
	idx = request("idx")

'//�������
if idx = "" then
	response.write "<script>"
	response.write "alert('idx �Ķ��Ÿ ���̾����ϴ�');"
	response.write "history.go(-1);"
	response.write "</script>"
end if

dim oip
set oip = new DiaryCls
	oip.frectidx = idx
	oip.geteventone
%>

�� �ش��̺�Ʈ�� ��뿩��Y , �̺�Ʈ������ �ϰ�츸 ����˴ϴ�. ������ ����� ����ϼŵ� �ڵ����� ������� �ʽ��ϴ�.
<table width="100%" border="0" align="center" class="a" cellpadding="0" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" action="/admin/diary2009/event_process.asp" mothod="get">
	<input type="hidden" name="mode" value="edit">
	<input type="hidden" name="idx" value="<%= oip.fitem.fidx %>">
	<input type="hidden" name="event_type" value="<%= oip.fitem.fevent_type %>">
	<tr bgcolor="FFFFFF" align="center">
		<td>
			��ũŸ��
		</td>
		<td>
			<select name="event_link">
				<option value="event" <% if oip.fitem.fevent_link = "event" or oip.fitem.fevent_link=""  then response.write " selected"%>>event</option>
				<option value="item" <% if oip.fitem.fevent_link = "item" then response.write " selected"%>>item</option>
			</select><br>
			item�� ���ý� �ϴܻ�ǰ�ڵ� ��ǰ�������� �̵�<br>
			event�� ���ý� �ϴ��̺�Ʈ�ڵ� �̺�Ʈ�������� �̵�
		</td>
	</tr>
	<tr bgcolor="FFFFFF" align="center">
		<td>�̺�Ʈ�ڵ�</td>
		<td><input type="text" name="evt_code" value="<%= oip.fitem.fevt_code %>"></td>
	</tr>
	<tr bgcolor="FFFFFF" align="center">
		<td>��ǰ�ڵ�</td>
		<td><input type="text" name="itemid" value="<%= oip.fitem.fitemid %>"></td>
	</tr>
	<tr bgcolor="FFFFFF" align="center">
		<td>����</td>
		<td><% SelectList "cate", oip.fitem.FCateCode %></td>
	</tr>
	<tr bgcolor="FFFFFF" align="center">
		<td>�������</td>
		<td><input type="text" name="idx_order" value="<%= oip.fitem.fidx_order %>">�⺻��0</td>
	</tr>

	<tr bgcolor="FFFFFF" align="center">
		<td>��뿩��</td>
		<td>
			<select name="isusing">
				<option value="Y" <% if oip.fitem.fisusing = "Y" then response.write " selected"%>>Y</option>
				<option value="N" <% if oip.fitem.fisusing = "N" or oip.fitem.fisusing="" then response.write " selected"%>>N</option>
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


