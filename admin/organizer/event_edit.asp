<%@ language=vbscript %>
<% option explicit %>
<%
'############### 2008�� 11�� 6�� �ѿ�� ����
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/organizer/organizer_cls.asp"-->
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
set oip = new organizerCls
	oip.frectidx = idx
	oip.geteventone
%>
  
�� �ش��̺�Ʈ�� ��뿩��Y , �̺�Ʈ������ �ϰ�츸 ����˴ϴ�. ������ ����� ����ϼŵ� �ڵ����� ������� �ʽ��ϴ�. 
<table width="100%" border="0" align="center" class="a" cellpadding="0" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" action="/admin/organizer/event_process.asp" mothod="get">
	<input type="hidden" name="mode" value="edit">
	<input type="hidden" name="idx" value="<%= oip.fitem.fidx %>">
	<input type="hidden" name="event_type" value="<%= oip.fitem.fevent_type %>">	
	<tr bgcolor="FFFFFF" align="center">
		<td>�̺�Ʈ�ڵ�</td>
		<td>�������</td>
		<td>��뿩��</td>
	</tr>
	
		
	<tr bgcolor="FFFFFF" align="center">
		<td><input type="text" name="evt_code" value="<%= oip.fitem.fevt_code %>"></td>
		<td><input type="text" name="idx_order" value="<%= oip.fitem.fidx_order %>"></td>
		<td>
			<select name="isusing">
				<option>�����ϼ���</option>
				<option value="Y" <% if oip.fitem.fisusing = "Y" then response.write " selected"%>>Y</option>
				<option value="N" <% if oip.fitem.fisusing = "N" then response.write " selected"%>>N</option>				
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