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

  
�� �ش��̺�Ʈ�� ��뿩��Y , �̺�Ʈ������ �ϰ�츸 ����˴ϴ�. ������ ����� ����ϼŵ� �ڵ����� ������� �ʽ��ϴ�. 
<table width="100%" border="0" align="center" class="a" cellpadding="0" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" action="/admin/diary2009/event_process.asp" mothod="get">
	<input type="hidden" name="mode" value="new">
	<input type="hidden" name="event_type" value="organizer">	
	<tr bgcolor="FFFFFF" align="center">
		<td>�̺�Ʈ�ڵ�</td>
		<td>�������</td>
		<td>��뿩��</td>
	</tr>
	
		
	<tr bgcolor="FFFFFF" align="center">
		<td><input type="text" name="evt_code" value=""></td>
		<td><input type="text" name="idx_order" value=""></td>
		<td>
			<select name="isusing">
				<option>�����ϼ���</option>
				<option value="Y">Y</option>
				<option value="N" checked>N</option>				
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