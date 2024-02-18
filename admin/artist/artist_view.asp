<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
'	History	:  2009.04.10 �ѿ�� ����
'	Description : artist gallery
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/admin/artistGalleryCls.asp" -->

<%
dim page ,i , idx
	page = request("page")
	idx = request("idx")
	if page="" then page=1

dim oip
	set oip = New Cinquiry_list
	oip.frectidx = idx
	oip.finquiry_oneitem()
%>

<table width="700" align="center" cellpadding="4" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%=adminColor("gray")%>">�۰���</td>
		<td align="center"><%= nl2br(oip.foneitem.fartist_name) %></td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%=adminColor("gray")%>">����</td>
		<td align="center"><%= oip.foneitem.fartist_name %></td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%=adminColor("gray")%>">�ּ�</td>
		<td align="center"><%= nl2br(oip.foneitem.faddress) %></td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%=adminColor("gray")%>">����ó</td>
		<td align="center"><%= oip.foneitem.fhp %></td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%=adminColor("gray")%>">�̸���</td>
		<td align="center"><%= oip.foneitem.fmail %></td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%=adminColor("gray")%>">����ڵ�Ϲ�ȣ</td>
		<td align="center"><%= oip.foneitem.flicense %></td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%=adminColor("gray")%>">Ȩ������</td>
		<td align="center"><a href="<%= oip.foneitem.fhomepage %>" class="a" target="_blank"><%= oip.foneitem.fhomepage %></a></td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%=adminColor("gray")%>">��Ƽ��Ʈ�Ұ�</td>
		<td align="center"><%= nl2br(oip.foneitem.fuser_info) %></td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%=adminColor("gray")%>">��ǰ��</td>
		<td align="center"><%= oip.foneitem.fsell_count %></td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%=adminColor("gray")%>">�����ǰ �Ǹ�ó</td>
		<td align="center"><%= oip.foneitem.fon_off_isusing %></td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%=adminColor("gray")%>">��ǰ�Ұ�</td>
		<td align="center"><%= nl2br(oip.foneitem.fitem_info) %></td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%=adminColor("gray")%>">÷������</td>
		<td align="center"><a href="<%=staticImgUrl%>/<%= oip.foneitem.ffile1 %>">�ٿ�ޱ�</a></td>
	</tr>
</table>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

<%
	set oip = nothing
%>	