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
dim page ,i , artist_idx
	page = request("page")
	artist_idx = request("artist_idx")
	if page="" then page=1

dim oip
	set oip = New Cinquiry_list
	oip.frectartist_idx = artist_idx
	oip.frecommend_oneitem()
%>

<table width="100%" align="center" cellpadding="4" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%=adminColor("gray")%>">��ȣ</td>
		<td align="center"><%= oip.foneitem.fartist_idx %></td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%=adminColor("gray")%>">��Ƽ��Ʈ</td>
		<td align="center"><%= oip.foneitem.fartist_name %></td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%=adminColor("gray")%>">��Ƽ��Ʈ�±�</td>
		<td align="center"><%= oip.foneitem.ftag %></td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%=adminColor("gray")%>">Ȩ������</td>
		<td align="center"><a href="<%= oip.foneitem.fhomepage %>" class="a" target="_blank"><%= nl2br(oip.foneitem.fhomepage) %></a></td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%=adminColor("gray")%>">��α�</td>
		<td align="center"><a href="<%= oip.foneitem.fblog %>" class="a" target="_blank"><%= oip.foneitem.fblog %></a></td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%=adminColor("gray")%>">��õ����</td>
		<td align="center"><%= nl2br(oip.foneitem.fwhyrecommend) %></td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%=adminColor("gray")%>">�ۼ���</td>
		<td align="center"><%= oip.foneitem.fuserid %></td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%=adminColor("gray")%>">�����</td>
		<td align="center"><%= oip.foneitem.fregdate %></td>
	</tr>
</table>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<%
	set oip = nothing
%>	