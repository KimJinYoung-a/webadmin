<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<!-- #include virtual="/academy/lib/classes/fingers_lecturecls.asp"-->

<!-- ���� ����� -->

<%
dim lec_idx
lec_idx=RequestCheckvar(request("lec_idx"),10)
dim wlec,w_i
set wlec = new CWaitLecture
wlec.GetWaitList lec_idx
%>
<body leftmargin="0" topmargin="0">
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
	<tr>
		<td colspan="6"align="center" bgcolor="#DDDDFF">����ڸ���Ʈ</td>
	</tr>

	<% for w_i = 1 to wlec.FResultCount %>
	<% if wlec.Flec_idx(w_i) = wlec.Flec_idx(w_i-1) then %>
	<% else %>
	<tr>
		<td colspan="6" bgcolor="#EEEEEE">
			<img src="<%= wlec.FLec_smallimg(w_i) %>" border="0"><%= wlec.FLec_title(w_i) %>(����	�ڵ� : <%= wlec.Flec_idx(w_i) %>)</td>
	</tr>
	<tr align="center" bgcolor="#DDDDFF">
		<td width="120" align="center">Userid</td>
		<td width="55" align="center">��û�μ�</td>
		<td width="60" align="center">�̸�</td>
		<td width="90" align="center">����ó</td>
		<td width="160" align="center">�̸���</td>
		<td width="140" align="center">��û��</td>
	</tr>

	<% end if %>
	<form name="wfrm_<%= w_i %>" method="get" action="">
	<tr>
		<td bgcolor="#FFFFFF" align="center"><% =wlec.FUserid(w_i) %></td>
		<td bgcolor="#FFFFFF" align="center"><% =wlec.FRegcount(w_i) %></td>
		<td bgcolor="#FFFFFF" align="center"><% =wlec.FUserName(w_i) %></td>
		<td bgcolor="#FFFFFF" align="left"><% =wlec.FPhone(w_i) %></td>
		<td bgcolor="#FFFFFF" align="left"><% =wlec.FEmail(w_i) %></td>
		<td bgcolor="#FFFFFF" align="left"><% =wlec.FRegdate(w_i) %></td>
	</tr>
	</form>
	<% next %>

	<tr>
		<td colspan="6" bgcolor="#FFFFFF" align="center"><input type="button" value="�ݱ�" onClick="self.close()"></td>
	</tr>

</table>

<%
set wlec= nothing
%>
</body>
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->