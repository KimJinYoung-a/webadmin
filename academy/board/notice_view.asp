<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/classes/notice_cls.asp"-->
<%
	'// ���� ���� //
	dim ntcId
	dim page, searchDiv, searchKey, searchString, param

	dim oNotice, i, lp

	'// �Ķ���� ���� //
	ntcId = RequestCheckvar(request("ntcId"),10)
	page = RequestCheckvar(request("page"),10)
	searchDiv = RequestCheckvar(request("searchDiv"),32)
	searchKey = RequestCheckvar(request("searchKey"),32)
	searchString = RequestCheckvar(request("searchString"),128)

	if page="" then page=1
	if searchKey="" then searchKey="titleLong"

	param = "&page=" & page & "&searchDiv=" & searchDiv & "&searchKey=" & searchKey & "&searchString=" & searchString	'������ ����

	'// ���� ����
	set oNotice = new CNotice
	oNotice.FRectntcId = ntcId

	oNotice.GetNoitceRead

%>
<script language="javascript">
<!--
	// �ۻ���
	function GotoNoticeDel(){
		if (confirm('���� �Ͻðڽ��ϱ�?')){
			document.frm_trans.submit();
		}
	}
//-->
</script>
<!-- ���� ȭ�� ���� -->
<table width="100%" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr align="center" bgcolor="#F0F0FD">
	<td colspan="2">
		<table width="100%" border="0" class="a" cellpadding="0" cellspacing="0">
		<tr>
			<td height="26" align="left"><b>�������� �� ����</b></td>
			<td height="26" align="right"><%=oNotice.FNoticeList(0).Fregdate%>&nbsp;</td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">�ۼ���</td>
	<td bgcolor="#FFFFFF"><%=oNotice.FNoticeList(0).Fusername & "(" & oNotice.FNoticeList(0).Fuserid & ")" %></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">����</td>
	<td bgcolor="#FFFFFF"><%=db2html(oNotice.FNoticeList(0).FcommNm)%></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">����</td>
	<td bgcolor="#F8F8FF"><%=db2html(oNotice.FNoticeList(0).Ftitle)%></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">����</td>
	<td bgcolor="#FFFFFF"><%=nl2br(db2html(oNotice.FNoticeList(0).Fcontents))%></td>
</tr>
<tr><td height="1" colspan="2" bgcolor="#D0D0D0"></td></tr>
<tr>
	<td colspan="2" height="32" bgcolor="#FAFAFA" align="center">
		<img src="/images/icon_modify.jpg" onClick="self.location='notice_modi.asp?menupos=<%=menupos%>&ntcId=<%=ntcId & param%>'" style="cursor:pointer" align="absmiddle"> &nbsp;
		<img src="/images/icon_delete.gif" onClick="GotoNoticeDel()" style="cursor:pointer" align="absmiddle"> &nbsp;
		<img src="/images/icon_list.gif" onClick="self.location='notice_list.asp?menupos=<%=menupos & param %>'" style="cursor:pointer" align="absmiddle">
	</td>
</tr>
<form name="frm_trans" method="POST" action="doNotice.asp">
<input type="hidden" name="ntcId" value="<%=ntcId%>">
<input type="hidden" name="mode" value="delete">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="searchDiv" value="<%=searchDiv%>">
<input type="hidden" name="searchKey" value="<%=searchKey%>">
<input type="hidden" name="searchString" value="<%=searchString%>">
</form>
</table>
<!-- ���� ȭ�� �� -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
