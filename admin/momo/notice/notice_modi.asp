<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ������� ��������
' Hieditor : 2009.11.27 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/momo/momo_cls.asp"-->
<%
'// ���� ���� //
dim oNotice, i, lp
dim ntcId,userid,commCd,title,contents,regdate,isusing
	ntcId = request("ntcId")

'// ���� ����
set oNotice = new CNotice
	oNotice.FRectntcId = ntcId
	
	'//������� �ϰ�쿡�� ����
	if ntcId <> "" then
		oNotice.GetNoitceRead
		
		if oNotice.FTotalCount > 0 then
			ntcId = oNotice.FNoticeList(0).fntcId
			userid = oNotice.FNoticeList(0).fuserid
			commCd = oNotice.FNoticeList(0).fcommCd
			title = oNotice.FNoticeList(0).ftitle
			contents = oNotice.FNoticeList(0).fcontents
			regdate = oNotice.FNoticeList(0).fregdate
			isusing = oNotice.FNoticeList(0).fisusing			
		end if
	end if
%>

<script language='javascript'>

	// �Է��� �˻�
	function chk_form(frm)
	{
		if(!frm.commcd.value)
		{
			alert("������ �ۼ����ֽʽÿ�.");			
			return false;
		}
		
		if(!frm.title.value)
		{
			alert("������ �Է����ֽʽÿ�.");
			frm.title.focus();
			return false;
		}

		if(!frm.contents.value)
		{
			alert("������ �ۼ����ֽʽÿ�.");
			frm.contents.focus();
			return false;
		}

		if(!frm.isusing.value)
		{
			alert("��뿩�θ� ������ �ֽʽÿ�.");			
			return false;
		}

		// �� ����
		return true;
	}

</script>

<table width="100%" border="0" align="center" class="a" cellpadding="0" cellspacing="1" bgcolor="#BABABA">
<form name="frm_write" method="POST" onSubmit="return chk_form(this)" action="doNotice.asp">
<input type="hidden" name="ntcId" value="<%=ntcId%>">
<input type="hidden" name="mode" value="edit">
<tr align="center" bgcolor="#FFFFFF">
	<td colspan="2" height="26" align="left"><b>�������� ���� ����</b></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#FFFFFF"><font color="darkred">* </font>����</td>
	<td bgcolor="#FFFFFF"><% drawnotics_gubun "commcd",commcd,"" %></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#FFFFFF"><font color="darkred">* </font>����</td>
	<td bgcolor="#FFFFFF"><input type="text" name="title" size="40" maxlength="40" value="<%= title %>"></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#FFFFFF">����</td>
	<td bgcolor="#FFFFFF"><textarea name="contents" rows="14" cols="80"><%= contents %></textarea></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#FFFFFF"><font color="darkred">* </font>��뿩��</td>
	<td bgcolor="#FFFFFF">
		<select name="isusing">
			<option value='' <% if isusing = "" then response.write " selected" %>>����</option>
			<option value='Y' <% if isusing = "Y" then response.write " selected" %>>Y</option>
			<option value='N' <% if isusing = "N" then response.write " selected" %>>N</option>
		</select>
	</td>
</tr>
<tr>
	<td colspan="2" height="32" bgcolor="#FAFAFA" align="center">
		<input type="image" src="/images/icon_save.gif" style="border:0px;cursor:pointer" align="absmiddle"> &nbsp;
		<img src="/images/icon_cancel.gif" onClick="history.back()" style="cursor:pointer" align="absmiddle">
	</td>
</tr>
</form>
</table>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
