<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
'	History	:  2009.09.10 �ѿ�� ����/�߰�
'	Description : ��Ʈ�ʽ�
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/classes/partner_lecturecls.asp"-->
<%
	'// ���� ���� //
	dim idx
	dim page, searchKey, searchString, searchConfirm, param

	dim oLecture, i, lp

	'// �Ķ���� ���� //
	idx = RequestCheckvar(request("idx"),10)
	page = RequestCheckvar(request("page"),10)
	searchKey = RequestCheckvar(request("searchKey"),16)
	searchString = request("searchString")
	searchConfirm = RequestCheckvar(request("searchConfirm"),1)
  	if searchString <> "" then
		if checkNotValidHTML(searchString) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
		response.write "</script>"
		response.End
		end if
	end if

	param = "&page=" & page & "&searchKey=" & searchKey  &_
			"&searchString=" & server.URLencode(searchString) & "&searchConfirm=" & searchConfirm	'������ ����

	'// ���� ����
	set oLecture = new CPartnerGroupLecture
	oLecture.FRectidx = idx

	oLecture.GetPartnerGroupView
%>

<script language='javascript'>

	// �Է��� �˻�
	function chk_form(frm)
	{
		if(!frm.confirmMemo.value)
		{
			alert("�亯 ������ �ۼ����ֽʽÿ�.");
			frm.confirmMemo.focus();
			return false;
		}

		// �� ����
		return true;
	}


	// �ۻ���
	function GotoLectureDel(){
		if (confirm('���� �Ͻðڽ��ϱ�?')){
			document.frm_write.mode.value="DelGroup";
			document.frm_write.submit();
		}
	}


</script>

<!-- ���� ȭ�� ���� -->
<table width="760" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<form name="frm_write" method="POST" onSubmit="return chk_form(this)" action="doPartnerLecture.asp">
<input type="hidden" name="mode" value="AnsGroup">
<input type="hidden" name="idx" value="<%=idx%>">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="searchKey" value="<%=searchKey%>">
<input type="hidden" name="searchString" value="<%=searchString%>">
<input type="hidden" name="searchConfirm" value="<%=searchConfirm%>">
<tr align="center" bgcolor="#F0F0FD">
	<td height="26" align="left" colspan="4"><b>��ü���� ���� �� ���� / �亯 �ۼ�</b></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">���¸�</td>
	<td bgcolor="#FDFDFD" width="260"><%=oLecture.FItemList(0).Flecturename%></td>
	<td align="center" width="120" bgcolor="#DDDDFF">���������</td>
	<td bgcolor="#FDFDFD" width="260"><%=oLecture.FItemList(0).Flecturedate%></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">��ü��</td>
	<td bgcolor="#FDFDFD" width="260"><%=oLecture.FItemList(0).Fpartyname%></td>
	<td align="center" width="120" bgcolor="#DDDDFF">�����ο���</td>
	<td bgcolor="#FDFDFD" width="260"><%=oLecture.FItemList(0).Fpartymannumber%></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">��û�ڸ�</td>
	<td bgcolor="#FDFDFD" width="260"><%=oLecture.FItemList(0).Fpartymastername%></td>
	<td align="center" width="120" bgcolor="#DDDDFF">�޴���ȭ</td>
	<td bgcolor="#FDFDFD" width="260"><%=oLecture.FItemList(0).Fpartymasterhp%></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">��ȭ</td>
	<td bgcolor="#FDFDFD" width="260"><%=oLecture.FItemList(0).fpartymastertel%></td>
	<td align="center" width="120" bgcolor="#DDDDFF">�������</td>
	<td bgcolor="#FDFDFD" width="260">
		<%
		if oLecture.FItemList(0).flecturearea = 0 then
			response.write "�������భ��"
		else
			response.write "�ܺ����భ��"
		end if
		%>	
	</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">�̸���</td>
	<td bgcolor="#FDFDFD" colspan="3"><%=oLecture.FItemList(0).Fpartymastermail%></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">�������� / �����Կ�</td>
	<td bgcolor="#FDFDFD" colspan="3"><%=oLecture.FItemList(0).Fchoiceyn%></td>
</tr>
<tr><td height="1" colspan="2" bgcolor="#D0D0D0"></td></tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>�亯 ����</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<textarea name="confirmMemo" rows="10" cols="80"><%=oLecture.FItemList(0).FconfirmMemo%></textarea><br>
		�� �亯 ������ ����� ���� ���Դϴ�. ������ ���� ���� �����Ƿ� ��������� ������ֽʽÿ�.
	</td>
</tr>
<tr>
	<td colspan="4" height="32" bgcolor="#FAFAFA" align="center">
		<input type="image" src="/images/icon_reply.gif" style="border:0px;cursor:pointer" align="absmiddle"> &nbsp;
		<img src="/images/icon_delete.gif" onClick="GotoLectureDel()" style="cursor:pointer" align="absmiddle"> &nbsp;
		<img src="/images/icon_cancel.gif" onClick="self.location='partnerGroupLecture_List.asp?menupos=<%=menupos & param%>'" style="cursor:pointer" align="absmiddle">
	</td>
</tr>
</form>
</table>
<!-- ���� ȭ�� �� -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->