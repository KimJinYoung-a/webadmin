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
	param = "&page=" & page & "&searchKey=" & searchKey &_
			"&searchString=" & server.URLencode(searchString) & "&searchConfirm=" & searchConfirm	'������ ����

	'// ���� ����
	set oLecture = new CPartnerLecture
	oLecture.FRectidx = idx

	oLecture.GetPartnerLectureView
%>
<script language='javascript'>
<!--
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
			document.frm_write.mode.value="DelLeturer";
			document.frm_write.submit();
		}
	}

	// ��â���� ���� ����
	function NewWindow(v)
	{
		  //var p = (v);
		  //w = window.open("http://www.thefingers.co.kr/myfingers/showimage.asp?img=" + v, "imageView", "left=10px,top=10px, width=560,height=400,status=no,resizable=yes,scrollbars=yes");
		  //w.focus();
		  window.open("http://www.thefingers.co.kr/myfingers/showimage.asp?img=" + v, "imageView", "left=10px,top=10px, width=560,height=400,status=no,resizable=yes,scrollbars=yes");
	}

//-->
</script>
<!-- ���� ȭ�� ���� -->
<table width="760" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<form name="frm_write" method="POST" onSubmit="return chk_form(this)" action="doPartnerLecture.asp">
<input type="hidden" name="mode" value="AnsLeturer">
<input type="hidden" name="idx" value="<%=idx%>">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="searchKey" value="<%=searchKey%>">
<input type="hidden" name="searchString" value="<%=searchString%>">
<input type="hidden" name="searchConfirm" value="<%=searchConfirm%>">
<tr align="center" bgcolor="#F0F0FD">
	<td height="26" align="left" colspan="4"><b>�����û ���� �� ���� / �亯 �ۼ�</b></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">���ºо�</td>
	<td bgcolor="#FDFDFD" width="260"><%=oLecture.FItemList(0).Flecarea%></td>
	<td align="center" width="120" bgcolor="#DDDDFF">�ۼ��Ͻ�</td>
	<td bgcolor="#FDFDFD" width="260"><%=oLecture.FItemList(0).Fregdate%></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">���� ����</td>
	<td bgcolor="#FDFDFD" colspan="3"><%= nl2br(oLecture.FItemList(0).Flectitle) %></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">���� �̸�</td>
	<td bgcolor="#FDFDFD" colspan="3"><%=oLecture.FItemList(0).Flecname%></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">���� �Ұ�(���)</td>
	<td bgcolor="#FDFDFD" colspan="3"><%=nl2br(oLecture.FItemList(0).Fleccontent)%></td>
</tr>
<!--
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">�������</td>
	<td bgcolor="#FDFDFD" colspan="3"><%=oLecture.FItemList(0).Flecbirthday%></td>
</tr>
-->
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">����ó</td>
	<td bgcolor="#FDFDFD" width="260"><%=oLecture.FItemList(0).Flectel%></td>
	<td align="center" width="120" bgcolor="#DDDDFF">�޴���</td>
	<td bgcolor="#FDFDFD" width="260"><%=oLecture.FItemList(0).Flechp%></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">�̸���</td>
	<td bgcolor="#FDFDFD" colspan="3"><%=oLecture.FItemList(0).Flecmail%></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">Ȩ������</td>
	<td bgcolor="#FDFDFD" colspan="3">
	<%
		if oLecture.FItemList(0).Flecurl<>"" then
			Response.Write "<a href='"& oLecture.FItemList(0).Flecurl & "' target='_blank'>" & oLecture.FItemList(0).Flecurl & "</a>"
		end if
	%>
	</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">�ּ�</td>
	<td bgcolor="#FDFDFD" colspan="3"><%=oLecture.FItemList(0).Flecaddress%></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">���ǰ���</td>
	<td bgcolor="#FDFDFD" colspan="3"><%=oLecture.FItemList(0).Flecwork%></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">������</td>
	<td bgcolor="#FDFDFD" colspan="3">
		<%
		if oLecture.FItemList(0).farea = 0 then
			response.write "�������భ��"
		else
			response.write "�ܺ����భ��"
		end if
		%>
	</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">÷������</td>
	<td bgcolor="#FDFDFD" colspan="3">
	<%
		if oLecture.FItemList(0).Flecfile<>"" then
			'���������� ���� ���� �ɼ� �߰�
			Select Case getFileExtention(oLecture.FItemList(0).Flecfile)
				Case "jpg", "gif", "png"
					Response.Write "<span onClick=""NewWindow('" & imgFingers & oLecture.upfolder & "lecturer/" & oLecture.FItemList(0).Flecfile & "')"" style='cursor:pointer'>" & oLecture.FItemList(0).Flecfile & "</span>"
				Case Else
					Response.Write "<a href='" & imgFingers & "/linkweb/download.asp?filepath=" & Server.URLencode(oLecture.upfolder & "lecturer/") & "&filename=" & Server.URLencode(oLecture.FItemList(0).Flecfile) & "'>" & oLecture.FItemList(0).Flecfile & "</a>"
			end Select
		end if
	%>
	</td>
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
		<img src="/images/icon_cancel.gif" onClick="self.location='partnerLecture_List.asp?menupos=<%=menupos & param%>'" style="cursor:pointer" align="absmiddle">
	</td>
</tr>
</form>
</table>
<!-- ���� ȭ�� �� -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
