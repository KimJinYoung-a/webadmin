<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lectureadmin/lib/classes/board_cls.asp"-->
<%
	'// ���� ���� //
	dim brdId, lecUserId
	dim page, searchDiv, searchKey, searchString, isanswer, param

	dim oBoard, i, lp

	'// �Ķ���� ���� //
	brdId = RequestCheckvar(request("brdId"),10)
	page = RequestCheckvar(request("page"),10)
	searchDiv = RequestCheckvar(request("searchDiv"),16)
	searchKey = RequestCheckvar(request("searchKey"),16)
	searchString = RequestCheckvar(request("searchString"),128)
	isanswer = RequestCheckvar(request("isanswer"),1)

	if page="" then page=1
	if searchKey="" then searchKey="qstTitle"

	param = "&page=" & page & "&searchDiv=" & searchDiv & "&searchKey=" & searchKey &_
			"&searchString=" & server.URLencode(searchString) & "&isanswer=" & isanswer	'������ ����

	'// ���� ����
	set oBoard = new CBoard
	oBoard.FRectbrdId = brdId

	oBoard.GetBoardRead

	if (oBoard.FResultCount = 0) then
	    response.write "<script>alert('�������� �ʴ� ���̰ų�, Ż���� ���Դϴ�.'); history.back();</script>"
	    dbget.close()	:	response.End
	end if
%>
<script language='javascript'>
<!--
	// �Է��� �˻�
	function chk_form(frm)
	{
		if(!frm.ansTitle.value)
		{
			alert("�亯 ������ �Է����ֽʽÿ�.");
			frm.ansTitle.focus();
			return false;
		}

		if(!frm.ansCont.value)
		{
			alert("�亯 ������ �ۼ����ֽʽÿ�.");
			frm.ansCont.focus();
			return false;
		}

		// �� ����
		return true;
	}


	// �亯 �Ӹ��� �ֱ�
	function chgCont(qcd, ccd)
	{
		FrameCHK.location="inc_board_cont.asp?brdId=<%=brdId%>&qcd=" + qcd + "&ccd=" + ccd;
	}

	// ���� ����
	function GotoBoardChange(){
		if (confirm('������ �����Ͻðڽ��ϱ�?')){
			document.frm_write.mode.value="change";
			document.frm_write.submit();
		}
	}

	// �ۻ���
	function GotoBoardDel(){
		if (confirm('���� �Ͻðڽ��ϱ�?')){
			document.frm_write.mode.value="delete";
			document.frm_write.submit();
		}
	}

//-->
</script>
<!-- ���� ȭ�� ���� -->
<table width="760" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<form name="frm_write" method="POST" onSubmit="return chk_form(this)" action="doLecBoard.asp">
<input type="hidden" name="mode" value="answer">
<input type="hidden" name="brdId" value="<%=brdId%>">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="searchDiv" value="<%=searchDiv%>">
<input type="hidden" name="searchKey" value="<%=searchKey%>">
<input type="hidden" name="searchString" value="<%=searchString%>">
<input type="hidden" name="isanswer" value="<%=isanswer%>">
<tr align="center" bgcolor="#F0F0FD">
	<td height="26" align="left" colspan="4"><b>����Խ��� �� ���� / �亯 �ۼ�</b></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#E8E8F1">����</td>
	<td colspan="3" bgcolor="#FFFFFF">
		<select name="commCd">
		<%=oBoard.optCommCd("'G000'", oBoard.FBoardList(0).FcommCd)%>
		</select>
		<img src="/images/icon_change.gif" onClick="GotoBoardChange()" style="cursor:pointer" align="absmiddle">
	</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#E8E8F1">�ۼ���</td>
	<td bgcolor="#FDFDFD" width="260"><%=oBoard.FBoardList(0).FlecUserId%></td>
	<td align="center" width="120" bgcolor="#E8E8F1">�ۼ��Ͻ�</td>
	<td bgcolor="#FDFDFD" width="260"><%=oBoard.FBoardList(0).Fregdate%></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#E8E8F1">���� ����</td>
	<td bgcolor="#FDFDFD" colspan="3"><%=db2html(oBoard.FBoardList(0).FqstTitle)%></td>
</tr>
<tr>
	<td colspan="4" align="center" bgcolor="#E8E8F1">���� ����</td>
</tr>
<tr>
	<td bgcolor="#FDFDFD" colspan="4" style="padding:10px"><%=nl2br(db2html(oBoard.FBoardList(0).FqstCont))%></td>
</tr>
<tr><td height="1" colspan="2" bgcolor="#D0D0D0"></td></tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>�亯 ����</td>
	<td bgcolor="#FFFFFF" colspan="3"><input type="text" name="ansTitle" value="<%=db2html(oBoard.FBoardList(0).FansTitle)%>" size="40" maxlength="40"></td>
</tr>
<% if oBoard.FBoardList(0).Fisanswer="���" then %>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">�Ӹ���/�λ縻</td>
	<td bgcolor="#FFFFFF" colspan="3">
		�Ӹ���
		<select name="preface" onchange="chgCont(this.value, compliment.value)">
			<%= oBoard.optCommCd("'G000'", oBoard.FBoardList(0).FcommCd)%>
		</select>
		/ �λ縻
		<select name="compliment" onchange="chgCont(preface.value, this.value)">
			<option value="">����</option>
			<%= oBoard.optCommCd("'E000'", "")%>
		</select>
	</td>
</tr>
<% end if %>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>�亯 ����</td>
	<td bgcolor="#FFFFFF" colspan="3"><textarea name="ansCont" rows="14" cols="80"><%=db2html(oBoard.inputAnswerCont(oBoard.FBoardList(0).FbrdId,"",""))%></textarea></td>
</tr>
<tr>
	<td colspan="4" height="32" bgcolor="#FAFAFA" align="center">
		<input type="image" src="/images/icon_reply.gif" style="border:0px;cursor:pointer" align="absmiddle"> &nbsp;
		<img src="/images/icon_delete.gif" onClick="GotoBoardDel()" style="cursor:pointer" align="absmiddle"> &nbsp;
		<img src="/images/icon_cancel.gif" onClick="self.location='lec_board_List.asp?menupos=<%=menupos & param%>'" style="cursor:pointer" align="absmiddle">
	</td>
</tr>
</form>
</table>
<iframe name="FrameCHK" src="" frameborder="0" width="0" height="0"></iframe>
<!-- ���� ȭ�� �� -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
