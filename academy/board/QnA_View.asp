<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
'	History	:  2009.09.16 �ѿ�� ����
'	Description : 1:1���
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/classes/QnA_cls.asp"-->
<%
	'// ���� ���� //
	dim qnaId, qstUserId
	dim page, searchDiv, searchKey, searchString, isanswer, param

	dim oQnA, oQnAList, oLec, i, lp

	'// �Ķ���� ���� //
	qnaId = RequestCheckvar(request("qnaId"),10)
	page = RequestCheckvar(request("page"),10)
	searchDiv = RequestCheckvar(request("searchDiv"),10)
	searchKey = RequestCheckvar(request("searchKey"),16)
	searchString = RequestCheckvar(request("searchString"),128)
	isanswer = RequestCheckvar(request("isanswer"),1)

	if page="" then page=1
	if searchKey="" then searchKey="qstTitle"

	param = "&page=" & page & "&searchDiv=" & searchDiv & "&searchKey=" & searchKey &_
			"&searchString=" & server.URLencode(searchString) & "&isanswer=" & isanswer	'������ ����

	'// ���� ����
	set oQnA = new CQnA
	oQnA.FRectqnaId = qnaId

	oQnA.GetQnARead

	if (oQnA.FResultCount = 0) then
	    response.write "<script>alert('�������� �ʴ� ���̰ų�, Ż���� ���Դϴ�.'); history.back();</script>"
	    dbget.close()	:	response.End
	end if
%>
<script language='javascript'>

	// �Է��� �˻�
	function chk_form(frm)
	{
		if(!frm.ansTitle.value)
		{
			alert("�亯 ������ �Է����ֽʽÿ�.");
			frm.ansTitle.focus();
			return false;
		}

		if(!frm.ansContents.value)
		{
			alert("�亯 ������ �ۼ����ֽʽÿ�.");
			frm.ansContents.focus();
			return false;
		}

		// �� ����
		return true;
	}


	// �亯 �Ӹ��� �ֱ�
	function chgCont(qcd, ccd)
	{
		FrameCHK.location="inc_qna_cont.asp?qnaId=<%=qnaId%>&qcd=" + qcd + "&ccd=" + ccd;
	}

	// ���� ����
	function GotoqnaChange(){
		if (confirm('������ �����Ͻðڽ��ϱ�?')){
			document.frm_write.mode.value="change";
			document.frm_write.submit();
		}
	}

	// �ۻ���
	function GotoqnaDel(){
		if (confirm('���� �Ͻðڽ��ϱ�?')){
			document.frm_write.mode.value="delete";
			document.frm_write.submit();
		}
	}

</script>
<!-- ���� ȭ�� ���� -->
<table width="760" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<form name="frm_write" method="POST" onSubmit="return chk_form(this)" action="doQnA.asp">
<input type="hidden" name="mode" value="answer">
<input type="hidden" name="qnaId" value="<%=qnaId%>">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="searchDiv" value="<%=searchDiv%>">
<input type="hidden" name="searchKey" value="<%=searchKey%>">
<input type="hidden" name="searchString" value="<%=searchString%>">
<input type="hidden" name="isanswer" value="<%=isanswer%>">
<input type="hidden" name="qstUserName" value="<%=oQnA.FQnAList(0).Fusername%>">
<input type="hidden" name="regdate" value="<%=oQnA.FQnAList(0).Fregdate%>">
<input type="hidden" name="qstContents" value="<%=db2html(oQnA.FQnAList(0).FqstContents)%>">
<input type="hidden" name="qstTitle" value="<%=db2html(oQnA.FQnAList(0).FqstTitle)%>">
<tr align="center" bgcolor="#F0F0FD">
	<td height="26" align="left" colspan="4"><b>QnA �� ���� / �亯 �ۼ�</b></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#E8E8F1">�з�</td>
	<td width="260" bgcolor="#FFFFFF"><%=oQnA.FQnAList(lp).FgroupNm%></td>
	<td align="center" width="120" bgcolor="#E8E8F1">����</td>
	<td width="260" bgcolor="#FFFFFF">
		<select name="commCd">
		<%=oQnA.optCommCd("'" & Left(oQnA.FQnAList(0).FcommCd,1) & "000'", oQnA.FQnAList(0).FcommCd)%>
		</select>
		<img src="/images/icon_change.gif" onClick="GotoqnaChange()" style="cursor:pointer" align="absmiddle">
	</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#E8E8F1">�ۼ���</td>
	<td bgcolor="#FDFDFD" width="260"><%=oQnA.FQnAList(0).Fusername & "(" & oQnA.FQnAList(0).FqstUserid & ")"%></td>
	<td align="center" width="120" bgcolor="#E8E8F1">�ۼ��Ͻ�</td>
	<td bgcolor="#FDFDFD" width="260"><%=oQnA.FQnAList(0).Fregdate%></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#E8E8F1">���� �̸���</td>
	<td bgcolor="#FDFDFD">
		<%=db2html(oQnA.FQnAList(0).FqstUserMail)%>
		<input type="hidden" name="qstUserMail" value="<%=oQnA.FQnAList(0).FqstUserMail%>">
	</td>
	<td align="center" width="120" bgcolor="#E8E8F1">���� ���ſ���</td>
	<td bgcolor="#FDFDFD">
		<%=oQnA.FQnAList(0).FmailOk%>
		<input type="hidden" name="mailOk" value="<%=oQnA.FQnAList(0).FmailOk%>">
	</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#E8E8F1">������(1-5��)</td>
	<td bgcolor="#FDFDFD" width="260">
		<% 
		if oQnA.FQnAList(0).fbestviewcount <> 0 then
		for i = 1 to oQnA.FQnAList(0).fbestviewcount 
		%>
		<img src="http://image.thefingers.co.kr/academy2009/lecture/star_on_gray.gif">
		<% 
		next 
		else
		%>
		�� ����
		<% end if %>
	</td>
	<td align="center" width="120" bgcolor="#E8E8F1"></td>
	<td bgcolor="#FDFDFD" width="260"></td>
</tr>
<%
	if oQna.FQnAList(0).FlecIdx<>"" then
		set oLec = new CQnA
		oLec.FRectlecIdx = oQna.FQnAList(0).FlecIdx

		oLec.GetLecRead

		if oLec.FlecList(0).FcateName<>"" then
%>
<tr>
	<td align="center" width="120" bgcolor="#E8E8F1">��������</td>
	<td bgcolor="#FDFDFD" width="640" colspan="3"><%= "[" & oLec.FlecList(0).FcateName & "] " & db2html(oLec.FlecList(0).FlecTitle)%></td>
</tr>
<%
		end if
	end if
%>
<tr>
	<td align="center" width="120" bgcolor="#E8E8F1">���� ����</td>
	<td bgcolor="#FDFDFD" colspan="3"><%=db2html(oQnA.FQnAList(0).FqstTitle)%></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#E8E8F1">�����ֹ���ȣ</td>
	<td bgcolor="#FDFDFD" colspan="3"><%=db2html(oQnA.FQnAList(0).Forderserial)%></td>
</tr>
<tr>
	<td colspan="4" align="center" bgcolor="#E8E8F1">���� ����</td>
</tr>
<tr>
	<td bgcolor="#FDFDFD" colspan="4" style="padding:10px"><%=nl2br(db2html(oQnA.FQnAList(0).FqstContents))%></td>
</tr>
<tr><td height="1" colspan="2" bgcolor="#D0D0D0"></td></tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>�亯 ����</td>
	<td bgcolor="#FFFFFF" colspan="3"><input type="text" name="ansTitle" value="<%=db2html(oQnA.FQnAList(0).FansTitle)%>" size="40" maxlength="40"></td>
</tr>
<% if oQnA.FQnAList(0).Fisanswer="���" then %>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">�Ӹ���/�λ縻</td>
	<td bgcolor="#FFFFFF" colspan="3">
		�Ӹ���
		<select name="preface" onchange="chgCont(this.value, compliment.value)">
			<%= oQnA.optPrfCd("'A000'", "H999")%>
		</select>
		/ �λ縻
		<select name="compliment" onchange="chgCont(preface.value, this.value)">
			<option value="">����</option>
			<%= oQnA.optCommCd("'E000'", "")%>
		</select>
	</td>
</tr>
<% end if %>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>�亯 ����</td>
	<td bgcolor="#FFFFFF" colspan="3"><textarea name="ansContents" rows="14" cols="80"><%=db2html(oQnA.inputAnswerCont(oQnA.FQnAList(0).FqnaId,"",""))%></textarea></td>
</tr>
<tr>
	<td colspan="4" height="32" bgcolor="#FAFAFA" align="center">
		<input type="image" src="/images/icon_reply.gif" style="border:0px;cursor:pointer" align="absmiddle"> &nbsp;
		<img src="/images/icon_delete.gif" onClick="GotoqnaDel()" style="cursor:pointer" align="absmiddle"> &nbsp;
		<img src="/images/icon_cancel.gif" onClick="self.location='QnA_List.asp?menupos=<%=menupos & param%>'" style="cursor:pointer" align="absmiddle">
	</td>
</tr>
</form>
</table>
<iframe name="FrameCHK" src="" frameborder="0" width="0" height="0"></iframe>
<!-- ���� ȭ�� �� -->
<%
		'������ ���̵� ����
		qstUserId = oQnA.FQnAList(0).FqstUserid
	set oQnA = Nothing
%>
<!-- ���� ����Ʈ ����  -->
<%
	'// Ŭ���� ����
	set oQnAList = new CQnA
	oQnAList.FCurrPage = 1 ''page <- ���������� �� �������� �ƴ�.
	oQnAList.FPageSize = 200
	oQnAList.FRectuserid = qstUserId

	oQnAList.GetQnAList
%>
<br><br>
<table width="760" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
	<tr align="center" bgcolor="#F0F0FD">
		<td colspan="6" align="center"><%= qstUserId %> ȸ���� ���� ���� ���</td>
	</tr>
	<tr align="center" bgcolor="#DDDDFF">
		<td align="center" width="40">��ȣ</td>
		<td align="center" width="120">����</td>
		<td align="center">����</td>
		<td align="center" width="70">�����</td>
		<td align="center" width="50">����</td>
		<td align="center" width="80">�����</td>
	</tr>
	<%
		for lp=0 to oQnAList.FResultCount - 1
	%>
	<tr align="center" bgcolor="#FFFFFF">
		<td><%= oQnAList.FQnAList(lp).FqnaId %></td>
		<td><%= oQnAList.FQnAList(lp).FcommNm %></td>
		<td align="left"><a href="QnA_view.asp?qnaId=<%= oQnAList.FQnAList(lp).FqnaId & param%>"><%= db2html(oQnAList.FQnAList(lp).FqstTitle) %></a></td>
		<td><%= oQnAList.FQnAList(lp).FqstUserId %></td>
		<td><%= oQnAList.FQnAList(lp).Fisanswer %></td>
		<td><%= FormatDate(oQnAList.FQnAList(lp).Fregdate,"0000.00.00") %></td>
	</tr>
	<%
		next
	%>
</table>
<!-- ���� ����Ʈ ��  -->
<%
	set oQnAList = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->