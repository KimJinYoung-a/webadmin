<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ������� qna
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
dim qnaId, qstUserId , searchDiv
dim oQnA, oQnAList, oLec, i, lp
	'// �Ķ���� ���� //
	qnaId = request("qnaId")
	searchDiv = request("searchDiv")

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

	// �ۻ���
	function GotoqnaDel(){
		if (confirm('���� �Ͻðڽ��ϱ�?')){
			document.frm_write.mode.value="delete";
			document.frm_write.submit();
		}
	}

</script>

<table width="100%" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<form name="frm_write" method="POST" onSubmit="return chk_form(this)" action="doQnA.asp">
<input type="hidden" name="mode" value="answer">
<input type="hidden" name="qnaId" value="<%=qnaId%>">
<input type="hidden" name="searchDiv" value="<%=searchDiv%>">
<input type="hidden" name="qstUserName" value="<%=oQnA.FQnAList(0).Fusername%>">
<input type="hidden" name="regdate" value="<%=oQnA.FQnAList(0).Fregdate%>">
<input type="hidden" name="qstContents" value="<%=db2html(oQnA.FQnAList(0).FqstContents)%>">
<input type="hidden" name="qstTitle" value="<%=db2html(oQnA.FQnAList(0).FqstTitle)%>">
<tr align="center" bgcolor="#F0F0FD">
	<td height="26" align="left" colspan="4"><b>QnA �� ���� / �亯 �ۼ�</b></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#FFFFFF">���� �̸���</td>
	<td bgcolor="#FFFFFF">
		<%=db2html(oQnA.FQnAList(0).FqstUserMail)%>
		<input type="hidden" name="qstUserMail" value="<%=oQnA.FQnAList(0).FqstUserMail%>">
	</td>
	<td align="center" width="120" bgcolor="#FFFFFF">����</td>
	<td width="260" bgcolor="#FFFFFF">
		<%=oQnA.FQnAList(0).Fisanswer%>
	</td>	
</tr>
<tr>
	<td align="center" width="120" bgcolor="#FFFFFF">�ۼ���</td>
	<td bgcolor="#FDFDFD" width="260">
		<%=oQnA.FQnAList(0).Fusername & "(" & oQnA.FQnAList(0).FqstUserid & ")"%>
	</td>
	<td align="center" width="120" bgcolor="#FFFFFF">�ۼ��Ͻ�</td>
	<td bgcolor="#FDFDFD" width="260"><%=oQnA.FQnAList(0).Fregdate%></td>
</tr>
<%
	if oQna.FQnAList(0).FlecIdx<>"" then
		set oLec = new CQnA
		oLec.FRectlecIdx = oQna.FQnAList(0).FlecIdx

		oLec.GetLecRead

		if oLec.FlecList(0).FcateName<>"" then
%>
<tr>
	<td align="center" width="120" bgcolor="#FFFFFF">��������</td>
	<td bgcolor="#FDFDFD" width="640" colspan="3"><%= "[" & oLec.FlecList(0).FcateName & "] " & db2html(oLec.FlecList(0).FlecTitle)%></td>
</tr>
<%
		end if
	end if
%>
<tr>
	<td align="center" width="120" bgcolor="#FFFFFF">���� ����</td>
	<td bgcolor="#FDFDFD" colspan="3"><%=db2html(oQnA.FQnAList(0).FqstTitle)%></td>
</tr>
<tr>
	<td colspan="4" align="center" bgcolor="#FFFFFF">���� ����</td>
</tr>
<tr>
	<td bgcolor="#FDFDFD" colspan="4" style="padding:10px"><%=nl2br(db2html(oQnA.FQnAList(0).FqstContents))%></td>
</tr>
</table>
<br>
<table width="100%" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>�亯 ����</td>
	<td bgcolor="#FFFFFF" colspan="3"><input type="text" name="ansTitle" value="<%=db2html(oQnA.FQnAList(0).FansTitle)%>" size="40" maxlength="40"></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>�亯 ����</td>
	<td bgcolor="#FFFFFF" colspan="3"><textarea name="ansContents" rows="14" cols="80"><%=db2html(oQnA.FQnAList(0).FansContents)%></textarea></td>
</tr>
<tr>
	<td colspan="4" height="32" bgcolor="#FAFAFA" align="center">
		<input type="image" src="/images/icon_reply.gif" style="border:0px;cursor:pointer" align="absmiddle"> &nbsp;
		<img src="/images/icon_delete.gif" onClick="GotoqnaDel()" style="cursor:pointer" align="absmiddle"> &nbsp;
		<img src="/images/icon_cancel.gif" onClick="self.location='QnA_List.asp'" style="cursor:pointer" align="absmiddle">
	</td>
</tr>
</form>
</table>

<!-- ���� ����Ʈ ����  -->
<%

'������ ���̵� ����
qstUserId = oQnA.FQnAList(0).FqstUserid
set oQnAList = Nothing

'//ȸ���� ���
if qstUserId <> "" then
	
	'// �ٽ� Ŭ���� ����
	set oQnAList = new CQnA
		oQnAList.FCurrPage = 1 
		oQnAList.FPageSize = 50
		oQnAList.FRectuserid = qstUserId
		oQnAList.GetQnAList
%>
	<br>
	<table width="100%" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
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
			<td><%= oQnAList.FQnAList(lp).Fcommcd %></td>
			<td align="left"><a href="QnA_view.asp?qnaId=<%= oQnAList.FQnAList(lp).FqnaId %>"><%= db2html(oQnAList.FQnAList(lp).FqstTitle) %></a></td>
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
end if	
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->