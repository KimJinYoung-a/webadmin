<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lectureadmin/lib/classes/board_cls.asp"-->
<%
	'// ���� ���� //
	dim brdId
	dim page, searchDiv, searchKey, searchString, param

	dim oBoard, i, lp

	'// �Ķ���� ���� //
	brdId = requestCheckVar(request("brdId"),10)
	page = requestCheckVar(request("page"),10)
	searchDiv = requestCheckVar(request("searchDiv"),10)
	searchKey = requestCheckVar(request("searchKey"),10)
	searchString = requestCheckVar(request("searchString"),128)
  	if searchString <> "" then
		if checkNotValidHTML(searchString) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');"
		response.write "</script>"
		response.End
		end if
	end if
	param = "&page=" & page & "&searchDiv=" & searchDiv & "&searchKey=" & searchKey & "&searchString=" & searchString	'������ ����

	'// ���� ����
	set oBoard = new Cboard
	oBoard.FRectbrdId = brdId

	oBoard.GetBoardRead

	if (oBoard.FResultCount = 0) then
	    response.write "<script>alert('�������� �ʴ� ���Դϴ�.'); history.back();</script>"
	    dbget.close()	:	response.End
	end if

%>
<script language="javascript">
<!--
	// �ۻ���
	function GotoBoardDel(){
		if (confirm('���� �Ͻðڽ��ϱ�?')){
			document.frm_trans.submit();
		}
	}
//-->
</script>
<!-- ���� ȭ�� ���� -->
<table width="750" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr align="center" bgcolor="#F0F0FD">
	<td colspan="4">
		<table width="100%" border="0" class="a" cellpadding="0" cellspacing="0">
		<tr>
			<td height="26" align="left"><b>�Խù� �� ����</b></td>
			<td height="26" align="right"><%=oBoard.FBoardList(0).Fregdate%>&nbsp;</td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">����</td>
	<td width="255" bgcolor="#FFFFFF"><%=db2html(oBoard.FBoardList(0).FcommNm)%></td>
	<td align="center" width="120" bgcolor="#DDDDFF">����</td>
	<td width="255" bgcolor="#FFFFFF"><%=oBoard.FBoardList(0).Fisanswer%></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">����</td>
	<td width="630" colspan="3" bgcolor="#F8F8FF"><%=db2html(oBoard.FBoardList(0).FqstTitle)%></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">����</td>
	<td colspan="3" bgcolor="#FFFFFF"><%=nl2br(db2html(oBoard.FBoardList(0).FqstCont))%></td>
</tr>
<tr><td height="1" colspan="4" bgcolor="#D0D0D0"></td></tr>
<% if oBoard.FBoardList(0).Fisanswer="�Ϸ�" then %>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">�亯��</td>
	<td width="255" bgcolor="#FFFFFF"><%=oBoard.FBoardList(0).FansUserId %></td>
	<td align="center" width="120" bgcolor="#DDDDFF">�亯�Ͻ�</td>
	<td width="255" bgcolor="#FFFFFF"><%=oBoard.FBoardList(0).FansDate %></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">�亯����</td>
	<td colspan="3" bgcolor="#F8F8FF"><%=db2html(oBoard.FBoardList(0).FansTitle)%></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">�亯����</td>
	<td colspan="3" bgcolor="#FFFFFF"><%=nl2br(db2html(oBoard.FBoardList(0).FansCont))%></td>
</tr>
<tr><td height="1" colspan="4" bgcolor="#D0D0D0"></td></tr>
<% end if %>
<tr>
	<td colspan="4" height="32" bgcolor="#FAFAFA" align="center">
		<% if oBoard.FBoardList(0).Fisanswer="���" then %>
		<img src="/images/icon_modify.jpg" onClick="self.location='board_modi.asp?menupos=<%=menupos%>&brdId=<%=brdId & param%>'" style="cursor:pointer" align="absmiddle"> &nbsp;
		<img src="/images/icon_delete.gif" onClick="GotoBoardDel()" style="cursor:pointer" align="absmiddle"> &nbsp;
		<% end if %>
		<img src="/images/icon_list.gif" onClick="self.location='board_list.asp?menupos=<%=menupos & param %>'" style="cursor:pointer" align="absmiddle">
	</td>
</tr>
<form name="frm_trans" method="POST" action="doBoard.asp">
<input type="hidden" name="brdId" value="<%=brdId%>">
<input type="hidden" name="mode" value="delete">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="searchDiv" value="<%=searchDiv%>">
<input type="hidden" name="searchKey" value="<%=searchKey%>">
<input type="hidden" name="searchString" value="<%=searchString%>">
</form>
</table>
<!-- ���� ȭ�� �� -->
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
