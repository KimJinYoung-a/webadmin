<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/classes/mardy_storycls.asp"-->
<%
	'// ���� ���� //
	dim storyId
	dim page, searchKey, searchString, param

	dim oStory, oStoryImage, i, lp

	'// �Ķ���� ���� //
	storyId = RequestCheckvar(request("storyId"),10)
	page = RequestCheckvar(request("page"),10)
	searchKey = RequestCheckvar(request("searchKey"),16)
	searchString = request("searchString")
  	if searchString <> "" then
		if checkNotValidHTML(searchString) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
		response.write "</script>"
		response.End
		end if
	end if
	if page="" then page=1
	if searchKey="" then searchKey="titleLong"

	param = "&page=" & page & "&searchKey=" & searchKey & "&searchString=" & searchString	'������ ����

	'// ���� ����
	set oStory = new CMardyStory
	oStory.FRectstoryId = storyId

	oStory.GetMardyStoryView

	'// ���� �̹��� ����
	set oStoryImage = new CMardyStory
	oStoryImage.FRectstoryId = storyId

	oStoryImage.GetMardyStoryImageList
%>
<script language="javascript">
<!--
	// �ۻ���
	function GotoStoryDel(){
		if (confirm('�� �Խù��� ������ ���� �Ͻðڽ��ϱ�?\n\n�� ���� �Ŀ��� �ٽ� ���� �� �� �����ϴ�.')){
			document.frm_trans.submit();
		}
	}


	// ��� ���� ����
	function GotoUse(md)
	{
		switch(md)
		{
			case "use" :
				if (confirm('����Ʈ ��Ͽ� ��µǵ��� ���¸� "���"���� �����Ͻðڽ��ϱ�?')){
					FrameCHK.location="inc_Mardy_Use.asp?Idx=<%=storyId%>&mode=StoryUse";
				}
				break;

			case "del" :
				if (confirm('����Ʈ���� �� �� ������ ���¸� "����"���� �����Ͻðڽ��ϱ�?')){
					FrameCHK.location="inc_Mardy_Use.asp?Idx=<%=storyId%>&mode=StoryDel";
				}
				break;
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
			<td height="26" align="left"><b>���� �̾߱� �� ����</b></td>
			<td height="26" align="right">
				<font color=gray>��뿩�� - </font>
				<%
					if oStory.FItemList(0).Fisusing="N" then
						Response.Write "<font color=darkred><b>����</b></font>"
					else
						Response.Write "<font color=darkblue><b>���</b></font>"
					end if
				%>&nbsp;
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">ª�� ����</td>
	<td bgcolor="#FFFFFF"><%=oStory.FItemList(0).FtitleShort%></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">�� ����</td>
	<td bgcolor="#FFFFFF"><%=oStory.FItemList(0).FtitleLong%></td>
</tr>
<tr><td height="1" colspan="2" bgcolor="#D0D0D0"></td></tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">�̹���</td>
	<td bgcolor="#FFFFFF">
		<table border="0" class="a" cellpadding="0" cellspacing="0">
		<%
			for lp=0 to oStoryImage.FTotalCount - 1
		%>
		<tr>
			<td align="center">
				<img src="<%=oStoryImage.FItemList(lp).FimgFile_full%>" ><br><br>
			</td>
		</tr>
		<% next %>
		</table>
	</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">���� ����</td>
	<td bgcolor="#FFFFFF"><%=nl2br(oStory.FItemList(0).Fcontents)%></td>
</tr>
<tr><td height="1" colspan="2" bgcolor="#D0D0D0"></td></tr>
<tr>
	<td colspan="2" height="32" bgcolor="#FAFAFA" align="center">
		<img src="/images/icon_modify.gif" onClick="self.location='mardyStory_modi.asp?menupos=<%=menupos%>&storyId=<%=storyId & param%>'" style="cursor:pointer" align="absmiddle"> &nbsp;
		<% if oStory.FItemList(0).Fisusing="N" then %>
		<img src="/images/icon_use.gif" onClick="GotoUse('use')" style="cursor:pointer" align="absmiddle"> &nbsp;
		<img src="/images/icon_delete.gif" onClick="GotoStoryDel()" style="cursor:pointer" align="absmiddle"> &nbsp;
		<% else %>
		<img src="/images/icon_hide.gif" onClick="GotoUse('del')" style="cursor:pointer" align="absmiddle"> &nbsp;
		<% end if %>
		<img src="/images/icon_list.gif" onClick="self.location='mardyStory_list.asp?menupos=<%=menupos & param %>'" style="cursor:pointer" align="absmiddle">
	</td>
</tr>
<form name="frm_trans" method="POST" action="http://image.thefingers.co.kr/linkweb/doMardyStory.asp" enctype="multipart/form-data">
<input type="hidden" name="storyId" value="<%=storyId%>">
<input type="hidden" name="mode" value="delete">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="searchKey" value="<%=searchKey%>">
<input type="hidden" name="searchString" value="<%=searchString%>">
<input type="hidden" name="adminId" value="<%=Session("ssBctId")%>">
</form>
</table>
<iframe name="FrameCHK" src="" frameborder="0" width="0" height="0"></iframe>
<!-- ���� ȭ�� �� -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->