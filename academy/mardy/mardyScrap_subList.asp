<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/classes/mardy_Scrapcls.asp"-->
<%
	'// ���� ���� //
	dim ScrapId
	dim page, searchKey, searchString, param

	dim oScrap, oScrapImage, i, lp

	'// �Ķ���� ���� //
	ScrapId = RequestCheckvar(request("ScrapId"),10)
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
	if searchKey="" then searchKey="title"

	param = "&page=" & page & "&searchKey=" & searchKey & "&searchString=" & searchString	'������ ����

	'// ���� ����
	set oScrap = new CMardyScrap
	oScrap.FRectScrapId = ScrapId

	oScrap.GetMardyScrapView

	'// ���� �̹��� ����
	set oScrapImage = new CMardyScrap
	oScrapImage.FRectScrapId = ScrapId

	oScrapImage.GetMardyScrapImageList
%>
<script language="javascript">
<!--
	// �ۻ���
	function GotoScrapDel(){
		if (confirm('�� �Խù��� ������ ���� �Ͻðڽ��ϱ�?\n\n�� ���� �Ŀ��� �ٽ� ���� �� �� �����ϴ�.')){
			document.frm_trans.mode.value = "delete_main";
			document.frm_trans.action="http://image.thefingers.co.kr/linkweb/doMardyScrap.asp";
			document.frm_trans.submit();
		}
	}

	// �������
	function GotoSubDel(sid){
		if (confirm('�����Ͻ� �ܰ踦 ���� �Ͻðڽ��ϱ�?')){
			document.frm_trans.subId.value = sid;
			document.frm_trans.mode.value = "delete_sub";
			document.frm_trans.action="http://image.thefingers.co.kr/linkweb/doMardyScrapSub.asp";
			document.frm_trans.submit();
		}
	}

	// ��â���� ���� ����
	function NewWindow(v)
	{
		  var p = (v);
		  w = window.open("http://thefingers.co.kr/photo_album/pop_photo.asp?img=" + v, "imageView", "left=10px,top=10px, width=560,height=400,status=no,resizable=yes,scrollbars=yes");
		  w.focus();
	}


	// ��� ���� ����
	function GotoUse(md)
	{
		switch(md)
		{
			case "use" :
				if (confirm('����Ʈ ��Ͽ� ��µǵ��� ���¸� "���"���� �����Ͻðڽ��ϱ�?')){
					FrameCHK.location="inc_Mardy_Use.asp?Idx=<%=ScrapId%>&mode=ScrapUse";
				}
				break;

			case "del" :
				if (confirm('����Ʈ���� �� �� ������ ���¸� "����"���� �����Ͻðڽ��ϱ�?')){
					FrameCHK.location="inc_Mardy_Use.asp?Idx=<%=ScrapId%>&mode=ScrapDel";
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
			<td height="26" align="left"><b>���� ��ũ�� �⺻ ����</b></td>
			<td height="26" align="right">
				<font color=gray>��뿩�� - </font>
				<%
					if oScrap.FItemList(0).Fisusing="N" then
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
	<td align="center" width="120" bgcolor="#DDDDFF">����</td>
	<td bgcolor="#FFFFFF"><%=oScrap.FItemList(0).Ftitle%></td>
</tr>
<!-- tr>
	<td align="center" width="120" bgcolor="#DDDDFF">Ÿ��Ʋ �̹���</td>
	<td bgcolor="#FFFFFF"><img src="<%=oScrap.FItemList(0).FimgTitle_full%>" style="border:1px solid #C0C0C0"></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">�� ����</td>
	<td bgcolor="#FFFFFF">
		<table border="0" class="a" cellpadding="0" cellspacing="0">
		<tr>
			<td width="284" valign="top">
				<img src="<%=oScrap.FItemList(0).FimgProd_full%>" width="282" onClick="NewWindow('<%=oScrap.FItemList(0).FimgProd_full%>')" style="cursor:pointer;border:1px solid #C0C0C0" alt="���� ����">
			</td>
			<td valign="top">
				<table border="0" class="a" cellpadding="0" cellspacing="0">
				<tr>
					<td><b>[<%=oScrap.FItemList(0).FscrName%>]</b></td>
				</tr>
				<tr>
					<td>[���̵�] <% for i=1 to oScrap.FItemList(0).FscrDef:Response.Write "��":next%></td>
				</tr>
				<tr>
					<td>[�ҿ�ð�] <%=oScrap.FItemList(0).FscrTime%></td>
				</tr>
				<tr>
					<td>[���] <%=oScrap.FItemList(0).FscrSource%></td>
				</tr>
				<tr>
					<td>[����] <%=oScrap.FItemList(0).FscrTool%></td>
				</tr>
				</table>
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">TIP</td>
	<td bgcolor="#FFFFFF"><%=nl2br(oScrap.FItemList(0).FscrTip)%></td>
</tr-->
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">��� ����</td>
	<td bgcolor="#FFFFFF">Type <%=oScrap.FItemList(0).FprintType%></td>
</tr>
<tr><td height="1" colspan="2" bgcolor="#D0D0D0"></td></tr>
<tr>
	<td colspan="2" height="32" bgcolor="#FAFAFA" align="center">
		<img src="/images/icon_modify.gif" onClick="self.location='mardyScrap_modi.asp?menupos=<%=menupos%>&ScrapId=<%=ScrapId & param%>'" style="cursor:pointer" align="absmiddle"> &nbsp;
		<% if oScrap.FItemList(0).Fisusing="N" then %>
		<img src="/images/icon_use.gif" onClick="GotoUse('use')" style="cursor:pointer" align="absmiddle"> &nbsp;
		<img src="/images/icon_delete.gif" onClick="GotoScrapDel()" style="cursor:pointer" align="absmiddle"> &nbsp;
		<% else %>
		<img src="/images/icon_hide.gif" onClick="GotoUse('del')" style="cursor:pointer" align="absmiddle"> &nbsp;
		<% end if %>
		<img src="/images/icon_list.gif" onClick="self.location='mardyScrap_list.asp?menupos=<%=menupos & param %>'" style="cursor:pointer" align="absmiddle">
	</td>
</tr>
<form name="frm_trans" method="POST" action="http://image.thefingers.co.kr/linkweb/doMardyScrapSub.asp" enctype="multipart/form-data">
<input type="hidden" name="ScrapId" value="<%=ScrapId%>">
<input type="hidden" name="subId" value="">
<input type="hidden" name="mode" value="delete_main">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="searchKey" value="<%=searchKey%>">
<input type="hidden" name="searchString" value="<%=searchString%>">
<input type="hidden" name="adminId" value="<%=Session("ssBctId")%>">
</form>
</table>
<!-- ���� ������ ��� ����  -->
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr align="center" bgcolor="#F0F0FD">
	<td colspan="4">
		<table width="100%" border="0" class="a" cellpadding="0" cellspacing="0">
		<tr>
			<td height="26" align="left"><b>����¹� �� �ܰ� ����</b></td>
			<td height="26" align="right">��ϰǼ� : <%= oScrapImage.FTotalCount %> ��</td>
		</tr>
		</table>
	</td>
</tr>
<tr align="center" bgcolor="#DDDDFF">
	<td width="40">��ȣ</td>
	<td width="50">�̹���</td>
	<td>�� ��</td>
	<td width="50">��Ʈ��</td>
</tr>
<%
	if oScrapImage.FTotalCount=0 then
%>
<tr><td height="46" colspan="4" align="center" bgcolor="#FFFFFF">��ϵ� ���� �����ϴ�.<br>�Ʒ� [�ű� ���] ��ư�� ���� �߰����ֽʽÿ�.</td></tr>
<%
	else
		for lp=0 to oScrapImage.FTotalCount-1
%>
<tr align="center" bgcolor="#FFFFFF">
	<td width="40"><%= oScrapImage.FItemList(lp).FsubSort %></td>
	<td width="10" align="left">
		<table border="0" align="center" cellpadding="0" cellspacing="0">
		<tr>
			<% if oScrapImage.FItemList(lp).FimgFile1<>"" then %><td width="104"><img src="<%= oScrapImage.FItemList(lp).FimgFile1_full %>" style="border:1px solid #C0C0C0"></td><% end if %>
			<% if oScrapImage.FItemList(lp).FimgFile2<>"" then %><td width="10"><img src="http://thefingers.co.kr/images/scrap_b_22.gif"></td><% end if %>
			<% if oScrapImage.FItemList(lp).FimgFile2<>"" then %><td width="104"><img src="<%= oScrapImage.FItemList(lp).FimgFile2_full %>" style="border:1px solid #C0C0C0"></td><% end if %>
			<% if oScrapImage.FItemList(lp).FimgFile3<>"" then %><td width="10"><img src="http://thefingers.co.kr/images/scrap_b_22.gif"></td><% end if %>
			<% if oScrapImage.FItemList(lp).FimgFile3<>"" then %><td width="104"><img src="<%= oScrapImage.FItemList(lp).FimgFile3_full %>" style="border:1px solid #C0C0C0"></td><% end if %>
		</tr>
		</table>
	</td>
	<td align="left">
		<%
			if oScrapImage.FItemList(lp).FsubName<>"" then
				Response.Write "<b>" & oScrapImage.FItemList(lp).FsubName & "</b><br><br>"
			end if
			Response.Write nl2br(oScrapImage.FItemList(lp).FsubCont)
		%>
	</td>
	<td width="50">
		<img src="/images/icon_modify.gif" onClick="self.location='mardyScrap_subModi.asp?menupos=<%=menupos%>&subId=<%=oScrapImage.FItemList(lp).FsubId%>&ScrapId=<%=ScrapId & param%>'" style="cursor:pointer" align="absmiddle" vspace="5"><br>
		<img src="/images/icon_delete.gif" onClick="GotoSubDel(<%=oScrapImage.FItemList(lp).FsubId%>)" style="cursor:pointer" align="absmiddle">
	</td>
</tr>
<%
		next
	end if
%>
<tr><td height="1" colspan="4" bgcolor="#D0D0D0"></td></tr>
<tr>
	<td colspan="4" height="32" bgcolor="#FAFAFA" align="right">
		<a href="mardyScrap_subWrite.asp?ScrapId=<%=ScrapId%>&menupos=<%=menupos & param%>"><img src="/images/icon_new_registration.gif" border="0" align="absmiddle"></a>
	</td>
</tr>
</table>
<iframe name="FrameCHK" src="" frameborder="0" width="0" height="0"></iframe>
<%
	Set oScrap = Nothing
	Set oScrapImage = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->