<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/classes/mardy_tipcls.asp"-->
<%
	'// ���� ���� //
	dim tipId
	dim page, searchKey, searchString, param

	dim oTip, oTipImage, i, lp

	'// �Ķ���� ���� //
	tipId = RequestCheckvar(request("tipId"),10)
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
	if searchKey="" then searchKey="tipName"

	param = "&page=" & page & "&searchKey=" & searchKey & "&searchString=" & searchString	'������ ����

	'// ���� ����
	set oTip = new CMardyTip
	oTip.FRectTipId = tipId

	oTip.GetMardyTipView

	'// ���� �̹��� ����
	set oTipImage = new CMardyTip
	oTipImage.FRectTipId = tipId

	oTipImage.GetMardyTipImageList
%>
<script language="javascript">
<!--
	// ����ȯ
	function chgTab(no)
	{
		var frm = document.all;
		var cnt = frm.tabico.length;
		
		if(cnt>1)
		{
			//��� ���� �����.
			for(i=0;i<cnt;i++)
			{
				frm.tab[i].style.display="none";
				frm.tabico[i].src="http://thefingers.co.kr/images/ico/tab_" + (i+1) + "_off.gif";
			}
			
			// ������ ���� ���δ�.
			frm.tab[no].style.display="";
			frm.tabico[no].src="http://thefingers.co.kr/images/ico/tab_" + (no+1) + "_on.gif";
		}
	}

	// �ۻ���
	function GotoTipDel(){
		if (confirm('�� �Խù��� ������ ���� �Ͻðڽ��ϱ�?\n\n�� ���� �Ŀ��� �ٽ� ���� �� �� �����ϴ�.')){
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
					FrameCHK.location="inc_Mardy_Use.asp?Idx=<%=tipId%>&mode=TipUse";
				}
				break;

			case "del" :
				if (confirm('����Ʈ���� �� �� ������ ���¸� "����"���� �����Ͻðڽ��ϱ�?')){
					FrameCHK.location="inc_Mardy_Use.asp?Idx=<%=tipId%>&mode=TipDel";
					//self.location="inc_Mardy_Use.asp?Idx=<%=tipId%>&mode=TipDel";
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
			<td height="26" align="left"><b>�����ø �� ����</b></td>
			<td height="26" align="right">
				<font color=gray>��뿩�� - </font>
				<%
					if oTip.FItemList(0).Fisusing="N" then
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
	<td align="center" width="120" bgcolor="#DDDDFF">�̹���</td>
	<td bgcolor="#FFFFFF">
		<table border="0" class="a" cellpadding="0" cellspacing="0">
		<tr>
			<td>
			<%
				for lp=0 to oTipImage.FTotalCount - 1
					if lp=0 then
						Response.Write "<img src='http://thefingers.co.kr/images/ico/tab_" & lp+1 & "_on.gif' id='tabico' onClick='chgTab(" & lp & ")' style='cursor:pointer' align='absmiddle'>"
					else
						Response.Write "<img src='http://thefingers.co.kr/images/ico/tab_" & lp+1 & "_off.gif' id='tabico' onClick='chgTab(" & lp & ")' style='cursor:pointer' align='absmiddle'>"
					end if
				next
			%>
			</td>
		</tr>
		<%
			for lp=0 to oTipImage.FTotalCount - 1
		%>
		<tr id="tab" <% if lp>0 then Response.Write "style='display:none'"%>>
			<td align="center"  bgcolor="#B6DD46" style="padding:5px">
				<img src="<%=oTipImage.FItemList(lp).FimgFile_full%>" width="560" onClick="NewWindow('<%=oTipImage.FItemList(lp).FimgFile_full%>')" style="cursor:pointer" alt="���� ����"><br><br>
				<%=nl2br(oTipImage.FItemList(lp).FimgCont)%>
			</td>
		</tr>
		<% next %>
		</table>
	</td>
</tr>
<tr><td height="1" colspan="2" bgcolor="#D0D0D0"></td></tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">�� ����</td>
	<td bgcolor="#FFFFFF"><%=oTip.FItemList(0).Ftitle%></td>
</tr>
<tr><td height="1" colspan="2" bgcolor="#D0D0D0"></td></tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">��ǰ��</td>
	<td bgcolor="#FFFFFF"><%=oTip.FItemList(0).FtipName%></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">��뵵</td>
	<td bgcolor="#FFFFFF"><%=oTip.FItemList(0).FtipUsage%></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">���̵�</td>
	<td bgcolor="#FFFFFF">
	<%
		for i=1 to oTip.FItemList(0).FtipDef
			Response.Write "��"
		next
	%>
	</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">�ҿ�ð�</td>
	<td bgcolor="#FFFFFF"><%=oTip.FItemList(0).FtipTime%></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">���</td>
	<td bgcolor="#FFFFFF"><%=oTip.FItemList(0).FtipPrice%></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">���ǻ���</td>
	<td bgcolor="#FFFFFF"><%=nl2br(oTip.FItemList(0).FtipAttent)%></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">Tip</td>
	<td bgcolor="#FFFFFF"><%=nl2br(oTip.FItemList(0).FtipCont)%></td>
</tr>
<tr><td height="1" colspan="2" bgcolor="#D0D0D0"></td></tr>
<tr>
	<td colspan="2" height="32" bgcolor="#FAFAFA" align="center">
		<img src="/images/icon_modify.gif" onClick="self.location='mardyTip_modi.asp?menupos=<%=menupos%>&tipId=<%=tipId & param%>'" style="cursor:pointer" align="absmiddle"> &nbsp;
		<% if oTip.FItemList(0).Fisusing="N" then %>
		<img src="/images/icon_use.gif" onClick="GotoUse('use')" style="cursor:pointer" align="absmiddle"> &nbsp;
		<img src="/images/icon_delete.gif" onClick="GotoTipDel()" style="cursor:pointer" align="absmiddle"> &nbsp;
		<% else %>
		<img src="/images/icon_hide.gif" onClick="GotoUse('del')" style="cursor:pointer" align="absmiddle"> &nbsp;
		<% end if %>
		<img src="/images/icon_list.gif" onClick="self.location='mardyTip_list.asp?menupos=<%=menupos & param %>'" style="cursor:pointer" align="absmiddle">
	</td>
</tr>
<form name="frm_trans" method="POST" action="http://image.thefingers.co.kr/linkweb/doMardyTip.asp" enctype="multipart/form-data">
<input type="hidden" name="tipId" value="<%=tipId%>">
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