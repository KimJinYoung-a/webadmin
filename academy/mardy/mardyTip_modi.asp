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
	dim page, searchKey, searchString

	dim oTip, oTipImage, i, lp, maxAddImg

	maxAddImg=10		'�߰� �̹��� �ִ�� ����

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

	'// ���� ����
	set oTip = new CMardyTip
	oTip.FRectTipId = tipId

	oTip.GetMardyTipView
%>
<script language='javascript'>
<!--
	// �̹��� ���� �˻�
	function checkImageSuffix (fileInput)
	{
	   var suffixPattern = /(gif|jpg|jpeg|png)$/i;

	   if (!suffixPattern.test(fileInput.value)) {
	     alert('GIF,JEPG,PNG ���ϸ� �����մϴ�.');
	     fileInput.focus();
	     fileInput.select();
	     return false;
	   }
	   return true;
	}

	// �Է��� �˻�
	function chk_form(frm)
	{
		if(!frm.title.value)
		{
			alert("������ �Է����ֽʽÿ�.");
			frm.title.focus();
			return false;
		}

		if(!frm.tipName.value)
		{
			alert("��ǰ�̸��� �Է����ֽʽÿ�.");
			frm.tipName.focus();
			return false;
		}

		if(!frm.tipUsage.value)
		{
			alert("��뵵�� �Է����ֽʽÿ�.");
			frm.tipUsage.focus();
			return false;
		}

		if(!frm.tipTime.value)
		{
			alert("�ҿ�ð��� �Է����ֽʽÿ�.");
			frm.tipTime.focus();
			return false;
		}

		for(var i=0;i<frm.imgFile.length;i++)
		{
			if ((frm.imgFile[i].value.length>0)&&(!checkImageSuffix(frm.imgFile[i]))){
				return false;
			}
		}

		// �� ����
		return true;

	}

	// �߰� �̹����� ����
	function addImgControl(lid, md, mx)
	{
		var frm = document.all;
		var btext = "";

		if(md=="add")
		{
			// ���̾� ���̱� �� ���� Ȱ��ȭ
			frm["addimg"+(lid+1)].style.display="";
			frm["btn"+(lid)].innerHTML = "";
			if(lid<mx)
				frm.filedelete[lid].value = "";

			lid++;

			// ��ưó��
			if(lid>1)
				btext += "<img src='/images/icon_minus.gif' alt='�̹��� ����' onClick=addImgControl(" + lid + ",'del'," + mx + ") style='cursor:pointer'> ";
			if(lid<<%=maxAddImg%>)
				btext += "<img src='/images/icon_plus.gif' alt='�̹��� �߰�' onClick=addImgControl(" + lid + ",'add'," + mx + ") style='cursor:pointer'>";

			frm["btn"+lid].innerHTML = btext;
		}
		else if(md=="del")
		{
			// ���̾� ����� �� ���� ����
			frm.imgFile[lid-1].select();
			document.execCommand('Delete');

			if(lid>mx)
				frm.imgCont[lid-1].value = "";
			else
				frm.filedelete[lid-1].value = "Y";

			frm["addimg"+lid].style.display="none";
			frm["btn"+lid].innerHTML = "";

			lid--;

			// ��ưó��
			if(lid>1)
				btext += "<img src='/images/icon_minus.gif' alt='�̹��� ����' onClick=addImgControl(" + lid + ",'del'," + mx + ") style='cursor:pointer'> ";
			if(lid<<%=maxAddImg%>)
				btext += "<img src='/images/icon_plus.gif' alt='�̹��� �߰�' onClick=addImgControl(" + lid + ",'add'," + mx + ") style='cursor:pointer'>";

			frm["btn"+lid].innerHTML = btext;
		}
	}
//-->
</script>
<!-- ���� ȭ�� ���� -->
<table width="100%" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<form name="frm_write" method="POST" onSubmit="return chk_form(this)" action="http://image.thefingers.co.kr/linkweb/doMardyTip.asp" enctype="multipart/form-data">
<input type="hidden" name="tipId" value="<%=tipId%>">
<input type="hidden" name="mode" value="modify">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="searchKey" value="<%=searchKey%>">
<input type="hidden" name="searchString" value="<%=searchString%>">
<input type="hidden" name="adminId" value="<%=Session("ssBctId")%>">
<tr align="center" bgcolor="#F0F0FD">
	<td colspan="2">
		<table width="100%" border="0" class="a" cellpadding="0" cellspacing="0">
		<tr>
			<td height="26" align="left"><b>�����ø ���� ����</b></td>
			<td height="26" align="right">
				<font color=gray><b>��뿩��</b></font>
				<select name="isusing">
					<option value="Y">���</option>
					<option value="N">����</option>
				</select>&nbsp;
				<script language="javascript">
					document.frm_write.isusing.value="<%=oTip.FItemList(0).Fisusing%>";
				</script>
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>����</td>
	<td bgcolor="#FFFFFF"><input type="text" name="title" size="80" maxlength="120" value="<%=oTip.FItemList(0).Ftitle%>"></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">��� �̹���</td>
	<td bgcolor="#FFFFFF">
		<table width="100%" cellpadding="0" cellspacing="0" border="0" class="a">
		<tr>
			<td width="52"><img src="<%=oTip.FItemList(0).FimgIcon_full%>" width="50" height="50"></td>
			<td>
				<input type="file" name="imgIcon" size="60"><br>
				<font color=darkred>�� 50px * 50px ũ���� JPG/GIF �̹��� ������ ���ε��Ͻʽÿ�.</font>
				<input type="hidden" name="orgIcon" value="<%=oTip.FItemList(0).FimgIcon%>">
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr><td height="1" colspan="2" bgcolor="#D0D0D0"></td></tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>��ǰ��</td>
	<td bgcolor="#FFFFFF"><input type="text" name="tipName" size="60" maxlength="100" value="<%=oTip.FItemList(0).FtipName%>"></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>��뵵</td>
	<td bgcolor="#FFFFFF"><input type="text" name="tipUsage" size="40" maxlength="60" value="<%=oTip.FItemList(0).FtipUsage%>"></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>���̵�</td>
	<td bgcolor="#FFFFFF">
		<select name="tipDef">
			<option value="1">[1]�ڡ١١١�</option>
			<option value="2">[2]�ڡڡ١١�</option>
			<option value="3">[3]�ڡڡڡ١�</option>
			<option value="4">[4]�ڡڡڡڡ�</option>
			<option value="5">[5]�ڡڡڡڡ�</option>
		</select>
		<script language="javascript">
			document.frm_write.tipDef.value="<%=oTip.FItemList(0).FtipDef%>";
		</script>
	</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>�ҿ�ð�</td>
	<td bgcolor="#FFFFFF"><input type="text" name="tipTime" size="30" maxlength="60" value="<%=oTip.FItemList(0).FtipTime%>"></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">���</td>
	<td bgcolor="#FFFFFF"><input type="text" name="tipPrice" size="40" maxlength="100" value="<%=oTip.FItemList(0).FtipPrice%>"></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">���ǻ���</td>
	<td bgcolor="#FFFFFF"><textarea name="tipAttent" rows="3" cols="80"><%=oTip.FItemList(0).FtipAttent%></textarea></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">Tip</td>
	<td bgcolor="#FFFFFF"><textarea name="tipCont" rows="3" cols="80"><%=oTip.FItemList(0).FtipCont%></textarea></td>
</tr>
<tr><td height="1" colspan="2" bgcolor="#D0D0D0"></td></tr>
<%
	'// �߰� �̹��� �� �ۼ�
	set oTipImage = new CMardyTip
	oTipImage.FRectTipId = tipId

	oTipImage.GetMardyTipImageList

	'// ���� ���
	for i=0 to oTipImage.FTotalCount - 1
		lp = lp+1
%>
<tr id="addimg<%=lp%>">
	<td align="center" width="120" bgcolor="#DDDDFF"><% if lp=1 then %><font color="darkred">* </font><% end if %>�̹��� #<%=lp%></td>
	<td bgcolor="#FFFFFF">
		<table width="100%" cellpadding="0" cellspacing="0" border="0" class="a">
		<tr>
			<td width="76" rowspan="2"><img src="<%=oTipImage.FItemList(i).FimgFile_full%>" width="74" height="74"></td>
			<td width="30" align="center">����</td>
			<td colspan="2">
				<input type="file" name="imgFile" size="50">
				<input type="hidden" name="orgImgId" value="<%=oTipImage.FItemList(i).FimgId%>">
				<input type="hidden" name="orgFile" value="<%=oTipImage.FItemList(i).FimgFile%>">
				<input type="hidden" name="filedelete" value="">
			</td>
		</tr>
		<tr>
			<td width="30" align="center">����</td>
			<td><textarea name="imgCont" rows="3" cols="70"><%=oTipImage.FItemList(i).FimgCont%></textarea></td>
			<td width="36" align="right" valign="bottom" id="btn<%=lp%>">
				<% if lp>1 then %><img src="/images/icon_minus.gif" alt="�̹��� ����" onClick="addImgControl(<%=lp%>,'del',<%=oTipImage.FTotalCount%>)" style="cursor:pointer"> <% end if %>
				<% if lp=oTipImage.FTotalCount then %><img src="/images/icon_plus.gif" alt="�̹��� �߰�" onClick="addImgControl(<%=lp%>,'add', <%=oTipImage.FTotalCount%>)" style="cursor:pointer"><% end if %>
			</td>
		</tr>
		</table>
	</td>
</tr>
<%
	next

	'// ���� ��ĭ
	for lp=i+1 to maxAddImg
%>
<tr id="addimg<%=lp%>" style='display:none;'>
	<td align="center" width="120" bgcolor="#DDDDFF"><% if lp=1 then %><font color="darkred">* </font><% end if %>�̹��� #<%=lp%></td>
	<td bgcolor="#FFFFFF">
		<table width="100%" cellpadding="0" cellspacing="0" border="0" class="a">
		<tr>
			<td width="30" align="center">����</td>
			<td colspan="2">
				<input type="file" name="imgFile" size="50">
				<input type="hidden" name="orgImgId" value="">
				<input type="hidden" name="orgFile" value="">
				<input type="hidden" name="filedelete" value="">
			</td>
		</tr>
		<tr>
			<td width="30" align="center">����</td>
			<td><textarea name="imgCont" rows="3" cols="70"></textarea></td>
			<td width="36" align="right" valign="bottom" id="btn<%=lp%>"></td>
		</tr>
		</table>
	</td>
</tr>
<%	next %>
<tr><td height="1" colspan="2" bgcolor="#D0D0D0"></td></tr>
<tr>
	<td colspan="2" height="32" bgcolor="#FAFAFA" align="center">
		<input type="image" src="/images/icon_save.gif" style="border:0px;cursor:pointer" align="absmiddle"> &nbsp;
		<img src="/images/icon_cancel.gif" onClick="history.back()" style="cursor:pointer" align="absmiddle">
	</td>
</tr>
</form>
</table>
<!-- ���� ȭ�� �� -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->