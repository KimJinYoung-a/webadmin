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
	dim oScrap, i, lp

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

	'// ���� ���� ����
	set oScrap = new CMardyScrap
	oScrap.FRectScrapId = ScrapId

	oScrap.GetMardyScrapView
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
		if(frm.imgTitle.value && !checkImageSuffix(frm.imgTitle))
		{
			return false;
		}

		if(!frm.title.value)
		{
			alert("������ �Է����ֽʽÿ�.");
			frm.title.focus();
			return false;
		}

		if(frm.imgProd.value && !checkImageSuffix(frm.imgProd))
		{
			return false;
		}

		if(!frm.scrName.value)
		{
			alert("��ǰ�̸��� �Է����ֽʽÿ�.");
			frm.scrName.focus();
			return false;
		}

		if(!frm.scrTime.value)
		{
			alert("�ҿ�ð��� �Է����ֽʽÿ�.");
			frm.scrTime.focus();
			return false;
		}
				
		if(frm.summ.value.length>150){
			alert("���ο�� ������ 150�� �̳��� �ۼ��� �ֽʽÿ�.");
			frm.summ.focus();
			return false;
		}

		// �� ����
		return true;

	}
//-->
</script>
<!-- ���� ȭ�� ���� -->
<table width="100%" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<form name="frm_write" method="POST" onSubmit="return chk_form(this)" action="http://image.thefingers.co.kr/linkweb/doMardyScrap.asp" enctype="multipart/form-data">
<input type="hidden" name="mode" value="modify_main">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="ScrapId" value="<%=ScrapId%>">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="searchKey" value="<%=searchKey%>">
<input type="hidden" name="searchString" value="<%=searchString%>">
<input type="hidden" name="adminId" value="<%=Session("ssBctId")%>">
<tr align="center" bgcolor="#F0F0FD">
	<td colspan="2">
		<table width="100%" border="0" class="a" cellpadding="0" cellspacing="0">
		<tr>
			<td height="26" align="left"><b>���� ��ũ�� ���� ����</b></td>
			<td height="26" align="right">
				<font color=gray><b>��뿩��</b></font>
				<select name="isusing">
					<option value="Y" <% if oScrap.FItemList(0).Fisusing="Y" then Response.Write "selected" %>>���</option>
					<option value="N" <% if oScrap.FItemList(0).Fisusing="N" then Response.Write "selected" %>>����</option>
				</select>&nbsp;
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>Ÿ��Ʋ �̹���</td>
	<td bgcolor="#FFFFFF">
		<table width="100%" border="0" class="a" cellpadding="0" cellspacing="0">
		<tr>
			<% if oScrap.FItemList(0).FimgTitle<>"" then %>
			<td width="124" align="center">
				<img src="<%=oScrap.FItemList(0).FimgTitle_full%>" style="width:120px;border:1px solid #C0C0C0">
			</td>
			<% end if %>
			<td>
				<% if oScrap.FItemList(0).FimgTitle<>"" then %>
					(���� : <%= oScrap.FItemList(0).FimgTitle%>)
				<% end if %>
				<input type="file" name="imgTitle" size="50"><br>
				<font color=darkred>�� �ʺ� 612px ũ���� JPG/GIF �̹��� ������ ���ε��Ͻʽÿ�.</font>
				<input type="hidden" name="orgTitle" value="<%=oScrap.FItemList(0).FimgTitle%>">
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>����</td>
	<td bgcolor="#FFFFFF"><input type="text" name="title" size="80" maxlength="120" value="<%=oScrap.FItemList(0).Ftitle%>"></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>�ϼ�ǰ �̹���</td>
	<td bgcolor="#FFFFFF">
		<table width="100%" border="0" class="a" cellpadding="0" cellspacing="0">
		<tr>
			<% if oScrap.FItemList(0).FimgProd<>"" then %>
			<td width="124" align="center">
				<img src="<%=oScrap.FItemList(0).FimgProd_full%>" style="width:120px;border:1px solid #C0C0C0">
			</td>
			<% end if %>
			<td>
				<% if oScrap.FItemList(0).FimgProd<>"" then %>
					(���� : <%= oScrap.FItemList(0).FimgProd%>)
				<% end if %>
				<input type="file" name="imgProd" size="50"><br>
				<font color=darkred>�� 10:7(����:����) ������ JPG/GIF �����Դϴ�.</font>
				<input type="hidden" name="orgProd" value="<%=oScrap.FItemList(0).FimgProd%>">
				<input type="hidden" name="orgIcon" value="<%=oScrap.FItemList(0).FimgIcon%>">
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr><td height="1" colspan="2" bgcolor="#D0D0D0"></td></tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>��ǰ��</td>
	<td bgcolor="#FFFFFF"><input type="text" name="scrName" size="60" maxlength="100" value="<%=oScrap.FItemList(0).FscrName%>"></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>���̵�</td>
	<td bgcolor="#FFFFFF">
		<select name="scrDef">
			<option value="1" <% if oScrap.FItemList(0).FscrDef="1" then Response.Write "selected" %>>[1]�ڡ١١١�</option>
			<option value="2" <% if oScrap.FItemList(0).FscrDef="2" then Response.Write "selected" %>>[2]�ڡڡ١١�</option>
			<option value="3" <% if oScrap.FItemList(0).FscrDef="3" then Response.Write "selected" %>>[3]�ڡڡڡ١�</option>
			<option value="4" <% if oScrap.FItemList(0).FscrDef="4" then Response.Write "selected" %>>[4]�ڡڡڡڡ�</option>
			<option value="5" <% if oScrap.FItemList(0).FscrDef="5" then Response.Write "selected" %>>[5]�ڡڡڡڡ�</option>
		</select>
	</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>�ҿ�ð�</td>
	<td bgcolor="#FFFFFF"><input type="text" name="scrTime" size="30" maxlength="60" value="<%=oScrap.FItemList(0).FscrTime%>"></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">���</td>
	<td bgcolor="#FFFFFF"><textarea name="scrSource" rows="2" cols="80"><%=oScrap.FItemList(0).FscrSource%></textarea></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">����</td>
	<td bgcolor="#FFFFFF"><textarea name="scrTool" rows="2" cols="80"><%=oScrap.FItemList(0).FscrTool%></textarea></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>��ũ�� ����</td>
	<td bgcolor="#FFFFFF">
		<table border="0" class="a" cellpadding="0" cellspacing="0">
		<tr>
			<td width="50%" align="center"><input type="radio" name="printType" value="A" <% if oScrap.FItemList(0).FprintType="A" then Response.Write "checked" %>> Type A</td>
			<td width="50%" align="center"><input type="radio" name="printType" value="B" <% if oScrap.FItemList(0).FprintType="B" then Response.Write "checked" %>> Type B</td>
		</tr>
		<tr>
			<td align="center"><img src="/images/tpl_typeA.gif" alt="������ �������� �̹����� �����ϴ� ����Դϴ�."></td>
			<td align="center"><img src="/images/tpl_typeB.gif" alt="������ �̹����� �����ϴ� ����Դϴ�."></td>
		</tr>
		<tr>
			<td align="center"><input type="radio" name="printType" value="C" <% if oScrap.FItemList(0).FprintType="C" then Response.Write "checked" %>> Type C</td>
			<td align="center"><input type="radio" name="printType" value="D" <% if oScrap.FItemList(0).FprintType="D" then Response.Write "checked" %>> Type D</td>
		</tr>
		<tr>
			<td align="center"><img src="/images/tpl_typeC.gif" alt="������ �̹����� �� �����ϴ� ����Դϴ�."></td>
			<td align="center"><img src="/images/tpl_typeD.gif" alt="HTML�� �ܺ��� �̹��� Ȥ�� ǥ�� ����Ͽ� �����Ӱ� �ۼ��մϴ�."></td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">TIP</td>
	<td bgcolor="#FFFFFF"><textarea name="scrTip" rows="4" cols="80"><%=oScrap.FItemList(0).FscrTip%></textarea></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">���ο��</td>
	<td bgcolor="#FFFFFF"><textarea name="summ" rows="2" cols="80"><%=oScrap.FItemList(0).Fsummary%></textarea> (150�� �̳�)</td>
</tr>
<tr><td height="1" colspan="2" bgcolor="#D0D0D0"></td></tr>
<tr>
	<td colspan="2" height="32" bgcolor="#FAFAFA" align="center">
		<input type="image" src="/images/icon_modify.gif" style="border:0px;cursor:pointer" align="absmiddle"> &nbsp;
		<img src="/images/icon_cancel.gif" onClick="history.back()" style="cursor:pointer" align="absmiddle">
	</td>
</tr>
</form>
</table>
<!-- ���� ȭ�� �� -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->