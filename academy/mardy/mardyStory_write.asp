<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
	dim lp
	dim maxAddImg

	maxAddImg=5		'�߰� �̹��� �ִ�� ����
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
		if(!frm.titleShort.value)
		{
			alert("ª�� ������ �Է����ֽʽÿ�.");
			frm.titleShort.focus();
			return false;
		}

		if(!frm.titleLong.value)
		{
			alert("��(��) ������ �Է����ֽʽÿ�.");
			frm.titleLong.focus();
			return false;
		}

		if(!frm.imgIcon.value && !frm.imgFile[0].value)
		{
			alert("��Ͽ� ������ �������� ����\n��� �̹��� Ȥ�� ÷�� �̹����� �ϳ��� �ݵ�� �������ֽʽÿ�.\n\n�� �̹����� JPG, GIF�������� �������ֽʽÿ�.");
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
	function addImgControl(lid, md)
	{
		var frm = document.all;
		var btext = "";

		if(md=="add")
		{
			// ���̾� ���̱� �� ���� Ȱ��ȭ
			frm["addimg"+(lid+1)].style.display="";
			frm["btn"+(lid)].innerHTML = "";

			lid++;
			
			// ��ưó��
			if(lid>1)
				btext += "<img src='/images/icon_minus.gif' alt='�̹��� ����' onClick=addImgControl(" + lid + ",'del') style='cursor:pointer'> ";
			if(lid<<%=maxAddImg%>)
				btext += "<img src='/images/icon_plus.gif' alt='�̹��� �߰�' onClick=addImgControl(" + lid + ",'add') style='cursor:pointer'>";
			
			frm["btn"+lid].innerHTML = btext;
		}
		else if(md=="del")
		{
			// ���̾� ����� �� ���� ����
			frm.imgFile[lid-1].select();
			document.execCommand('Delete');
			frm["addimg"+lid].style.display="none";
			frm["btn"+lid].innerHTML = "";

			lid--;

			// ��ưó��
			if(lid>1)
				btext += "<img src='/images/icon_minus.gif' alt='�̹��� ����' onClick=addImgControl(" + lid + ",'del') style='cursor:pointer'> ";
			if(lid<<%=maxAddImg%>)
				btext += "<img src='/images/icon_plus.gif' alt='�̹��� �߰�' onClick=addImgControl(" + lid + ",'add') style='cursor:pointer'>";

			frm["btn"+lid].innerHTML = btext;
		}
	}
//-->
</script>
<!-- ���� ȭ�� ���� -->
<table width="100%" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<form name="frm_write" method="POST" onSubmit="return chk_form(this)" action="http://image.thefingers.co.kr/linkweb/doMardyStory.asp" enctype="multipart/form-data">
<input type="hidden" name="mode" value="write">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="adminId" value="<%=Session("ssBctId")%>">
<tr align="center" bgcolor="#F0F0FD">
	<td height="26" align="left" colspan="2"><b>���� �̾߱� �ű� ���</b></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>ª�� ����</td>
	<td bgcolor="#FFFFFF"><input type="text" name="titleShort" size="40" maxlength="40"></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>�� ����</td>
	<td bgcolor="#FFFFFF"><input type="text" name="titleLong" size="80" maxlength="120"></td>
</tr>
<tr><td height="1" colspan="2" bgcolor="#D0D0D0"></td></tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">��� �̹���</td>
	<td bgcolor="#FFFFFF">
		<input type="file" name="imgIcon" size="60"><br>
		<font color=darkred>�� 50px * 50px ũ���� JPG/GIF �̹��� ������ ���ε��Ͻʽÿ�.</font>
	</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">���� ����</td>
	<td bgcolor="#FFFFFF"><textarea name="contents" rows="2" cols="80"></textarea></td>
</tr>
<%
	'// �߰� �̹��� �� �ۼ�
	for lp=1 to maxAddImg
%>
<tr id="addimg<%=lp%>" <% if lp>1 then Response.Write "style='display:none;'"%>>
	<td align="center" width="120" bgcolor="#DDDDFF">�̹��� #<%=lp%></td>
	<td bgcolor="#FFFFFF">
		<table width="100%" cellpadding="0" cellspacing="0" border="0" class="a">
		<tr>
			<td><input type="file" name="imgFile" size="50"></td>
			<td width="36" align="right" valign="bottom" id="btn<%=lp%>">
				<% if lp=1 then %><img src="/images/icon_plus.gif" alt="�̹��� �߰�" onClick="addImgControl(<%=lp%>,'add')" style="cursor:pointer"><% end if %>
			</td>
		</tr>
		</table>
	</td>
</tr>
<% next %>
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