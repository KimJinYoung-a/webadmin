<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
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
		if(!frm.imgTitle.value)
		{
			alert("Ÿ��Ʋ �̹����� �������ֽʽÿ�.\n\n�� �̹����� JPG, GIF�������� �������ֽʽÿ�.");
			return false;
		}
		else if(!checkImageSuffix(frm.imgTitle))
		{
			return false;
		}

		if(!frm.title.value)
		{
			alert("������ �Է����ֽʽÿ�.");
			frm.title.focus();
			return false;
		}

		if(!frm.imgProd.value)
		{
			alert("�ϼ��� ��ǰ�� �̹����� �������ֽʽÿ�.\n\n�� �̹����� JPG, GIF�������� �������ֽʽÿ�.");
			return false;
		}
		else if(!checkImageSuffix(frm.imgProd))
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
<input type="hidden" name="mode" value="wirte_main">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="adminId" value="<%=Session("ssBctId")%>">
<tr align="center" bgcolor="#F0F0FD">
	<td height="26" align="left" colspan="2"><b>���� ��ũ�� �ű� ���</b></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>Ÿ��Ʋ �̹���</td>
	<td bgcolor="#FFFFFF">
		<input type="file" name="imgTitle" size="60"><br>
		<font color=darkred>�� �ʺ� 612px ũ���� JPG/GIF �̹��� ������ ���ε��Ͻʽÿ�.</font>
	</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>����</td>
	<td bgcolor="#FFFFFF"><input type="text" name="title" size="80" maxlength="120"></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>�ϼ�ǰ �̹���</td>
	<td bgcolor="#FFFFFF">
		<input type="file" name="imgProd" size="60">
		<font color=darkred>�� 10:7(����:����) ������ JPG/GIF �����Դϴ�.</font>
	</td>
</tr>
<tr><td height="1" colspan="2" bgcolor="#D0D0D0"></td></tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>��ǰ��</td>
	<td bgcolor="#FFFFFF"><input type="text" name="scrName" size="60" maxlength="100"></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>���̵�</td>
	<td bgcolor="#FFFFFF">
		<select name="scrDef">
			<option value="1">[1]�ڡ١١١�</option>
			<option value="2" selected>[2]�ڡڡ١١�</option>
			<option value="3">[3]�ڡڡڡ١�</option>
			<option value="4">[4]�ڡڡڡڡ�</option>
			<option value="5">[5]�ڡڡڡڡ�</option>
		</select>
	</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>�ҿ�ð�</td>
	<td bgcolor="#FFFFFF"><input type="text" name="scrTime" size="30" maxlength="60"></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">���</td>
	<td bgcolor="#FFFFFF"><textarea name="scrSource" rows="2" cols="80"></textarea></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">����</td>
	<td bgcolor="#FFFFFF"><textarea name="scrTool" rows="2" cols="80"></textarea></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>��ũ�� ����</td>
	<td bgcolor="#FFFFFF">
		<table border="0" class="a" cellpadding="0" cellspacing="0">
		<tr>
			<td width="50%" align="center"><input type="radio" name="printType" value="A" checked> Type A</td>
			<td width="50%" align="center"><input type="radio" name="printType" value="B"> Type B</td>
		</tr>
		<tr>
			<td align="center"><img src="/images/tpl_typeA.gif" alt="������ �������� �̹����� �����ϴ� ����Դϴ�."></td>
			<td align="center"><img src="/images/tpl_typeB.gif" alt="������ �̹����� �����ϴ� ����Դϴ�."></td>
		</tr>
		<tr>
			<td align="center"><input type="radio" name="printType" value="C"> Type C</td>
			<td align="center"><input type="radio" name="printType" value="D"> Type D</td>
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
	<td bgcolor="#FFFFFF"><textarea name="scrTip" rows="4" cols="80"></textarea></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">���ο��</td>
	<td bgcolor="#FFFFFF"><textarea name="summ" rows="2" cols="80"></textarea> (150�� �̳�)</td>
</tr>
<tr><td height="1" colspan="2" bgcolor="#D0D0D0"></td></tr>
<tr>
	<td colspan="2" height="32" bgcolor="#FAFAFA" align="center">
		<input type="image" src="/images/icon_next.gif" style="border:0px;cursor:pointer" align="absmiddle"> &nbsp;
		<img src="/images/icon_cancel.gif" onClick="history.back()" style="cursor:pointer" align="absmiddle">
	</td>
</tr>
</form>
</table>
<!-- ���� ȭ�� �� -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->