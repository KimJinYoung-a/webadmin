<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : �������� ������
' History : ���ʻ����ڸ�
'			2017.04.13 �ѿ�� ����(���Ȱ���ó��)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
dim yyyy1,mm1,dd1
dim nowdate

nowdate = Left(CStr(now()),10)

yyyy1 = requestCheckVar(request("yyyy1"),4)
mm1 = requestCheckVar(request("mm1"),2)
dd1 = requestCheckVar(request("dd1"),2)

if (yyyy1="") then
	yyyy1 = Left(nowdate,4)
	mm1   = Mid(nowdate,6,2)
	dd1   = Mid(nowdate,9,2)
end if

%>
<script language="JavaScript">
<!--

function checkok(frm){
      frm.submit();
}

//-->
</script>
<table cellpadding="0" cellspacing="0" border="1" align="center" bordercolordark="White" bordercolorlight="black">
<form method="post" name="monthly" action="<%=uploadUrl%>/ftp/offshop_mailzine_ok.asp" enctype="multipart/form-data">
<input type="hidden" name="nothing" value="">
<input type="hidden" name="mode" value="write">
<tr class="a">
	<td align="center" height="35" colspan="2"><b>������ �ۼ�</b></td>
</tr>
<tr class="a">
	<td align="center" width="100" height="35">������ �����</td>
	<td>&nbsp;<% DrawOneDateBox yyyy1,mm1,dd1 %></td>
</tr>
<tr class="a">
	<td align="center" width="100" height="35">��������</td>
	<td>&nbsp;<input type="text" name="title" class="input" size="55"></td>
</tr>
<tr class="a">
	<td align="center" height="25" colspan="2">��������<br><font color="#FF7D7D">(���� �������� ������ ���̸�ŭ ����˴ϴ�. �ٹٲ��� �� ��������ּ���)</font></td>
</tr>
<tr>
	<td colspan="2">
	   <table border="0" cellpadding="0" cellspacing="0" class="a">
	   <tr>
		<td>
			<textarea name="news" rows="10" cols="75" class="textarea"></textarea>
		</td>
	   </tr>
	   </table>
	</td>
</tr>
<tr class="a">
	<td align="center" width="100" height="50">�ΰŽ� �̹���</td>
	<td>&nbsp;<input type="file" name="img1" class="input" size="40"><br>&nbsp;<input type="text" name="url1" class="input" size="60"></td>
</tr>
<tr class="a">
	<td align="center" width="100" height="35">MD��õ��ǰ���</td>
	<td>&nbsp;<input type="file" name="img2" class="input" size="40"></td>
</tr>
<tr class="a">
	<td align="center" width="100" height="50">�̺�Ʈ���1</td>
	<td>&nbsp;<input type="file" name="img3" class="input" size="40"><br>&nbsp;<input type="text" name="url2" class="input" size="60"></td>
</tr>
<tr class="a">
	<td align="center" width="100" height="50">�̺�Ʈ���2</td>
	<td>&nbsp;<input type="file" name="img4" class="input" size="40"><br>&nbsp;<input type="text" name="url3" class="input" size="60"></td>
</tr>
<tr class="a">
	<td align="center" width="100" height="50">�̺�Ʈ���3</td>
	<td>&nbsp;<input type="file" name="img5" class="input" size="40"><br>&nbsp;<input type="text" name="url4" class="input" size="60"></td>
</tr>
<tr class="a">
	<td align="center" width="100" height="35">POP�̹���</td>
	<td>&nbsp;<input type="file" name="img6" class="input" size="40"></td>
</tr>
<tr class="a">
	<td align="center" width="100" height="35">POP�̹���Big</td>
	<td>&nbsp;<input type="file" name="img7" class="input" size="40"></td>
</tr>
<tr class="a">
	<td align="center" width="100" height="35">��õ�귣�� 6</td>
	<td>&nbsp;<input type="text" name="brand" class="input" size="60">&nbsp;<input type="button" class="button" value="�̹����ø���" onclick="TnFtpUpload('D:/home/cube1010/imgstatic/main/brand/','/main/brand/');"><br>(�������� �޸�(,)�� �־��ּ��� ex:mmmg,ia,heewoo,)</td>
</tr>
<tr>
	<td align="right" colspan="2" height="30"><input type="button" value="������ ���" onclick="checkok(this.form);" class="button">&nbsp;&nbsp;&nbsp;</td>
</tr>
</form>
</table>


<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->