<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  ������湮ī��Ʈ
' History : 2012.05.10 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/guest/shop_guestcount_cls.asp"-->
<%
dim iMaxLength , memupos
	memupos = requestCheckVar(request("memupos"),10)
	iMaxLength = 5
%>

<script language="javascript">

function fnChkFile(sFile, arrExt){
    //���� ���ε� ����Ȯ��
     if (!sFile){
    	 return true;
    	}

    var blnResult = false;

    //���� Ȯ���� Ȯ��
	var pPoint = sFile.lastIndexOf('.');
	var fPoint = sFile.substring(pPoint+1,sFile.length);
	var fExet = fPoint.toLowerCase();

	for (var i = 0; i < arrExt.length; i++)
	   	{
	    	if (arrExt[i].toLowerCase() == fExet)
	    	{
	   			blnResult =  true;
	   		}
		}

	return blnResult;
}

function XLSumbit(){
	var frm = document.frmFile;
    
	arrFileExt = new Array();			
	arrFileExt[arrFileExt.length]  = "xls";
	
	if (frm.sFile.value==''){
		alert('������ �Է��� �ּ���');
		frm.sFile.focus();
		return;
	}
	
	//������ȿ�� üũ
	if (!fnChkFile(frm.sFile.value, arrFileExt)){
		alert("������ xls���ϸ� ���ε� �����մϴ�.");
		return;
	}
	
	frm.submit();
}

</script>

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmFile" method="post" action="<%=uploadUrl%>/linkweb/offshop/guest/shopguestcount/shop_guestcount_upload.asp"  enctype="MULTIPART/FORM-DATA">
<input type="hidden" name="iML" value="<%=iMaxLength%>">
<input type="hidden" name="mode" value="excelupload">
<tr bgcolor="#FFFFFF" align="center">
	<td colspan="2">
		<b>���� ���湮 �������</b>
	</td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
	<td>�ʵ�!!</td>
	<td align="left">
		���� ���湮 Ŭ���̾�Ʈ(remote manager)�� ����,
		<Br><Br>�α���(����Ÿ ������ �α��� üũ�� ��й�ȣ 1)�Ϸ���, ���� �ϴܿ� ����ī��Ʈ�� Ŭ��,
		<br><Br>��� �����ʿ� �����͸���Ʈ Ŭ����, �׷����� �ð��뺰 ����Ʈ�� ����,
		<br><Br>�ϴܿ� �����(����)�� ����� ���� ������ ���� Ŭ��,
		<Br><Br>�׷����� ���������� Ŭ���ؼ� �ٿ�ε���, �ٿ�ε� ���� ������,
		<br><Br><font color="red">���ο� ������Ʈ�� ����&�ٿ��ֱ��� Excel 97 -2003 ���չ����� ����</font>�� ����Ͻø� �˴ϴ�
	</td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
	<td>����</td>
	<td align="left"><a href="/common/offshop/guest/sample.xls" onfocus="this.blur()">sample.xls</a></td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
	<td>���ϸ�:</td>
	<td align="left"><input type="file" name="sFile" class="button"></td>
</tr>				
<tr bgcolor="#FFFFFF" align="center">
	<td colspan="2">
	    <input type="button" class="button" value="���" onClick="XLSumbit();">
	    <input type="button" class="button" value="���" onClick="self.close();">
	</td>
</tr>
</form>	
</table>

<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/common/lib/poptail.asp"-->