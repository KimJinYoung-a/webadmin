<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/md5.asp"-->
<!-- #include virtual="/common/incSessionAdminorShop.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->

<script language="javascript">
document.domain = "10x10.co.kr";

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

<form name="frmFile" method="post" action="<%=uploadUrl%>/linkweb/offshop/workschedule/upload_work_schedule_proc.asp"  enctype="MULTIPART/FORM-DATA">
<input type="hidden" name="reguserid" value="<%=session("ssBctId")%>">
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="#999999">
<tr bgcolor="#FFFFFF" align="center">
	<td colspan="2">
		<b>������������ ������ �������</b>
	</td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
	<td width="60">����</td>
	<td align="left"><a href="/common/offshop/staff/schedule.xls" onfocus="this.blur()">schedule.xls</a></td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
	<td>���ǻ���</td>
	<td align="left">
	* �� ����� ��, �� ���� �ִ� <b>1���� �״�� �μ���.</b><br>
	* �̸��� �߸� �Է��ص� �����ϳ� <b>����� ���� Ʋ���� �ȵ˴ϴ�.</b> �̸��� ���� ������ ���������� <b>����� �����Ͽ� ��� �����͸� ������� ��ȸ�Ǳ⿡ �����ؼ� �Է�</b>�ϱ� �ٶ��ϴ�.<br>
	* 1 ~ 31 ĭ���� �ش��Ͽ� �ش�� �����ڵ带 �����ø� �˴ϴ�. <b>������ �����ڵ� �ܿ� �Է½� ������ ���ϴ�.</b><br>
	* ��, ���� �ش�Ǵ� ��¥�� �ݵ�� <b>�޷¿� �ִ� �״���� �� ��</b> ��ŭ �Է��ϼ���.
	</td>
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
</table>
</form>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->