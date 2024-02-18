<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'####################################################
' Description : AGV ��ǰ ���� �ϰ� ���ε�
' History : 2020.05.22 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/md5.asp"-->
<!-- #include virtual="/common/incSessionAdminorShop.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->

<script type="text/javascript">

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

	frm.target='view';
	frm.submit();
}

//�̹��� �������� Ÿ�� ����
function openerreload(){
	opener.location.reload();
	self.close();
}

</script>

<style>
html, body { margin: 0; padding: 0; }
</style>

<p />

<form name="frmFile" method="post" action="<%= uploadUrl %>/linkweb/item/agv/upload_AGV_stock_invest_item_excel.asp"  enctype="MULTIPART/FORM-DATA" style="margin:0px;">

<table width="98%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="#999999">
<tr bgcolor="#FFFFFF" align="center">
	<td colspan="2">
		<b>AGV ��ǰ ���� �ϰ� ���ε�</b>
	</td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
	<td width="60">����</td>
	<td align="left"><a href="<%= uploadUrl %>/offshop/sample/item/logics_agv_stock_invest_sample_v1.xls" target="_blank">logics_agv_stock_invest_sample_v1.xls</a></td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
	<td><font color="red"><b>���ǻ���</b></font></td>
	<td align="left">
		�� �������� ���������� <font color="red"><b>Save As Excel 97 -2003 ���չ���</b></font> ���¸� �ν� �մϴ�.
		<br><br>* ���� ������ �ٿ� �������� ����Ͻø� �˴ϴ�.
		<br><br>* �� ����� �����ڵ尡 �ִ� <font color="red"><b>1���� �״�� �μ���.</b></font>
		<br><br>* �����ڵ带 �������� ������Ʈ �Ǳ� ������ <font color="red"><b>�����ڵ�� ���� �����̰ų�, Ʋ���� �ȵ˴ϴ�.</b></font>
		<br>�����ڵ� �߰��� �ٽñ�ȣ(-) �� ���� �ϼŵ� ����� �˴ϴ�. ��)10-01239286-0000 -> 10012392860000
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

<p />

<iframe id="view" name="view" src="" width="98%" height=30 frameborder="0" scrolling="no" align="center" style="display:block;"></iframe>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
