<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'####################################################
' Description :  ������ ������ ���� ���ε� �ϰ� ����
' History : 2018.09.28 ������ ����
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

<form name="frmFile" method="post" action="<%= uploadUrl %>/linkweb/offshop/upload_shopitemlist_excel.asp"  enctype="MULTIPART/FORM-DATA" style="margin:0px;">

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="#999999">
<tr bgcolor="#FFFFFF" align="center">
	<td colspan="2">
		<b>�������� ��ǰ ���� �ϰ� ���ε� ����</b>
	</td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
	<td width="60">����</td>
	<td align="left"><a href="<%= uploadUrl %>/offshop/sample/item/shop_item_list_sample_v1.xls" target="_blank">shop_item_list_sample.xls</a></td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
	<td><font color="red"><b>���ǻ���</b></font></td>
	<td align="left">
		�� �������� ���������� <font color="red"><b>Save As Excel 97 -2003 ���չ���</b></font> ���¸� �ν� �մϴ�.
		<br><br>* ���� ������ �ٿ� ��������, �Է� �Ͻðų�, ����Ʈ�� �ִ� <font color="red"><b>�����ٿ�ε�</b></font>�� �����ż� ����Ͻø� �˴ϴ�.
		<br><br>* �� ����� �귣��ID, ��ǰ�ڵ�, ��ǰ�� ���� �ִ� <font color="red"><b>1���� �״�� �μ���.</b></font>
		<br><br>* ��ǰ�ڵ带 �������� ������Ʈ �Ǳ� ������ <font color="red"><b>��ǰ�ڵ�� ���� �����̰ų�, Ʋ���� �ȵ˴ϴ�.</b></font>
		<br><br><font color="red"><b>* �Һ��ڰ�, �ǸŰ�, ���԰�, ������ް�, ���͸��Ա���, ������ڵ�, ��뿩��, ON/OFF ���ݿ���</b></font>�ʵ带 �Է��Ͻø� �Ǹ�, �״�� ���� �˴ϴ�.
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

<iframe id="view" name="view" src="" width=1280 height=30 frameborder="0" scrolling="no"></iframe>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->