<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ��ǰ ������������ �ϰ����� Excel ���ε�
' Hieditor : 2015.05.22 ������ ����
'###########################################################
%>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<link rel="stylesheet" href="/css/tpl.css" type="text/css">
<script type="text/javascript">
<!--
	function fnFileDownload() {
		window.open("/designer/itemmaster/itemInfoFile/infoFileDownload.asp?fn=990");
	}

	function fnFileDownloadWithItem() {
		window.open("/designer/itemmaster/itemInfoFile/item_safetyInfo_xls.asp");
	}

	function XLSumbit() {
		var frm = document.frmFile;
		//���� Ȯ��
		if(!frm.sFile.value){
			alert("������ �Է��Ͻʽÿ�.");
			frm.sFile.focus();
			return;
		}

		arrFileExt = new Array();			
		arrFileExt[arrFileExt.length]  = "xls";

		//������ȿ�� üũ
		if (!fnChkFile(frm.sFile.value, arrFileExt)){
			alert("������ ����(*.xls)���ϸ� ���ε� �����մϴ�.");
			return;
		}

		if(confirm("�����Ͻ� ���Ϸ� [�������� ���] ������ �ϰ� ����Ͻðڽ��ϱ�?")) {
			frm.submit();
		}
	}

	function fnChkFile(sFile, arrExt) {
		//���� ���ε� ����Ȯ��
		if (!sFile) return true;

		var blnResult = false;

		//���� Ȯ���� Ȯ��
		var pPoint = sFile.lastIndexOf('.');
		var fPoint = sFile.substring(pPoint+1,sFile.length);
		var fExet = fPoint.toLowerCase();

		for (var i = 0; i < arrExt.length; i++) {
			if (arrExt[i].toLowerCase() == fExet) {
				blnResult =  true;
			}
		}
		return blnResult;
	}
//-->
</script>
<form name="frmFile" method="post" action="itemSafetyInfoFileUpload_process.asp"  enctype="MULTIPART/FORM-DATA">
<table width="100%" align="left" cellpadding="3" cellspacing="0" class="table_tl">
<tr height="25">
	<td class="td_br" colspan="2">
		<b>��ǰ [�������� ���] ���� �뷮���</b>
	</td>
</tr>
<tr height="30">
	<td width="90" align="right" class="td_br_tablebar">�ٿ�ε�:</td>
	<td class="td_br">
		<input type="button" class="button" value="��� �ٿ�ε�" onclick="fnFileDownload()">
		<input type="button" class="button" value="���+��ǰ���" onclick="fnFileDownloadWithItem()">
	</td>
</tr>
<tr height="30">
	<td width="90" align="right" class="td_br_tablebar">���ε�:</td>
	<td class="td_br">
		<input type="file" name="sFile" class="file" style="width:350px;">
	</td>
</tr>
<tr>
    <td colspan="2">
    	* <b>������������(KC��ũ)</b>�� �Է°����մϴ�.<br />
    	* ���� �ٿ�ε� > �躯�泻����� > �����Ͼ��ε�<br />
    	* �ݵ�� �� ���ε� ������� ���ε� (���¸� ������������)<br />
    	* ���� ���ε� ���� ���� (Email : kobula@10x10.co.kr �ش� ���� ÷�� �� ���� ���ּ���)
    </td>
</tr>
<tr>
	<td align="center" colspan="2" class="td_br">
	    <input type="button" class="button" value=" �� �� " onClick="XLSumbit();" style="background-color:#FFDDDD"> &nbsp;
	    <input type="button" class="button" value=" ��� " onClick="self.close();">
	</td>
</tr>
</table>
</form>
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->