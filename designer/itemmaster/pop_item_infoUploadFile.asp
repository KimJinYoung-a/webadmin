<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ��ǰ ǰ������ �ϰ����� Excel ���ε�
' Hieditor : 2012.10.25 ������ ����
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
		var frm = document.frmFile;
		if(!frm.infoDiv.value) {
			alert("�ٿ�ε� �Ͻ� ǰ�������� �������ֽʽÿ�.")
			frm.infoDiv.focus();
			return;
		}
		window.open("/designer/itemmaster/itemInfoFile/infoFileDownload.asp?fn="+frm.infoDiv.value);
	}

	function XLSumbit() {
		var frm = document.frmFile;
		if(!frm.infoDiv.value) {
			alert("�ϰ������Ͻ� ��ǰ�� ǰ�������� �������ֽʽÿ�.")
			frm.infoDiv.focus();
			return;
		}

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

		if(confirm("�����Ͻ� ���Ϸ� [��ǰ������ð���] �߰������� �ϰ� ����Ͻðڽ��ϱ�?")) {
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
<form name="frmFile" method="post" action="itemInfoFileUpload_process.asp"  enctype="MULTIPART/FORM-DATA">
<table width="100%" align="left" cellpadding="3" cellspacing="0" class="table_tl">
<tr height="25">
	<td class="td_br" colspan="2">
		<b>[��ǰ������ð���] �߰����� �뷮���</b>
	</td>
</tr>
<tr height="30">
	<td width="90" align="right" class="td_br_tablebar">ǰ������ ���� :</td>
	<td class="td_br">
		<select name="infoDiv" class="select">
		<option value="">::��ǰǰ��::</option>
		<option value="01">01.�Ƿ�</option>
		<option value="02">02.����/�Ź�</option>
		<option value="03">03.����</option>
		<option value="04">04.�м���ȭ(����/��Ʈ/�׼�����)</option>
		<option value="05">05.ħ����/Ŀư</option>
		<option value="06">06.����(ħ��/����/��ũ��/DIY��ǰ)</option>
		<option value="07">07.������(TV��)</option>
		<option value="08">08.������ ������ǰ(�����/��Ź��/�ı⼼ô��/���ڷ�����)</option>
		<option value="09">09.��������(������/��ǳ��)</option>
		<option value="10">10.�繫����(��ǻ��/��Ʈ��/������)</option>
		<option value="11">11.���б��(������ī�޶�/ķ�ڴ�)</option>
		<option value="12">12.��������(MP3/���ڻ��� ��)</option>
		<option value="14">14.������̼�</option>
		<option value="15">15.�ڵ�����ǰ(�ڵ�����ǰ/��Ÿ �ڵ�����ǰ)</option>
		<option value="16">16.�Ƿ���</option>
		<option value="17">17.�ֹ��ǰ</option>
		<option value="18">18.ȭ��ǰ</option>
		<option value="19">19.�ͱݼ�/����/�ð��</option>
		<option value="20">20.��ǰ(����깰)</option>
		<option value="21">21.������ǰ</option>
		<option value="22">22.�ǰ���ɽ�ǰ/ü��������ǰ</option>
		<option value="23">23.�����ƿ�ǰ</option>
		<option value="24">24.�Ǳ�</option>
		<option value="25">25.��������ǰ</option>
		<option value="26">26.����</option>
		<option value="35">35.��Ÿ</option>
		</select>
	</td>
</tr>
<tr height="30">
	<td align="right" class="td_br_tablebar">�ٿ�ε�:</td>
	<td class="td_br">
		<input type="button" class="button" value="��� �ٿ�ε�" onclick="fnFileDownload()">
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
    	* ��ǰ�������� ��� �ٿ�ε� > �躯�泻����� > �����Ͼ��ε�<br />
    	* �ݵ�� �� ���ε� ������� ���ε� (���¸� ������������)<br />
    	* [��ǰ�ڵ�]�� ������ [�Ϲ�] �Ǵ� [����]�� �������ּ���.<br />
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