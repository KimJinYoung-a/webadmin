<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ���޸� ��ǰ����
' Hieditor : 2018.11.06 eastone
'###########################################################
%>
<!-- #include virtual="/admin/incSessionSTadmin.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language="javascript">

function fnChkFile(sFile, arrExt) {
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

function jsChkNull(type,obj,msg){
     switch (type) {
        // text, password, textarea, hidden
        case "text" :
        case "password" :
        case "textarea" :
        case "hidden" :
            if (jsChkBlank(obj.value)) {
				alert(msg);
				//obj.focus();
                return false;
            }
            else {
                return true;
            }
            break;
        // checkbox
        case "checkbox" :
            if (!obj.checked) {
				alert(msg);
                return false;
            }
            else {
                return true;
            }
            break;
        // radiobutton
        case "radio" :
            var objlen = obj.length;

            for (i=0; i < objlen; i++) {
                if (obj[i].checked == true)
                    return true;
			}
            if (i == objlen) {
				alert(msg);
                return false;
            }else{
				return true;
            }
            break;

		// ���ڰ˻�
        case "numeric" :
            if (!jsChkNumber(obj.value)||jsChkBlank(obj.value)) {
				alert(msg);
                return false;
            }
            else {
                return true;
            }
            break;
	}

        // select list
        if (obj.type.indexOf("select") != -1) {
            if (obj.options[obj.selectedIndex].value == 0 || obj.options[obj.selectedIndex].value == ""){
				alert(msg);
                return false;
            }else{
                return true;
			}
        }

        return true;
}

function XLSumbit(){
	var frm = document.frmFile;

	arrFileExt = new Array();
	arrFileExt[arrFileExt.length] = "xls";
	// arrFileExt[arrFileExt.length] = "xlsx";

	arrFileExt2 = new Array();
	arrFileExt2[arrFileExt2.length] = "xlsx";

	if (frm.extsellsite.value == "") {
		alert("���� ����Ÿ������ �����ϼ���.");
		return;
	}


	//���� Ȯ��
	if(!jsChkNull("text",frm.sFile,"������ �Է��Ͻʽÿ�.")){
		frm.sFile.focus();
		return;
	}

	if ((frm.extsellsite.value == "kakaogift")) {
		if (!fnChkFile(frm.sFile.value, arrFileExt2)) {
			alert("XLSX ���ϸ� ���ε� �����մϴ�.");
			return;
		}
		frm.action="/admin/maechul/extjungsandata/extJungsanUpload_process.asp";

		<% IF NOT (application("Svr_Info")="Dev") then %>
			frm.action="http://stscm.10x10.co.kr/admin/etc/difforder/popExtItemCheckUpload_process.asp";
		<% end if %>

	}else if ((frm.extsellsite.value == "lotteCom")){
		if (!fnChkFile(frm.sFile.value, arrFileExt)) {
			alert("XLS ���ϸ� ���ε� �����մϴ�.");
			return;
		}

		frm.action="/admin/etc/difforder/popExtItemCheckUpload_process.asp";
		<% IF NOT (application("Svr_Info")="Dev") then %>
			frm.action="http://stscm.10x10.co.kr/admin/etc/difforder/popExtItemCheckUpload_process.asp";
		<% end if %>
	}else{
		//������ȿ�� üũ
		if (!fnChkFile(frm.sFile.value, arrFileExt)) {
			alert("XLS ���ϸ� ���ε� �����մϴ�.");
			return;
		}
		
		frm.action="/admin/etc/difforder/popExtItemCheckUpload_process.asp";
	}

	frm.submit();
}

function jsBySite(s){
	if(s == "lotteimall"){
		$("#extMeachulDate_span").show();
	}else{
		$("#extMeachulDate_span").hide();
	}

	if (s == "cjmallbeasongpay"){
		$("#extMeachulMonth_span").show();
	}else{
		$("#extMeachulMonth_span").hide();
	}
}

function jsPopCal(sName){
	var winCal;

	winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
	winCal.focus();
}
</script>

<link rel="stylesheet" href="/css/tpl.css" type="text/css">

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<b>���޸� ��ǰ ���� ����</b>
	</td>
	<td align="right">
	</td>
</tr>
</table>
<!-- �׼� �� -->

<form name="frmFile" method="post" action="<%= ItemUploadUrl %>/linkweb/extjungsandata/extJungsanUpload_process.asp"  enctype="MULTIPART/FORM-DATA">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">����:</td>
	<td align="left">
		<select class="select" name="extsellsite" onChange="jsBySite(this.value);">
			<option></option>
            <option value="lotteCom">�Ե�����</option>
            <!--
			<option value="interpark">������ũ</option>
			<option value="lotteimall">�Ե����̸�</option>
			<option value="auction1010">����</option>
			<option value="gmarket1010">������(NEW)</option>
			<option value="11st1010">11����</option>
			<option value="gseshop">GS��</option>
			<option value="cjmall">CJ��</option>
			<option value="nvstorefarm">�������</option>
			<option value="ezwel">���������</option>
			<option value="kakaogift">kakaogift</option>
			<option value="coupang">coupang</option>
			<option value="ssg6006">ssg</option>
			<option value="halfclub">����Ŭ��</option>
            -->
		</select>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">���ϸ�:</td>
	<td align="left">
		<input type="file" name="sFile" class="file"  >
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td align="center" colspan="2" height="35">

		<span>
	    <input type="button" class="button" value="���" onClick="XLSumbit();">
	    <input type="button" class="button" value="���" onClick="self.close();">
	    </span>

	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan="2" >
    <!--
	* interpark : �������&gt;���곻����ȸ&gt;�Ⱓ�� �󼼳��� ��������&gt;(XLSX ������)<br>
	* ssg : �������&gt;����Ź��������Ʈ&gt;������ȸ&gt;����ȸ&gt;�����ٿ�(����/����/�鼼) (XLSX)<br>
	* 11���� : �������&gt;�Ǹ�������Ȳ (XLS) <br>
	* coupang : �������&gt;���⳻��&gt;[�߰�]����(����Ȯ��)����&gt;�󼼴ٿ�ε�(��û�� ������������ٿ�ε���) (XLSX)<br>
	(�������ڸ� �ø��� �ֹ�/��ǰ �ݾ��� 0���� ó���Ǿ� +-������ �� �� ����. - �ݾ��� �ʸ���.)<br>
	* cjmall��ǰ : �������&gt;��������&gt;�����������Ȳ&gt;��ȸ&gt;�ֹ���ȣ���󼼳��� (XLS)<br>
	* cjmall��ۺ� : �������&gt;�������&gt;�������� ���� ���ܰ���ۺ�,��ȯ�ù��, ��ǰ�ù��, ����������?, A/S�ù�� (XLS)<br>
	* gmarket ��ǰ : �ֹ�����&gt;G���� �Ǹ����೻��&gt;�˻�����:<strong>��ۿϷ���</strong>&gt;��ۿϷ�Ŭ��<br>
	* auction ��ǰ/��ۺ�(����) : �������&gt;�ΰ���ġ���Ű���&gt;�󼼳����ٿ�(XLS)<br>
	* ezwel ��ǰ/��ۺ� : �������&gt;��ȸ&gt;�Ƿ��ڷ�Ŭ��&gt;�����ٿ�(XLS)<br>
	* nvstorefarm ��ǰ/��ۺ� : �������&gt;���곻����&gt;��¥,���������&gt;�����ٿ�(XLSX)<br>
	* lotteCom ��ǰ : ���곻����ȸ&gt;����Ȯ��������ȸ&gt;���������&gt;��Ź�Ǹ�&gt;����(�ִ�)5,0000��,�����ٿ�(XLS)<br>
	* lotteCom ��ۺ� : ��ۺ����곻����ȸ&gt;���������&gt;����(�ִ�)1,0000��,�����ٿ�(XLS)<br>
	* halfclub ��ǰ : �Ǹ����޳���&gt;<strong>����Ʈ:�Ϲ�</strong>,�󼼺���,������ȸ&gt;����������,�ٸ��̸�����(XLSX)<br>
	* halfclub ��ۺ� : ��ۺ�����&gt;����������,�ٸ��̸�����(XLSX)<br>
	* gsshop ��ǰ : �������&gt;������޳���&gt;�ŷ��󼼳���(���ֹ���),�ٸ��̸�����(XLSX)<br>
	* gsshop ��ۺ� : �������&gt;������޳���&gt;�����ۺ�/LOSS(�����ۺ�),�Ϲݳ���,�ٸ��̸�����(XLSX)<br>
	* lotteimall ��ǰ/��ۺ� : ����/���ݰ�꼭&gt;�����������Ȳ&gt;����󼼳���(�Ϻ��δٿ�ε�),�ٸ��̸�����(XLS)<br>
	* kakaogift : ������� &gt; �Ǹ� Ȯ�� �� ��Ȳ&gt;���������(XLSX)<br>
    -->
	</td>
</tr>
</table>
</form>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
