<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ���޸� �Ǹ� ��� ����
' Hieditor : 2011.04.22 �̻� ����
'			 2012.08.24 �ѿ�� ����
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

	/*
	if (frm.extsellsite.value == "lotteimall") {
		if (frm.etcPrice.value == "") {
			alert("�Ե����̸��� ��� ������ ���Ժ� �ΰ��� �ݾ��� �Է��ϼ���.");
			return;
		}

		if (frm.etcPrice.value*0 != 0) {
			alert("�ݾ��� ��Ȯ�� �Է��ϼ���.");
			return;
		}
	}
	*/

	//���� Ȯ��
	if(!jsChkNull("text",frm.sFile,"������ �Է��Ͻʽÿ�.")){
		frm.sFile.focus();
		return;
	}

	if ((frm.extsellsite.value == "kakaogift") || (frm.extsellsite.value == "goodwearmall10beasongpay")||(frm.extsellsite.value == "wconcept1010")||(frm.extsellsite.value == "goodshop1010")||(frm.extsellsite.value == "kakaostore")||(frm.extsellsite.value == "coupang")||(frm.extsellsite.value == "ssg6006")||(frm.extsellsite.value == "ssg6007")||(frm.extsellsite.value == "nvstorefarm")||(frm.extsellsite.value == "nvstorefarmclass")||(frm.extsellsite.value == "nvstoremoonbangu")||(frm.extsellsite.value == "Mylittlewhoopee")||(frm.extsellsite.value == "nvstoregift")||(frm.extsellsite.value == "wadsmartstore")||(frm.extsellsite.value == "lotteon")||(frm.extsellsite.value == "yes24")||(frm.extsellsite.value == "halfclubproduct")||(frm.extsellsite.value == "halfclubbeasongpay")||(frm.extsellsite.value == "gsshopproduct")||(frm.extsellsite.value == "gsshopbeasongpay")||(frm.extsellsite.value == "gsshopproductday")||(frm.extsellsite.value == "WMP")||(frm.extsellsite.value == "WMPbeasongpay")||(frm.extsellsite.value == "wmpfashion")||(frm.extsellsite.value == "wmpfashionbeasongpay")||(frm.extsellsite.value == "ohou1010") ||(frm.extsellsite.value == "LFmall")||(frm.extsellsite.value == "LFmallbeasongpay") ) {
		if (!fnChkFile(frm.sFile.value, arrFileExt2)) {
			alert("XLSX ���ϸ� ���ε� �����մϴ�.");
			return;
		}
		frm.action="/admin/maechul/extjungsandata/extJungsanUpload_process.asp";

		<% IF NOT (application("Svr_Info")="Dev") then %>
			//frm.action="https://stscm.10x10.co.kr/admin/maechul/extjungsandata/extJungsanUpload_process.asp";
		<% end if %>
	}else if ((frm.extsellsite.value == "interpark")){
		if (!fnChkFile(frm.sFile.value, arrFileExt2)) {
			alert("XLSX ���ϸ� ���ε� �����մϴ�.");
			return;
		}

		frm.action="/admin/maechul/extjungsandata/extJungsanUpload_process_multisheet.asp";

		<% IF NOT (application("Svr_Info")="Dev") then %>
			//frm.action="https://stscm.10x10.co.kr/admin/maechul/extjungsandata/extJungsanUpload_process_multisheet.asp";
		<% end if %>
	}else if ((frm.extsellsite.value == "interparkrenewal")){
		if (!fnChkFile(frm.sFile.value, arrFileExt2)) {
			alert("XLSX ���ϸ� ���ε� �����մϴ�.");
			return;
		}
		frm.action="/admin/maechul/extjungsandata/extJungsanUpload_process_multisheet2.asp";

		<% IF NOT (application("Svr_Info")="Dev") then %>
			//frm.action="https://stscm.10x10.co.kr/admin/maechul/extjungsandata/extJungsanUpload_process_multisheet.asp";
		<% end if %>
	}else if ((frm.extsellsite.value == "11st1010") || (frm.extsellsite.value == "goodwearmall10") || (frm.extsellsite.value == "withnature1010") || (frm.extsellsite.value == "GS25") || (frm.extsellsite.value == "boriboriproduct") || (frm.extsellsite.value == "boriboribeasongpay") || (frm.extsellsite.value == "cjmallbeasongpay")||(frm.extsellsite.value == "cjmallproduct")||(frm.extsellsite.value == "gmarket1010")||(frm.extsellsite.value == "gmarket1010beasongpay")||(frm.extsellsite.value == "auction1010")||(frm.extsellsite.value == "auction1010beasongpay")||(frm.extsellsite.value == "ezwel")||(frm.extsellsite.value == "lotteimall")||(frm.extsellsite.value == "alphamallMaechul")||(frm.extsellsite.value == "alphamallHuanBool")||(frm.extsellsite.value == "casamia_good_com")||(frm.extsellsite.value == "shintvshopping")||(frm.extsellsite.value == "shintvshoppingbeasongpay") || (frm.extsellsite.value == "wetoo1300k")||(frm.extsellsite.value == "wetoo1300kbeasongpay") ||(frm.extsellsite.value == "skstoa")||(frm.extsellsite.value == "skstoabeasongpay") ){
		if (!fnChkFile(frm.sFile.value, arrFileExt)) {
			alert("XLS ���ϸ� ���ε� �����մϴ�.");
			return;
		}

		frm.action="/admin/maechul/extjungsandata/extJungsanUpload_process.asp";
		<% IF NOT (application("Svr_Info")="Dev") then %>
			//frm.action="https://stscm.10x10.co.kr/admin/maechul/extjungsandata/extJungsanUpload_process.asp";
		<% end if %>
	}else if ((frm.extsellsite.value == "lotteCom")||(frm.extsellsite.value == "lotteCombeasongpay")||(frm.extsellsite.value == "hmallproduct")||(frm.extsellsite.value == "hmallbeasongpay")){
		if (!fnChkFile(frm.sFile.value, arrFileExt)) {
			alert("XLS ���ϸ� ���ε� �����մϴ�.");
			return;
		}

		frm.action="/admin/maechul/extjungsandata/extJungsanUpload_process.asp";
		<% IF NOT (application("Svr_Info")="Dev") then %>
			//frm.action="https://stscm.10x10.co.kr/admin/maechul/extjungsandata/extJungsanUpload_process.asp";
		<% end if %>

	}else if ((frm.extsellsite.value == "cookatmall")){
		if (!fnChkFile(frm.sFile.value, arrFileExt2)) {
			alert("XLSX ���ϸ� ���ε� �����մϴ�.");
			return;
		}
		frm.action="/admin/maechul/extjungsandata/extJungsanUpload_process_multisheet_cookatmall.asp";
	}else if ((frm.extsellsite.value == "aboutpet")){
		if (!fnChkFile(frm.sFile.value, arrFileExt2)) {
			alert("XLSX ���ϸ� ���ε� �����մϴ�.");
			return;
		}
		frm.action="/admin/maechul/extjungsandata/extJungsanUpload_process.asp";

		<% IF NOT (application("Svr_Info")="Dev") then %>
			//frm.action="https://stscm.10x10.co.kr/admin/maechul/extjungsandata/extJungsanUpload_process_multisheet.asp";
		<% end if %>

	}else{
		//������ȿ�� üũ
		if (!fnChkFile(frm.sFile.value, arrFileExt)) {
			alert("XLS ���ϸ� ���ε� �����մϴ�.");
			return;
		}

		frm.action="<%= ItemUploadUrl %>/linkweb/extjungsandata/extJungsanUpload_process.asp";
	}

	frm.submit();
}

function jsBySite(s){
	if((s == "lotteimall")||(s == "LFmall")){
		$("#extMeachulDate_span").show();
	}else{
		$("#extMeachulDate_span").hide();
	}

	if (s == "cjmallbeasongpay"||s == "shintvshoppingbeasongpay"||s == "skstoabeasongpay"){
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

function isValidDate (d) {
	var date = new Date(d);
	var day = ""+date.getDate();
	if( day.length == 1)day = "0"+day;
	var month = "" +( date.getMonth() + 1);
	if( month.length == 1)month = "0"+month;
	var year = "" + date.getFullYear();

	return ((year + "-" + month + "-" + day ) == d);
}

$(document).ready(function(){
	var fileTarget = $("#sFile");
	fileTarget.on('change', function(){ // ���� ����Ǹ�
		if (document.getElementById("extMeachulDate_span").style.display!="none"){
			if(window.FileReader){ // modern browser
				var filename = $(this)[0].files[0].name;
			} else { // old IE
				var filename = $(this).val().split('/').pop().split('\\').pop(); // ���ϸ� ����
			}
			// ������ ���ϸ� ����
			filename = filename.split(".")[0];
			if (filename.length==8){
				filename = filename.substring(0,4)+"-"+filename.substring(4,6)+"-"+filename.substring(6,8);
				if (isValidDate(filename)){
					$("#extMeachulDate").val(filename);
				}

			}

		}
	});
});


</script>

<link rel="stylesheet" href="/css/tpl.css" type="text/css">

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<b>���޸� ���굥��Ÿ �������</b>
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
			<!-- <option value="interpark">������ũ ��ǰ+��ۺ� ����</option> -->
			<option value="interparkrenewal">������ũ ��ǰ+��ۺ� ����(RenewalAdminPage)</option>
			<option>---------</option>
			<option value="lotteimall">�Ե����̸� ��ǰ+��ۺ� ����</option>
			<option value="auction1010">���� ��ǰ����</option>
			<option value="auction1010beasongpay">���� ��ۺ�����</option>
			<option value="gmarket1010">������(NEW)</option>
			<option value="gmarket1010beasongpay">������(NEW) ��ۺ�</option>
			<option>---------</option>
			<option value="lotteCom">�Ե�����</option>
			<option value="lotteCombeasongpay">�Ե����� ��ۺ�</option>
			<!-- option value="lotteComM">�Ե�����(������)</option -->
			<option>---------</option>
			<option value="11st1010">11����</option>
			<option>---------</option>
			<option value="gsshopproduct">GS�� ��ǰ����(��)</option>
			<option value="gsshopbeasongpay">GS�� ��ۺ�����(��)</option>
			<option value="gsshopproductday">GS�� ��ǰ����(��)-��ǰ����</option>
			<option>---------</option>
			<option value="cjmallproduct">CJ�� ��ǰ����</option>
			<option value="cjmallbeasongpay">CJ�� ��ۺ�����</option>
			<option>---------</option>
			<option value="wconcept1010">����������</option>
			<option>---------</option>
			<option value="withnature1010">�ڿ��̶�</option>
			<option>---------</option>
			<option value="nvstorefarm">������� ��ǰ+��ۺ� ����</option>
			<option value="Mylittlewhoopee">������� Ĺ�ص� ��ǰ+��ۺ� ����</option>
<!--
			<option value="nvstorefarmclass">�������-Ŭ���� ��ǰ ����</option>
			<option value="nvstoremoonbangu">������� ���汸 ��ǰ+��ۺ� ����</option>
-->
			<option value="nvstoregift">������� �����ϱ� ��ǰ+��ۺ� ����</option>
			<option value="wadsmartstore">�͵彺��Ʈ����� ��ǰ+��ۺ� ����</option>
			<option value="ezwel">��������� ��ǰ+��ۺ� ����(����)</option>
			<option>---------</option>
			<option value="kakaogift">kakaogift ����</option>
			<option value="kakaostore">kakaostore ����</option>
			<option>---------</option>
			<option value="boriboriproduct">�������� ��ǰ����</option>
			<option value="boriboribeasongpay">�������� ��ۺ�����</option>
			<option>---------</option>
			<option value="GS25">GS25ī�޷α� ����</option>
			<option>---------</option>
			<option value="coupang">coupang ����(�Ϻ�)</option>
			<option>---------</option>
			<option value="ssg6006">SSG</option>
			<!-- <option value="ssg6007">SSG-ssg ����</option> �ٽ� ������-->
			<option>---------</option>
			<option value="halfclubproduct">����Ŭ�� ��ǰ����</option>
			<option value="halfclubbeasongpay">����Ŭ�� ��ۺ�����</option>
			<option>---------</option>
			<option value="hmallproduct">Hmall ��ǰ����</option>
			<option value="hmallbeasongpay">Hmall ��ۺ�����</option>
			<option>---------</option>
			<option value="WMP">WMP ��ǰ����</option>
			<option value="WMPbeasongpay">WMP ��ۺ�����</option>
			<option>---------</option>
			<option value="wmpfashion">WMPW�м� ��ǰ����</option>
			<option value="wmpfashionbeasongpay">WMPW�м� ��ۺ�����</option>
			<option>---------</option>
			<option value="LFmall">LFmall ����</option>
			<!-- <option value="LFmallbeasongpay">LFmall ��ۺ�����</option> -->
			<option>---------</option>
			<option value="lotteon">�Ե�On</option>
			<option>---------</option>
			<option value="yes24">yes24</option>
			<option>---------</option>
			<option value="alphamallMaechul">���ĸ� ����</option>
			<option value="alphamallHuanBool">���ĸ� ȯ��</option>
			<option>---------</option>
			<option value="ohou1010">��������</option>
			<option>---------</option>
			<option value="casamia_good_com">���̾�</option>
			<option>---------</option>
			<option value="cookatmall">��Ĺ</option>
			<option>---------</option>
			<option value="aboutpet">��ٿ���</option>
			<option>---------</option>
			<option value="goodshop1010">�¼�</option>
			<option>---------</option>
			<option value="shintvshopping">�ż���TV���� ��ǰ</option>
			<option value="shintvshoppingbeasongpay">�ż���TV���� ��ۺ�</option>
			<option>---------</option>
			<option value="wetoo1300k">1300k</option>
			<option value="wetoo1300kbeasongpay">1300k ��ۺ�</option>
			<option>---------</option>
			<option value="skstoa">SKSTOA ��ǰ</option>
			<option value="skstoabeasongpay">SKSTOA ��ۺ�</option>
			<option>---------</option>
			<option value="goodwearmall10">�¿���� ��ǰ</option>
			<option value="goodwearmall10beasongpay">�¿���� ��ۺ�</option>
		</select>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">���ϸ�:</td>
	<td align="left">
		<input type="file" name="sFile" id="sFile" class="file" >
	</td>
</tr>
<!--
<tr align="center" bgcolor="#FFFFFF">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">��Ÿ�ݾ�:</td>
	<td align="left">
		<input type="text" class="text" name="etcPrice" value = "">
	</td>
</tr>
-->
<tr align="center" bgcolor="#FFFFFF">
	<td align="center" colspan="2" height="35">
		<span id="extMeachulDate_span" style="margin-right:20px;display:none;">
			������(Default:������¥) :
			<input type="text" name="extMeachulDate" id="extMeachulDate" value="<%=DateAdd("d",-1,Date())%>" onClick="jsPopCal('extMeachulDate');" style="cursor:pointer;" size="10" maxlength="10" readonly>
		</span>

		<span id="extMeachulMonth_span" style="margin-right:20px;display:none;">
			�����(Default:������) :
			<input type="text" name="extMeachulMonth" id="extMeachulMonth" value="<%=LEFT(DateAdd("m",-1,Date()),7)%>" size="10" maxlength="10" >
			<select name ="cjbeasongGubun" class="select">
				<option value="1">���������Ȳ(��ȯ�ù��)</option>
				<option value="2">���������Ȳ(��ǰ�ù��)</option>
				<option value="3">���������Ȳ(���ܰ���ۺ�)</option>
			</select>
		</span>

		<span>
	    <input type="button" class="button" value="���" onClick="XLSumbit();">
	    <input type="button" class="button" value="���" onClick="self.close();">
	    </span>



	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan="2" >
	* interpark : �������&gt;���곻����ȸ&gt;�Ⱓ�� �󼼳��� ��������&gt;(XLSX ������)<br>
	* ssg : �������&gt;����Ź��������Ʈ&gt;������ȸ&gt;����ȸ&gt;�����ٿ�(����/����/�鼼) (XLSX)<br>
	* 11���� : �������&gt;�Ǹ�������Ȳ (XLS) <br>
	* coupang : �������&gt;���⳻��&gt;[�߰�]����(����Ȯ��)����&gt;�󼼴ٿ�ε�(��û�� ������������ٿ�ε���) (XLSX)<br>
	(�������ڸ� �ø��� �ֹ�/��ǰ �ݾ��� 0���� ó���Ǿ� +-������ �� �� ����. - �ݾ��� �ʸ���.)<br>
	* cjmall��ǰ : �������&gt;��������&gt;�����������Ȳ&gt;��ȸ&gt;�ֹ���ȣ���󼼳��� (XLS)<br>
	* cjmall��ۺ� : �������&gt;�������&gt;�������� ���� ���ܰ���ۺ�,��ȯ�ù��, ��ǰ�ù��, ����������?, A/S�ù�� (XLS)<br>
	* gmarket ��ǰ : �ֹ�����&gt;G���� �Ǹ����೻��&gt;�˻�����:<strong>��ۿϷ���</strong>&gt;��ۿϷ�Ŭ��<br>
	<!-- * auction ��ǰ/��ۺ�(����) : �������&gt;�ΰ���ġ���Ű���&gt;�󼼳����ٿ�(XLS)<br> 2019/05/08 �ּ�ó�� -->
	* auction ��ǰ / ��ۺ� : �ֹ�����&gt;���� �Ǹ� ���೻��&gt;�˻�����:<strong>���������</strong>&gt;�������Ŭ��<br>
	* ezwel ��ǰ/��ۺ� : �������&gt;��ȸ&gt;�Ƿ��ڷ�Ŭ��&gt;�����ٿ�(XLS)<br>
	* nvstorefarm ��ǰ/��ۺ� : �������&gt;���곻����&gt;��¥,���������&gt;�����ٿ�(XLSX)<br>
	* Mylittlewhoopee ��ǰ : �������&gt;���곻����&gt;��¥,���������&gt;�����ٿ�(XLSX) ���̵𱸺�����<br>
	<!-- * nvstorefarmclass ��ǰ : �������&gt;���곻����&gt;��¥,���������&gt;�����ٿ�(XLSX) ���̵𱸺�����<br> -->
	<!-- * nvstoremoonbangu ��ǰ : �������&gt;���곻����&gt;��¥,���������&gt;�����ٿ�(XLSX) ���̵𱸺�����<br> -->
	* nvstoregift ��ǰ : �������&gt;���곻����&gt;��¥,���������&gt;�����ٿ�(XLSX) ���̵𱸺�����<br>
	* lotteCom ��ǰ : ���곻����ȸ&gt;����Ȯ��������ȸ&gt;���������&gt;��Ź�Ǹ�&gt;����(�ִ�)5,0000��,�����ٿ�(XLS)<br>
	* lotteCom ��ۺ� : ��ۺ����곻����ȸ&gt;���������&gt;����(�ִ�)1,0000��,�����ٿ�(XLS)<br>
	* halfclub ��ǰ : �Ǹ����޳���&gt;<strong>����Ʈ:�Ϲ�</strong>,�󼼺���,������ȸ&gt;����������,�ٸ��̸�����(XLSX)<br>
	* halfclub ��ۺ� : ��ۺ�����&gt;����������,�ٸ��̸�����(XLSX)<br>
	* gsshop ��ǰ(����) : �������&gt;������޳���&gt;�ŷ��󼼳���(���ֹ���),�ٸ��̸�����(XLSX)<br>
	* gsshop ��ۺ�(����) : �������&gt;������޳���&gt;�����ۺ�/LOSS(�����ۺ�),�Ϲݳ���,�ٸ��̸�����(XLSX)<br>
	* gsshop ��ǰ(�Ϻ�) : �ֹ�/���/��ǰ/���&gt;���»���&gt;�����ֹ�����&gt;����Ϸ��ϱ���,�ٸ��̸�����(XLSX)<br>
	* lotteimall ��ǰ/��ۺ� : ����/���ݰ�꼭&gt;�����������Ȳ&gt;����󼼳���(�Ϻ��δٿ�ε�),�ٸ��̸�����(XLS)<br>
	* kakaogift / kakaostore : ������� &gt; �Ǹ� Ȯ�� �� ��Ȳ&gt;���������(XLSX)<br>
	* hmall ��ǰ : �޴��˻� &gt; ������Ȳ(������_����������) (XLS)<br>
	* hmall ��ۺ� : �޴��˻� &gt; �Ҿ׹�ۺ� ���� (XLS)<br>
	* WMP / WMPW�м� ��ǰ : ������� &gt; ������Ȳ &gt; �˻� �� [�� �ֹ����� �ٿ�] ��ư(XLSX)<br>
	* WMP / WMPW�м� ��ǰ : ������� &gt; ������Ȳ &gt; �˻� �� [��ۺ� ���� �˻����] ����(XLSX)<br>
	* �Ե�On : �������&gt;�߰��ŷ�������� (XLSX)<br>
	* ���ĸ� ���� : ����&gt;SCM ���� (XLS) / �˻����� - ����<br>
	* ���ĸ� ȯ�� : ����&gt;SCM ���� (XLS) / �˻����� - ȯ��<br>
	* �������� : �������&gt;������Ȳ (XLSX) / �˻����� - ����(����Ȯ��)<br>
	* wetoo1300k : �������&gt;���곻��(��꼭����)&gt;������ ����, �����ֹ���ȣ ����(����), �ٸ��̸����� (XLS)<br>
	* wetoo1300k ��ۺ� : �������&gt;���곻��(��꼭����)&gt;��ۺ�, �����ֹ���ȣ ����(����), �ٸ��̸����� (XLS)<br>
	* skstoa : �������&gt;������޻󼼳�����ȸ / �ٸ��̸����� (XLS)<br>
	* skstoa ��ۺ� : �������&gt;���δ��ۺ������ȸ / �ٸ��̸����� (XLS)<br>
	</td>
</tr>
</table>
</form>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
