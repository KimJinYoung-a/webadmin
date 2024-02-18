<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
Dim iMaxLength
Dim mallid
mallid = requestCheckVar(Request("mallid"),32)

IF iMaxLength = "" THEN iMaxLength = 10
%>
<script language="javascript">
function fnChkFile(sFile, sMaxSize, arrExt){
	//파일 업로드 유무확인
	if (!sFile){
		return true;
	}
	var blnResult = false;
	var maxsize = sMaxSize * 1024 * 1024;
	
	//파일 확장자 확인
	var pPoint = sFile.lastIndexOf('.');
	var fPoint = sFile.substring(pPoint+1,sFile.length);
	var fExet = fPoint.toLowerCase();
	for (var i = 0; i < arrExt.length; i++){
		if (arrExt[i].toLowerCase() == fExet){
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
		}else {
			return true;
		}
		break;
		// checkbox
		case "checkbox" :
		if (!obj.checked) {
			alert(msg);
			return false;
		}else {
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
		// 숫자검사
		case "numeric" :
		if (!jsChkNumber(obj.value)||jsChkBlank(obj.value)) {
			alert(msg);
			return false;
		}else {
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
	arrFileExt[arrFileExt.length]  = "xls";

	//파일 확인
	if(!jsChkNull("text",frm.sFile,"파일을 입력하십시오.")){
		frm.sFile.focus();
		return;
	}

	//파일유효성 체크
	if (!fnChkFile(frm.sFile.value, <%=iMaxLength%>, arrFileExt)){
		alert("파일은 <%=iMaxLength%>MB이하의 xls파일만 업로드 가능합니다.");
		return;
	}
	document.getElementById("lodingtb").style.display="block";
	setInterval("doBlink()",350) 
	frm.submit();
}
function doBlink() { 
	var blink = document.all.tags("BLINK") 
	for (var i=0; i < blink.length; i++) 
	blink[i].style.visibility = blink[i].style.visibility == "" ? "hidden" : "" 
} 
</script>
<link rel="stylesheet" href="/css/tpl.css" type="text/css">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmFile" method="post" action="/admin/etc/orderInput/procEtcExcelFileUpload.asp"  enctype="MULTIPART/FORM-DATA">
<input type="hidden" name="mallid" value="<%=mallid%>">
<tr align="center" bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">파일명:</td>
	<td align="left">
		<input type="file" name="sFile" class="file">
		&nbsp;&nbsp;&nbsp;<a href="/admin/etc/orderInput/sample.xls"><font color="BLUE"><strong>샘플다운</strong></font></a><br>
		1.엑셀은 <font color="red">97-2003 통합문서</font>여야 하고, <font color="red">첫라인</font>은 고정.<br>
		2.샘플 형식과 동일해야 함
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td align="center" colspan="2">
	    <input type="button" class="button" value="등록" onClick="XLSumbit();">
	    <input type="button" class="button" value="취소" onClick="self.close();">
	</td>
</tr>
</form>
</table>
<br><br>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" id="lodingtb" style="display:none">
	<tr align="center" bgcolor="#FFFFFF">
		<td><font color="RED"><blink>해당 요청 처리 중입니다. 잠시만 기다려주세요</blink></font></td>
	</tr>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->