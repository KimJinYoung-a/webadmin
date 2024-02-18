<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db_TPLOpen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<%

Dim  iMaxLength
IF iMaxLength = "" THEN iMaxLength = 10

Dim partnercompanyid  : partnercompanyid = requestCheckVar("partnercompanyid",32)


%>
<script language="javascript">
function XLSumbit(){
	var frm = document.frmFile;

	arrFileExt = new Array();
	arrFileExt[arrFileExt.length]  = "xls";

	if (frm.tplcompanyid.value == '') {
		alert('고객사를 선택하세요.');
		return;
	}

	if (frm.xltype.value == '') {
		alert('엑셀포멧을 선택하세요.');
		return;
	}

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

	frm.submit();
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

		// 숫자검사
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

function fnChkFile(sFile, sMaxSize, arrExt){
    //파일 업로드 유무확인
     if (!sFile){
    	 return true;
    	}

    var blnResult = false;

    //파일 용량 확인
    var maxsize = sMaxSize * 1024 * 1024;

     //	var img = new Image();
    //	img.dynsrc = sFile;
    //var fSize = img.fileSize ;
    	//if (fSize > maxsize){
    		//alert("파일크기는 "+sMaxSize+"MB이하만 가능합니다.");
    		//return false;
    	//}

    //파일 확장자 확인
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
</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="left" bgcolor="#FFFFFF" height="25">
		<td bgcolor="<%= adminColor("tabletop") %>" colspan="2">
			<b>1. 주문데이타 엑셀등록</b>
		</td>
	</tr>
	<form name="frmFile" method="post" action="procFileUpload.asp"  enctype="MULTIPART/FORM-DATA">
	<input type="hidden" name="iML" value="<%=iMaxLength%>">
	<tr bgcolor="#FFFFFF">
		<td width="80" align="right" bgcolor="<%= adminColor("tabletop") %>"> 고객사 :</td>
		<td>
			<% CALL drawPartner3plCompany("tplcompanyid","","") %>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td width="80" align="right" bgcolor="<%= adminColor("tabletop") %>"> 엑셀포멧 :</td>
		<td>
			<select class="select" name="xltype">
				<option></option>
				<option value="sabangnet">사방넷 주문</option>
				<option value="default">기본 포멧</option>
			</select>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td width="80" align="right" bgcolor="<%= adminColor("tabletop") %>">파일명:</td>
		<td><input type="file" name="sFile" ></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center" colspan="2">
		    <input type="button" class="button" value="등록" onClick="XLSumbit();">
		    <input type="button" class="button" value="취소" onClick="self.close();">
		</td>
	</tr>

	</form>
</table>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/db_TPLClose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
