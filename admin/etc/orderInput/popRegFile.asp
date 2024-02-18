<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 제휴몰 판매 등록 관리
' Hieditor : 2011.04.22 이상구 생성
'			 2012.08.24 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/etc/xSiteTempOrderCls.asp"-->
<%
Dim iMaxLength	,sellsite
	sellsite = requestCheckVar("sellsite",32)

IF iMaxLength = "" THEN iMaxLength = 10
%>

<script language="javascript">

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

function XLSumbit(){
	var frm = document.frmFile;
    if (frm.sellsite.value.length<1){
        alert('제휴몰을 선택하세요.');
        frm.sellsite.focus();
        return;
    }
    
    /*
    if (frm.sellsite.value=="interpark"){
        alert('인터파크 수정중');
        return;
    }
    */

	if (frm.sellsite.value=="gmarket1010"){
		if (!confirm('Gmarket XL 등록시 실판매금액 계산에 문제가 있습니다. 그래도 계속하시겠습니까? 등록 후 관리자에게 알려주세요.')){
			return;
		}
	}

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

	frm.submit();
}

</script>

<link rel="stylesheet" href="/css/tpl.css" type="text/css">

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<b>1. 주문데이타 엑셀등록</b>
	</td>
	<td align="right">
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmFile" method="post" action="procFileUpload.asp"  enctype="MULTIPART/FORM-DATA">
<input type="hidden" name="iML" value="<%=iMaxLength%>">
<tr align="center" bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">쇼핑몰 선택 :</td>
	<td align="left">
		<% call drawSelectBoxXSiteOrderInputPartner("sellsite", sellsite) %>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">파일명:</td>
	<td align="left">
		<input type="file" name="sFile" class="file">
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

<br>
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td align="center" width=70>
		업로드샘플
	</td>
	<td align="left">
		* 인터파크 엑셀: 발송관리>주문별엑셀다운로드(편집하지말것)
		<Br>
		<% get_xsite_excel_order_sample() %>
	</td>
</tr>
</table>
<br/>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td align="center" width=70>
		공통 샘플
	</td>
	<td align="left">
		*엑셀 내용은 전부 셀서식이 텍스트로 되어있습니다.<br />
		서식 변경하지 마시고 샘플데이터대로 입력하시면 됩니다.<br />
		붉은 글씨는 필수입력 값이고, 검정 글씨는 옵션입니다.<br/> 
	<%
		response.write "<a href='http://imgstatic.10x10.co.kr/offshop/sample/xsite/Common_Sample.xls' onfocus='this.blur()'>"
		response.write "<font color='red'>*공통 양식</font></a>"
	%>
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->