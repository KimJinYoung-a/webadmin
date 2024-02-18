<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 상품 안정인증정보 일괄변경 Excel 업로드
' Hieditor : 2015.05.22 허진원 생성
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
		//파일 확인
		if(!frm.sFile.value){
			alert("파일을 입력하십시오.");
			frm.sFile.focus();
			return;
		}

		arrFileExt = new Array();			
		arrFileExt[arrFileExt.length]  = "xls";

		//파일유효성 체크
		if (!fnChkFile(frm.sFile.value, arrFileExt)){
			alert("파일은 엑셀(*.xls)파일만 업로드 가능합니다.");
			return;
		}

		if(confirm("선택하신 파일로 [안전인증 대상] 정보를 일괄 등록하시겠습니까?")) {
			frm.submit();
		}
	}

	function fnChkFile(sFile, arrExt) {
		//파일 업로드 유무확인
		if (!sFile) return true;

		var blnResult = false;

		//파일 확장자 확인
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
		<b>상품 [안전인증 대상] 정보 대량등록</b>
	</td>
</tr>
<tr height="30">
	<td width="90" align="right" class="td_br_tablebar">다운로드:</td>
	<td class="td_br">
		<input type="button" class="button" value="양식 다운로드" onclick="fnFileDownload()">
		<input type="button" class="button" value="양식+상품목록" onclick="fnFileDownloadWithItem()">
	</td>
</tr>
<tr height="30">
	<td width="90" align="right" class="td_br_tablebar">업로드:</td>
	<td class="td_br">
		<input type="file" name="sFile" class="file" style="width:350px;">
	</td>
</tr>
<tr>
    <td colspan="2">
    	* <b>국가통합인증(KC마크)</b>만 입력가능합니다.<br />
    	* ①양식 다운로드 > ②변경내용기입 > ③파일업로드<br />
    	* 반드시 위 업로드 양식으로 업로드 (형태를 편집하지말것)<br />
    	* 파일 업로드 오류 문의 (Email : kobula@10x10.co.kr 해당 파일 첨부 후 문의 해주세요)
    </td>
</tr>
<tr>
	<td align="center" colspan="2" class="td_br">
	    <input type="button" class="button" value=" 등 록 " onClick="XLSumbit();" style="background-color:#FFDDDD"> &nbsp;
	    <input type="button" class="button" value=" 취소 " onClick="self.close();">
	</td>
</tr>
</table>
</form>
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->