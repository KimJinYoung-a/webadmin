<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 상품해외배송정보 일괄변경 Excel 업로드
' Hieditor : 2016.06.03 정윤정 생성
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
		window.open("/designer/itemmaster/itemInfoFile/infoFileDownload.asp?fn=900");
	}

	function fnFileDownloadWithItem() {
		window.open("/designer/itemmaster/itemInfoFile/item_overseaInfo_xls.asp");
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

		if(confirm("선택하신 파일로 [해외배송] 정보를 일괄 등록하시겠습니까?")) {
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
<form name="frmFile" method="post" action="itemOverseaInfoFileUpload_process.asp"  enctype="MULTIPART/FORM-DATA">
<table width="100%" align="left" cellpadding="3" cellspacing="0" class="table_tl">
<tr height="25">
	<td class="td_br" colspan="2">
		<b>상품 [해외배송] 정보 대량등록</b>
	</td>
</tr>
<tr height="30">
	<td width="90" align="right" class="td_br_tablebar">다운로드:</td>
	<td class="td_br">
		<input type="button" class="button" value="양식 다운로드" onclick="fnFileDownload()"> &nbsp;
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
    	* 항목모두 필수입력사항입니다.<br />
    	* ①양식 다운로드 > ②변경내용기입 > ③파일업로드<br />
    	* 반드시 위 업로드 양식으로 업로드 (형태를 편집하지말것)<br />
    	* Excel 97-2003통합 문서로 저장
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