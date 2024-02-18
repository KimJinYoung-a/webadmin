<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 상품 품목정보 일괄변경 Excel 업로드
' Hieditor : 2012.10.25 허진원 생성
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
			alert("다운로드 하실 품목유형을 선택해주십시오.")
			frm.infoDiv.focus();
			return;
		}
		window.open("/designer/itemmaster/itemInfoFile/infoFileDownload.asp?fn="+frm.infoDiv.value);
	}

	function XLSumbit() {
		var frm = document.frmFile;
		if(!frm.infoDiv.value) {
			alert("일괄변경하실 상품의 품목유형을 선택해주십시오.")
			frm.infoDiv.focus();
			return;
		}

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

		if(confirm("선택하신 파일로 [상품정보고시관련] 추가정보를 일괄 등록하시겠습니까?")) {
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
<form name="frmFile" method="post" action="itemInfoFileUpload_process.asp"  enctype="MULTIPART/FORM-DATA">
<table width="100%" align="left" cellpadding="3" cellspacing="0" class="table_tl">
<tr height="25">
	<td class="td_br" colspan="2">
		<b>[상품정보고시관련] 추가정보 대량등록</b>
	</td>
</tr>
<tr height="30">
	<td width="90" align="right" class="td_br_tablebar">품목유형 선택 :</td>
	<td class="td_br">
		<select name="infoDiv" class="select">
		<option value="">::상품품목::</option>
		<option value="01">01.의류</option>
		<option value="02">02.구두/신발</option>
		<option value="03">03.가방</option>
		<option value="04">04.패션잡화(모자/벨트/액세서리)</option>
		<option value="05">05.침구류/커튼</option>
		<option value="06">06.가구(침대/소파/싱크대/DIY제품)</option>
		<option value="07">07.영상가전(TV류)</option>
		<option value="08">08.가정용 전기제품(냉장고/세탁기/식기세척기/전자레인지)</option>
		<option value="09">09.계절가전(에어컨/온풍기)</option>
		<option value="10">10.사무용기기(컴퓨터/노트북/프린터)</option>
		<option value="11">11.광학기기(디지털카메라/캠코더)</option>
		<option value="12">12.소형전자(MP3/전자사전 등)</option>
		<option value="14">14.내비게이션</option>
		<option value="15">15.자동차용품(자동차부품/기타 자동차용품)</option>
		<option value="16">16.의료기기</option>
		<option value="17">17.주방용품</option>
		<option value="18">18.화장품</option>
		<option value="19">19.귀금속/보석/시계류</option>
		<option value="20">20.식품(농수산물)</option>
		<option value="21">21.가공식품</option>
		<option value="22">22.건강기능식품/체중조절식품</option>
		<option value="23">23.영유아용품</option>
		<option value="24">24.악기</option>
		<option value="25">25.스포츠용품</option>
		<option value="26">26.서적</option>
		<option value="35">35.기타</option>
		</select>
	</td>
</tr>
<tr height="30">
	<td align="right" class="td_br_tablebar">다운로드:</td>
	<td class="td_br">
		<input type="button" class="button" value="양식 다운로드" onclick="fnFileDownload()">
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
    	* ①품목유형별 양식 다운로드 > ②변경내용기입 > ③파일업로드<br />
    	* 반드시 위 업로드 양식으로 업로드 (형태를 편집하지말것)<br />
    	* [상품코드]의 서식은 [일반] 또는 [숫자]로 저장해주세요.<br />
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