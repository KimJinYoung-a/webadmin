<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 제휴몰 상품검증
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
    //파일 업로드 유무확인
     if (!sFile){
    	 return true;
    	}

    var blnResult = false;

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

	arrFileExt = new Array();
	arrFileExt[arrFileExt.length] = "xls";
	// arrFileExt[arrFileExt.length] = "xlsx";

	arrFileExt2 = new Array();
	arrFileExt2[arrFileExt2.length] = "xlsx";

	if (frm.extsellsite.value == "") {
		alert("먼저 데이타구분을 지정하세요.");
		return;
	}


	//파일 확인
	if(!jsChkNull("text",frm.sFile,"파일을 입력하십시오.")){
		frm.sFile.focus();
		return;
	}

	if ((frm.extsellsite.value == "kakaogift")) {
		if (!fnChkFile(frm.sFile.value, arrFileExt2)) {
			alert("XLSX 파일만 업로드 가능합니다.");
			return;
		}
		frm.action="/admin/maechul/extjungsandata/extJungsanUpload_process.asp";

		<% IF NOT (application("Svr_Info")="Dev") then %>
			frm.action="http://stscm.10x10.co.kr/admin/etc/difforder/popExtItemCheckUpload_process.asp";
		<% end if %>

	}else if ((frm.extsellsite.value == "lotteCom")){
		if (!fnChkFile(frm.sFile.value, arrFileExt)) {
			alert("XLS 파일만 업로드 가능합니다.");
			return;
		}

		frm.action="/admin/etc/difforder/popExtItemCheckUpload_process.asp";
		<% IF NOT (application("Svr_Info")="Dev") then %>
			frm.action="http://stscm.10x10.co.kr/admin/etc/difforder/popExtItemCheckUpload_process.asp";
		<% end if %>
	}else{
		//파일유효성 체크
		if (!fnChkFile(frm.sFile.value, arrFileExt)) {
			alert("XLS 파일만 업로드 가능합니다.");
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

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<b>제휴몰 상품 오류 검증</b>
	</td>
	<td align="right">
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<form name="frmFile" method="post" action="<%= ItemUploadUrl %>/linkweb/extjungsandata/extJungsanUpload_process.asp"  enctype="MULTIPART/FORM-DATA">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">구분:</td>
	<td align="left">
		<select class="select" name="extsellsite" onChange="jsBySite(this.value);">
			<option></option>
            <option value="lotteCom">롯데닷컴</option>
            <!--
			<option value="interpark">인터파크</option>
			<option value="lotteimall">롯데아이몰</option>
			<option value="auction1010">옥션</option>
			<option value="gmarket1010">지마켓(NEW)</option>
			<option value="11st1010">11번가</option>
			<option value="gseshop">GS샵</option>
			<option value="cjmall">CJ몰</option>
			<option value="nvstorefarm">스토어팜</option>
			<option value="ezwel">이지웰페어</option>
			<option value="kakaogift">kakaogift</option>
			<option value="coupang">coupang</option>
			<option value="ssg6006">ssg</option>
			<option value="halfclub">하프클럽</option>
            -->
		</select>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">파일명:</td>
	<td align="left">
		<input type="file" name="sFile" class="file"  >
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td align="center" colspan="2" height="35">

		<span>
	    <input type="button" class="button" value="등록" onClick="XLSumbit();">
	    <input type="button" class="button" value="취소" onClick="self.close();">
	    </span>

	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan="2" >
    <!--
	* interpark : 정산관리&gt;정산내역조회&gt;기간별 상세내역 엑셀저장&gt;(XLSX 로저장)<br>
	* ssg : 정산관리&gt;위수탁마감리스트&gt;몰별조회&gt;상세조회&gt;엑셀다운(몰별/과세/면세) (XLSX)<br>
	* 11번가 : 정산관리&gt;판매정산현황 (XLS) <br>
	* coupang : 정산관리&gt;매출내역&gt;[중계]매출(구매확정)내역&gt;상세다운로드(요청후 정산관리엑셀다운로드목록) (XLSX)<br>
	(여러일자를 올리면 주문/반품 금액이 0으로 처리되어 +-내역을 알 수 없음. - 금액이 않맞음.)<br>
	* cjmall상품 : 정산관리&gt;실적관리&gt;수수료매출현황&gt;조회&gt;주문번호별상세내역 (XLS)<br>
	* cjmall배송비 : 정산관리&gt;대금정산&gt;공제내역 에서 저단가배송비,교환택배비, 반품택배비, 무료배송쿠폰?, A/S택배비 (XLS)<br>
	* gmarket 상품 : 주문관리&gt;G마켓 판매진행내역&gt;검색조건:<strong>배송완료일</strong>&gt;배송완료클릭<br>
	* auction 상품/배송비(월별) : 정산관리&gt;부가가치세신고내역&gt;상세내역다운(XLS)<br>
	* ezwel 상품/배송비 : 정산관리&gt;조회&gt;의뢰자료클릭&gt;엑셀다운(XLS)<br>
	* nvstorefarm 상품/배송비 : 정산관리&gt;정산내역상세&gt;날짜,정산기준일&gt;엑셀다운(XLSX)<br>
	* lotteCom 상품 : 정산내역조회&gt;정산확정내역조회&gt;정산기준일&gt;위탁판매&gt;수량(최대)5,0000건,엑셀다운(XLS)<br>
	* lotteCom 배송비 : 배송비정산내역조회&gt;정산기준일&gt;수량(최대)1,0000건,엑셀다운(XLS)<br>
	* halfclub 상품 : 판매지급내역&gt;<strong>사이트:일반</strong>,상세보기,내역조회&gt;엑셀로저장,다른이름저장(XLSX)<br>
	* halfclub 배송비 : 배송비정산&gt;엑셀로저장,다른이름저장(XLSX)<br>
	* gsshop 상품 : 정산관리&gt;대금지급내역&gt;거래상세내역(고객주문별),다른이름저장(XLSX)<br>
	* gsshop 배송비 : 정산관리&gt;대금지급내역&gt;유료배송비/LOSS(유료배송비),일반내역,다른이름저장(XLSX)<br>
	* lotteimall 상품/배송비 : 정산/세금계산서&gt;수수료매출현황&gt;매출상세내역(일별로다운로드),다른이름저장(XLS)<br>
	* kakaogift : 정산관리 &gt; 판매 확정 상세 현황&gt;정산기준일(XLSX)<br>
    -->
	</td>
</tr>
</table>
</form>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
