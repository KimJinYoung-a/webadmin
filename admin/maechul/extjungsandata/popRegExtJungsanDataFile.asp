<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 제휴몰 판매 등록 관리
' Hieditor : 2011.04.22 이상구 생성
'			 2012.08.24 한용민 수정
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

	/*
	if (frm.extsellsite.value == "lotteimall") {
		if (frm.etcPrice.value == "") {
			alert("롯데아이몰의 경우 수수료 매입분 부가세 금액을 입력하세요.");
			return;
		}

		if (frm.etcPrice.value*0 != 0) {
			alert("금액을 정확히 입력하세요.");
			return;
		}
	}
	*/

	//파일 확인
	if(!jsChkNull("text",frm.sFile,"파일을 입력하십시오.")){
		frm.sFile.focus();
		return;
	}

	if ((frm.extsellsite.value == "kakaogift") || (frm.extsellsite.value == "goodwearmall10beasongpay")||(frm.extsellsite.value == "wconcept1010")||(frm.extsellsite.value == "goodshop1010")||(frm.extsellsite.value == "kakaostore")||(frm.extsellsite.value == "coupang")||(frm.extsellsite.value == "ssg6006")||(frm.extsellsite.value == "ssg6007")||(frm.extsellsite.value == "nvstorefarm")||(frm.extsellsite.value == "nvstorefarmclass")||(frm.extsellsite.value == "nvstoremoonbangu")||(frm.extsellsite.value == "Mylittlewhoopee")||(frm.extsellsite.value == "nvstoregift")||(frm.extsellsite.value == "wadsmartstore")||(frm.extsellsite.value == "lotteon")||(frm.extsellsite.value == "yes24")||(frm.extsellsite.value == "halfclubproduct")||(frm.extsellsite.value == "halfclubbeasongpay")||(frm.extsellsite.value == "gsshopproduct")||(frm.extsellsite.value == "gsshopbeasongpay")||(frm.extsellsite.value == "gsshopproductday")||(frm.extsellsite.value == "WMP")||(frm.extsellsite.value == "WMPbeasongpay")||(frm.extsellsite.value == "wmpfashion")||(frm.extsellsite.value == "wmpfashionbeasongpay")||(frm.extsellsite.value == "ohou1010") ||(frm.extsellsite.value == "LFmall")||(frm.extsellsite.value == "LFmallbeasongpay") ) {
		if (!fnChkFile(frm.sFile.value, arrFileExt2)) {
			alert("XLSX 파일만 업로드 가능합니다.");
			return;
		}
		frm.action="/admin/maechul/extjungsandata/extJungsanUpload_process.asp";

		<% IF NOT (application("Svr_Info")="Dev") then %>
			//frm.action="https://stscm.10x10.co.kr/admin/maechul/extjungsandata/extJungsanUpload_process.asp";
		<% end if %>
	}else if ((frm.extsellsite.value == "interpark")){
		if (!fnChkFile(frm.sFile.value, arrFileExt2)) {
			alert("XLSX 파일만 업로드 가능합니다.");
			return;
		}

		frm.action="/admin/maechul/extjungsandata/extJungsanUpload_process_multisheet.asp";

		<% IF NOT (application("Svr_Info")="Dev") then %>
			//frm.action="https://stscm.10x10.co.kr/admin/maechul/extjungsandata/extJungsanUpload_process_multisheet.asp";
		<% end if %>
	}else if ((frm.extsellsite.value == "interparkrenewal")){
		if (!fnChkFile(frm.sFile.value, arrFileExt2)) {
			alert("XLSX 파일만 업로드 가능합니다.");
			return;
		}
		frm.action="/admin/maechul/extjungsandata/extJungsanUpload_process_multisheet2.asp";

		<% IF NOT (application("Svr_Info")="Dev") then %>
			//frm.action="https://stscm.10x10.co.kr/admin/maechul/extjungsandata/extJungsanUpload_process_multisheet.asp";
		<% end if %>
	}else if ((frm.extsellsite.value == "11st1010") || (frm.extsellsite.value == "goodwearmall10") || (frm.extsellsite.value == "withnature1010") || (frm.extsellsite.value == "GS25") || (frm.extsellsite.value == "boriboriproduct") || (frm.extsellsite.value == "boriboribeasongpay") || (frm.extsellsite.value == "cjmallbeasongpay")||(frm.extsellsite.value == "cjmallproduct")||(frm.extsellsite.value == "gmarket1010")||(frm.extsellsite.value == "gmarket1010beasongpay")||(frm.extsellsite.value == "auction1010")||(frm.extsellsite.value == "auction1010beasongpay")||(frm.extsellsite.value == "ezwel")||(frm.extsellsite.value == "lotteimall")||(frm.extsellsite.value == "alphamallMaechul")||(frm.extsellsite.value == "alphamallHuanBool")||(frm.extsellsite.value == "casamia_good_com")||(frm.extsellsite.value == "shintvshopping")||(frm.extsellsite.value == "shintvshoppingbeasongpay") || (frm.extsellsite.value == "wetoo1300k")||(frm.extsellsite.value == "wetoo1300kbeasongpay") ||(frm.extsellsite.value == "skstoa")||(frm.extsellsite.value == "skstoabeasongpay") ){
		if (!fnChkFile(frm.sFile.value, arrFileExt)) {
			alert("XLS 파일만 업로드 가능합니다.");
			return;
		}

		frm.action="/admin/maechul/extjungsandata/extJungsanUpload_process.asp";
		<% IF NOT (application("Svr_Info")="Dev") then %>
			//frm.action="https://stscm.10x10.co.kr/admin/maechul/extjungsandata/extJungsanUpload_process.asp";
		<% end if %>
	}else if ((frm.extsellsite.value == "lotteCom")||(frm.extsellsite.value == "lotteCombeasongpay")||(frm.extsellsite.value == "hmallproduct")||(frm.extsellsite.value == "hmallbeasongpay")){
		if (!fnChkFile(frm.sFile.value, arrFileExt)) {
			alert("XLS 파일만 업로드 가능합니다.");
			return;
		}

		frm.action="/admin/maechul/extjungsandata/extJungsanUpload_process.asp";
		<% IF NOT (application("Svr_Info")="Dev") then %>
			//frm.action="https://stscm.10x10.co.kr/admin/maechul/extjungsandata/extJungsanUpload_process.asp";
		<% end if %>

	}else if ((frm.extsellsite.value == "cookatmall")){
		if (!fnChkFile(frm.sFile.value, arrFileExt2)) {
			alert("XLSX 파일만 업로드 가능합니다.");
			return;
		}
		frm.action="/admin/maechul/extjungsandata/extJungsanUpload_process_multisheet_cookatmall.asp";
	}else if ((frm.extsellsite.value == "aboutpet")){
		if (!fnChkFile(frm.sFile.value, arrFileExt2)) {
			alert("XLSX 파일만 업로드 가능합니다.");
			return;
		}
		frm.action="/admin/maechul/extjungsandata/extJungsanUpload_process.asp";

		<% IF NOT (application("Svr_Info")="Dev") then %>
			//frm.action="https://stscm.10x10.co.kr/admin/maechul/extjungsandata/extJungsanUpload_process_multisheet.asp";
		<% end if %>

	}else{
		//파일유효성 체크
		if (!fnChkFile(frm.sFile.value, arrFileExt)) {
			alert("XLS 파일만 업로드 가능합니다.");
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
	fileTarget.on('change', function(){ // 값이 변경되면
		if (document.getElementById("extMeachulDate_span").style.display!="none"){
			if(window.FileReader){ // modern browser
				var filename = $(this)[0].files[0].name;
			} else { // old IE
				var filename = $(this).val().split('/').pop().split('\\').pop(); // 파일명만 추출
			}
			// 추출한 파일명 삽입
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

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<b>제휴몰 정산데이타 엑셀등록</b>
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
			<!-- <option value="interpark">인터파크 상품+배송비 정산</option> -->
			<option value="interparkrenewal">인터파크 상품+배송비 정산(RenewalAdminPage)</option>
			<option>---------</option>
			<option value="lotteimall">롯데아이몰 상품+배송비 정산</option>
			<option value="auction1010">옥션 상품정산</option>
			<option value="auction1010beasongpay">옥션 배송비정산</option>
			<option value="gmarket1010">지마켓(NEW)</option>
			<option value="gmarket1010beasongpay">지마켓(NEW) 배송비</option>
			<option>---------</option>
			<option value="lotteCom">롯데닷컴</option>
			<option value="lotteCombeasongpay">롯데닷컴 배송비</option>
			<!-- option value="lotteComM">롯데닷컴(직매출)</option -->
			<option>---------</option>
			<option value="11st1010">11번가</option>
			<option>---------</option>
			<option value="gsshopproduct">GS샵 상품정산(월)</option>
			<option value="gsshopbeasongpay">GS샵 배송비정산(월)</option>
			<option value="gsshopproductday">GS샵 상품정산(일)-반품제외</option>
			<option>---------</option>
			<option value="cjmallproduct">CJ몰 상품정산</option>
			<option value="cjmallbeasongpay">CJ몰 배송비정산</option>
			<option>---------</option>
			<option value="wconcept1010">더블유컨셉</option>
			<option>---------</option>
			<option value="withnature1010">자연이랑</option>
			<option>---------</option>
			<option value="nvstorefarm">스토어팜 상품+배송비 정산</option>
			<option value="Mylittlewhoopee">스토어팜 캣앤독 상품+배송비 정산</option>
<!--
			<option value="nvstorefarmclass">스토어팜-클래스 상품 정산</option>
			<option value="nvstoremoonbangu">스토어팜 문방구 상품+배송비 정산</option>
-->
			<option value="nvstoregift">스토어팜 선물하기 상품+배송비 정산</option>
			<option value="wadsmartstore">와드스마트스토어 상품+배송비 정산</option>
			<option value="ezwel">이지웰페어 상품+배송비 정산(월별)</option>
			<option>---------</option>
			<option value="kakaogift">kakaogift 정산</option>
			<option value="kakaostore">kakaostore 정산</option>
			<option>---------</option>
			<option value="boriboriproduct">보리보리 상품정산</option>
			<option value="boriboribeasongpay">보리보리 배송비정산</option>
			<option>---------</option>
			<option value="GS25">GS25카달로그 정산</option>
			<option>---------</option>
			<option value="coupang">coupang 정산(일별)</option>
			<option>---------</option>
			<option value="ssg6006">SSG</option>
			<!-- <option value="ssg6007">SSG-ssg 정산</option> 다시 합쳐짐-->
			<option>---------</option>
			<option value="halfclubproduct">하프클럽 상품정산</option>
			<option value="halfclubbeasongpay">하프클럽 배송비정산</option>
			<option>---------</option>
			<option value="hmallproduct">Hmall 상품정산</option>
			<option value="hmallbeasongpay">Hmall 배송비정산</option>
			<option>---------</option>
			<option value="WMP">WMP 상품정산</option>
			<option value="WMPbeasongpay">WMP 배송비정산</option>
			<option>---------</option>
			<option value="wmpfashion">WMPW패션 상품정산</option>
			<option value="wmpfashionbeasongpay">WMPW패션 배송비정산</option>
			<option>---------</option>
			<option value="LFmall">LFmall 정산</option>
			<!-- <option value="LFmallbeasongpay">LFmall 배송비정산</option> -->
			<option>---------</option>
			<option value="lotteon">롯데On</option>
			<option>---------</option>
			<option value="yes24">yes24</option>
			<option>---------</option>
			<option value="alphamallMaechul">알파몰 매출</option>
			<option value="alphamallHuanBool">알파몰 환불</option>
			<option>---------</option>
			<option value="ohou1010">오늘의집</option>
			<option>---------</option>
			<option value="casamia_good_com">까사미아</option>
			<option>---------</option>
			<option value="cookatmall">쿠캣</option>
			<option>---------</option>
			<option value="aboutpet">어바웃펫</option>
			<option>---------</option>
			<option value="goodshop1010">굿샵</option>
			<option>---------</option>
			<option value="shintvshopping">신세계TV쇼핑 상품</option>
			<option value="shintvshoppingbeasongpay">신세계TV쇼핑 배송비</option>
			<option>---------</option>
			<option value="wetoo1300k">1300k</option>
			<option value="wetoo1300kbeasongpay">1300k 배송비</option>
			<option>---------</option>
			<option value="skstoa">SKSTOA 상품</option>
			<option value="skstoabeasongpay">SKSTOA 배송비</option>
			<option>---------</option>
			<option value="goodwearmall10">굿웨어몰 상품</option>
			<option value="goodwearmall10beasongpay">굿웨어몰 배송비</option>
		</select>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">파일명:</td>
	<td align="left">
		<input type="file" name="sFile" id="sFile" class="file" >
	</td>
</tr>
<!--
<tr align="center" bgcolor="#FFFFFF">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">기타금액:</td>
	<td align="left">
		<input type="text" class="text" name="etcPrice" value = "">
	</td>
</tr>
-->
<tr align="center" bgcolor="#FFFFFF">
	<td align="center" colspan="2" height="35">
		<span id="extMeachulDate_span" style="margin-right:20px;display:none;">
			정산일(Default:어제날짜) :
			<input type="text" name="extMeachulDate" id="extMeachulDate" value="<%=DateAdd("d",-1,Date())%>" onClick="jsPopCal('extMeachulDate');" style="cursor:pointer;" size="10" maxlength="10" readonly>
		</span>

		<span id="extMeachulMonth_span" style="margin-right:20px;display:none;">
			정산월(Default:지난달) :
			<input type="text" name="extMeachulMonth" id="extMeachulMonth" value="<%=LEFT(DateAdd("m",-1,Date()),7)%>" size="10" maxlength="10" >
			<select name ="cjbeasongGubun" class="select">
				<option value="1">대금지불현황(교환택배비)</option>
				<option value="2">대금지불현황(반품택배비)</option>
				<option value="3">대금지불현황(저단가배송비)</option>
			</select>
		</span>

		<span>
	    <input type="button" class="button" value="등록" onClick="XLSumbit();">
	    <input type="button" class="button" value="취소" onClick="self.close();">
	    </span>



	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan="2" >
	* interpark : 정산관리&gt;정산내역조회&gt;기간별 상세내역 엑셀저장&gt;(XLSX 로저장)<br>
	* ssg : 정산관리&gt;위수탁마감리스트&gt;몰별조회&gt;상세조회&gt;엑셀다운(몰별/과세/면세) (XLSX)<br>
	* 11번가 : 정산관리&gt;판매정산현황 (XLS) <br>
	* coupang : 정산관리&gt;매출내역&gt;[중계]매출(구매확정)내역&gt;상세다운로드(요청후 정산관리엑셀다운로드목록) (XLSX)<br>
	(여러일자를 올리면 주문/반품 금액이 0으로 처리되어 +-내역을 알 수 없음. - 금액이 않맞음.)<br>
	* cjmall상품 : 정산관리&gt;실적관리&gt;수수료매출현황&gt;조회&gt;주문번호별상세내역 (XLS)<br>
	* cjmall배송비 : 정산관리&gt;대금정산&gt;공제내역 에서 저단가배송비,교환택배비, 반품택배비, 무료배송쿠폰?, A/S택배비 (XLS)<br>
	* gmarket 상품 : 주문관리&gt;G마켓 판매진행내역&gt;검색조건:<strong>배송완료일</strong>&gt;배송완료클릭<br>
	<!-- * auction 상품/배송비(월별) : 정산관리&gt;부가가치세신고내역&gt;상세내역다운(XLS)<br> 2019/05/08 주석처리 -->
	* auction 상품 / 배송비 : 주문관리&gt;옥션 판매 진행내역&gt;검색조건:<strong>매출기준일</strong>&gt;매출기준클릭<br>
	* ezwel 상품/배송비 : 정산관리&gt;조회&gt;의뢰자료클릭&gt;엑셀다운(XLS)<br>
	* nvstorefarm 상품/배송비 : 정산관리&gt;정산내역상세&gt;날짜,정산기준일&gt;엑셀다운(XLSX)<br>
	* Mylittlewhoopee 상품 : 정산관리&gt;정산내역상세&gt;날짜,정산기준일&gt;엑셀다운(XLSX) 아이디구분주의<br>
	<!-- * nvstorefarmclass 상품 : 정산관리&gt;정산내역상세&gt;날짜,정산기준일&gt;엑셀다운(XLSX) 아이디구분주의<br> -->
	<!-- * nvstoremoonbangu 상품 : 정산관리&gt;정산내역상세&gt;날짜,정산기준일&gt;엑셀다운(XLSX) 아이디구분주의<br> -->
	* nvstoregift 상품 : 정산관리&gt;정산내역상세&gt;날짜,정산기준일&gt;엑셀다운(XLSX) 아이디구분주의<br>
	* lotteCom 상품 : 정산내역조회&gt;정산확정내역조회&gt;정산기준일&gt;위탁판매&gt;수량(최대)5,0000건,엑셀다운(XLS)<br>
	* lotteCom 배송비 : 배송비정산내역조회&gt;정산기준일&gt;수량(최대)1,0000건,엑셀다운(XLS)<br>
	* halfclub 상품 : 판매지급내역&gt;<strong>사이트:일반</strong>,상세보기,내역조회&gt;엑셀로저장,다른이름저장(XLSX)<br>
	* halfclub 배송비 : 배송비정산&gt;엑셀로저장,다른이름저장(XLSX)<br>
	* gsshop 상품(월별) : 정산관리&gt;대금지급내역&gt;거래상세내역(고객주문별),다른이름저장(XLSX)<br>
	* gsshop 배송비(월별) : 정산관리&gt;대금지급내역&gt;유료배송비/LOSS(유료배송비),일반내역,다른이름저장(XLSX)<br>
	* gsshop 상품(일별) : 주문/배송/반품/재고&gt;협력사배송&gt;직송주문관리&gt;매출완료일기준,다른이름저장(XLSX)<br>
	* lotteimall 상품/배송비 : 정산/세금계산서&gt;수수료매출현황&gt;매출상세내역(일별로다운로드),다른이름저장(XLS)<br>
	* kakaogift / kakaostore : 정산관리 &gt; 판매 확정 상세 현황&gt;정산기준일(XLSX)<br>
	* hmall 상품 : 메뉴검색 &gt; 매출현황(수수료_에누리포함) (XLS)<br>
	* hmall 배송비 : 메뉴검색 &gt; 소액배송비 내역 (XLS)<br>
	* WMP / WMPW패션 상품 : 정산관리 &gt; 매출현황 &gt; 검색 후 [총 주문내역 다운] 버튼(XLSX)<br>
	* WMP / WMPW패션 상품 : 정산관리 &gt; 매출현황 &gt; 검색 후 [배송비 매출 검색결과] 엑셀(XLSX)<br>
	* 롯데On : 정산관리&gt;중개거래정산관리 (XLSX)<br>
	* 알파몰 매출 : 매출&gt;SCM 정산 (XLS) / 검색조건 - 매출<br>
	* 알파몰 환불 : 매출&gt;SCM 정산 (XLS) / 검색조건 - 환불<br>
	* 오늘의집 : 정산관리&gt;매출현황 (XLSX) / 검색조건 - 기준(구매확정)<br>
	* wetoo1300k : 정산관리&gt;정산내역(계산서발행)&gt;수수료 정산, 엑셀주문번호 서식(숫자), 다른이름저장 (XLS)<br>
	* wetoo1300k 배송비 : 정산관리&gt;정산내역(계산서발행)&gt;배송비, 엑셀주문번호 서식(숫자), 다른이름저장 (XLS)<br>
	* skstoa : 정산관리&gt;대금지급상세내역조회 / 다른이름저장 (XLS)<br>
	* skstoa 배송비 : 정산관리&gt;고객부담배송비공제조회 / 다른이름저장 (XLS)<br>
	</td>
</tr>
</table>
</form>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
