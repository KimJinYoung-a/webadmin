<%@  codepage="65001" language="VBScript" %>
<% option explicit %>
<%
session.codePage = 65001
%>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : 수기랙코드인덱스출력
' Hieditor : 2020.01.09 한용민 생성
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheaderUTF8.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function_utf8.asp"-->
<!-- #include virtual="/lib/offshop_function_utf8.asp"-->
<%
dim menupos, barcodeprintyn
	menupos = requestCheckVar(getNumeric(request("onoffgubun")), 10)
	barcodeprintyn = requestCheckVar(request("barcodeprintyn"), 1)

if barcodeprintyn="" then barcodeprintyn="Y"
%>
<script type="text/javascript" src="/js/ttpbarcode_utf8.js"></script>
<script type="text/javascript" src="/js/barcode.js"></script>
<script type="text/javascript" src="/js/DOSHIBAbarcode_utf8.js"></script>
<script type="text/javascript">

function jsIndexSudongrackBarcodePrint() {
    var isforeignprint; var domainname; var showdomainyn; var showdomainyn; var ttptype;
    var barcodetype;
    var shopbrandyn; var currencychar; var showpriceyn;

	isforeignprint = "N";

	shopbrandyn		= "N";
	ttptype			= "TTP-243_80x50";
	barcodetype		= "";
	var paperwidth = "80";
	var paperheight = "50";
	var papermargin = "3";
	var heightoffset = 0;
	showpriceyn = "N";

	currencychar = "￦";
	domainname		= "www.10x10.co.kr";
	showdomainyn	= "Y";

	var barcodeprintyn = document.frm.barcodeprintyn.value;	
	var msg = document.frm.message.value;
	var itemno = document.frm.itemno.value;
	var lines = msg.split("\n");
	var MAX_LINE_COUNT = 50;
	var MAX_MESSAGE_LENGTH = 8;

    if (msg==""){
        alert("출력메세지를 입력해주세요.");
        return;
    }
	if (lines.length > MAX_LINE_COUNT) {
		alert("\n========== 에러 ==========\n\n출력메시지는 " + MAX_LINE_COUNT + "줄 이상을 넘을 수 없습니다. ");
		return;
	}

	for (var i = 0; i < lines.length; i++) {
		if (lines[i].replace("\r", "").length > MAX_MESSAGE_LENGTH) {
			alert("\n========== 에러 ==========\n\n출력메시지는 각 줄마다 " + MAX_MESSAGE_LENGTH + "글자를 이상을 넘을 수 없습니다. ");
			return;
		}
	}

	var fontName = document.frm.fontName.value;
	//var radios = document.getElementsByName('fontName');
	//for (var i = 0, length = radios.length; i < length; i++) {
	//	if (radios[i].checked) {
	//		fontName = radios[i].value;
	//		break;
	//	}
	//}
    if (itemno=="" || itemno==0){
        alert("출력하실 수량을 입력해주세요.");
        frm.itemno.focus();
        return;
    }
    if (fontName==""){
        alert("폰트를 선택해주세요.");
        return;
    }

	//TEC B-FV4		//2020.01.09 한용민 생성
	if (TEC_DO3.IsDriver == 1){
		if (confirm("직접입력 랙코드 인덱스를 출력합니다.\n\nTEC B-FV4 로 출력하시겠습니까?") == true) {
			TOSHIBA_DOMAINNAME = domainname;
			TOSHIBA_SHOWDOMAINYN = showdomainyn;
			TOSHIBA_PAPERWIDTH = 800;
			TOSHIBA_PAPERHEIGHT = 500;
			TOSHIBA_PAPERMARGIN = 3;
			TOSHIBA_SHOWPRICEYN = showpriceyn;
			TOSHIBA_currencyChar = currencychar;
			TOSHIBA_SHOPBRANDYN = shopbrandyn;
			TOSHIBA_BARCODETYPE = barcodetype;

			printTOSHIBARackIndexSudongLabel(msg,itemno,fontName,barcodeprintyn);
		}

	// /js/barcode.js 참조
	}else if (initTTPprinter(ttptype, barcodetype, showdomainyn, domainname, showpriceyn, currencychar, shopbrandyn, papermargin, heightoffset) == true) {
		if (confirm("직접입력 랙코드 인덱스를 출력합니다.\n\nTTP-243 로 출력하시겠습니까?") == true) {
			printTTPRackIndexSudongLabel(msg,itemno,fontName,barcodeprintyn);
		}

	}else {
	    alert("TTP-243(구)나 TEC B-FV4 드라이버를 설치해 주세요");
	}
	return;
}

</script>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="left">
		※ 수기랙코드인덱스출력
	</td>
	<td align="right"></td>
</tr>
</table>
<!-- 액션 끝 -->

<form name="frm" method="post" action="" style="margin:0px;">
<input type="hidden" name="mode">
<input type="hidden" name="fontName" value="10X10">
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA">
<tr>
	<td bgcolor="#e1e1e1" align="center">출력메세지</td>
	<td bgcolor="#FFFFFF">
		<textarea name="message" style="width:400px;" rows="20"></textarea>
	</td>
</tr>
<tr height="30">
	<td bgcolor="#e1e1e1" align="center">수량</td>
	<td bgcolor="#FFFFFF">
        <input type="text" name="itemno" value="1" size=8 maxlength=10>
	</td>
</tr>
<tr height="30">
	<td bgcolor="#e1e1e1" align="center">바코드출력여부</td>
	<td bgcolor="#FFFFFF">
        <% drawSelectBoxisusingYN "barcodeprintyn", barcodeprintyn, "" %>
	</td>
</tr>
<!--<tr>
	<td bgcolor="#e1e1e1" align="center">폰트</td>
	<td bgcolor="#FFFFFF">
		<input type="radio" name="fontName" value="Batang"> 바탕체
		<input type="radio" name="fontName" value="10X10" checked> 텐바이텐 폰트
		<input type="radio" name="fontName" value="Gulim"> 굴림체
		<input type="radio" name="fontName" value="Malgun Gothic"> 맑은고딕
	</td>
</tr>-->
<tr>
	<td bgcolor="#FFFFFF" colspan=2 align="center">
        <input type="button" class="button" onClick="jsIndexSudongrackBarcodePrint();" value=" 출 력 ">
	</td>
</tr>
</table>
</form>

<%
session.codePage = 949
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/admin/lib/poptail.asp"-->
