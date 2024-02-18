<%@ language=vbscript %>
<% option explicit %>
<%
session.codePage = 949
Response.CharSet = "EUC-KR"
%>
<%
'###########################################################
' Description :  SCM 우편번호 찾기
' History : 2016.07.01 한용민 프론트 이전 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAnalopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/search/Zipsearchcls.asp" -->
<!-- #include virtual="/lib/util/pageformlib.asp" -->
<%
	dim fiximgPath
	'이미지 경로 지정(SSL 처리)
	if request.ServerVariables("SERVER_PORT_SECURE")<>1 then
		fiximgPath = "http://fiximage.10x10.co.kr"
	else
		fiximgPath = "/fiximage"
	end If
	
	Dim strTarget
	Dim strMode
	strTarget	= requestCheckVar(Request("target"),32)
	strMode     = requestCheckVar(Request("strMode"),32)

	dim PageSize	: PageSize = getNumeric(requestCheckVar(request("psz"),5))
	dim CurrPage : CurrPage = getNumeric(requestCheckVar(request("cpg"),8))
	if CurrPage="" then CurrPage=1
	if PageSize="" then PageSize=10
%>
<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="euc-kr" />
<meta http-equiv="X-UA-Compatible" content="IE=edge" />
<meta http-equiv='Content-Type' content='text/html;charset=euc-kr' />
<title>텐바이텐 10X10 : 우편번호찾기</title>
<style type="text/css">
html, body, blockquote, caption, dd, div, dl, dt, fieldset, form, frame, h1, h2, h3, h4, h5, h6, hr, iframe, input, legend, li, object, ol, p, pre, q, select, table, textarea, tr, td, ul, button {margin:0; padding:0;}
ol, ul {list-style:none;}
fieldset, img {border:0;}
h1, h2, h3, h4, h5, h6 {font-style:normal; font-size:12px;}
hr {display:none;}
table {border-collapse:collapse; border:0; empty-cells:show;}
textarea {resize:none;}
input, button {border:0;}
button {overflow:visible;}

body, h1, h2, h3 ,h4 {font-size:12px; font-family:dotum, dotumche, '돋움', '돋움체', verdana, tahoma, sans-serif; line-height:1.6; color:#555;}
a {color:inherit; text-decoration:none;}
a:link, a:active, a:visited {color:#555;}
a:hover {text-decoration:none;}
a:hover {text-decoration:none;}
legend {visibility:hidden; width:0; height:0;}
caption {overflow:hidden; width:0; height:0; font-size:0; line-height:0; text-indent:-9999px;}
button {border:0; cursor:pointer;}
input[type=number]::-webkit-inner-spin-button {-webkit-appearance:none;}

html, body {height:100%;}

/* Popup layout */
body > .heightgird {min-height:100%; height:auto;}
.heightgird {position:relative;}
.popWrap {padding-bottom:45px;}
.popWrap .popHeader {padding:27px 15px 15px; background:#d50c0c; color:#fff;}
.popContent {padding:30px; font-size:11px;}
.popFooter {position:absolute; bottom:0; width:100%; padding:0; border-top:1px solid #ddd; background:#f5f5f5;}
.popFooter .btnArea {float:right; padding:8px 30px 11px 0;}
.popFooter .btnArea .btn {padding:5px 11px 3px 24px; border:0; border-bottom:1px solid #efefef; background:#999 url(http://fiximage.10x10.co.kr/web2013/common/btn_close_popup.gif) 11px center no-repeat;}
.popFooter .btnArea .btn:hover {border:0; border-bottom:1px solid #efefef; background:#8a8a8a url(http://fiximage.10x10.co.kr/web2013/common/btn_close_popup.gif) 11px center no-repeat;}
.popFooter button {font-family:Dotum; font-weight:normal;}

/* button */
.btn {display:inline-block; text-align:center; font-weight:bold; vertical-align:middle; cursor:pointer; font-family:/*verdana, tahoma,*/ dotum, dotumche, '돋움', '돋움체', sans-serif;}
.btn:link, .btn:active, .btn:visited {color:#fff;}
.btn:hover {text-decoration:none;}
.btnRed {color:#fff; background:#d50c0c; border:1px solid #d50c0c;}
.btnRed:hover {background:#b20202; border:1px solid #b20202;}
.btnWhite {color:#d50c0c; background:#fff; border:1px solid #d50c0c;}
.btnWhite:link, .btnWhite:active, .btnWhite:visited {color:#d50c0c;}
.btnM2 {font-size:12px; line-height:15px; padding:8px 40px 5px;}
.btnW220 {width:218px; padding-left:0; padding-right:0;}

/* Zip code */
.popContent .finder {margin:20px 0 36px; padding:16px 0; text-align:center;}
.popContent .finder .field {margin:12px auto 6px; width:260px; padding:6px 7px 4px 20px ; border:1px solid #d50c0c; background:#fff; text-align:left;}
.popContent .finder .field input {width:240px; height:18px; color:#d50c0c; font-size:12px; vertical-align:top;}
.popContent .finder .field .btnSearch {width:13px; height:14px; background:url(http://fiximage.10x10.co.kr/web2013/common/btn_search3.gif) left top no-repeat;}
.popContent .finder .field .btnSearch:hover {background:url(http://fiximage.10x10.co.kr/web2013/common/btn_search3.gif) left bottom no-repeat ;}
.popContent .zipcode {max-height:269px; height:auto !important; height:269px; overflow-y:auto; border-bottom:1px solid #ccc;}
.popContent .zipcode table {border-bottom:0;}

/* paging */
.pageWrapV15 {position:relative; padding-top : 20px;}
.paging {width:100%; text-align:center; height:25px;}
.paging a {display:inline-block; height:23px; line-height:22px; border:1px solid #ccc; background-color:#fff; text-decoration:none; vertical-align:top; overflow:hidden;}
.paging a span {display:block; height:23px; vertical-align:middle; font-size:12px; font-family:verdana, tahoma, sans-serif; color:#555; min-width:8px; padding:0 8px 0 7px; letter-spacing:-1px;}
.paging a.arrow {background-color:#fff;}
.paging a.arrow span {background-image:url(http://fiximage.10x10.co.kr/web2015/common/paging_arrow.gif); background-repeat:no-repeat; text-indent:-9999px; width:23px; padding:0;}
.paging a.current {background-color:#fff; border:1px solid #d50c0c; color:#d50c0c; font-weight:bold;}
.paging a.current span {color:#d50c0c;}
.paging a.current:hover {background-color:#fff;}
.paging a.first span {background-position:6px 8px;}
.paging a.prev span {background-position:-22px 8px;}
.paging a.next span {background-position:-348px 8px;}
.paging a.end span {background-position:-378px 8px;}
.paging a:hover {background-color:#ececec;}
.pageMove {position:absolute; right:0; top:0; font-size:11px;}
.pageMove input {padding:3px 5px; border:1px solid #ccc; text-align:right; vertical-align:middle; font-size:11px;}

.popWrap .popHeader h1 img {vertical-align:top;}

/* zipcode 2017 */
.zipcodeV17 {min-width:500px;}
.zipcodeV17 .hidden {visibility:hidden; width:0; height:0; overflow:hidden; position:absolute; top:-1000%; line-height:0;}
.zipcodeV17 .tabs {overflow:hidden; margin-left:-1px; border-left:1px solid #ddd;}
.zipcodeV17 .tabs li {float:left; width:50%;}
.zipcodeV17 .tabs li a {display:block; border:1px solid #ddd; border-left:0; background-color:#f5f5f5; color:#969696; font-size:13px; font-weight:bold; line-height:33px; text-align:center;}
.zipcodeV17 .tabs li a:hover {text-decoration:none;}
.zipcodeV17 .tabs .on {border-bottom:0; background-color:#fff; color:#555;}

.zipcodeV17 input {font-family:'Dotum', '?¸¿?', 'Verdana'; font-size:11px;}
.zipcodeV17 select {appearance:none; -webkit-appearance:none; -moz-appearance:none; height:30px; padding-left:10px; border:1px solid #bbb; padding-right:25px; background:url(/fiximage/web2015/giftcard/bg_select_arr.gif) no-repeat 100% 50%; color:#555; font-family:'Dotum', '?¸¿?'; font-size:11px;}
.zipcodeV17 select {padding-right:0\9;background:none\9;}
.zipcodeV17 select::-ms-expand {
	display:none;
}
.zipcodeV17 .itext {display:block; height:28px; padding:0 10px; border:1px solid #bbb;}
.zipcodeV17 .itext input {width:100%; height:28px; background-color:transparent; line-height:28px;}
.zipcodeV17 .itext input[type=search] {-webkit-appearance:none; line-height: normal;}
.zipcodeV17 .searchForm input::-webkit-input-placeholder {color:#888;}
.zipcodeV17 .searchForm input::-moz-placeholder {color:#888;} /* firefox 19+ */
.zipcodeV17 .searchForm input:-ms-input-placeholder {color:#888;} /* ie */
.zipcodeV17 .searchForm input:-moz-placeholder {color:#888;}

.zipcodeV17 .searchForm {margin-top:30px;}
.zipcodeV17 .searchForm input[type=search] {border:0;}
.zipcodeV17 .searchForm .finder {position:relative; margin:0; padding:20px; border:5px solid #fafafa;}
.zipcodeV17 .searchForm .inner {position:relative; padding-right:130px; text-align:left;}
.zipcodeV17 .searchForm .finder .btn {position:absolute; top:20px; right:20px; width:120px; padding-right:0; padding-left:0;}
.zipcodeV17 .searchForm .btnReset {position:absolute; top:1px; right:131px; width:36px; height:28px; background:#fff url(/fiximage/web2017/common/btn_reset.png) 50% 50% no-repeat; color:transparent; cursor:pointer;}

.zipcodeV17 .searchForm ul {padding:25px 20px; border:5px solid #fafafa;}
.zipcodeV17 .searchForm ul li {position:relative; margin-top:10px; padding-left:80px; *zoom:1;}
.zipcodeV17 .searchForm ul li:first-child {margin-top:0;}
.zipcodeV17 .searchForm ul li label {position:absolute; top:0; left:0; width:80px; height:30px; color:#555; font-weight:bold; line-height:30px; text-align:left;}
.zipcodeV17 .searchForm ul li select {width:100%;}

.zipcodeV17 .guide {padding:27px 0 26px; color:#888; font-weight:bold; text-align:center;}

.zipcodeV17 .tip {padding:5px; background-color:#fafafa;}
.zipcodeV17 .tip h3 {padding:13px 0 17px 27px; color:#010000; font-weight:bold;}
.zipcodeV17 .tip h3 span {display:inline-block; width:27px; height:15px; border:1px solid #000; border-radius:13px; font-size:10px; font-weight:normal; line-height:15px; text-transform:uppercase; text-align:center;}
.zipcodeV17 .tip ul {padding:27px 28px 25px; background-color:#fff;}
.zipcodeV17 .tip ul li {margin-top:12px; color:#888; font-size:12px; font-weight:bold;}
.zipcodeV17 .tip ul li:first-child {margin-top:0;}
.zipcodeV17 .tip ul li span {font-size:11px; font-weight:normal;}

.zipcodeV17 .total {margin-top:17px; padding-bottom:8px; color:#555; }
.zipcodeV17 .total em {color:#000; font-weight:bold;}
.zipcodeV17 .result ul {overflow-y:auto; position: relative; max-height:260px; border-top:1px solid #ddd; border-bottom:1px solid #ddd;}
.zipcodeV17 .result ul li {position:relative; padding:12px 0 13px; margin-right:12px; border-top:1px solid #eee; *zoom:1;}
.zipcodeV17 .result ul li:first-child {border-top:0;}
.zipcodeV17 .result ul li .zipcode,
.zipcodeV17 .result ul li a {overflow:hidden; display:block; margin:5px 70px 0 0; color:#555;}
.zipcodeV17 .result ul li .postcode + a {margin-top:0;}
.zipcodeV17 .result ul li .postcode {position:absolute; top:0; *top:25px; right:0; width:50px; height:100%; font-family:'Verdana'; font-weight:bold; text-align:center;}
.zipcodeV17 .result ul li .postcode span {display:table; width:100%; height:100%;}
.zipcodeV17 .result ul li .postcode span i {display:table-cell; width:100%; height:100%; vertical-align:middle; font-style:normal;}
.zipcodeV17 .result ul li a:hover,
.zipcodeV17 .result ul li a:hover em {color:#d50c0c; text-decoration:none;}
.zipcodeV17 .result ul li em {float:left; width:9.8%; color:#000; font-weight:bold;}
.zipcodeV17 .result ul li div {float:left; width:90.2%; cursor:pointer;}

.zipcodeV17 .pageWrapV15 {margin-top:30px;}
.zipcodeV17 .pageWrapV15 .pageMove {display:none;}

.zipcodeV17 .btnAreaV16a {margin-top:30px; text-align:center;}
.zipcodeV17 .btnAreaV16a .btn {width:198px; margin:0 3px; padding-right:0; padding-left:0; font-size:12px;}

</style>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script>
	$(function(){
		/* tab onoff */
		$(".tabonoff .tabcontainer .tabcont").css("display", "none");
		$(".tabonoff .tabcontainer .tabcont:first-child").css("display", "block");
		$(".tabonoff .tabs li:first-child a").addClass("on");
		$(".tabonoff").delegate(".tabs li", "click", function() {
			var index = $(this).parent().children().index(this);
			$(this).siblings().children().removeClass();
			$(this).children().addClass("on");
			$(this).parent().next(".tabcontainer").children().hide().eq(index).show();
			return false;
		});

		$(".finder .btnReset").hide();
		$(".finder input[type=search]" ).focus(function() {
			$(".finder .btnReset").show();
		});
	});

	function SubmitForm(stype) {

		<%'// 통합 검색일 경우 %>
		if (stype=="jibun")
		{
			if ($("#tJibundong").val().length < 2) { alert("검색어를 두 글자 이상 입력하세요."); return; }
			$("#sGubun").val(stype);
			$("#sJibundong").val($("#tJibundong").val());
			$("#cpg").val(1);
			$("#keyword").val("");
		}

		$.ajax({
			type:"get",
			url:"/lib/searchzip_newproc.asp",
		   data: $("#searchProcFrm").serialize(),
		   dataType: "text",
			async:false,
			cache:false,
			success : function(Data, textStatus, jqXHR){
				if (jqXHR.readyState == 4) {
					if (jqXHR.status == 200) {
						if(Data!="") {
							res = Data.split("|");
							if (res[0]=="OK")
							{
								if (stype=="jibun")
								{
									if (res[1]=="<p>검색된 주소가 없습니다</p>")
									{
										SubmitFormAPI();
									}
									else
									{
										$("#resultJibun").show();
										$("#guideTxtVal").hide();
										$("#noResultData").hide();
										$("#tipTxtVal").hide();
										setTimeout(function () {
											window.$('html,body').animate({scrollTop:$("#resultJibun").offset().top}, 0);
										}, 10);
										$("#jibunaddrList").empty().html(res[1]);
										if (res[3]!="")
										{
											$("#addrpaging").empty().html(res[3]);
										}
										$("#jibuntotalcntView").empty().html("총 <em>"+numberWithCommas(res[2])+"</em> 건");
									}
								}
							}
							else
							{
								errorMsg = res[1].replace(">?n", "\n");
								alert(errorMsg );
								return false;
							}
						} else {
							alert("잘못된 접근 입니다.");
							return false;
						}
					}
				}
			},
			error:function(jqXHR, textStatus, errorThrown){
				alert("잘못된 접근 입니다!!");
				return false;
			}
		});
	}

	<%'// 시군구 리스트 가져옴 %>
	function getgunguList(v, stype) {

		$("#sGubun").val("gungureturn");
		$("#sSidoGubun").val(v);

		if (v=="")
		{
			alert("시/도를 선택해 주세요.");
			return false;
		}

		<%'// 세종특별자치시는 시군구가 없으므로 안타도됨 %>
		if (v=="세종특별자치시")
		{
			$("#"+stype).empty().html("<option value=''>시/군/구 없음</option>");
		}
		else
		{
			$.ajax({
				type:"POST",
				url:"/lib/searchzip_newProc.asp",
			   data: $("#searchProcFrm").serialize(),
			   dataType: "text",
				async:false,
				cache:true,
				success : function(Data, textStatus, jqXHR){
					if (jqXHR.readyState == 4) {
						if (jqXHR.status == 200) {
							if(Data!="") {
								res = Data.split("|");
								if (res[0]=="OK")
								{
									$("#"+stype).empty().html(res[1]);
								}
								else
								{
									errorMsg = res[1].replace(">?n", "\n");
									alert(errorMsg );
									return false;
								}
							} else {
								alert("잘못된 접근 입니다.");
								return false;
							}
						}
					}
				},
				error:function(jqXHR, textStatus, errorThrown){
					alert("잘못된 접근 입니다!");
					return false;
				}
			});
		}
	}

	function numberWithCommas(x) {
		return x.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",");
	}

	function setBackAction(x, y, z) {
		$("#"+x).hide();
		$("#"+y).show();
		$("#"+z).show();
	}

	<%'// form에 각 값들 넣고 기본, 상세 주소 입력값 만듦 %>
	function setAddr(zip, sido, gungu, dong, eupmyun, ri, official_bld, jibun, road, building_no, type, wp, uwp) {

		var basicAddr; // 기본주소
		var basicAddr2; // 상세주소
		var roadbasicAddr; // 도로명으로 검색할시 표시할 지번주소

		$("#zip").val(zip);
		$("#sido").val(sido);
		$("#gungu").val(gungu);
		$("#dong").val(dong);
		$("#eupmyun").val(eupmyun);
		$("#ri").val(ri);
		$("#official_bld").val(official_bld);
		$("#jibun").val(jibun);
		$("#road").val(road);
		$("#building_no").val(building_no);
		$("#gubun").val(type);

		if (type=="jibun")
		{
			<%'// 기본주소 입력값을 만든다.%>
			basicAddr = "["+zip+"] "+sido+" "+gungu;
			if (dong=="")
			{
				basicAddr = basicAddr + " "+eupmyun;
			}
			else
			{
				basicAddr = basicAddr + " "+dong;
			}
			if (ri!="")
			{
				basicAddr = basicAddr + " "+ri;
			}
			if (official_bld!="")
			{
				basicAddr = basicAddr + jibun +" "+ official_bld;
			}
			else
			{
				basicAddr = basicAddr + " "+jibun;
			}
			$("#Jibunfinder").hide();
			$("#resultJibun").hide();
			$("#jibunDetail").show();
			$("#jibunDetailAddr").val(basicAddr);
		}

		if (type=="road")
		{
			<%'// 기본주소 입력값을 만든다.%>
			basicAddr = "["+zip+"] "+sido+" "+gungu;
			if (eupmyun!="")
			{
				basicAddr = basicAddr + " "+eupmyun+" "+road;
			}
			else
			{
				basicAddr = basicAddr + " "+road;
			}
			if (building_no!="")
			{
				basicAddr = basicAddr + " "+building_no;
			}
			<%'// 상세주소 입력값을 만든다.%>
			if (official_bld!="")
			{
				basicAddr2 = ""+official_bld+"";
			}

			$("#Jibunfinder").hide();
			$("#resultJibun").hide();
			$("#jibunDetail").show();
			$("#jibunDetailAddr").val(basicAddr);
		}

		$("#"+wp).empty().html(basicAddr);
		if (basicAddr2!="")
		{
			$("#"+uwp).val(basicAddr2);
		}
		$("#"+uwp).focus();
	}

	<%'// 모창에 값 던져줌(api 검색 또는 검색엔진) %>
	function CopyZip(x)	{
		
		<%'// api로 검색시에는 CopyZipAPI로 던져줌 %>
		if ($("#keyword").val()!="")
		{
			CopyZipAPI(x);
			return false;
		}

		var frm = eval("opener.document.<%=strTarget%>");
		var basicAddr;
		var basicAddr2;

		<%'// 기본주소 입력값을 만든다.%>
		basicAddr = $("#sido").val()+" "+$("#gungu").val();

		if ($("#gubun").val()=="jibun")
		{
			<%'// 상세주소 입력값을 만든다.%>
			if ($("#dong").val()=="")
			{
				basicAddr2 = $("#eupmyun").val();
			}
			else
			{
				basicAddr2 = $("#dong").val();
			}
			if ($("#ri").val()!="")
			{
				basicAddr2 = basicAddr2 + " "+$("#ri").val();
			}
			if ($("#official_bld").val()!="")
			{
				basicAddr2 = basicAddr2 + " "+$("#jibun").val()+" "+$("#official_bld").val();
			}
			else
			{
				basicAddr2 = basicAddr2 + " "+$("#jibun").val();
			}
			if ($("#"+x).val()!="")
			{
				basicAddr2 = basicAddr2 + " "+$("#"+x).val();
			}
		}
		if ($("#gubun").val()=="road")
		{
			if ($("#eupmyun").val()!="")
			{
				basicAddr2 = $("#eupmyun").val()+" "+$("#road").val();
			}
			else
			{
				basicAddr2 = $("#road").val();
			}
			if ($("#building_no").val()!="")
			{
				basicAddr2 = basicAddr2 + " "+$("#building_no").val();
			}
			if ($("#"+x).val()!="")
			{
				basicAddr2 = basicAddr2 + " "+$("#"+x).val();
			}
		}

		<% if strMode="A" then %>
			frm.reqzipcode.value = $("#zip").val();
			frm.reqzipaddr.value = basicAddr;
			frm.reqaddress.value = basicAddr2;
			frm.reqaddress.focus();
		<% elseif (strMode="B") then %>
			frm.zipcode.value = $("#zip").val();
			frm.zipaddr.value = basicAddr;
			frm.useraddr.value = basicAddr2;
			frm.useraddr.focus();
		<% elseif (strMode="C") then %>
			frm.company_zipcode.value = $("#zip").val();
			frm.company_address.value = basicAddr;
			frm.company_address2.value = basicAddr2;
			frm.company_address2.focus();
		<% elseif (strMode="D") then %>
			frm.return_zipcode.value = $("#zip").val();
			frm.return_address.value = basicAddr;
			frm.return_address2.value = basicAddr2;
			frm.return_address2.focus();
		<% elseif (strMode="E") then %>
			frm.zipcode.value = $("#zip").val();
			frm.addr1.value = basicAddr;
			frm.addr2.value = basicAddr2;
			frm.addr2.focus();
		<% elseif (strMode="F") then %>
			frm.shopzipcode.value = $("#zip").val();
			frm.shopaddr1.value = basicAddr;
			frm.shopaddr2.value = basicAddr2;
			frm.shopaddr2.focus();
		<% elseif (strMode="G") then %>
			frm.sPCd.value = $("#zip").val();
			frm.sAddr.value = basicAddr + " " + basicAddr2;
			frm.sAddr.focus();
		<% elseif (strMode="I") then %>
			frm.p_return_zipcode.value = $("#zip").val();
			frm.p_return_address.value = basicAddr;
			frm.p_return_address2.value = basicAddr2;
			frm.p_return_address2.focus();
		<% elseif (strMode="J") then %>
			frm.returnZipcode.value = $("#zip").val();
			frm.returnZipaddr.value = basicAddr;
			frm.returnEtcaddr.value = basicAddr2;
			frm.returnEtcaddr.focus();
		<% End if %>

		// close this window
		window.close();

	}

	function numkeyCheck(e) 
	{ 
		if (e.length>7)
		{
			alert("우편번호는 7자리까지 입력가능합니다.");
			$("#zipcode").val(e.substr(0, 7));
			return false;
		}
		else
		{
			return true; 
		}
	}

	<%' 직접입력용 %>
	function CopyZipUserInput()
	{

		var frm = eval("opener.document.<%=strTarget%>");

		if ($("#zipcode").val()=="")
		{
			alert("우편번호를 입력해주세요.");
			$("#zipcode").focus();
			return false;
		}

		if(!(/^\d{3}-?\d{3}$/.test($("#zipcode").val()) || /^\d{5}/.test($("#zipcode").val()))){
			alert("우편번호 형식이 아닙니다. 우편번호를 확인해주세요.");
			$("#zipcode").focus();
			return false;
		}

		if ($("#city1").val()=="")
		{
			alert("시/도를 선택해주세요.");
			$("#city1").focus()
			return false;
		}
		if ($("#city1").val()!="세종특별자치시")
		{
			if ($("#city2").val()=="")
			{
				alert("시/군/구를 선택해주세요.");
				$("#city2").focus()
				return false;
			}
		}
		if ($("#DetailAddr").val()=="")
		{
			alert("도로명/지번을 입력해주세요.");
			$("#DetailAddr").focus()
			return false;
		}

		<% if strMode="A" then %>
			frm.reqzipcode.value		= $("#zipcode").val();
			frm.reqzipaddr.value		= $("#city1").val()+" "+$("#city2").val();
			frm.reqaddress.value		= $("#DetailAddr").val()+" "+$("#DetailAddr2").val()
			frm.reqaddress.focus();
		<% elseif (strMode="B") then %>
			frm.zipcode.value		= $("#zipcode").val();
			frm.zipaddr.value		= $("#city1").val()+" "+$("#city2").val();
			frm.useraddr.value		= $("#DetailAddr").val()+" "+$("#DetailAddr2").val()
			frm.useraddr.focus();
		<% elseif (strMode="C") then %>
			frm.company_zipcode.value		= $("#zipcode").val();
			frm.company_address.value		= $("#city1").val()+" "+$("#city2").val();
			frm.company_address2.value		= $("#DetailAddr").val()+" "+$("#DetailAddr2").val()
			frm.company_address2.focus();
		<% elseif (strMode="D") then %>
			frm.return_zipcode.value		= $("#zipcode").val();
			frm.return_address.value		= $("#city1").val()+" "+$("#city2").val();
			frm.return_address2.value		= $("#DetailAddr").val()+" "+$("#DetailAddr2").val()
			frm.return_address2.focus();
		<% elseif (strMode="E") then %>
			frm.zipcode.value		= $("#zipcode").val();
			frm.addr1.value		= $("#city1").val()+" "+$("#city2").val();
			frm.addr2.value		= $("#DetailAddr").val()+" "+$("#DetailAddr2").val()
			frm.addr2.focus();
		<% elseif (strMode="F") then %>
			frm.shopzipcode.value		= $("#zipcode").val();
			frm.shopaddr1.value		= $("#city1").val()+" "+$("#city2").val();
			frm.shopaddr2.value		= $("#DetailAddr").val()+" "+$("#DetailAddr2").val()
			frm.shopaddr2.focus();
		<% elseif (strMode="G") then %>
			frm.sPCd.value		= $("#zipcode").val();
			frm.sAddr.value		= $("#city1").val()+" "+$("#city2").val()+" "+$("#DetailAddr").val()+" "+$("#DetailAddr2").val();
			frm.sAddr.focus();
		<% elseif (strMode="I") then %>
			frm.p_return_zipcode.value		= $("#zipcode").val();
			frm.p_return_address.value		= $("#city1").val()+" "+$("#city2").val();
			frm.p_return_address2.value		= $("#DetailAddr").val()+" "+$("#DetailAddr2").val()
			frm.p_return_address2.focus();
		<% elseif (strMode="J") then %>
			frm.returnZipcode.value		= $("#zipcode").val();
			frm.returnZipaddr.value		= $("#city1").val()+" "+$("#city2").val();
			frm.returnEtcaddr.value		= $("#DetailAddr").val()+" "+$("#DetailAddr2").val()
			frm.returnEtcaddr.focus();
		<% End if %>

		// close this window
		window.close();
	}

	function jsPageGo(icpg){
		var frm = document.searchProcFrm;
		frm.cpg.value=icpg;

		$.ajax({
			type:"get",
			url:"/lib/searchzip_newProc.asp",
		   data: $("#searchProcFrm").serialize(),
		   dataType: "text",
			async:false,
			cache:false,
			success : function(Data, textStatus, jqXHR){
				if (jqXHR.readyState == 4) {
					if (jqXHR.status == 200) {
						if(Data!="") {
							res = Data.split("|");
							if (res[0]=="OK")
							{
								$("#resultJibun").show();
								$("#jibunaddrList").empty().html(res[1]);
								if (res[3]!="")
								{
									$("#addrpaging").empty().html(res[3]);
								}
								$("#jibunaddrList").scrollTop(0);
							}
							else
							{
								errorMsg = res[1].replace(">?n", "\n");
								alert(errorMsg );
								return false;
							}
						} else {
							alert("잘못된 접근 입니다.");
							return false;
						}
					}
				}
			},
			error:function(jqXHR, textStatus, errorThrown){
				alert("잘못된 접근 입니다!");
				return false;
			},
			complete:function(){
				$(this).scrollTop(0);
			}

		});

	}

	<%' 검색 juso.go.kr api 사용영역 %>
	function SubmitFormAPI()
	{
		if ($("#tJibundong").val().length < 2) { alert("검색어를 두 글자 이상 입력하세요."); return; }
		$("#keyword").val($("#tJibundong").val());
		$("#currentPage").val(1);
		$.ajax({
/*
		     url :"http://www.juso.go.kr/addrlink/addrLinkApiJsonp.do"
			,type:"post"
			,data:$("#searchProcApi").serialize()
			,dataType:"jsonp"
			,cache:true
			,crossDomain:true
*/
			 url : "/lib/sz_gate.asp" 
			,type:"get"
			,data:$("#searchProcApi").serialize()
			,dataType:"jsonp"
			,cache:true
			,success:function(xmlStr){
				if(navigator.appName.indexOf("Microsoft") > -1){
					var xmlData = new ActiveXObject("Microsoft.XMLDOM");
					xmlData.loadXML(xmlStr.returnXml)
				}else{
					var xmlData = xmlStr.returnXml;
				}
				$("#jibunaddrList").html("");
				var errCode = $(xmlData).find("errorCode").text();
				var errDesc = $(xmlData).find("errorMessage").text();
				if(errCode != "0"){
					alert(errCode+"="+errDesc);
				}else{
					if ($(xmlData).find("totalCount").text()=="0")
					{
						$("#Jibunfinder").show();
						$("#guideTxtVal").hide();
						$("#tipTxtVal").show();
						$("#noResultData").show();
						$("#noResultData").empty().html("<p>검색된 주소가 없습니다</p>");
						$("#resultJibun").hide();
					}
					else
					{

						if(xmlStr != null){
							$("#resultJibun").show();
							$("#guideTxtVal").hide();
							$("#noResultData").hide();
							$("#tipTxtVal").hide();
							$("#jibuntotalcntView").empty().html("총 <em>"+$(xmlData).find("totalCount").text()+"</em> 건");
							window.$('html,body').animate({scrollTop:$("#resultJibun").offset().top}, 0);
							fnDisplayPaging_New_nottextboxdirectJS($("#currentPage").val(),$(xmlData).find("totalCount").text(),$("#countPerPage").val(),5,'jsPageGoAPI');
							makeList(xmlData);
						}
					}
				}
			}
		});
	}

	<%'// 페이징 자바스크립트 버전 %>
	function fnDisplayPaging_New_nottextboxdirectJS(strCurrentPage, intTotalRecord, intRecordPerPage, intBlockPerPage, strJsFuncName) {
		var intCurrentPage;
		var strCurrentPath;
		var vPageBody;
		var intStartBlock;
		var intEndBlock;
		var intTotalPage;
		var strParamName;
		var intLoop;

		<%'// 현재 페이지 설정 %>
		intCurrentPage = strCurrentPage;

		<%'// 해당 페이지에 표시되는 시작페이지와 마지막페이지 설정 %>
		intStartBlock = parseInt((intCurrentPage - 1) / intBlockPerPage) * intBlockPerPage + 1;
		intEndBlock = parseInt((intCurrentPage - 1) / intBlockPerPage) * intBlockPerPage + intBlockPerPage;

		<%'// 총 페이지 수 설정 %>
		intTotalPage = parseInt((intTotalRecord - 1)/intRecordPerPage) + 1

		if (intTotalPage < 1)
		{
			intTotalPage = 1;
		}

		vPageBody = "";
		vPageBody = vPageBody + "<div class='paging'>";
//		vPageBody = vPageBody + "<a href='#' title='첫 페이지' class='first arrow' onclick='"+(strJsFuncName)+"(1);return false;'><span style='cursor:pointer;'>맨 처음 페이지로 이동</span></a>&nbsp;";

		<%'// 이전 페이지 %>
		if (intStartBlock > 1)
		{
			vPageBody = vPageBody + "<a href='#' title='이전 페이지' class='prev arrow' onclick='"+strJsFuncName+"("+(intStartBlock-1)+");return false;'><span style='cursor:pointer;'>이전페이지로 이동</span></a>&nbsp;";
		}
		else
		{
			vPageBody = vPageBody + "<a href='#' title='이전 페이지' class='prev arrow' onclick='return false;'><span style='cursor:pointer;'>이전페이지로 이동</span></a>&nbsp;";
		}

		<%'// 현재 페이지 %>
		if (intTotalPage > 1)
		{
			for (intLoop = intStartBlock; intLoop<(intEndBlock+1); intLoop++)
			{
				if (intLoop > intTotalPage)
				{
					break;
				}
				if (intLoop == intCurrentPage) 
				{
					vPageBody = vPageBody + "<a href='#' title='"+intLoop+" 페이지' class='current' onclick='"+strJsFuncName+"("+(intLoop)+");return false;'><span style='cursor:pointer;'>"+intLoop+"</span></a>&nbsp;";
				}
				else
				{
					vPageBody = vPageBody + "<a href='#' title='"+intLoop+" 페이지' onclick='"+strJsFuncName+"("+(intLoop)+");return false;'><span style='cursor:pointer;'>"+intLoop+"</span></a>&nbsp;";
				}

			}
		}
		else
		{
			vPageBody = vPageBody + "<a href='#' title='1 페이지' class='current' onclick='"+strJsFuncName+"(1);return false;'><span style='cursor:pointer;'>1</span></a>&nbsp;";
		}
		<%'// 다음 페이지 %>
		if (intEndBlock < intTotalPage)
		{
			vPageBody = vPageBody + "<a href='#' title='다음 페이지' class='next arrow' onclick='"+strJsFuncName+"("+(intEndBlock+1)+");return false;'><span style='cursor:pointer;'>다음 페이지로 이동</span></a>&nbsp;";
		}
		else
		{
			vPageBody = vPageBody + "<a href='#' title='다음 페이지' class='next arrow' onclick='return false;'><span style='cursor:pointer;'>다음 페이지로 이동</span></a>&nbsp;";
		}

		<%'// 마지막 페이지 %>
//		vPageBody = vPageBody + "<a href='#' title='마지막 페이지' class='end arrow' onclick='"+strJsFuncName+"("+(intTotalPage)+");return false;'><span style='cursor:pointer;'>맨 마지막 페이지로 이동</span></a>&nbsp;";
		vPageBody = vPageBody + "</div>";

		$("#addrpaging").empty().html(vPageBody);

	}

	function jsPageGoAPI(icomp)
	{
		$("#currentPage").val(icomp);
		$.ajax({
/*
		     url :"http://www.juso.go.kr/addrlink/addrLinkApiJsonp.do"
			,type:"post"
			,data:$("#searchProcApi").serialize()
			,dataType:"jsonp"
			,crossDomain:true
			,cache:true
*/
			 url : "/lib/sz_gate.asp" 
			,type:"get"
			,data:$("#searchProcApi").serialize()
			,dataType:"jsonp"
			,cache:true
			,success:function(xmlStr){
				if(navigator.appName.indexOf("Microsoft") > -1){
					var xmlData = new ActiveXObject("Microsoft.XMLDOM");
					xmlData.loadXML(xmlStr.returnXml)
				}else{
					var xmlData = xmlStr.returnXml;
				}
				$("#jibunaddrList").html("");
				var errCode = $(xmlData).find("errorCode").text();
				var errDesc = $(xmlData).find("errorMessage").text();
				if(errCode != "0"){
					alert(errCode+"="+errDesc);
				}else{
					if ($(xmlData).find("totalCount").text()=="0")
					{
						$("#Jibunfinder").show();
						$("#guideTxtVal").hide();
						$("#tipTxtVal").show();
						$("#noResultData").show();
						$("#noResultData").empty().html("<p>검색된 주소가 없습니다</p>");
					}
					else
					{
						if(xmlStr != null){
							$("#Jibunfinder").show();
							$("#resultJibun").show();
							$("#JibunHelp").show();
							$("#jibuntotalcntView").empty().html("총 <em>"+$(xmlData).find("totalCount").text()+"</em> 건");
							window.$('html,body').animate({scrollTop:$("#resultJibun").offset().top}, 0);
							$("#jibunaddrList").scrollTop(0);
							fnDisplayPaging_New_nottextboxdirectJS($("#currentPage").val(),$(xmlData).find("totalCount").text(),$("#countPerPage").val(),5,'jsPageGoAPI');
							makeList(xmlData);
						}
					}
				}
			}
		});

	}

	function makeList(xmlStr){
		var htmlStr = "";
		$(xmlStr).find("juso").each(function(){
			var r = "'"+$(this).find('zipNo').text()+"','"+$(this).find('roadAddr').text()+"','jibunDetailAddr','jibunDetailAddr2'";
			var s = "'"+$(this).find('zipNo').text()+"','"+$(this).find('jibunAddr').text()+"','jibunDetailAddr','jibunDetailAddr2'";
			htmlStr += '<li><span class="postcode"><span><i>'+$(this).find('zipNo').text()+'</i></span></span>';
			htmlStr += '<a href="" onclick="setAddrAPI('+r+');return false;"><em>[도로]</em><div>'+$(this).find('roadAddr').text()+'</div></a>';
			htmlStr += '<a href="" onclick="setAddrAPI('+s+');return false;"><em>[지번]</em><div>'+$(this).find('jibunAddr').text();
			htmlStr += '</div></a></li>';

		});
		$("#jibunaddrList").empty().html(htmlStr);
	}

	function setAddrAPI(zip, addr, wp, uwp)
	{
		var basicAddr; // 기본주소

		basicAddr = "["+zip+"] "+addr;

		$("#Jibunfinder").hide();
		$("#resultJibun").hide();
		$("#jibunDetail").show();

		basicAddr = basicAddr.replace("  "," ");
		addr = addr.replace("  "," ");

		$("#tzip").val(zip);
		$("#taddr1").val(addr);

		$("#"+wp).val(basicAddr);
		$("#"+uwp).focus();
	}

	<%'// 모창에 값 던져줌 %>
	function CopyZipAPI(x)	{
		var frm = eval("opener.document.<%=strTarget%>");
		var basicAddr;
		var basicAddr2;
		var chkAddr;
		var tmpaddr;
		basicAddr = "";
		basicAddr2 = "";

		<%'// 기본주소 입력값을 만든다.%>
		tmpaddr = $("#taddr1").val().split(" ");

		if (tmpaddr.length >= 3)
		{
			if (tmpaddr[2].substring(tmpaddr[2].length-1, tmpaddr[2].length)=="구")
			{
				basicAddr = tmpaddr[0]+" "+tmpaddr[1]+" "+tmpaddr[2];
				chkAddr = "2";
			}
			else
			{
				basicAddr = tmpaddr[0]+" "+tmpaddr[1];
				chkAddr = "1";
			}
		}
		else
		{
			basicAddr = tmpaddr[0]+" "+tmpaddr[1];
			chkAddr = "1";
		}

		<%'// 상세주소 입력값을 만든다.%>
		for (var iadd=parseInt(chkAddr)+1;iadd < parseInt(tmpaddr.length);iadd++)
		{
			basicAddr2 += tmpaddr[iadd]+" ";
		}
		if ($("#"+x).val()!="")
		{
			basicAddr2 = basicAddr2 + $("#"+x).val();
		}

		<% if strMode="A" then %>
			frm.reqzipcode.value = $("#tzip").val();
			frm.reqzipaddr.value = basicAddr;
			frm.reqaddress.value = basicAddr2;
			frm.reqaddress.focus();
		<% elseif (strMode="B") then %>
			frm.zipcode.value = $("#tzip").val();
			frm.zipaddr.value = basicAddr;
			frm.useraddr.value = basicAddr2;
			frm.useraddr.focus();
		<% elseif (strMode="C") then %>
			frm.company_zipcode.value = $("#tzip").val();
			frm.company_address.value = basicAddr;
			frm.company_address2.value = basicAddr2;
			frm.company_address2.focus();
		<% elseif (strMode="D") then %>
			frm.return_zipcode.value = $("#tzip").val();
			frm.return_address.value = basicAddr;
			frm.return_address2.value = basicAddr2;
			frm.return_address2.focus();
		<% elseif (strMode="E") then %>
			frm.zipcode.value = $("#tzip").val();
			frm.addr1.value = basicAddr;
			frm.addr2.value = basicAddr2;
			frm.addr2.focus();
		<% elseif (strMode="F") then %>
			frm.shopzipcode.value = $("#tzip").val();
			frm.shopaddr1.value = basicAddr;
			frm.shopaddr2.value = basicAddr2;
			frm.shopaddr2.focus();
		<% elseif (strMode="G") then %>
			frm.sPCd.value = $("#tzip").val();
			frm.sAddr.value = basicAddr + " " + basicAddr2;
			frm.sAddr.focus();
		<% elseif (strMode="I") then %>
			frm.p_return_zipcode.value = $("#tzip").val();
			frm.p_return_address.value = basicAddr;
			frm.p_return_address2.value = basicAddr2;
			frm.p_return_address2.focus();
		<% elseif (strMode="J") then %>
			frm.returnZipcode.value = $("#tzip").val();
			frm.returnZipaddr.value = basicAddr;
			frm.returnEtcaddr.value = basicAddr2;
			frm.returnEtcaddr.focus();
		<% End if %>

		// close this window
		window.close();

	}
	<%'// 검색 juso.go.kr api 사용영역 %>

	function setResetVal()
	{
		$("#zipcode").val("");	
		$("#city1").val("");
		$("#city2").empty().html("<option value=''>시/군/구 선택</option>");
		$("#DetailAddr").val("");
		$("#DetailAddr2").val("");
	}

</script>
</head>
<body>
<%' for dev msg : 팝업 창 사이즈 width=578, height=690 %>
<div class="heightgird">
	<div class="popWrap">
		<div class="popHeader">
			<h1><img src="http://fiximage.10x10.co.kr/web2013/common/tit_zipcode_find.gif" alt="우편번호 찾기" /></h1>
		</div>
		<div class="popContent">

			<div class="tabonoff zipcodeV17">
				<ul class="tabs">
					<li><a href="#tabcont1">통합 검색</a></li>
					<li><a href="#tabcont2">직접 입력</a></li>
				</ul>

				<div class="tabcontainer">
					<%' tab1 통합 검색 %>
					<div id="tabcont1" class="tabcont">
						<h2 class="hidden">통합 검색</h2>

						<%' 검색 %>
						<div class="searchForm">
							<div class="finder" id="Jibunfinder">
								<form onsubmit="return false">
									<fieldset>
										<legend>주소 검색 폼</legend>
										<div class="inner">
											<span class="itext"><input type="search" id="tJibundong" title="검색어 입력" placeholder="예) 동숭동 1-45" onkeydown="javascript: if (event.keyCode == 13) {SubmitForm('jibun');}" /></span>
											<input type="reset" value="초기화" class="btnReset" />
										</div>
										<input type="submit" value="검색" class="btn btnM2 btnRed" onclick="SubmitForm('jibun');" />
									</fieldset>
								</form>
							</div>
						</div>

						<div class="guide" id="guideTxtVal">
							<p>도로명, 건물명, 지번을 이용해 주소를 검색해주세요</p>
						</div>

						<div class="guide noData" id="noResultData" style="display:none;"></div>

						<div class="tip" id="tipTxtVal">
							<h3><span>Tip</span> 효과적인 우편번호 검색방법</h3>
							<ul>
								<li>① 도로명 + 건물번호 검색 <span>대학로12길 31 , 사직로 161</span></li>
								<li>② 지역명(동/리) + 번지 검색 <span>동숭동 1-45 , 세종로 1-91</span></li>
								<li>③ 지역명(동/리) + 건물명(아파트명) 검색 <span>번동 주공</span></li>
							</ul>
						</div>

						<%' 검색결과 %>
						<div class="result" id="resultJibun" style="display:none;">
							<div class="total" id="jibuntotalcntView"></div>
							<ul id="jibunaddrList"></ul>
							<div id="addrpaging" class="pageWrapV15 tMar20"></div>
						</div>

						<%' 상세 주소 입력 %>
						<div class="searchForm" id="jibunDetail" style="display:none;">
							<form onsubmit="return false">
								<fieldset>
								<legend>상세 주소 입력</legend>
									<ul>
										<li>
											<label for="defaultAddress">선택주소</label>
											<span class="itext"><input type="text" id="jibunDetailAddr" readonly="readonly" /></span>
										</li>
										<li>
											<label for="detailAddress">상세주소</label>
											<span class="itext"><input type="text" id="jibunDetailAddr2" onkeydown="javascript: if (event.keyCode == 13) {CopyZip('jibunDetailAddr2', 'jibun');}"/></span>
										</li>
									</ul>

									<div class="btnAreaV16a">
										<a href="" class="btn btnM2 btnWhite" onclick="setBackAction('jibunDetail','resultJibun','Jibunfinder');return false;">이전</a>
										<input type="submit" class="btn btnM2 btnRed" value="확인" onclick="CopyZip('jibunDetailAddr2', 'jibun');" id="btnonsubmitSearchaddr" />
									</div>
								</fieldset>
							</form>
						</div>
					</div>
					<%' //tab1 %>

					<%' tab2 직접 입력 %>
					<div id="tabcont2" class="tabcont">
						<h2 class="hidden">직접 입력</h2>
						<div class="searchForm">
							<form onsubmit="return false">
								<fieldset>
								<legend>우편번호, 시/도, 시/군/구 및 도로명 또는 지번, 상세주소 입력 폼</legend>
									<ul>
										<li>
											<label for="zipcodeNo">우편번호</label>
											<span class="itext"><input type="text" id="zipcode" onkeyup="numkeyCheck(this.value);" maxlength="7" /></span>
										</li>
										<li>
											<label for="city1">시/도</label>
											<select id="city1" onchange="getgunguList(this.value, 'city2')">
												<option value="">시/도 선택</option>
												<option value="서울특별시">서울특별시</option>
												<option value="경기도">경기도</option>
												<option value="강원도">강원도</option>
												<option value="인천광역시">인천광역시</option>
												<option value="충청북도">충청북도</option>
												<option value="충청남도">충청남도</option>
												<option value="대전광역시">대전광역시</option>
												<option value="경상북도">경상북도</option>
												<option value="경상남도">경상남도</option>
												<option value="세종특별자치시">세종특별자치시</option>
												<option value="대구광역시">대구광역시</option>
												<option value="부산광역시">부산광역시</option>
												<option value="울산광역시">울산광역시</option>
												<option value="전라북도">전라북도</option>
												<option value="전라남도">전라남도</option>
												<option value="광주광역시">광주광역시</option>
												<option value="제주특별자치도">제주특별자치도</option>
											</select>
										</li>
										<li>
											<label for="city2">시/군/구</label>
											<select id="city2">
												<option value="">시/군/구 선택</option>
											</select>
										</li>
										<li>
											<label for="town">도로명/지번</label>
											<span class="itext"><input type="text" id="DetailAddr" /></span>
										</li>
										<li>
											<label for="address">상세주소</label>
											<span class="itext"><input type="text" id="DetailAddr2" /></span>
										</li>
									</ul>

									<div class="btnAreaV16a">
										<input type="reset" value="초기화" class="btn btnM2 btnWhite" onclick="setResetVal();return false;" />
										<input type="submit" value="확인" class="btn btnM2 btnRed" onclick="CopyZipUserInput();return false;" />
									</div>
								</fieldset>
							</form>
						</div>
					</div>
					<!-- //tab2 -->
				</div>
			</div>

		</div>
	</div>
	<div class="popFooter">
		<div class="btnArea">
			<button type="button" class="btn btnS1 btnGry2" onclick="window.close();">닫기</button>
		</div>
	</div>
</div>

<form name="searchProcFrm" id="searchProcFrm" method="post" style="margin:0px;" >
	<input type="hidden" name="sGubun" id="sGubun">
	<input type="hidden" name="sJibundong" id="sJibundong">
	<input type="hidden" name="sSidoGubun" id="sSidoGubun">
	<input type="hidden" name="sSido" id="sSido">
	<input type="hidden" name="sGungu" id="sGungu">
	<input type="hidden" name="sRoadName" id="sRoadName">
	<input type="hidden" name="sRoadBno" id="sRoadBno">
	<input type="hidden" name="sRoaddong" id="sRoaddong">
	<input type="hidden" name="sRoadjibun" id="sRoadjibun">
	<input type="hidden" name="sRoadBname" id="sRoadBname">
	<input type="hidden" name="cpg" id="cpg" value="<%=currpage%>">
	<input type="hidden" name="psz" id="psz" value="<%=pagesize%>">
</form>

<form name="tranFrm" id="tranFrm" method="post" style="margin:0px;" >
	<input type="hidden" name="zip" id="zip">
	<input type="hidden" name="sido" id="sido">
	<input type="hidden" name="gungu" id="gungu">
	<input type="hidden" name="dong" id="dong">
	<input type="hidden" name="eupmyun" id="eupmyun">
	<input type="hidden" name="ri" id="ri">
	<input type="hidden" name="official_bld" id="official_bld">
	<input type="hidden" name="jibun" id="jibun">
	<input type="hidden" name="road" id="road">
	<input type="hidden" name="building_no" id="building_no">
	<input type="hidden" name="gubun" id="gubun">
</form>

<form name="searchProcApi" id="searchProcApi" method="post" style="margin:0px;" >
	<input type="hidden" name="currentPage" id="currentPage" value="1"/>
	<input type="hidden" name="countPerPage" id="countPerPage" value="10"/> 
	<input type="hidden" name="confmKey" id="confmKey" value="U01TX0FVVEgyMDE2MDcwNDIwMjE0NDEzNTk5"/>
	<input type="hidden" name="keyword" id="keyword" value=""/>
</form>

<form name="tranFrmApi" id="tranFrmApi" method="post" style="margin:0px;" >
	<input type="hidden" name="tzip" id="tzip">
	<input type="hidden" name="taddr1" id="taddr1">
	<input type="hidden" name="taddr2" id="taddr2">
</form>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAnalclose.asp" -->