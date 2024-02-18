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
<!-- #include virtual="/partner/incSessionDesigner.asp" -->
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
<meta http-equiv="X-UA-Compatible" content="IE=edge" />
<meta http-equiv='Content-Type' content='text/html;charset=euc-kr' />
<title>우편번호 찾기</title>
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

/* popup layout */
html, body {height:100%;}
body > .heightgird {min-height:100%; height:auto;}
.heightgird {position:relative;}
.popWrap {padding-bottom:45px;}
.popWrap .popHeader {padding:27px 15px 15px; background:#d50c0c; color:#fff;}
.popWrap .popHeader h1 img {vertical-align:top;}
.popContent {padding:30px; font-size:11px;}
.popFooter {position:absolute; bottom:0; width:100%; padding:0; border-top:1px solid #ddd; background:#f5f5f5;}
.popFooter .btnArea {float:right; padding:8px 30px 11px 0;}
.popFooter .btnArea .btn {padding:5px 11px 3px 24px; border:0; border-bottom:1px solid #efefef; background:#999 url(http://fiximage.10x10.co.kr/web2013/common/btn_close_popup.gif) 11px center no-repeat; color:#fff;}
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

/* zipcode */
.zipcodeV16 .hidden {visibility:hidden; width:0; height:0; overflow:hidden; position:absolute; top:-1000%; line-height:0;}
.zipcodeV16 legend {visibility:hidden; width:0; height:0; overflow:hidden; position:absolute; top:-1000%; line-height:0;}
.zipcodeV16 .tabs {overflow:hidden; margin-left:-1px; border-left:1px solid #ddd;}
.zipcodeV16 .tabs li {float:left; width:50%;}
.zipcodeV16 .tabs li a {display:block; border:1px solid #ddd; border-left:0; background-color:#f5f5f5; color:#969696; font-size:13px; font-weight:bold; line-height:33px; text-align:center;}
.zipcodeV16 .tabs li a:hover {text-decoration:none;}
.zipcodeV16 .tabs .on {border-bottom:0; background-color:#fff; color:#555;}

.zipcodeV16 .tabsLine {width:100%; margin-top:20px; border-left:0;}
.zipcodeV16 .tabsLine li {float:left; width:33.333%;}
.zipcodeV16 .tabsLine li a {background-color:#fff; color:#888; font-size:11px; font-weight:normal;}
.zipcodeV16 .tabsLine li:first-child a {border-left:1px solid #ddd;}
.zipcodeV16 .tabsLine a.on {border-bottom:1px solid #999; border-color:#999; background-color:#999; color:#fff;}
.zipcodeV16 .tabsLine li:first-child a.on {border-color:#999;}

.zipcodeV16 input {font-family:'Dotum', '돋움'; font-size:11px;}
.zipcodeV16 select {appearance:none; -webkit-appearance:none; -moz-appearance:none; height:30px; padding-left:10px; border:1px solid #bbb; padding-right:25px; background:url(http://fiximage.10x10.co.kr/web2015/giftcard/bg_select_arr.gif) no-repeat 100% 50%; color:#555; font-family:'Dotum', '돋움'; font-size:11px;}
.zipcodeV16 select {padding-right:0\9;background:none\9;}
.zipcodeV16 select::-ms-expand {
	display:none;
}
.zipcodeV16 .itext {display:block; height:28px; padding:0 10px; border:1px solid #bbb;}
.zipcodeV16 .itext input {width:100%; height:28px; line-height:28px; background-color:transparent;}

.zipcodeV16 .finder {margin:0; padding:0;}
.zipcodeV16 .help {margin-top:15px; padding:18px 0 17px; background-color:#f5f5f5; color:#959595; text-align:center;}
.zipcodeV16 .help p:first-child {font-weight:bold;}

.zipcodeV16 .finder ul {overflow:hidden; padding:12px 16px 20px 0; border:5px solid #f5f5f5; border-top:0;}
.zipcodeV16 .finder ul li {float:left; margin-top:8px; width:50%;}
.zipcodeV16 .finder ul li div {position:relative; padding-left:66px;}
.zipcodeV16 .finder ul li label {position:absolute; top:0; left:13px; width:50px; height:30px; line-height:30px; text-align:left;}
.zipcodeV16 .finder ul li.child2 label,
.zipcodeV16 .finder ul li.child4 label {left:16px;}
.zipcodeV16 .finder ul li select {width:100%;}

.zipcodeV16 .btnAreaV16a {margin-top:30px; text-align:center;}
.zipcodeV16 .btnAreaV16a .btn {margin:0 3px; font-size:12px;}

.zipcodeV16 .reference {margin-top:32px; color:#888; font-size:11px; text-align:center;}
.zipcodeV16 .reference p {margin-top:16px;}
.zipcodeV16 .reference a {color:#888; font-weight:bold;}
.zipcodeV16 .reference a:hover {text-decoration:underline;}

.zipcodeV16 .result li {margin:0 5px; border-top:1px solid #eee;}
.zipcodeV16 .result li:first-child {border-top:0;}
.zipcodeV16 .result li a, .result li span {display:block; color:#555;}
.zipcodeV16 .result li a {padding:16px 0 15px; font-weight:bold;}
.zipcodeV16 .result li span {font-weight:normal;}
.zipcodeV16 .result li.nodata {padding:30px 0; color:#888; font-weight:bold; text-align:center;}

.zipcodeV16 .scrollbarwrap {overflow-y:auto; width:100%; min-height:65px; max-height:268px; border-bottom:1px solid #ddd;}

.zipcodeV16 .form .address {padding:18px 15px 20px; border:5px solid #f5f5f5; border-top:0;}
.zipcodeV16 .form .address p {color:#888;}
.zipcodeV16 .form .address p span {display:block;}
.zipcodeV16 .form .address p span:first-child {color:#000;}
.zipcodeV16 .form .address .itext {margin-top:13px;}

.jibeon .help {margin-top:33px; padding:0; background-color:#fff;}
.jibeon .scrollbarwrap {margin-top:33px; border-top:1px solid #ddd;}
.jibeon .finder .address {margin-top:30px; padding:18px 15px 20px; border:5px solid #ebebeb;}
.jibeon .finder .address .row {position:relative; padding-left:60px; margin-right:2px; text-align:left;}
.jibeon .finder .address .row label {position:absolute; top:0; left:0; width:60px; height:30px; line-height:30px;}
.jibeon .finder .address .row input {width:100%;}
.jibeon .form .address {margin-top:33px; border-top:5px solid #ebebeb;}

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

</style>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script>
	document.title="텐바이텐 우편번호 검색";
	$(function(){
		/* tab onoff */
		$(".tabonoff .tabcontainer .tabcont").css("display", "none");
		$(".tabonoff .tabcontainer .tabcont:nth-child(2)").css("display", "block");
		$(".tabonoff .parent li:nth-child(2) a").addClass("on");
		$(".tabonoff").delegate(".tabs li", "click", function() {
			var index = $(this).parent().children().index(this);
			if ( $(this).hasClass("first")) {
				$(".tabonoff .tabcontainer .tabcont").css("display", "none");
				$(".tabonoff .tabcontainer .tabcont:first-child").css("display", "block");
				$(this).siblings().children().removeClass();
				$(".tabonoff .parent li:first a").addClass("on");
				$(".tabsLine li a").removeClass("on");
				$(".tabsLine li:first-child a").addClass("on");
				return false;
			} else {
				$(this).siblings().children().removeClass();
				$(this).children().addClass("on");
				$(this).parent().next(".tabcontainer").children().hide().eq(index).show();
				return false;
			}
		});
	});


	<%'// 검색 %>
	function SubmitFormAPI()
	{
		if ($("#tJibundong").val().length < 2) { alert("검색어를 두 글자 이상 입력하세요."); return; }
		$("#keyword").val( $("#tJibundong").val());
		$("#currentPage").val(1);
		
		$("#keyword").val(encodeURIComponent($("#keyword").val())); // 인코딩
		$.ajax({
		    /*
		     url :"http://www.juso.go.kr/addrlink/addrLinkApiJsonp.do"
			,type:"post"
			,data:$("#searchProcApi").serialize()
			,dataType:"jsonp"
			,contentType: "application/x-www-form-urlencoded;charset=euc-kr"		
			,cache:false
			,crossDomain:true
			*/
			 url : "/lib/sz_gate.asp" 
			,type:"get"
			,data:$("#searchProcApi").serialize()
			,dataType:"jsonp"
			,cache:false
			//,crossDomain:true
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
						$("#Jibunfinder").hide();
						$("#resultJibun").show();
						window.$('html,body').animate({scrollTop:$("#resultJibun").offset().top}, 0);
						$("#addrpaging").empty();
						$("#jibunaddrList").empty().html("<li class='nodata'>검색된 주소가 없습니다.</li>");
					}
					else
					{
						if(xmlStr != null){
							$("#Jibunfinder").hide();
							$("#resultJibun").show();
							window.$('html,body').animate({scrollTop:$("#resultJibun").offset().top}, 0);
							if (parseInt($(xmlData).find("totalCount").text())>=100)
							{
								$("#cautionTxtJibun").empty().html("<p></p><p>검색 결과가 많을 경우 지번 또는 건물명과 함께 검색해주세요</p><p class='ex'>예) 동숭동 1-45, 동숭동 동숭아트센타</p>");
								$("#cautionTxtJibun").show();
							}
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
		vPageBody = vPageBody + "<a href='#' title='첫 페이지' class='first arrow' onclick='"+(strJsFuncName)+"(1);return false;'><span style='cursor:pointer;'>맨 처음 페이지로 이동</span></a>&nbsp;";

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
			,cache:false
			*/
			 url : "/lib/sz_gate.asp" 
			,type:"get"
			,data:$("#searchProcApi").serialize()
			,dataType:"jsonp"
			,cache:false
			//,crossDomain:true
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
						
						$("#Jibunfinder").hide();
						$("#resultJibun").show();
						window.$('html,body').animate({scrollTop:$("#resultJibun").offset().top}, 0);
						$("#jibunaddrList").empty().html("<li class='nodata'>검색된 주소가 없습니다.</li>");
					}
					else
					{
						if(xmlStr != null){
							$("#Jibunfinder").hide();
							$("#resultJibun").show();
							window.$('html,body').animate({scrollTop:$("#resultJibun").offset().top}, 0);
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
			var s = "'"+$(this).find('zipNo').text()+"','"+$(this).find('jibunAddr').text()+"','jibunDetailtxt','jibunDetailAddr2'";
			htmlStr += '<li><a href="" onclick="setAddrAPI('+s+');return false;">'+$(this).find('jibunAddr').text() +'<br>';
			htmlStr += "<span>도로명주소 : "+ $(this).find('roadAddr').text() +"</span></a></li>";
		});
		$("#jibunaddrList").empty().html(htmlStr);
	}

	function setAddrAPI(zip, addr, wp, uwp)
	{
		var basicAddr; // 기본주소

		basicAddr = "["+zip+"] "+addr;

		$("#resultJibun").hide();
		$("#jibunDetail").show();

		basicAddr = basicAddr.replace("  "," ");
		addr = addr.replace("  "," ");

		$("#tzip").val(zip);
		$("#taddr1").val(addr);

		$("#"+wp).empty().html(basicAddr);
		$("#"+uwp).focus();
	}

	<%'// 모창에 값 던져줌 %>
	function CopyZipAPI(x, y)	{
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
		<% End if %>

		// close this window
		window.close();

	}

	function SubmitForm(stype) {

		<%'// 지번 일 경우 %>
		if (stype=="jibun")
		{
			if ($("#tJibundong").val().length < 2) { alert("검색어를 두 글자 이상 입력하세요."); return; }
			$("#sGubun").val(stype);
			$("#sJibundong").val($("#tJibundong").val());
			$("#cpg").val(1);
		}

		<%'// 도로명+건물번호 일 경우 %>
		if (stype=="RoadBnumber")
		{
			if ($("#ctiy11").val()=="")
			{
				alert('시/도를 선택해 주세요.');
				return;
			}

			<%'// 세종특별자치시는 시군구가 없어서 체크안함 %>
			if ($("#ctiy11").val()!="세종특별자치시")
			{
				if ($("#ctiy12").val()=="")
				{
					alert('시/군/구를 선택해 주세요.');
					return;
				}
			}
			if ($("#NameRoadBnumber").val()=="")
			{
				alert('도로명을 입력해 주세요.');
				$("#NameRoadBnumber").focus();
				return;	
			}
			if ($("#NumberRoadBnumber").val()=="")
			{
				alert('건물번호를 입력해 주세요.');
				$("#NumberRoadBnumber").focus();
				return;	
			}

			$("#sGubun").val(stype);
			$("#sSido").val($("#ctiy11").val());
			$("#sGungu").val($("#ctiy12").val());
			$("#sRoadName").val($("#NameRoadBnumber").val());
			$("#sRoadBno").val($("#NumberRoadBnumber").val());
		}

		<%'// 도로명에 동(읍/면)+지번 일 경우 %>
		if (stype=="RoadBjibun")
		{
			if ($("#ctiy21").val()=="")
			{
				alert('시/도를 선택해 주세요.');
				return;
			}

			<%'// 세종특별자치시는 시군구가 없어서 체크안함 %>
			if ($("#ctiy21").val()!="세종특별자치시")
			{
				if ($("#ctiy22").val()=="")
				{
					alert('시/군/구를 선택해 주세요.');
					return;
				}
			}
			if ($("#DongRoadBjibun").val()=="")
			{
				alert('동(읍/면)을 입력해 주세요.');
				$("#DongRoadBjibun").focus();
				return;	
			}
			if ($("#JibunRoadBjibun").val()=="")
			{
				alert('지번을 입력해 주세요.');
				$("#JibunRoadBjibun").focus();
				return;	
			}
			$("#sGubun").val(stype);
			$("#sSido").val($("#ctiy21").val());
			$("#sGungu").val($("#ctiy22").val());
			$("#sRoaddong").val($("#DongRoadBjibun").val());
			$("#sRoadjibun").val($("#JibunRoadBjibun").val());
		}

		<%'// 도로명에 건물명 일 경우 %>
		if (stype=="RoadBname")
		{
			if ($("#ctiy31").val()=="")
			{
				alert('시/도를 선택해 주세요.');
				return;
			}

			<%'// 세종특별자치시는 시군구가 없어서 체크안함 %>
			if ($("#ctiy31").val()!="세종특별자치시")
			{
				if ($("#ctiy32").val()=="")
				{
					alert('시/군/구를 선택해 주세요.');
					return;
				}
			}
			if ($("#NameRoadBname").val()=="")
			{
				alert('건물명을 입력해 주세요.');
				$("#NameRoadBname").focus();
				return;	
			}
			$("#sGubun").val(stype);
			$("#sSido").val($("#ctiy31").val());
			$("#sGungu").val($("#ctiy32").val());
			$("#sRoadBname").val($("#NameRoadBname").val());
		}

		$.ajax({
			type:"get",
			url:"/designer/lib/searchzip_newProc.asp",
		   data: $("#searchProcFrm").serialize(),
		   dataType: "text",
		   contentType: "application/x-www-form-urlencoded;charset=euc-kr",		
			async:false,
			cache:false,
			success : function(Data, textStatus, jqXHR){
				if (jqXHR.readyState == 4) {
					if (jqXHR.status == 200) {
						if(Data!="") {
							var str;
							for(var i in Data)
							{
								 if(Data.hasOwnProperty(i))
								{
									str += Data[i];
								}
							}
							str = str.replace("undefined","");
							res = str.split("|");
							if (res[0]=="OK")
							{
								if (stype=="jibun")
								{
									$("#Jibunfinder").hide();
									$("#resultJibun").show();
									window.$('html,body').animate({scrollTop:$("#resultJibun").offset().top}, 0);
									$("#jibunaddrList").empty().html(res[1]);
									if (res[3]!="")
									{
										$("#addrpaging").empty().html(res[3]);
									}
									if (res[2] > 100)
									{
										$("#cautionTxtJibun").empty().html("<p></p><p>검색 결과가 많을 경우 지번 또는 건물명과 함께 검색해주세요</p><p class='ex'>예) 동숭동 1-45, 동숭동 동숭아트센타</p>");
										$("#cautionTxtJibun").show();
									}
									else
									{
										$("#cautionTxtJibun").empty();
									}
								}

								if (stype=="RoadBnumber")
								{
									$("#RoadBnumberfinder").hide();
									$("#resultRoadBnumber").show();
									window.$('html,body').animate({scrollTop:$("#resultRoadBnumber").offset().top}, 0);
									$("#RoadBnumberaddrList").empty().html(res[1]);
								}

								if (stype=="RoadBjibun")
								{
									$("#RoadBjibunfinder").hide();
									$("#resultRoadBjibun").show();
									window.$('html,body').animate({scrollTop:$("#resultRoadBjibun").offset().top}, 0);
									$("#RoadBjibunaddrList").empty().html(res[1]);
								}

								if (stype=="RoadBname")
								{
									$("#RoadBnamefinder").hide();
									$("#resultRoadBname").show();
									window.$('html,body').animate({scrollTop:$("#resultRoadBname").offset().top}, 0);
									$("#RoadBnameaddrList").empty().html(res[1]);
								}
							}
							else
							{
								errorMsg = res[1].replace(">?n", "\n");
								alert(errorMsg );
								return false;
							}
						} else {
							alert("잘못된 접근 입니다[1].");
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
				url:"/designer/lib/searchzip_newProc.asp",
			   data: $("#searchProcFrm").serialize(),
			   dataType: "text",
			   contentType: "application/x-www-form-urlencoded;charset=euc-kr",		
				async:false,
				cache:false,
				success : function(Data, textStatus, jqXHR){
					if (jqXHR.readyState == 4) {
						if (jqXHR.status == 200) {
							if(Data!="") {
								var str;
								for(var i in Data)
								{
									 if(Data.hasOwnProperty(i))
									{
										str += Data[i];
									}
								}
								str = str.replace("undefined","");
								res = str.split("|");
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
								alert("잘못된 접근 입니다[2].");
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
	}

	function numberWithCommas(x) {
		return x.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",");
	}

	function setBackAction(x, y) {
		$("#"+x).hide();
		$("#"+y).show();
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
			<%'// 상세주소 입력값을 만든다.%>
			if (official_bld!="")
			{
				basicAddr2 = official_bld+" "+jibun;
			}
			else
			{
				basicAddr2 = jibun;
			}
			$("#resultJibun").hide();
			$("#jibunDetail").show();
		}

		if (type=="RoadBnumber")
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

			<%' // 지번주소 입력값을 만든다.%>
			roadbasicAddr = sido+" "+gungu;
			if (dong=="")
			{
				roadbasicAddr = roadbasicAddr + " "+eupmyun;
			}
			else
			{
				roadbasicAddr = roadbasicAddr + " "+dong;
			}
			if (ri!="")
			{
				roadbasicAddr = roadbasicAddr + " "+ri;
			}
			if (official_bld!="")
			{
				roadbasicAddr = roadbasicAddr + " "+official_bld+" "+jibun;
			}
			else
			{
				roadbasicAddr = roadbasicAddr + " "+jibun;
			}
			$("#RoadBnumberJibunDetail").empty().html("지번 주소 : "+roadbasicAddr);
			$("#resultRoadBnumber").hide();
			$("#RoadBnumberDetail").show();
		}

		if (type=="RoadBjibun")
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

			<%' // 지번주소 입력값을 만든다.%>
			roadbasicAddr = sido+" "+gungu;
			if (dong=="")
			{
				roadbasicAddr = roadbasicAddr + " "+eupmyun;
			}
			else
			{
				roadbasicAddr = roadbasicAddr + " "+dong;
			}
			if (ri!="")
			{
				roadbasicAddr = roadbasicAddr + " "+ri;
			}
			if (official_bld!="")
			{
				roadbasicAddr = roadbasicAddr + " "+official_bld+" "+jibun;
			}
			else
			{
				roadbasicAddr = roadbasicAddr + " "+jibun;
			}
			$("#RoadBjibunJibunDetail").empty().html("지번 주소 : "+roadbasicAddr);
			$("#resultRoadBjibun").hide();
			$("#RoadBjibunDetail").show();
		}

		if (type=="RoadBname")
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

			<%' // 지번주소 입력값을 만든다.%>
			roadbasicAddr = sido+" "+gungu;
			if (dong=="")
			{
				roadbasicAddr = roadbasicAddr + " "+eupmyun;
			}
			else
			{
				roadbasicAddr = roadbasicAddr + " "+dong;
			}
			if (ri!="")
			{
				roadbasicAddr = roadbasicAddr + " "+ri;
			}
			if (official_bld!="")
			{
				roadbasicAddr = roadbasicAddr + " "+official_bld+" "+jibun;
			}
			else
			{
				roadbasicAddr = roadbasicAddr + " "+jibun;
			}
			$("#RoadBnameJibunDetail").empty().html("지번 주소 : "+roadbasicAddr);
			$("#resultRoadBname").hide();
			$("#RoadBnameDetail").show();
		}

		$("#"+wp).empty().html(basicAddr);
		if (basicAddr2!="")
		{
			$("#"+uwp).val(basicAddr2);
		}
		$("#"+uwp).focus();
	}


	<%'// 모창에 값 던져줌 %>
	function CopyZip(x, y)	{
		var frm = eval("opener.document.<%=strTarget%>");
		var basicAddr;
		var basicAddr2;

		<%'// 기본주소 입력값을 만든다.%>
		basicAddr = $("#sido").val()+" "+$("#gungu").val();

		if (y=="jibun")
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
			if ($("#"+x).val()!="")
			{
				basicAddr2 = basicAddr2 + " "+$("#"+x).val();
			}
		}
		if (y=="RoadBnumber")
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
		if (y=="RoadBjibun")
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
		if (y=="RoadBname")
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
		<% End if %>

		// close this window
		window.close();

	}

	function jsPageGo(icpg){
		var frm = document.searchProcFrm;
		frm.cpg.value=icpg;

		$.ajax({
			type:"get",
			url:"/designer/lib/searchzip_newProc.asp",
		   data: $("#searchProcFrm").serialize(),
		   dataType: "text",
			async:false,
			cache:false,
			success : function(Data, textStatus, jqXHR){
				if (jqXHR.readyState == 4) {
					if (jqXHR.status == 200) {
						if(Data!="") {
							var str;
							for(var i in Data)
							{
								 if(Data.hasOwnProperty(i))
								{
									str += Data[i];
								}
							}
							str = str.replace("undefined","");
							res = str.split("|");
							if (res[0]=="OK")
							{
								$("#Jibunfinder").hide();
								$("#resultJibun").show();
								window.$('html,body').animate({scrollTop:$("#resultJibun").offset().top}, 0);
								$("#jibunaddrList").empty().html(res[1]);
								if (res[3]!="")
								{
									$("#addrpaging").empty().html(res[3]);
								}
								if (res[2] > 100)
								{
									$("#cautionTxtJibun").empty().html("<p></p><p>검색 결과가 많을 경우 지번 또는 건물명과 함께 검색해주세요</p><p class='ex'>예) 동숭동 1-45, 동숭동 동숭아트센타</p>");
									$("#cautionTxtJibun").show();
								}
								else
								{
									$("#cautionTxtJibun").empty();
								}
							}
							else
							{
								errorMsg = res[1].replace(">?n", "\n");
								alert(errorMsg );
								return false;
							}
						} else {
							alert("잘못된 접근 입니다[3].");
							return false;
						}
					}
				}
			},
			error:function(jqXHR, textStatus, errorThrown){
				alert("잘못된 접근 입니다!!!");
				return false;
			}
		});

	}
</script>
</head>
<body>
	<div class="heightgird">
		<!-- ------------------------------------------------ -->
		<div class="popWrap">
			<div class="popHeader">
				<h1><img src="http://fiximage.10x10.co.kr/web2013/common/tit_zipcode_find.gif" alt="우편번호 찾기" /></h1>
			</div>
			<div class="popContent">

				<div class="tabonoff zipcodeV16">
					<ul class="tabs parent">
						<li class="first"><a href="#tabcont1">도로명 주소</a></li>
						<li><a href="#tabcont2">지번 주소</a></li>
					</ul>

					<div class="tabcontainer">
						<%' tab1 도로명 주소 %>
						<div id="tabcont1" class="tabcont">
							<h2 class="hidden">도로명 주소</h2>
							<div class="tabonoff">
								<ul class="tabs tabsLine">
									<li class="tabs1"><a href="#tabcont1-1">도로명 + 건물번호</a></li>
									<li class="tabs2"><a href="#tabcont1-2">동(읍/면) + 지번</a></li>
									<li class="tabs3"><a href="#tabcont1-3">건물명</a></li>
								</ul>
								<div class="tabcontainer">
									<%' tab1-1 도로명 + 건물번호 %>
									<div id="tabcont1-1" class="tabcont">
										<h3 class="hidden">도로명 + 건물번호</h3>

										<%' 검색 %>
										<div class="finder" id="RoadBnumberfinder">
											<fieldset>
												<legend>도로명 + 건물번호로 우편번호 찾기</legend>
												<div class="help">
													<p>도로명, 건물번호 를 입력 후 검색해주세요</p>
													<p class="ex">예) 대학로12길(도로명) 31 (건물번호)</p>
												</div>

												<ul>
													<li class="child1">
														<div>
															<label for="ctiy11">시/도</label>
															<select id="ctiy11" onchange="getgunguList(this.value, 'ctiy12')">
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
														</div>
													</li>
													<li class="child2">
														<div>
															<label for="ctiy12">시/군/구</label>
															<select id="ctiy12">
																<option>시/군/구 선택</option>
															</select>
														</div>
													</li>
													<li class="child3">
														<div>
															<label for="road">도로명</label>
															<span class="itext"><input type="text" id="NameRoadBnumber" /></span>
														</div>
													</li>
													<li class="child4">
														<div>
															<label for="buildingno">건물번호</label>
															<span class="itext"><input type="text" id="NumberRoadBnumber" onkeydown="javascript: if (event.keyCode == 13) {SubmitForm('RoadBnumber');}" /></span>
														</div>
													</li>
												</ul>

												<div class="btnAreaV16a">
													<input type="submit" class="btn btnM2 btnRed btnW220" value="검색" onclick="SubmitForm('RoadBnumber');" />
												</div>
											</fieldset>

											<div class="reference">
												<p>도로명 주소 검색 결과가 없을 경우,<br /> 도로명 주소 안내시스템을 참고해주시길 바랍니다</p>
												<p><a href="http://www.juso.go.kr" target="_blank">http://www.juso.go.kr</a></p>
											</div>
										</div>

										<%' 검색결과 %>
										<div class="result" id="resultRoadBnumber" style="display:none;">
											<div class="help">
												<p>아래 주소중 해당하는 주소를 선택해주세요</p>
											</div>

											<div class="scrollbarwrap">
												<ul class="list" id="RoadBnumberaddrList"></ul>
											</div>

											<div class="btnAreaV16a">
												<a href="" class="btn btnM2 btnWhite btnW220" onclick="setBackAction('resultRoadBnumber','RoadBnumberfinder');return false;">이전</a>
											</div>
										</div>

										<%' 상세주소 입력 %>
										<div class="form" id="RoadBnumberDetail" style="display:none;">
											<fieldset>
												<legend>상세주소 입력</legend>
												<div class="help">
													<p>상세 주소를 입력하신 후 &apos;주소입력&apos; 버튼을 눌러주세요</p>
												</div>

												<div class="address">
													<p><span id="RoadBnumberDetailTxt"></span><span id="RoadBnumberJibunDetail"></span></p>
													<div class="itext"><input type="text" title="상세주소 입력" id="RoadBnumberDetailAddr2" placeholder="상세 주소를 입력해주세요" onkeydown="javascript: if (event.keyCode == 13) {CopyZip('RoadBnumberDetailAddr2', 'RoadBnumber');}" /></div>
												</div>

												<div class="btnAreaV16a">
													<a href="" class="btn btnM2 btnWhite btnW150" onclick="setBackAction('RoadBnumberDetail','resultRoadBnumber');return false;">이전</a>
													<input type="submit" class="btn btnM2 btnRed btnW150" value="주소입력" onclick="CopyZip('RoadBnumberDetailAddr2', 'RoadBnumber');" />
												</div>
											</fieldset>
										</div>
									</div>
									<%' //tab1-1 %>

									<%' tab1-2 동(읍/면) + 지번 %>
									<div id="tabcont1-2" class="tabcont">
										<h3 class="hidden">동(읍/면) + 지번</h3>

										<%' 검색 %>
										<div class="finder" id="RoadBjibunfinder">
											<fieldset>
												<legend>동(읍/면) + 지번으로 우편번호 찾기</legend>
												<div class="help">
													<p>동(읍/면), 지번 입력 후 검색해주세요</p>
													<p class="ex">예) 동숭동(동) 1-45 (지번)</p>
												</div>

												<ul>
													<li class="child1">
														<div>
															<label for="ctiy21">시/도</label>
															<select id="ctiy21" onchange="getgunguList(this.value, 'ctiy22')">
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
														</div>
													</li>
													<li class="child2">
														<div>
															<label for="ctiy22">시/군/구</label>
															<select id="ctiy22">
																<option>시/군/구 선택</option>
															</select>
														</div>
													</li>
													<li class="child3">
														<div>
															<label for="town">동(읍/면)</label>
															<span class="itext"><input type="text" id="DongRoadBjibun" /></span>
														</div>
													</li>
													<li class="child4">
														<div>
															<label for="addressno">지번</label>
															<span class="itext"><input type="text" id="JibunRoadBjibun" onkeydown="javascript: if (event.keyCode == 13) {SubmitForm('RoadBjibun');}"/></span>
														</div>
													</li>
												</ul>

												<div class="btnAreaV16a">
													<input type="submit" class="btn btnM2 btnRed btnW220" value="검색" onclick="SubmitForm('RoadBjibun');" />
												</div>
											</fieldset>
											<div class="reference">
												<p>도로명 주소 검색 결과가 없을 경우,<br /> 도로명 주소 안내시스템을 참고해주시길 바랍니다</p>
												<p><a href="http://www.juso.go.kr" target="_blank">http://www.juso.go.kr</a></p>
											</div>
										</div>

										<%' 검색결과 %>
										<div class="result" id="resultRoadBjibun" style="display:none;">
											<div class="help">
												<p>아래 주소중 해당하는 주소를 선택해주세요</p>
											</div>

											<div class="scrollbarwrap">
												<ul class="list" id="RoadBjibunaddrList"></ul>
											</div>

											<div class="btnAreaV16a">
												<a href="" class="btn btnM2 btnWhite btnW220" onclick="setBackAction('resultRoadBjibun','RoadBjibunfinder');return false;">이전</a>
											</div>
										</div>

										<%' 상세주소 입력 %>
										<div class="form" id="RoadBjibunDetail" style="display:none;">
											<fieldset>
												<legend>상세주소 입력</legend>
												<div class="help">
													<p>상세 주소를 입력하신 후 &apos;주소입력&apos; 버튼을 눌러주세요</p>
												</div>

												<div class="address">
													<p><span id="RoadBjibunDetailTxt"></p><span id="RoadBjibunJibunDetail"></span></p>
													<div class="itext"><input type="text" title="상세주소 입력" placeholder="상세 주소를 입력해주세요" id="RoadBjibunDetailAddr2" onkeydown="javascript: if (event.keyCode == 13) {CopyZip('RoadBjibunDetailAddr2', 'RoadBjibun');}" /></div>
												</div>

												<div class="btnAreaV16a">
													<a href="" class="btn btnM2 btnWhite btnW150" onclick="setBackAction('RoadBjibunDetail','resultRoadBjibun');return false;">이전</a>
													<input type="submit" class="btn btnM2 btnRed btnW150" value="주소입력" onclick="CopyZip('RoadBjibunDetailAddr2', 'RoadBjibun');" />
												</div>
											</fieldset>
										</div>
									</div>
									<%' //tab1-2 %>

									<%' tab1-3 건물명 %>
									<div id="tabcont1-3" class="tabcont">
										<h3 class="hidden">건물명</h3>

										<%' 검색 %>
										<div class="finder" id="RoadBnamefinder">
											<fieldset>
												<legend>건물명으로 우편번호 찾기</legend>
												<div class="help">
													<p>건물명을 입력 후 검색해주세요</p>
													<p class="ex">예) 동숭아트센타 (건물번호)</p>
												</div>

												<ul>
													<li class="child1">
														<div>
															<label for="ctiy31">시/도</label>
															<select id="ctiy31"  onchange="getgunguList(this.value, 'ctiy32')">
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
														</div>
													</li>
													<li class="child2">
														<div>
															<label for="ctiy32">시/군/구</label>
															<select id="ctiy32">
																<option>시/군/구 선택</option>
															</select>
														</div>
													</li>
													<li class="child3">
														<div>
															<label for="building">건물명</label>
															<span class="itext"><input type="text" id="NameRoadBname" onkeydown="javascript: if (event.keyCode == 13) {SubmitForm('RoadBname');}"/></span>
														</div>
													</li>
												</ul>

												<div class="btnAreaV16a">
													<input type="submit" class="btn btnM2 btnRed btnW220" value="검색" onclick="SubmitForm('RoadBname');" />
												</div>
											</fieldset>
											<div class="reference">
												<p>도로명 주소 검색 결과가 없을 경우,<br /> 도로명 주소 안내시스템을 참고해주시길 바랍니다</p>
												<p><a href="http://www.juso.go.kr" target="_blank">http://www.juso.go.kr</a></p>
											</div>
										</div>

										<%' 검색결과 %>
										<div class="result" id="resultRoadBname" style="display:none;">
											<div class="help">
												<p>아래 주소중 해당하는 주소를 선택해주세요</p>
											</div>

											<div class="scrollbarwrap">
												<ul class="list" id="RoadBnameaddrList"></ul>
											</div>

											<div class="btnAreaV16a">
												<a href="" class="btn btnM2 btnWhite btnW220" onclick="setBackAction('resultRoadBname','RoadBnamefinder');return false;">이전</a>
											</div>
										</div>

										<%' 상세주소 입력 %>
										<div class="form" id="RoadBnameDetail" style="display:none;">
											<fieldset>
												<legend>상세주소 입력</legend>
												<div class="help">
													<p>상세 주소를 입력하신 후 &apos;주소입력&apos; 버튼을 눌러주세요</p>
												</div>

												<div class="address">
													<p><span id="RoadBnameDetailTxt"></p><span id="RoadBnameJibunDetail"></span></p>
													<div class="itext"><input type="text" title="상세주소 입력" placeholder="상세 주소를 입력해주세요" id="RoadBnameDetailAddr2" onkeydown="javascript: if (event.keyCode == 13) {CopyZip('RoadBnameDetailAddr2', 'RoadBname');}"/></div>
												</div>

												<div class="btnAreaV16a">
													<a href="" class="btn btnM2 btnWhite btnW150" onclick="setBackAction('RoadBnameDetail','resultRoadBname');return false;">이전</a>
													<input type="submit" class="btn btnM2 btnRed btnW150" value="주소입력" onclick="CopyZip('RoadBnameDetailAddr2', 'RoadBname');" />
												</div>
											</fieldset>
										</div>
									</div>
									<%' //tab1-3 %>
								</div>
							</div>
						</div>
						<%' //tab1 %>

						<%' tab2 지번 주소 %>
						<div id="tabcont2" class="tabcont jibeon">
							<h2 class="hidden">지번 주소</h2>

							<%' 검색 %>
							<div class="finder" id="Jibunfinder">
								<fieldset>
									<legend>동(읍/면)으로 우편번호 찾기</legend>
									<div class="help">
										<p>찾고 싶으신 주소의 동(읍/면) 또는 동(읍/면) 지번, 건물명을 입력해주세요</p>
										<p class="ex">예) 동숭동, 동숭동 1-45, 동숭동 동숭아트센타</p>
									</div>

									<div class="address">
										<div class="row">
											<label for="dong">동(읍/면)</label>
											<%'// 만약 api서버가 죽으면 아래 주석 풀고 내부서버로 돌린다. %>
											<!--<span class="itext"><input type="text" id="tJibundong" placeholder="동숭동" onkeydown="javascript: if (event.keyCode == 13) {SubmitForm('jibun');}" /></span>-->
											<span class="itext"><input type="text" id="tJibundong" placeholder="동숭동" onkeydown="javascript: if (event.keyCode == 13) {SubmitFormAPI();}" /></span>
										</div>
									</div>

									<div class="btnAreaV16a">
										<%'// 만약 api서버가 죽으면 아래 주석 풀고 내부서버로 돌린다. %>
										<!--<input type="submit" class="btn btnM2 btnRed btnW220" value="검색" onclick="SubmitForm('jibun');"/>-->
										<input type="submit" class="btn btnM2 btnRed btnW220" value="검색" onclick="SubmitFormAPI();"/>
									</div>
								</fieldset>
							</div>

							<%' 검색결과 %>
							<div class="result" id="resultJibun" style="display:none;">
								<div class="help">
									<p>아래 주소중 해당하는 주소를 선택해주세요</p>
									<span id="cautionTxtJibun"></span>
								</div>

								<div class="scrollbarwrap">
									<ul class="list" id="jibunaddrList"></ul>
								</div>
								
								<div id="addrpaging" class="pageWrapV15 tMar20"></div>

								<div class="btnAreaV16a">
									<a href="" class="btn btnM2 btnWhite btnW220" onclick="setBackAction('resultJibun','Jibunfinder');return false;">이전</a>
								</div>
							</div>

							<%' 상세주소 입력 %>
							<div class="form" id="jibunDetail" style="display:none;">
								<fieldset>
									<div class="help">
										<p>상세 주소를 입력하신 후 &apos;주소입력&apos; 버튼을 눌러주세요</p>
									</div>

									<div class="address">
										<p><div id="jibunDetailtxt"></div></p>
										<span class="itext"><input type="text" title="상세주소 입력" id="jibunDetailAddr2" value="" placeholder="상세 주소를 입력해주세요" onkeydown="javascript: if (event.keyCode == 13) {CopyZipAPI('jibunDetailAddr2', 'jibun');}"  /></span>
									</div>

									<div class="btnAreaV16a">
										<a href="" class="btn btnM2 btnWhite btnW150" onclick="setBackAction('jibunDetail','resultJibun');return false;">이전</a>
										<%'// 만약 api서버가 죽으면 아래 주석 풀고 내부서버로 돌린다. %>
										<!--<input type="submit" class="btn btnM2 btnRed btnW150" onclick="CopyZip('jibunDetailAddr2', 'jibun');" value="주소입력" />-->
										<input type="submit" class="btn btnM2 btnRed btnW150" onclick="CopyZipAPI('jibunDetailAddr2', 'jibun');" value="주소입력" />
									</div>
								</fieldset>
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
		<form name="searchProcFrm" id="searchProcFrm" method="post">
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

		<form name="tranFrm" id="tranFrm" method="post">
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
		</form>

		<form name="searchProcApi" id="searchProcApi" method="post">
			<input type="hidden" name="currentPage" id="currentPage" value="1"/>
			<input type="hidden" name="countPerPage" id="countPerPage" value="10"/> 
			<input type="hidden" name="confmKey" id="confmKey" value="U01TX0FVVEgyMDE2MDcwNDIwMjE0NDEzNTk5"/>
			<input type="hidden" name="keyword" id="keyword" value=""/>
		</form>

		<form name="tranFrmApi" id="tranFrmApi" method="post">
			<input type="hidden" name="tzip" id="tzip">
			<input type="hidden" name="taddr1" id="taddr1">
			<input type="hidden" name="taddr2" id="taddr2">
		</form>
		<!-- ------------------------------------------------ -->
	</div>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAnalclose.asp" -->