<%@ language=vbscript %>
<% option explicit %>
<%
session.codePage = 949
Response.CharSet = "EUC-KR"

'###########################################################
' Description :  우편번호 찾기(카카오 API)
' History : 2019.06.13 원승현 생성
'           2019.07.30 한용민 프론트 이전 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->

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
.popWrap .popHeader h1 img {vertical-align:top;}

</style>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">
    $(function(){
        searchZipKakao();
	});

	<%'// 모창에 값 던져줌 %>
	function CopyZipAPI()	{
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
		if ($("#extraAddr").val()!="")
		{
			basicAddr2 = basicAddr2 + $("#extraAddr").val();
		}

		<% if strMode="A" then %>
			// copy
			frm.reqzipcode.value		= $("#tzip").val();
			frm.reqzipaddr.value		= basicAddr
			frm.reqaddress.value		= basicAddr2
			// focus
			frm.reqaddress.focus();

		<% elseif strMode="B" then %>
			// copy
			frm.zipcode.value		= $("#tzip").val();
			frm.zipaddr.value		= basicAddr
			frm.useraddr.value		= basicAddr2
			// focus
			frm.useraddr.focus();

		<% elseif strMode="C" then %>
			// copy
			frm.company_zipcode.value		= $("#tzip").val();
			frm.company_address.value		= basicAddr
			frm.company_address2.value		= basicAddr2
			// focus
			frm.company_address2.focus();

		<% elseif strMode="D" then %>
			// copy
			frm.return_zipcode.value		= $("#tzip").val();
			frm.return_address.value		= basicAddr
			frm.return_address2.value		= basicAddr2
			// focus
			frm.return_address2.focus();

		<% elseif strMode="E" then %>
			// copy
			frm.zipcode.value		= $("#tzip").val();
			frm.addr1.value		= basicAddr
			frm.addr2.value		= basicAddr2
			// focus
			frm.addr2.focus();

		<% elseif strMode="F" then %>
			// copy
			frm.shopzipcode.value		= $("#tzip").val();
			frm.shopaddr1.value		= basicAddr
			frm.shopaddr2.value		= basicAddr2
			// focus
			frm.shopaddr2.focus();

		<% elseif strMode="G" then %>
			// copy
			frm.sPCd.value		= $("#tzip").val();
			frm.sAddr.value		= basicAddr + " " + basicAddr2;
			// focus
			frm.sAddr.focus();

		<% elseif strMode="I" then %>
			// copy
			frm.p_return_zipcode.value		= $("#tzip").val();
			frm.p_return_address.value		= basicAddr
			frm.p_return_address2.value		= basicAddr2
			// focus
			frm.p_return_address2.focus();

		<% elseif strMode="J" then %>
			// copy
			frm.returnZipcode.value		= $("#tzip").val();
			frm.returnZipaddr.value		= basicAddr
			frm.returnEtcaddr.value		= basicAddr2
			// focus
			frm.returnEtcaddr.focus();

		<% end if %>

		// close this window
		window.close();
	}

    // 우편번호 찾기 찾기 화면을 넣을 element
    var element_wrap = $("#searchZipWrap");

    function searchZipKakao() {
        // 현재 scroll 위치를 저장해놓는다.
        var currentScroll = Math.max(document.body.scrollTop, document.documentElement.scrollTop);
		daum.postcode.load(function(){
			new daum.Postcode({
				oncomplete: function(data) {
					var addr = ''; // 주소 변수
					var extraAddr = ''; // 참고항목 변수

					<%'//사용자가 선택한 주소 타입에 따라 해당 주소 값을 가져온다.%>
					if (data.userSelectedType === 'R') { // 사용자가 도로명 주소를 선택했을 경우
						addr = data.roadAddress;
					} else { // 사용자가 지번 주소를 선택했을 경우(J)
						addr = data.jibunAddress;
					}

					<%'// 사용자가 선택한 주소가 도로명 타입일때 참고항목을 조합한다.%>
					if(data.userSelectedType === 'R'){
						<%'// 법정동명이 있을 경우 추가한다. (법정리는 제외)%>
						<%'// 법정동의 경우 마지막 문자가 "동/로/가"로 끝난다.%>
						if(data.bname !== '' && /[동|로|가]$/g.test(data.bname)){
							extraAddr += data.bname;
						}
						<%'// 건물명이 있고, 공동주택일 경우 추가한다.%>
						if(data.buildingName !== '' && data.apartment === 'Y'){
							extraAddr += (extraAddr !== '' ? ', ' + data.buildingName : data.buildingName);
						}
						<%'// 표시할 참고항목이 있을 경우, 괄호까지 추가한 최종 문자열을 만든다.%>
						if(extraAddr !== ''){
							extraAddr = ' (' + extraAddr + ')';
						}
						<%'// 조합된 참고항목을 해당 필드에 넣는다.%>
						$("#extraAddr").val(extraAddr);
					} else {
						$("#extraAddr").val("");
					}

					<%'// 우편번호와 주소 정보를 해당 필드에 넣는다.%>
					$("#tzip").val(data.zonecode);
					$("#taddr1").val(addr);

					<%'// iframe을 넣은 element를 안보이게 한다.%>
					<%'// (autoClose:false 기능을 이용한다면, 아래 코드를 제거해야 화면에서 사라지지 않는다.)%>
					<%'//element_wrap.style.display = 'none';%>

					<%'// 우편번호 찾기 화면이 보이기 이전으로 scroll 위치를 되돌린다.%>
					document.body.scrollTop = currentScroll;
				},
				<%'// 사용자가 주소를 클릭했을때%>
				onclose : function(state) {
					if(state === 'COMPLETE_CLOSE'){
						CopyZipAPI();
					}
				},
				width : '100%',
				height : '89%',
				hideMapBtn : true,
				hideEngBtn : true,
				shorthand : false
			}).embed(element_wrap);
	    });
        <%'// iframe을 넣은 element를 보이게 한다.%>
        element_wrap.style.display = 'block';
    }
</script>
</head>
<body>
<img src="//fiximage.10x10.co.kr/web2019/common/tit_post.jpg" style="width:100%">
<div id="searchZipWrap" style="display:none;border:1px solid;width:500px;height:700px;margin:5px 0;position:relative">
</div>
<form name="tranFrmApi" id="tranFrmApi" method="post">
	<input type="hidden" name="tzip" id="tzip">
	<input type="hidden" name="taddr1" id="taddr1">
	<input type="hidden" name="taddr2" id="taddr2">
    <input type="hidden" name="extraAddr" id="extraAddr">
</form>
<script src="https://ssl.daumcdn.net/dmaps/map_js_init/postcode.v2.js"></script>
</body>
</html>