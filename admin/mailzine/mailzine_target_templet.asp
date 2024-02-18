<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<%
session.codepage = 65001
response.Charset="UTF-8"
%>
<%
'###########################################################
' Description :  텐바이텐 메일진
' History : 2018.04.27 이상구 생성(메일러 연동 생성 메일러로 발송 내역 전송. 메일 가져오기 생성.)
'			2019.06.24 정태훈 수정(템플릿 기능 신규 추가)
'			2020.05.28 한용민 수정(TMS 메일러 추가)
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib_UTF8.asp"-->
<!-- #include virtual="/lib/function_utf8.asp"-->
<!-- #include virtual="/admin/lib/adminbodyheadUTF8.asp"-->
<!-- #include virtual="/admin/incSessionAdmin_UTF8.asp" -->
<style type="text/css">
box-sizing:border-box;
body {height:100%;}
.contWrap {height:100%; border:1px solid #BABABA; background-color:#FFF; padding: 5px;}
.leftArea {width:49%;height:99%; display:inline-block;}
.rightArea {width:49%;height:99%; border:1px solid #666; display:inline-block;}
</style>
<script src='/js/jquery-1.11.0.min.js'></script>
<script>
$(function(){
	showPage();
	$("#mailSource").on("keyup",function(){
		showPage();
	});
});

function showPage() {
	var cont = $("#mailSource").val();
	$("#previewArea").contents().find("body").empty().append(cont);
}
</script>
<div class="contWrap">
<div class="leftArea">
<textarea name="mailcontents" id="mailSource" class="input" style="width:100%; height:100%;">
<html>
<head>
<title>[텐바이텐]</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
</head>
<body style="background-color:#fff; color:#000;">
	
<table align="center" border="0" cellpadding="0" cellspacing="0" style="max-width:700px; margin-left:auto; margin-right:auto;">
<!-- 상단 영역 -->
<thead>
	<!-- 20220915 헤더 수정 -->
	<tr>
		<td style="padding:24px 4px 16px 4px;"><img style="width:100%; vertical-align:top;" src="http://mailzine.10x10.co.kr/2022/common/tit_main_head_03.png" alt="" /></td>
	</tr>
	<!-- //20220915 헤더 수정 -->
</thead>
<!-- //상단 영역 -->

<!-- 컨텐츠 영역 -->
<tbody>
	<tr>
		<td style="margin-left:auto; margin-right:auto; background-color:#fff; border:solid 1px #eaeaea;">

			<table border="0" cellpadding="0" cellspacing="0" style="width:100%;">
				<tr>
					<td style="margin:0; padding:50px 0 50px; font-size:14px; line-height:22px; font-family:dotum, '돋움', sans-serif; color:#707070; text-align:center;">
						내용입니다.
					</td>
				</tr>
				<tr>
					<td style="margin:0; padding:25px 0; border-top:solid 1px #eaeaea; font-size:12px; line-height:12px; font-family:dotum, '돋움', sans-serif; color:#707070; text-align:center;">끝까지 기분 좋은 쇼핑이 될 수 있도록 최선을 다하겠습니다.</td>
				</tr>
			</table>

		</td>
	</tr>
</tbody>
<!-- //컨텐츠 영역 -->

<!-- 하단 공통 영역 -->
<tfoot>
	<tr>
		<td style="background:#f4f4f4; font-size:18px;">
			<!-- 아이디/비밀번호 찾기 등 시스템 강제로 발송할 경우 사용하는 Footer -->
			<table border="0" cellpadding="0" cellspacing="0" style="width:100%;">
				<tr>
					<td style="margin:0; padding:67px 37px 8px 37px; font-family:'맑은고딕','Malgun Gothic','돋움', dotum, sans-serif; font-size:18px !important; line-height:27px; letter-spacing:-0.3px; color:#838383; text-align:left;">* 본 메일은 발신전용 입니다. <br />&nbsp;&nbsp;문의사항이 있으시면 <a href="http://www.10x10.co.kr/cscenter/" target="_blank" style="margin:0; padding:0; color:#707070; font-weight:bold; text-align:left;">고객상담문의</a>를 이용해 주시기 바랍니다.<br />
					</td> 
				</tr>
				<tr>
					<td style="margin:0; padding:8px 35px; font-family:'맑은고딕','Malgun Gothic','돋움', dotum, sans-serif; font-size:18px !important; line-height:27px; letter-spacing:-0.3px; color:#838383; text-align:left;">(03086) 서울시 종로구 대학로12길 31 자유빌딩 5층 텐바이텐 / 대표이사: 최은희 <br /> 사업자등록번호 : 211-87-00620 / 통신판매업신고 : 제01-1968호 / 개인정보 보호 및 청소년 보호책임자 : 이문재 / 고객행복센터 TEL : <b>1644-6030</b> / E-mail : <a href="mailto:customer@10x10.co.kr" style="color:#838383; text-decoration:none; font-style:bold;"><b>customer@10x10.co.kr</b></a></td>
				</tr>
				<tr>
					<td style="margin:0; padding:8px 35px; font-family:Verdana, sans-serif; font-size:18px; line-height:1.39; letter-spacing:-0.3px; color:#838383; text-align:left;">COPYRIGHTS 10x10. ALL RIGHTS RESERVED.</td>
				</tr>
				<tr>
					<td style="padding:35px 35px 72px 35px; line-height:28px; text-align:center;">
						<a href="http://www.facebook.com/your10x10/" target="_blank"><img src="http://mailzine.10x10.co.kr/2017/ico_facebook.png" alt="텐바이텐 공식 Facebook으로 이동" style="margin:0 25px; border:0;" /></a>
						<a href="http://www.instagram.com/your10x10/" target="_blank"><img src="http://mailzine.10x10.co.kr/2017/ico_instargram.png" alt="텐바이텐 공식 Instargram으로 이동" style="margin:0 25px; border:0;" /></a>
						<a href="https://www.pinterest.com/your10x10/" target="_blank"><img src="http://mailzine.10x10.co.kr/2017/ico_pinterest.png" alt="텐바이텐 공식 Pinterest로 이동" style="margin:0 25px; border:0;" /></a>
						<a href="http://www.10x10shop.com/" target="_blank"><img src="http://mailzine.10x10.co.kr/2017/ico_china.png" alt="텐바이텐 공식 china 사이트로 이동" style="margin:0 25px; border:0;" /></a>
					</td>
				</tr>
			</table>
		</td>
	</tr>
</tfoot>
<!-- //하단 공통 영역 -->
</table>
</body>
</html>
</textarea>
</div>
<div class="rightArea">
	<iframe id="previewArea" frameBorder="0" style="width:100%; height:100%;"></iframe>
</div>
</div>

<%
session.codePage = 949
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
