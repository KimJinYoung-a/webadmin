<% Response.Addheader "P3P","policyref=""/w3c/p3p.xml"", CP=""CONi NOI DSP LAW NID PHY ONL OUR IND COM"""%>
<%
	'세션 UTF-8 지정
	Session.CodePage = "65001"

	'/서버 주기적 업데이트 위한 공사중 처리 '2011.11.11 한용민 생성
	'/리뉴얼시 이전해 주시고 지우지 말아 주세요
	Call serverupdate_underconstruction()

	'# 현재 페이지명, TITLE
	Dim nowViewPage, strPageTitle, strPageKeyword, strPageUrl, strPageImage, strHeaderAddMetaTag, strPageDesc, strHeadTitle
	nowViewPage = request.ServerVariables("SCRIPT_NAME")

	'// 로그인 유효기간 확인 및 처리
	Select Case lcase(Request.ServerVariables("URL"))
		Case "/_index.asp", "/index.asp"
			Call chk_ValidLogin()
	End Select

	'// 자동로그인 확인
	Call chk_AutoLogin()
	
	
	'// 페이지 검색 키워드
	if strPageKeyword="" then
		strPageKeyword = "더핑거스, 핑거스 아카데미, Fingers Academy, 텐바이텐, 10x10, 만지기, 꿰매기, 꾸미기, 맛보기, 그리기, 즐기기 강좌"
	else
		strPageKeyword = "더핑거스, Fingers Academy, 핑거스 아카데미, " & strPageKeyword
	end if
	if strPageDesc="" then strPageDesc = "특별함을 전하는, 더핑거스"

	if strPageTitle="" then : strPageTitle = "더핑거스"

	if strHeadTitle="" then : strHeadTitle = "더핑거스"

	'// Facebook 오픈그래프 메타태그 작성 (필요에 따라 변경요망)
	if strHeaderAddMetaTag = "" then
		strHeaderAddMetaTag = "<meta property=""og:title"" content=""" & strHeadTitle & """ />" & vbCrLf &_
							"	<meta property=""og:type"" content=""website"" />" & vbCrLf
	end if
	if strPageUrl<>"" then
		strHeaderAddMetaTag = strHeaderAddMetaTag & "	<meta property=""og:url"" content=""" & strPageUrl & """ />" & vbCrLf
	end if
	if Not(strPageImage="" or isNull(strPageImage)) then
		strHeaderAddMetaTag = strHeaderAddMetaTag & "	<meta property=""og:image"" content=""" & strPageImage & """ />" & vbCrLf &_
													"	<link rel=""image_src"" href=""" & strPageImage & """ />" & vbCrLf
	else
		'기본 이미지
		strHeaderAddMetaTag = strHeaderAddMetaTag & "	<meta property=""og:image"" content=""http://m.thefingers.co.kr/lib/inc/theFingersicon.gif"" />" & vbCrLf &_
													"	<link rel=""image_src"" href=""http://m.thefingers.co.kr/lib/inc/theFingersicon.gif"" />" & vbCrLf
	end If


%>
<!doctype html>
<html lang="ko">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, minimum-scale=1.0, maximum-scale=1.0, user-scalable=no">
<meta name="description" content="<%=strPageDesc%>" />
<meta name="keywords" content="<%=strPageKeyword%>" />
<meta name="format-detection" content="telephone=no" />
<%=strHeaderAddMetaTag%>
<link REL="SHORTCUT ICON" href="http://m.thefingers.co.kr/lib/ico/thefingers.ico">
<link REL="apple-touch-icon" href="http://m.thefingers.co.kr/lib/ico/TouchIcon_180x180.png"/>
<link rel="canonical" href="http://www.thefingers.co.kr/">
<title><%=strHeadTitle%></title>
<link rel="stylesheet" type="text/css" href="http://m.thefingers.co.kr/lib/css/common.css" />
<link rel="stylesheet" type="text/css" href="http://m.thefingers.co.kr/lib/css/content.css" />
<link rel="stylesheet" type="text/css" href="http://m.thefingers.co.kr/lib/css/myfingers.css" />
<script type="text/javascript" src="http://m.thefingers.co.kr/lib/js/jquery-2.2.2.min.js"></script>
<script type="text/javascript" src="http://m.thefingers.co.kr/lib/js/jquery.swiper-3.3.1.min.js"></script>
<script type="text/javascript" src="http://m.thefingers.co.kr/lib/js/common.js"></script>
<script type="text/javascript" SRC="http://m.thefingers.co.kr/lib/js/fingerscommon.js"></script>