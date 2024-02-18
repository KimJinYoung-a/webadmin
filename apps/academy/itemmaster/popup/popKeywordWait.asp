<%@ codepage="65001" language="VBScript" %>
<% option explicit %>
<%
Session.CodePage = 65001
Response.Charset = "UTF-8"
%>
<%
Dim pageTitle
pageTitle="2016 The Fingers Artist Admin App - 검색 키워드 입력"
%>
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/apps/academy/lib/head.asp" -->
<!-- #include virtual="/apps/academy/itemmaster/DIYitemCls.asp"-->
<!-- #include virtual="/lib/util/base64Lib.asp"-->
<script language="jscript" runat="server">
function jsDecodeURIComponent(v) { return decodeURIComponent(v); }
function jsEncodeURIComponent(v) { return encodeURIComponent(v); }
</script>
<%
Dim param
param = Base64Decode(jsDecodeURIComponent(request("param")),"UTF-8")
''param = URLDecode(decSpecialCharNativeFun(request("param")))
%>
<script>
<!--
jQuery(document).ready(function(){
	//$("#keywords").keyup(function(){
		fnAPPShowRightConfirmBtns();
	//});
});

function fnAppCallWinConfirm(){
	var jsontxt;
	if($("[name='keywords']").val()==""){
		alert('검색 키워드를 입력해주세요.');
	}
	else{
		var keyword = $("#keywords").val();
		keyword = keyword.replace(/\n/g,",");
		keyword = keyword.replace(/,,/g,",");
		keyword = keyword.replace(/,,/g,",");
		
		if(keyword.substring(keyword.length,keyword.length-1)==","){
			keyword = keyword.substring(0,keyword.length-1);
		}
		//jsontxt = encSpecialCharNativeFun(keyword);
		jsontxt = Base64.encode(keyword);
		fnAPPopenerJsCallClose("fnKeyWordSet(\"" + jsontxt + "\")");
	}
}
//-->
</script>
</head>
<body>
<div class="wrap bgGry">
	<div class="container">
		<!-- content -->
		<form method="post" name="sform" autocomplete="off">
		<div class="content bgGry">
			<h1 class="hidden">검색 키워드 입력</h1>
			<div class="keywordInput">
				<div class="linkInsert">
					<textarea rows="7" name="keywords" id="keywords" placeholder="검색 키워드가 여러 개일 경우 쉼표(,)로 구분해 주세요."><%=param%></textarea>
				</div>
				<div class="linkNotice">
					<p>예) 키워드1, 키워드2, 키워드3</p>
					<p class="tMar1-5r"><img src="http://image.thefingers.co.kr/apps/2016/img_keyword_ex.png" alt="화면에 이렇게 보여지게 됩니다." /></p>
				</div>
			</div>
		</div>
		</form>
		<!--// content -->
		<div id="layerMask" class="layerMask"></div>
	</div>
</div>
</body>
</html>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->