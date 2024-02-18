<%@ codepage="65001" language="VBScript" %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/apps/academy/itemmaster/DIYitemCls.asp"-->
<!-- #include virtual="/lib/util/base64Lib.asp"-->
<%
Session.CodePage = 65001
Response.Charset = "UTF-8"

Dim pageTitle
pageTitle="2016 The Fingers Artist Admin App - 제작 특이사항 입력"
%>
<!-- #include virtual="/apps/academy/lib/head.asp" -->
<script language="jscript" runat="server">
function jsDecodeURIComponent(v) { return decodeURIComponent(v); }
function jsEncodeURIComponent(v) { return encodeURIComponent(v); }
</script>
<%
Dim param
param = Base64Decode(jsDecodeURIComponent(request("param")),"UTF-8")
''param = URLDecode(request("param"))
%>
<script>
jQuery(document).ready(function(){
	//$("#requirecontents").keyup(function(){
		fnAPPShowRightConfirmBtns();
	//});
});
function fnAppCallWinConfirm(){
	if($("#requirecontents").val()==""){
		alert("제작 특이사항을 입력해주세요.");
	}
	else{
		//var requirecontentstxt = encSpecialCharNativeFun($("#requirecontents").val());
		var requirecontentstxt = Base64.encode($("#requirecontents").val());
		fnAPPopenerJsCallClose("fnMakeUnusualSet(\""+requirecontentstxt+"\")");
	}
}
</script>
</head>
<body>
<div class="wrap bgGry">
	<div class="container">
		<!-- content -->
		<form method="post" name="sform" autocomplete="off">
		<div class="content bgGry">
			<h1 class="hidden">제작 특이사항 입력</h1>
			<div class="spcNote">
				<div class="linkInsert">
					<textarea rows="5" name="requirecontents" id="requirecontents" placeholder="특이사항이 있을 경우 입력해주세요"><%=param%></textarea>
				</div>
				<div class="linkNotice">
					<p class="fs1-5r">고객이 알아야할 <br />제작 특이사항을 입력해주세요.</p>
				</div>
			</div>
		</div>
		<!--// content -->
		</form>
		<div id="layerMask" class="layerMask"></div>
	</div>
</div>
</body>
</html>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->