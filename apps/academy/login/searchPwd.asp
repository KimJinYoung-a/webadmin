<%@ codepage="65001" language="VBScript" %>
<% Option Explicit %>
<% response.Charset="UTF-8" %>
<%
Session.CodePage = 65001
Dim pageTitle
pageTitle="2016 The Fingers Artist Admin App - 비밀번호 찾기"
%>
<!-- #include virtual="/apps/academy/lib/head.asp" -->
<script>
function FnGoPassSearch(frm,type){
	frm.rdoAType.value=type;
	frm.action="/apps/academy/login/searchPwd_Step2.asp"
	frm.submt();
}
</script>
</head>
<body>
<div class="wrap">
	<div class="container">
		<!-- content -->
		<div class="content">
			<div class="pwrSearch">
				<h1 class="tit">비밀번호를 잊어버리셨나요?</h1>
				<p class="tMar1r">몇가지의 인증절차를 거친 뒤, 안전하게 비밀번호를 <br />재설정할 수 있도록 도와드리겠습니다.</p>

				<div class="inputUnit">
					<form name="frmID" id="frmID" method="post">
					<input type="hidden" id="rdoAType" name="rdoAType">
						<fieldset>
							<div class="tMar0-5r grid1"><button class="btnB1 btnGrn" onClick="FnGoPassSearch(this.form,1)">대표자 이름으로 찾기</button></div>
							<div class="tMar0-5r grid1"><button class="btnB1 btnGrn" onClick="FnGoPassSearch(this.form,2)">사업자 등록번호로 찾기</button></div>
						</fieldset>
					</form>
				</div>
			</div>
		</div>
		<!--// content -->
		<div id="layerMask" class="layerMask"></div>
	</div>
</div>
</body>
</html>