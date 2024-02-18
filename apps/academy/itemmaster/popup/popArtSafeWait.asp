<%@ codepage="65001" language="VBScript" %>
<% option explicit %>
<%
Session.CodePage = 65001
Response.Charset = "UTF-8"

Dim pageTitle
pageTitle="2016 The Fingers Artist Admin App - 동영상 삽입"
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
param = URLDecode(request("param"))

Dim ArrSafe, safetyYn, safetyDiv, safetyNum
If param <> "" Then
ArrSafe = Split(param,",")
safetyYn = ArrSafe(0)
safetyDiv = ArrSafe(1)
safetyNum = Base64Decode(jsDecodeURIComponent(ArrSafe(2)),"UTF-8")
Else
safetyYn = "N"
safetyDiv = ""
safetyNum = ""
End If
%>
<script type="text/javascript" src="/apps/academy/lib/waititemreg.js"></script>
<script>
jQuery(document).ready(function(){
	<% if safetyDiv="" Then %>
	fnAPPShowRightConfirmBtns();
	<% else %>
	$("#safetyDiv").val('<%=safetyDiv%>').attr("selected","selected");
	fnAPPShowRightConfirmBtns();
	<% end if %>
});

function fnAppCallWinConfirm(){
	var jsontxt;
	if($("#safetyYn").val()=="Y"){
		if($("#safeCertify option:selected").val()==""){
			alert("안전인증구분을 선택해주세요.");
		}else if($("#safetyNum").val()==""){
			alert("인증번호를 입력해주세요.");
		}
		else{
			var safetyNum = Base64.encode($("input[name='safetyNum']").val());
			jsontxt = $("input[name='safetyYn']").val() + "!" + $("#safeCertify option:selected").val() + "!" + safetyNum + "!" + $("#safeCertify option:selected").text();
			fnAPPopenerJsCallClose("fnSafeInfoSet(\""+ jsontxt + "\")");
		}
	}else{
		jsontxt = $("input[name='safetyYn']").val() + "!!!";
		fnAPPopenerJsCallClose("fnSafeInfoSet(\""+ jsontxt + "\")");
	}
}
function fnSafetyDel(){
	$("#safetyDiv").val("");
	$("#safetyNum").val("");
}
</script>
</head>
<body>
<div class="wrap bgGry">
	<div class="container">
		<!-- content -->
		<form name="safe" method="post" onsubmit="return false;" autocomplete="off">
		<input type="hidden" name="safetyYn" id="safetyYn" value="<%=safetyYn%>">
		<div class="content bgGry">
			<h1 class="hidden">안전인증 대상</h1>
			<div class="artSafeSet">
				<div class="selectBtn">
					<div class="grid2"><button type="button" class="btnM1 btnGry<% If safetyYn="Y" Then %> selected<% End If %>" name="bsafetyYn" id="bsafetyYn" value="Y" onClick="chgodr('safeCertify',2,'safetyYn','Y');">대상</button></div>
					<div class="grid2"><button type="button" class="btnM1 btnGry<% If safetyYn="N" Then %> selected<% End If %>" name="bsafetyYn" id="bsafetyYn" value="N" onClick="chgodr('safeCertify',1,'safetyYn','N');fnSafetyDel();">대상 아님</button></div>
				</div>
				<div class="safeCertify" id="safeCertify" style="display:<% If safetyYn="N" Then %>none<% End If %>"><!-- for dev msg : 안전인증 대상 선택시 노출됩니다.-->
					<ul class="list">
						<li class="selectBtn">
							<select name="safetyDiv" id="safetyDiv" class="select" onChange="chgodr('safeCertify2',2,'','');">
								<option value="">안전인증구분을 선택해주세요</option>
								<option value="10">국가통합인증(KC마크)</option>
								<option value="20">전기용품 안전인증</option>
								<option value="30">KPS 안전인증 표시</option>
								<option value="40">KPS 자율안전 확인 표시</option>
								<option value="50">KPS 어린이 보호포장 표시</option>
							</select>
						</li>
						<li id="safeCertify2" style="display:<% If ArrSafe(1)="" Then %>none<% End If %>">
							<dfn><b>인증번호</b></dfn>
							<div><input type="text" name="safetyNum" id="safetyNum" placeholder="인증번호를 입력해주세요" value="<%=safetyNum%>" /></div>
						</li>
					</ul>
					<div class="optionUnit fs1-2r cGy1 rt">※ 유아용품이나 전기용품일 경우 필수 입력</div>
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