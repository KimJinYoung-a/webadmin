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
<%
Dim waititemid
waititemid = requestCheckVar(request("waititemid"),10)

Dim oitem, safetyYn, safetyDiv, safetyNum
set oitem = new CWaitItemDetail
oitem.FRectDesignerID = request.cookies("partner")("userid")
if (waititemid<>"") then
oitem.WaitProductDetail(waititemid)
safetyYn=oitem.FsafetyYn
safetyDiv=oitem.FsafetyDiv
safetyNum=oitem.FsafetyNum
End If
If safetyYn="" Then safetyYn="N"
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
	if($("#safetyYn").val()=="Y"){
		if($("#safeCertify option:selected").val()==""){
			alert("안전인증구분을 선택해주세요.");
		}else if($("#safetyNum").val()==""){
			alert("인증번호를 입력해주세요.");
		}
		else{
			document.safe.action="/apps/academy/itemmaster/popup/WaitDIYItemPopupDetailinfoEdit_Process.asp";
			document.safe.target="FrameCKP";
			document.safe.submit();
		}
	}else{
		jsontxt = $("input[name='safetyYn']").val() + "!!!";
		fnAPPopenerJsCallClose("fnSafeInfoSet(\""+ jsontxt + "\")");
	}
}
function fnDetailInfoEnd(){
	var selecttxt;
	var jsontxt;
	if($("#safeCertify option:selected").text()=="안전인증구분을 선택해주세요"){
		selecttxt = "";
	}else{
		selecttxt = $("#safeCertify option:selected").text();
	}
	var safetyNum = Base64.encode($("input[name='safetyNum']").val());
	jsontxt = $("input[name='safetyYn']").val() + "!" + $("#safeCertify option:selected").val() + "!" + safetyNum + "!" + $("#safeCertify option:selected").text();
	fnAPPopenerJsCallClose("fnSafeInfoSet(\""+ jsontxt + "\")");
}

function fnselectReset(){
	$("#safetyYn").val("N");
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
		<input type="hidden" name="waititemid" value="<%=waititemid%>">
		<div class="content bgGry">
			<h1 class="hidden">안전인증 대상</h1>
			<div class="artSafeSet">
				<div class="selectBtn">
					<div class="grid2"><button type="button" class="btnM1 btnGry<% If safetyYn="Y" Then %> selected<% End If %>" name="bsafetyYn" id="bsafetyYn" value="Y" onClick="chgodr('safeCertify',2,'safetyYn','Y');">대상</button></div>
					<div class="grid2"><button type="button" class="btnM1 btnGry<% If safetyYn="N" Then %> selected<% End If %>" name="bsafetyYn" id="bsafetyYn" value="N" onClick="chgodr('safeCertify',1,'safetyYn','N');fnselectReset();">대상 아님</button></div>
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
						<li id="safeCertify2" style="display:<% If safetyDiv="" Then %>none<% End If %>">
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
<iframe name="FrameCKP" src="about:blank" frameborder="0" width="0" height="0"></iframe>
<%
set oitem = nothing
%>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->