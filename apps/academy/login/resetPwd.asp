<%@ codepage="65001" language="VBScript" %>
<% Option Explicit %>
<% response.Charset="UTF-8" %>
<%
Session.CodePage = 65001
Dim pageTitle
pageTitle="2016 The Fingers Artist Admin App - 비밀번호 설정"
%>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/md5.asp"-->
<!-- #include virtual="/apps/academy/lib/head.asp" -->
<%
dim userid, searchID
dim manager_name, manager_hp ,jungsan_name, jungsan_hp, deliver_name, deliver_hp, manager_shp, jungsan_shp,deliver_shp
dim sql
userid  = session("AuthUID")
 
If userid ="" Or session("AuthChk") <>"Y" Then
	response.write("<script>alert('인증정보가 올바르지 않습니다.확인 후 다시 시도해주세요') ;location.href='/apps/academy/login/searchPwd.asp';</script>") 
	response.End
End If
%>
<script> 

$(document).ready(function() {
	// memberFrm폼에 submit이벤트가 일어날때 반응
	// jquery 해당폼 해당이벤트 이런식으로 함수 작성
	$("form#frmPWD").bind("submit", function(){
		if ($("input#upwd").val().trim() == ""){
		   alert("1차 비밀번호를 입력해주세요.");
		   $("input#upwd").focus();
		   return false;
		}
		if ($("input#upwd").val().length < 8 || $("input#upwd").val().length > 16) {
			alert("1차 비밀번호는 공백 없는 8~16 영문/숫자로만 사용 가능합니다.");
			$("input#upwd").focus();
			return false;
		}
		if ($("input#upwd").val().trim() == $("input#uid").val().trim()) {
		   alert("아이디와 다른 1차 비밀번호를 사용해주세요.");
		   $("input#upwd").focus();
		   return false;
		}
		if (!fnChkComplexPassword($("input#upwd").val().trim())) {
			alert('1차 비밀번호는 영문/숫자/특수문자 중 두 가지 이상의 조합으로 입력해주세요.');
			$("input#upwd").focus();
			return false;
		}
		if ($("input#upwd2").val().trim() == "") {
		   alert("1차 비밀번호 확인을 입력해주세요.");
		   $("input#upwd2").focus();
		   return false;
		}
		if ($("input#upwd").val().trim() != $("input#upwd2").val().trim()) {
		   alert("처음 입력된 비밀번호와 일치하지 않습니다.\n비밀번호를 정확히 입력해주세요.");
		   $("input#upwd2").focus();
		   return false;
		}
		if ($("input#upwdS1").val().trim() == "") {
		   alert("2차 비밀번호를 입력해주세요.");
		   $("input#upwdS1").focus();
		   return false;
		}
		if ($("input#upwdS1").val().length < 8 || $("input#upwdS1").val().length > 16) {
			alert("2차 비밀번호는 공백 없는 8~16 영문/숫자로만 사용 가능합니다.");
			$("input#upwdS1").focus();
			return false;
		}
		if ($("input#upwdS1").val().trim() == $("input#uid").val().trim()) {
		   alert("아이디와 다른 2차 비밀번호를 사용해주세요.");
		   $("input#upwdS1").focus();
		   return false;
		}
		if (!fnChkComplexPassword($("input#upwdS1").val().trim())) {
			alert('2차 비밀번호는 영문/숫자/특수문자 중 두 가지 이상의 조합으로 입력해주세요.');
			$("input#upwdS1").focus();
			return false;
		}
		if ($("input#upwdS2").val().trim() == "") {
		   alert("2차 비밀번호 확인을 입력해주세요.");
		   $("input#upwdS2").focus();
		   return false;
		}
		if ($("input#upwdS1").val().trim() != $("input#upwdS2").val().trim()) {
		   alert("처음 입력된 비밀번호와 일치하지 않습니다.\n비밀번호를 정확히 입력해주세요.");
		   $("input#upwdS2").focus();
		   return false;
		}
	return true;
	});

});


// 패스워드 복잡도 검사
function fnChkComplexPassword(pwd) {
    var aAlpha = /[a-z]|[A-Z]/;
    var aNumber = /[0-9]/;
    var aSpecial = /[!|@|#|$|%|^|&|*|(|)|-|_]/;
    var sRst = true;

    if(pwd.length < 8){
        sRst=false;
        return sRst;
    }

    var numAlpha = 0;
    var numNums = 0;
    var numSpecials = 0;
    for(var i=0; i<pwd.length; i++){
        if(aAlpha.test(pwd.substr(i,1)))
            numAlpha++;
        else if(aNumber.test(pwd.substr(i,1)))
            numNums++;
        else if(aSpecial.test(pwd.substr(i,1)))
            numSpecials++;
    }

    if((numAlpha>0&&numNums>0)||(numAlpha>0&&numSpecials>0)||(numNums>0&&numSpecials>0)) {
    	sRst=true;
    } else {
    	sRst=false;
    }
    return sRst;
}





</script>
</head>
<body>
<div class="wrap">
	<div class="container">
		<!-- content -->
		<div class="content">
			<div class="pwrReset">
				<h1 class="tit">인증완료! 비밀번호를 재설정해주세요.</h1>
				<p class="tMar1r">공백 없는 8~16자 영문/숫자로 입력해주세요.</p>

				<div class="idView">
					<span><%=userid%></span>
				</div>

				<div class="inputUnit">
					<form id="frmPWD" name="frmPWD" method="post" action="/apps/academy/login/searchPwdProc.asp">
					<input type="hidden" name="hidM" value="C">
					<input type="hidden" id="uid" name="uid" value="<%=userid%>">
					<fieldset>
						<div>
							<div class="textForm2"><label>첫번째 비밀번호</label><input type="password" id="upwd" name="upwd" placeholder="1차 비밀번호" style="width:100%;" AUTOCOMPLETE="off" onKeyPress="if (event.which == 13) $('input#upwd2').focus();" /></div>
							<div class="textForm2"><label>첫번째 비밀번호 확인</label><input type="password" id="upwd2" name="upwd2" placeholder="1차 비밀번호 확인" style="width:100%;" AUTOCOMPLETE="off" onKeyPress="if (event.which == 13) $('input#upwdS1').focus();" /></div>
						</div>
						<div class="tPad2r">
							<div class="textForm2"><label>두번째 비밀번호</label><input type="password" id="upwdS1" name="upwdS1" placeholder="2차 비밀번호" style="width:100%;" AUTOCOMPLETE="off" onKeyPress="if (event.which == 13) $('input#upwdS2').focus();" /></div>
							<div class="textForm2"><label>두번째 비밀번호 확인</label><input type="password" id="upwdS2" name="upwdS2" placeholder="2차 비밀번호 확인" style="width:100%;" AUTOCOMPLETE="off" onKeyPress="if (event.which == 13) $('#submit').click();" /></div>
						</div>
					</fieldset>
					<div class="grid1 tMar1-5r"><button class="btnB1 btnGrn">확 인</button></div>
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