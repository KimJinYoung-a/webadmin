<%@ codepage="65001" language="VBScript" %>
<% Option Explicit %>
<% response.Charset="UTF-8" %>
<%
Session.CodePage = 65001

Dim pageTitle
pageTitle="2016 The Fingers Artist Admin App - 비밀번호 찾기"
%>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/md5.asp"-->
<!-- #include virtual="/apps/academy/lib/head.asp" -->
<%
Dim searchType

searchType= requestCheckVar(request("rdoAType"),1)
If searchType ="" Then searchType = 2 '기본값:사업자등록번호 검색
	

%>
<script type="text/javascript">
// jquery에서 function호출할때 형식
//onload이벤트라고 보시면 됩니다.
$(document).ready(function() {
	// memberFrm폼에 submit이벤트가 일어날때 반응
	// jquery 해당폼 해당이벤트 이런식으로 함수 작성
	$("form#frmID").bind("submit", function () {
	if ($("input#rdoAType").val().trim() == 2) {
		if ($("input#uid").val() == "") {
			alert("아이디를 입력해주세요.");
			$("input#uid").focus();
			return false;
		}
	  if ($("input#BNo").val().trim() == "") {
	   alert("사업자 등록번호를 입력해주세요.");
	   $("input#BNo").focus();
	   return false;
	  }
	}
	else{
		if ($("input#uid").val() == "") {
			alert("아이디를 입력해주세요.");
			$("input#uid").focus();
			return false;
		}
	  if ($("input#Cnm").val().trim() == "") {
	   alert("대표자명을 입력해주세요.");
	   $("input#Cnm").focus();
	   return false;
	  }
	}
	return true;
	});

});

function number_format(num) {
     num = num.replace(/-/g, "");
     var num_str = num.toString();
     var result = '';
 
      for(var i=0; i<num_str.length; i++) {
            var tmp = num_str.length-(i+1);
            if(i==5){
				result = '-' + result;
			}
			else if(i==7){
				result = '-' + result;
			}
            result = num_str.charAt(tmp) + result;
       }
       return result;
}

</script>
</head>
<body>
<div class="wrap">
	<div class="container">
		<!-- content -->
		<div class="content">
			<div class="pwrSearch">
				<% If searchType=2 Then %>
				<h1 class="tit">사업자 등록번호로 찾기</h1>
				<p class="tMar1r">아이디에 등록된 담당자의 휴대폰 번호 인증을 통해 <br />비밀번호를 안전하게 재설정하실 수 있습니다.</p>
				<% Else %>
				<h1 class="tit">대표자 이름으로 찾기</h1>
				<p class="tMar1r">아이디에 등록된 담당자의 휴대폰 번호 인증을 통해 <br />비밀번호를 안전하게 재설정하실 수 있습니다.</p>
				<% End If %>
				<div class="inputUnit">
					<form name="frmID" id="frmID" method="post" action="/apps/academy/login/certifyPwd.asp">
					<input type="hidden" id="rdoAType" name="rdoAType" value="<% =searchType %>">
						<% If searchType=2 Then %>
						<fieldset>
							<div class="textForm1"><label>아이디 입력</label><input type="text" id="uid" name="uid" placeholder="아이디를 입력해주세요" style="width:100%;" autocomplete="off" onKeyPress="if (event.keyCode == 13) document.frmID.BNo.focus();" /></div>
							<div class="textForm1 tMar0-5r"><label>사업자 등록번호 입력</label><input type="text" id="BNo" name="BNo" placeholder="사업자 등록번호를 입력해주세요" onkeyup="this.value=number_format(this.value)" style="width:100%;" /></div>
							<div class="tMar0-5r grid1"><button class="btnB1 btnGrn">확 인</button></div>
						</fieldset>
						<% Else %>
						<fieldset>
							<div class="textForm1"><label>아이디 입력</label><input type="text" id="uid" name="uid" placeholder="아이디를 입력해주세요" style="width:100%;" autocomplete="off" onKeyPress="if (event.keyCode == 13) document.frmID.Cnm.focus();" /></div>
							<div class="textForm1 tMar0-5r"><label>대표자 이름 입력</label><input type="text" id="Cnm" name="Cnm" placeholder="대표자 이름을 입력해주세요" style="width:100%;" /></div>
							<div class="tMar0-5r grid1"><button class="btnB1 btnGrn">확 인</button></div>
						</fieldset>
						<% End If %>
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
<!-- #include virtual="/lib/db/dbclose.asp" -->