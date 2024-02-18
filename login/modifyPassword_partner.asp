<%@ language="vbscript" %>
<% option explicit %>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
	'로그인 확인
	''if session("ssBctId")="" or isNull(session("ssBctId")) then
	if session("ssnTmpUIDPartner")="" or isNull(session("ssnTmpUIDPartner")) then
		Call Alert_Return("잘못된 접속입니다.")
		response.End
	end if
%>
 
<!-- #include virtual="/partner/lib/adminHead.asp" -->
 
<script language='JavaScript'>
<!--
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

	function chkForm() {
		var frm = document.frmLogin;
		
		if(!frm.upwd.value) {
			alert("비밀번호를 입력해주세요.");
			frm.upwd.focus();
			return  ;
		}
		
	
		if (frm.upwd.value.length < 8 || frm.upwd.value.length > 16){
			alert("비밀번호는 공백없이 8~16자입니다.");
			frm.upwd.focus();
			return ;
		 }
		 
		 	if(frm.upwd.value==frm.uid.value) {
			alert("아이디와 다른 비밀번호를 사용해주세요.");
			frm.upwd.focus();
			return  ;
		}
		
		 if (!fnChkComplexPassword(frm.upwd.value)) {
				alert('패스워드는 영문/숫자/특수문자 중 두 가지 이상의 조합으로 입력해주세요.');
				frm.upwd.focus();
				return;
			}
		 
		 	if(!frm.upwd2.value) {
			alert("비밀번호 확인을 입력해주세요.");
			frm.upwd2.focus();
			return  ;
		}
		
			if(frm.upwd.value!=frm.upwd2.value) {
			alert("비밀번호 확인이 틀립니다.\n정확한 비밀번호를 입력해주세요.");
			frm.upwd.focus();
			return  ;
			} 
			
		
			if(!frm.upwdS1.value) {
			alert("2차 비밀번호를 입력해주세요.");
			frm.upwdS1.focus();
			return  ;
		}
		
	
		if (frm.upwdS1.value.length < 8 || frm.upwdS1.value.length > 16){
			alert("비밀번호는 공백없이 8~16자입니다.");
			frm.upwdS1.focus();
			return ;
		 }
		 
		 	if(frm.upwdS1.value==frm.uid.value) {
			alert("아이디와 다른 비밀번호를 사용해주세요.");
			frm.upwdS1.focus();
			return  ;
		}
		
			if(frm.upwdS1.value==frm.upwd.value) {
			alert("1차 비밀번호와  다른 비밀번호를 사용해주세요.");
			frm.upwdS1.focus();
			return  ;
		}

		if (!fnChkComplexPassword(frm.upwd.value)) {
			alert('1차 새로운 패스워드는 영문/숫자/특수문자 등 두가지 이상의 조합으로 입력하세요. 최소길이 10자(2조합) , 8자(3조합)');
			frm.upwd.focus();
			return;
		}
		if (!fnChkComplexPassword(frm.upwdS1.value)) {
			alert('2차 새로운 패스워드는 영문/숫자/특수문자 등 두가지 이상의 조합으로 입력하세요. 최소길이 10자(2조합) , 8자(3조합)');
			frm.upwdS1.focus();
			return;
		}

		 	if(!frm.upwdS2.value) {
			alert("비밀번호 확인을 입력해주세요.");
			frm.upwdS2.focus();
			return  ;
		}
		
			if(frm.upwdS1.value!=frm.upwdS2.value) {
			alert("비밀번호 확인이 틀립니다.\n정확한 비밀번호를 입력해주세요.");
			frm.upwdS1.focus();
			return  ;
			}  

		 frm.submit(); 
	}
//-->
</script>
</head>
<body onLoad="document.frmLogin.upwd.focus()">  
<div class="wrap" id="login">
	<div class="container scrl">
		<div class="pwrBoxV16">
			<div class="titWrap">
				<h1>비밀번호 변경</h1>
			</div>
			<div class="pwrContWrap">
				<p> 2008년 12월 15일부터 <span class="cRd3">비밀번호 강화 정책</span>으로 <br />보안에 취약한 패스워드는 변경하셔야 텐바이텐 어드민을 사용하실 수 있습니다. 
			    또한 비밀번호는 최소 3개월에 한번 이상 변경해 주시기 바랍니다.<br><br>
			    강화된 비밀번호 정책은 아래와 같습니다.<br />
						<span class="cBd3"> &nbsp; 1. 최소 8자리 이상의 비밀번호 사용<br />
			    &nbsp; 2. 아이디와 동일하거나 아이디를 포함한 패스워드 금지<br />
			    &nbsp; 3. 같은 문자를 연속으로 3자 이상 금지<br />
			    &nbsp; 4. 새로운 패스워드는 영문/숫자/특수문자 등 두가지 이상의 조합으로 입력하세요. 최소길이 10자(2조합) , 8자(3조합)<br /><br /></span>
			</p>
				<form name="frmLogin" method="post" action="/login/doPasswordModi_partner.asp" target="FrameCKP"  >
    		<input type="hidden" name="backpath" value="<%= request("backpath") %>">				 
						<strong class="fs14 cBk1">ID:<%=session("ssnTmpUIDPartner")%><input type=hidden name=uid value='<%=session("ssnTmpUIDPartner")%>'></strong> 
						<div class="sectionWrap">
							<div class="partitionZone">
								<h2>1차 비밀번호</h2>
								<div class="ftRt" style="width:265px;">
									<p class="inputArea"><label for="id">1차 비밀번호</label><input type="password" id="upwd" name="upwd" class="formTxt" placeholder="1차 비밀번호" style="width:100%;" maxlength="32" AUTOCOMPLETE="off" onKeyPress="if (event.keyCode == 13) frmLogin.upwd2.focus();"/></p>
									<p class="inputArea tPad10"><label for="pwr">1차 비밀번호 확인</label><input type="password" id="upwd2" name="upwd2" class="formTxt" placeholder="1차 비밀번호 확인" style="width:100%;" maxlength="32" AUTOCOMPLETE="off" onKeyPress="if (event.keyCode == 13) frmLogin.upwdS1.focus();"/></p> 
									<p class="tPad10 cRd3">공백 없는 8~16자 영문/숫자로 입력해주세요. <!-- 비밀번호가 일치하지 않습니다. 다시 확인해주세요. --></p>
								</div>
						</div>
						<div class="partitionZone tMar20">
							<h2>2차 비밀번호</h2>
							<div class="ftRt" style="width:265px;">
								<p class="inputArea"><label for="id2">2차 비밀번호</label><input type="password" id="upwdS1" name="upwdS1" class="formTxt" placeholder="2차 비밀번호" style="width:100%;" AUTOCOMPLETE="off" onKeyPress="if (event.keyCode == 13) frmLogin.upwdS2.focus();"/></p>
								<p class="inputArea tPad10"><label for="pwr2">2차 비밀번호 확인</label><input type="password" id="upwdS2" name="upwdS2" class="formTxt" placeholder="2차 비밀번호 확인" style="width:100%;" AUTOCOMPLETE="off" onKeyPress="if (event.keyCode == 13) chkForm();"/></p> 
							</div>
						</div>
					</div>
				<button type="button" class="viewBtnV16 tMar20" style="width:100%;" onClick="chkForm();">비밀번호 저장</button>
			</form>
			<iframe name="FrameCKP" src="" frameborder="0" width="0" height="0"></iframe>
			</div>
			<div class="copy">COPYRIGHT&copy; 10x10.co.kr ALL RIGHTS RESERVED.</div>
		</div>
	</div>
</div>

</body>
</html>


 