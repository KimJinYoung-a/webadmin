<%@ language="vbscript" %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/tenEncUtil.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/login/partner_loginCheck_function.asp"-->
<%
'/서버 주기적 업데이트 위한 공사중 처리 '2011.11.11 한용민 생성
'/리뉴얼시 이전해 주시고 지우지 말아 주세요
Call serverupdate_underconstruction()

dim manageUrl
IF application("Svr_Info")="Dev" THEN
	manageUrl 	 = "http://testwebadmin.10x10.co.kr"
ELSE
	manageUrl 	 = "http://webadmin.10x10.co.kr"
END IF
dim UserOsInfo
dim vSavedID,saved_id
dim vUserid, vPassWd, vChkAuth, vIsSec
  
UserOsInfo = Request.ServerVariables("HTTP_USER_AGENT") 
vUserid = session("tmpUID")
vPassWd = session("tmpUPWD")
vChkAuth = requestCheckVar(trim(request.Form("chkAuth")),1)
vIsSec  = requestCheckVar(trim(request.Form("hidSec")),1)
saved_id= requestCheckVar(trim(request.Form("saved_id")),1) 
 
if vChkAuth <> "Y" or vUserid="" or vPassWd="" then
	response.write("<script>window.alert('아이디/1차 비밀번호 확인 후 2차 인증이 가능합니다.');</script>")
  response.write("<script>self.location='"&manageUrl&"'</script>")
end if

if inStr(UserOsInfo,"Windows CE")>0 then
	response.redirect manageUrl&"/PDAadmin/indexPDA.asp"
	dbget.close()	:	response.End
end if
%> 
<!-- #include virtual="/partner/lib/adminHead_NoJs.asp" -->
<link REL="SHORTCUT ICON" href="http://fiximage.10x10.co.kr/icons/10x10SCM.ico">
<script type="text/javascript"> 
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

	function validate(){  
		var frm = document.frmLogin;		
		<%if vIsSec ="N" then '2차 비번 미설정시 설정가능하도록%>
			
			if(!frm.upwdS1.value){
				 alert("비밀번호가 입력되지 않았습니다.");
				 frm.upwdS1.focus();
				 return;
			}  
			
			if (frm.upwdS1.value.length < 8 || frm.upwdS1.value.length > 16){
			alert("비밀번호는 공백없이 8~16자입니다.");
			frm.upwdS1.focus();
			return ;
		 }
	
	
			if (!fnChkComplexPassword(frm.upwdS1.value)) {
				alert('패스워드는 영문/숫자/특수문자 중 두 가지 이상의 조합으로 입력해주세요.');
				frm.upwdS1.focus();
				return;
			}
	
		 	if(!frm.upwdS2.value){
					 alert("비밀번호를 확인해주세요");
					 frm.upwdS2.focus();
					 return;
				}  
				
			if (frm.upwdS1.value!=frm.upwdS2.value){
				alert("비밀번호가 일치하지 않습니다.");
				frm.upwdS1.focus();
				return ;
			} 
	
	<%else%>
	
		if(!document.frmLogin.upwdS.value){
			 alert("비밀번호가 입력되지 않았습니다.");
			 document.frmLogin.upwdS.focus();
			 return;
		}  
		
	<%end if%> 
	 
		document.frmLogin.submit();
	}
 
 
 function jsSearchPWD(){
 	location.href = "/login/searchPwd.asp";
}

$(function(){
	var contH = $('.loginBoxV16').outerHeight();	
	$('.loginBoxV16').css('margin-top',-contH/2+70+'px');
});

</script>
 
</head> 
<body <%if vIsSec <>"N" then%>onLoad="document.frmLogin.upwdS.focus()"<%end if%>>  
<div   id="login">
	<div class="container scrl">
		<% if (application("Svr_Info")="Dev") then %>
		<h1>This is 2009  Test Server...</h1> 
		<% end if %> 
		<div class="loginBoxV16">
			<h1><img src="/images/partner/admin_login_logo_2016.png" alt="Partner Login - 디자인감성채널 텐바이텐의 협력사 페이지입니다." /></h1>
			<div class="loginCont">
				<form method="post" name="frmLogin" action="<%=getSCMSSLURL%>/login/dologinByPartner.asp">
    			<input type="hidden" name="backpath" value="<%= request("backpath") %>">
    			<input type="hidden" name="loginNo" value="2">
    			<input type="hidden" name="hidSec" value="<%=vIsSec%>"> 
				<div class="loginInput">
					<fieldset>
						<p class="inputArea"><label for="id">아이디</label><input type="text" id="id"  class="formTxt" value="<%=vUserid%>" disabled="disabled" maxlength="32" /></p>
						<p class="inputArea tPad10"><label for="pwr">1차 비밀번호</label><input type="password" id="pwr"  class="formTxt" value="********" disabled="disabled" /></p>
						<%if vIsSec ="N" then
							dim islongtimeNotUsingID
							islongtimeNotUsingID = IsLongTimeNotLoginUserid(vUserid)
							
							    if (islongtimeNotUsingID) then
							%>
							<div class="cautionMsg">
								<p class="cRd3">장기간 로그인 정보가 없습니다.  </p>
								<p class="cRd3"> 2차 비밀번호 인증 정보를 위해<br/> 고객센터로 연락주세요</p>
								 <div class="cBl3"><br/>고객센터: 070-4868-1799</div></div>
							<%			
							    else
							%>
						
										<div class="cautionMsg">보안을 위해 2단계 인증을 시행합니다.<br />로그인에 사용하실 <span class="cRd3">2차 비밀번호를 설정</span>해주세요.</div>
										<p class="inputArea tPad10"><label for="pwr2">2차 비밀번호</label>
											<input type="password" id="upwdS1" name="upwdS1" class="formTxt" placeholder="2차 비밀번호" maxlength="32"  AUTOCOMPLETE="off" onKeyPress="if (event.keyCode == 13) frmLogin.upwdS2.focus();"/></p>  
									 	<p class="inputArea tPad10"><label for="pwr2">2차 비밀번호 확인</label>
									 	<input type="password" id="upwdS2" name="upwdS2" class="formTxt" placeholder="2차 비밀번호 확인" maxlength="32"  AUTOCOMPLETE="off" onKeyPress="if (event.keyCode == 13) validate();"/></p>
									 	<div class="alertMsg">공백없는 8~16자의 영문/숫자 조합<br /> 대소문자구분</div>
									 	<!-- for dev msg : 로그인정보 잘못입력의 경우 노출 //-->
										<div id="divMsg" class="alertMsg" style="display:none;">아이디나 비밀번호가 올바르지 않습니다.</div>
										<button type="button" class="loginBtnV16" onClick="validate();">Login</button>
										<p class="tPad10 fs11 cGy3"><input type="checkbox" id="saved_id" class="formCheck" name="saved_id" value="o" <%=chkIIF(saved_id="o","checked","")%>/> <label for="idSave">아이디저장</label></p>
										<span class="helpTxt" onclick="jsSearchPWD();">1 / 2차 비밀번호 찾기</span>
					 	    <%	end if %>
						<%else%>
						
							<div class="cautionMsg">2차 비밀번호를 입력해주세요</div>
						<p class="inputArea tPad10"><label for="pwr2">2차 비밀번호</label><input type="password" id="upwdS" name="upwdS" class="formTxt"    AUTOCOMPLETE="off" onKeyPress="if (event.keyCode == 13) validate();"/></p>
						<!-- for dev msg : 로그인정보 잘못입력의 경우 노출 //-->
						<div id="divMsg" class="alertMsg" style="display:none;">아이디나 비밀번호가 올바르지 않습니다.</div>
						<button type="button" class="loginBtnV16" onClick="validate();">Login</button>
						<p class="tPad10 fs11 cGy3"><input type="checkbox" id="saved_id" class="formCheck" name="saved_id" value="o" <%=chkIIF(saved_id="o","checked","")%>/> <label for="idSave">아이디저장</label></p>
						<span class="helpTxt" onclick="jsSearchPWD();">1 / 2차 비밀번호 찾기</span>
						<%end if%> 						
					</fieldset>
				</div>
				</form>
				<div class="helpTxtBox" style="display:;">
					<dl>
						<dt>2단계 인증은 왜 하나요?</dt>
						<dd>로그인 보안강화를 위해 2단계 인증이 시행됩니다.<br />기존의 아이디와 비밀번호 외에 2차 비밀번호를 입력하는 이중보안 서비스입니다.</dd>
					</dl>
				</div>
			</div>
			<div class="linkWrapV16">
				<ul class="goLink">
					<li class="link01"><a href="http://company.10x10.co.kr/inquiry_write.asp" target="_blank">신규입점</a></li>
					<li class="link02"><a href="http://www.10x10.co.kr" target="_blank">온라인샵</a></li>
					<li class="link03"><a href="http://www.10x10.co.kr/offshop/index.asp" target="_blank">오프라인샵</a></li>
					<li class="link04"><a href="http://company.10x10.co.kr/company_04.htm" target="_blank">오시는길</a></li>
				</ul>
			</div>
			<div class="copy">COPYRIGHT&copy; 10x10.co.kr ALL RIGHTS RESERVED.</div>
		</div>
	</div>
</div>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->

 
	  
 
 

 
  
              	  
         