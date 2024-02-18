<%@ language="vbscript" %>
<% option explicit %>
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/NoUSBAllowIpList.asp"-->
<!-- #include virtual="/lib/util/tenEncUtil.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/classes/admin/scmBGCls.asp" -->
<%
'/서버 주기적 업데이트 위한 공사중 처리 '2011.11.11 한용민 생성
'/리뉴얼시 이전해 주시고 지우지 말아 주세요.
Call serverupdate_underconstruction()

dim UserOsInfo
dim vSavedID,vSavedEno
dim sBGImg, ClsscmBG
dim lgnMethod
lgnMethod = requestCheckVar(trim(request("lgnMethod")),1)
UserOsInfo = Request.ServerVariables("HTTP_USER_AGENT")
vSavedID = tenDec(request.cookies("SCMSave")("SAVED_ID"))
vSavedEno = tenDec(request.cookies("SCMSave")("SAVED_Eno"))


''USB 인증없이 로그인 체크(인증범위내 접속)
Dim NoUsbValidIP
NoUsbValidIP = fncheckAllowIPWithByDB("Y", "", "")

' 방화벽에서 컨트롤 하고 있음. 주석처리요청(유재규) 2021.07.09 한용민
' if Not(NoUsbValidIP) and date>="2021-03-15" then
' 	call alert_move("외부에서의 웹어드민 접속이 제한됩니다.\n리모트뷰를 통해 접속해주세요.","http://www.10x10.co.kr")
' 	dbget.close: Response.End
' end if

if lgnMethod ="" then
	lgnMethod = CHKIIF(NoUsbValidIP,"U","S")
end if


if Application("scmBGdiv") = 0 then
	set ClsscmBG = new CscmBG
	ClsscmBG.fnGetBGUrl
	sBGImg = ClsscmBG.FBGImg
	set ClsscmBG = nothing
	Application.lock
	Application("scmBG") = sBGImg
	Application("scmBGdiv") = 1
	Application.unlock
end if


Function fnExistFile(filePath)
  Dim fso, result
  Set fso = CreateObject("Scripting.FileSystemObject")
  If fso.FileExists(filePath) Then
    result = 1
  Else
    result = 0
  End If
  fnExistFile = result
End Function

'if fnExistFile(Application("scmBG")) = 0 Then Application("scmBG")=""

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="ko" xml:lang="ko">
<head>
<title>10x10 WEBADMIN LOGIN</title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<meta http-equiv="X-UA-Compatible" content="IE=edge" />
<meta name="viewport" content="width=device-width, initial-scale=1.0, minimum-scale=1.0, maximum-scale=1.0, user-scalable=no">
<meta name="Robots" content="noindex,nofollow">
<link rel="stylesheet" type="text/css" href="/css/adminDefault.css" />
<link rel="stylesheet" type="text/css" href="/css/adminCommon.css" />
<link REL="SHORTCUT ICON" href="http://fiximage.10x10.co.kr/icons/10x10SCM.ico">
<link REL="apple-touch-icon" href="/images/iphone_icon_SCM.png"/>
<style>
html {overflow:auto;}
</style>
<!--[if lt IE 9]>
	<script src="/js/respond.min.js"></script>
	<link rel="stylesheet" type="text/css" href="/css/adminIe.css" />
<![endif]-->
<script type="text/javascript" src="/js/jquery-1.7.2.min.js"></script>
<SCRIPT  type="text/javascript">
// 로그인 폼검사/실행
function chkForm() {
	if(document.frmLogin.lgnMethod.value=="S") {

			 if(!document.frmLogin.usid.value) {
					alert('아이디를 입력해주세요.');
					document.frmLogin.usid.focus();
					return  ;
				}

				if(!document.frmLogin.uspwd.value) {
				alert('비밀번호를 입력해주세요.');
				document.frmLogin.uspwd.focus();
				return ;
				}

				if(document.frmLogin.sAuthNo.value.length<6) {
				alert('휴대폰으로 받으신 인증번호를 입력해주세요.');
				document.frmLogin.sAuthNo.focus();
				return  ;
				}

				document.frmLogin.usn.value="";
				document.frmLogin.action="<%=getSCMSSLURL%>/login/dologin.asp";

	} else if(document.frmLogin.lgnMethod.value=="N") {

				if(!document.frmLogin.usn.value) {
					alert('사번을 입력해주세요.');
					document.frmLogin.usn.focus();
					return  ;
				}


			if(!document.frmLogin.unpwd.value) {
				alert('비밀번호를 입력해주세요.');
				document.frmLogin.unpwd.focus();
				return  ;
			}

			document.frmLogin.<%=CHKIIF(lgnMethod="U","uid","usid")%>.value="";
			document.frmLogin.action="<%=getSCMSSLURL%>/login/dologinbyempno.asp";
	} else{
		if(!document.frmLogin.uid.value) {
			alert('아이디를 입력해주세요.');
			document.frmLogin.uid.focus();
			return;
		}

		if(!document.frmLogin.upwd.value) {
			alert('비밀번호를 입력해주세요.');
			document.frmLogin.upwd.focus();
			return ;
		}

		document.frmLogin.usn.value="";
		document.frmLogin.action="<%=getSCMSSLURL%>/login/dologin.asp";
	}

	document.frmLogin.submit();
 }


// SMS로그인 인증번호 발송
function popSMSAuthNo() {
	if(!document.frmLogin.usid.value) {
		alert('아이디를 입력해주세요.');
		document.frmLogin.usid.focus();
		return;
	}

	if(!document.frmLogin.uspwd.value) {
		alert('비밀번호를 입력해주세요.');
		document.frmLogin.uspwd.focus();
		return ;
	}

	hidFrm.location.href="/admin/member/tenbyten/iframe_adminLogin_SendSMS.asp?uid="+document.frmLogin.usid.value;

}

//sms 인증단계
function jsSetStep(iValue){
 if(iValue==2){
 	 document.all.dvid.style.display = 'none';
 	 document.all.dvAuth.style.display = ''
 }else{
 	document.all.dvid.style.display = '';
 	 document.all.dvAuth.style.display = 'none'
 }
}

// SMS입력 카운터 작동(3분간:180초)
var iSecond=180;
var timerchecker = null;

function startLimitCounter(cflg) {
	if(cflg=="new") {
		if(timerchecker != null) {
			alert("이미 인증번호를 발송하였습니다.\n휴대폰의 SMS를 확인해주세요.");
			return ;
		}
		iSecond=180;
	}
    rMinute = parseInt(iSecond / 60);
    rSecond = iSecond % 60;
    if(rSecond<10) {rSecond="0"+rSecond};

    if(iSecond > 0)
    {
        document.frmLogin.sLimitTime.value = rMinute+":"+rSecond;
        iSecond--;
        timerchecker = setTimeout("startLimitCounter()", 1000); // 1초 간격으로 체크
    }
    else
    {
        clearTimeout(timerchecker);
        document.frmLogin.sLimitTime.value = "0:00";
        timerchecker = null;
        alert("인증번호 입력 시간이 종료되었습니다.\n\nSMS를 받지 못했다면 다시 번호를 받아주세요.");
    }
}

// 휴대폰번호 변경/본인확인 팝업
function PopChgHPNum() {
	alert("IP 로그인 후 나의정보에서 휴대폰 본인확인 후 이용가능합니다.");
	return;
<% if (false) then %>
//	if(confirm("본인확인을 아직 받지 않은 아이디입니다.\n본인 확인을 받으시겠습니까?")) {
//		if(!document.frmLogin.usid.value) {
//			alert('아이디를 입력해주세요.');
//			document.frmLogin.usid.focus();
//			return;
//		} else {
//			var popwin = window.open("pop_ChangeHPIdentify.asp?uid="+document.frmLogin.usid.value,"PopChgHPNum","width=400 height=270 scrollbars=yes");
//			popwin.focus();
//		}
//	}
<% end if %>
}

// 인증안내 팝업
function popSecLgnInfo(flg) {
	if(flg=="S") {
		var InfoPop = window.open("/login/SMS_Auth_Info.htm","LoginInfoPop","width=690,height=600,scrollbars=yes");
		InfoPop.focus();
	}
}

 $(function(){
	/* tab */
	$(".tabCont").hide();
	$(".tabNav").find("li:first").addClass("current");
	$(".tabContainer").find(".tabCont:first").show();
	$(".tabNav li").click(function(){
		$(this).siblings("li").removeClass("current");
		$(this).addClass("current");
		$(this).closest(".tabNav").nextAll(".tabContainer:first").find(".tabCont").hide();
		var activeTab = $(this).find("a").attr("href");
		$(activeTab).show();
		var tMtd =	$(this).attr("hdmethod");
		if (tMtd=="S"){
			document.frmLogin.lgnMethod.value="S";
			document.frmLogin.usid.focus()
		}else if(tMtd=="N"){
			document.frmLogin.lgnMethod.value="N";
			document.frmLogin.usn.focus()
		}else{
			document.frmLogin.lgnMethod.value="U";
			document.frmLogin.uid.focus()
		}
		return false;
	});

	// input action
	$(".inpForm input").focus(function(){
		$(this).addClass('onInput');
		$(this).siblings("label").hide();
	});
	$(".inpForm input").focusout(function(){
		$(this).removeClass('onInput');
		if($(this).val() == ""){
			$(this).siblings("label").show();
		}
	});

	// family site
	$(".tenFamily dt").click(function(){
		if($(".tenFamily dd").is(":hidden")){
			$(this).parent().children('dd').show();
		}else{
			$(this).parent().children('dd').hide();
		};
	});
	$(".tenFamily dd li").click(function(){
		var evtName = $(this).text();
		$(this).parent().parent().parent().children('dt').text(evtName);
		$(this).parent().parent().hide();
		 document.getElementById("hidL").value = evtName;
	});
	$(".tenFamily dl").mouseleave(function(){
		$(this).children("dd").hide();
	});

	<%
	' 방화벽에서 컨트롤 하고 있음. 주석처리요청(유재규) 2021.07.09 한용민
	'if Not(NoUsbValidIP) and date>="2021-02-11" then
	%>
	<% 'alert("2021년 2월 22일부터 외부에서의 웹어드민 접속이 제한됩니다.\n리모트뷰를 통해 접속해주세요.\n(설치 문의 : 운영유닛)"); %>
	<% 'end if %>
});

function jsGoUrl(){
	var strUrl;
	if( document.getElementById("hidL").value=="ONLINE"){
		strUrl = "http://www.10x10.co.kr/"
	}else if(document.getElementById("hidL").value=="OFFLINE"){
		strUrl = "http://www.10x10.co.kr/offshop/index.asp"

	}else if(document.getElementById("hidL").value=="THE FINGERS"){
		strUrl = "http://www.thefingers.co.kr/"
	}

	var winOp = window.open("about:blank");
	winOp.location.href = strUrl;
}


$(document).ready(function() {
	if (document.frmLogin.<%=CHKIIF(lgnMethod="U","uid","usid")%>.value != "") {
		document.frmLogin.<%=CHKIIF(lgnMethod="U","upwd","uspwd")%>.focus();
	} else {
		document.frmLogin.<%=CHKIIF(lgnMethod="U","uid","usid")%>.focus();
	}

	// set default tab2 on mobile acess
	var filter = "win16|win32|win64|mac|macintel";
	if ( navigator.platform ) {
		if ( filter.indexOf( navigator.platform.toLowerCase() ) < 0 ) {
			document.getElementById("lgnMethod").value = "S";
		}
	}

	if (document.frmLogin.<%=CHKIIF(lgnMethod="U","upwd","uspwd")%>.value!=""){
		$(".inpForm input").siblings("label").hide();
	}
	var iniactiveTab ;
	if(document.getElementById("lgnMethod").value=="N"){
		$(".tab2").siblings("li").removeClass("current");
		$(".tab2").addClass("current");
		$(".tab2").closest(".tabNav").nextAll(".tabContainer:first").find(".tabCont").hide();
		iniactiveTab = "#tab2";
		$(iniactiveTab).show();
		document.frmLogin.usn.focus();
	} else if(document.getElementById("lgnMethod").value=="S"){
		document.frmLogin.usid.focus();
	}
});

</SCRIPT>
</head>
<body class="scmLogin <%if Application("scmBG") = "" then%>noImage<%end if%>" <%if Application("scmBG") <> "" then%>style="background-image:url(<%=Application("scmBG")%>);"<%end if%>>
	<h1><span></span>
		<% if (application("Svr_Info")="Dev") then %>[Dev <%=request.ServerVariables("REMOTE_ADDR")%> | <%=G_IsLocalDev%>]<% end if %>
		<% if (application("Svr_Info")="Staging") then %>[Staging]<% end if %>
		10X10 WEBADMIN LOGIN</h1>
	<form name="frmLogin" method="post" action="<%=getSCMSSLURL%>/login/dologin.asp"  >
    <input type="hidden" name="backpath" value="<%= request("backpath") %>">
    <input type="hidden" name="tokenSn" value="">
    <input type="hidden" name="lgnMethod" id="lgnMethod" value="<%=lgnMethod%>">
	<div class="loginBox">
		<div class="tabNav">
			<ul>
				<li class="tab1" hdmethod="<%=CHKIIF(lgnMethod="U","U","S")%>"><a href="#tab1"><%=CHKIIF(lgnMethod="U","IP","SMS")%></a></li>
				<li class="tab2" hdmethod="N"><a href="#tab2">사번</a></li>
			</ul>
		</div>
		<div class="tabContainer">
			<div id="tab1" class="tabCont">
			<% if lgnMethod="U" then %>
				<!-- IP -->
				<p class="inpForm">
					<!--<label for="memId1">아이디</label> -->
					<input type="text" id="uid" name="uid" placeholder="아이디"  value="<%=vSavedID%>"  AUTOCOMPLETE="off" onKeyPress="if (event.keyCode == 13) frmLogin.upwd.focus();"/>
				</p>
				<p class="inpForm">
					<!--<label for="memPw1">비밀번호</label>-->
					<input type="password" id="upwd" name="upwd" placeholder="비밀번호"  AUTOCOMPLETE="off" onKeyPress="if (event.keyCode == 13) chkForm();"/>
				</p>
				<p class="rt">
					<input type="checkbox" id="saved_id" name="saved_id" value="o" <%=chkIIF(vSavedID<>"","checked","")%>/>
					<label for="saveId1">아이디저장</label>
				</p>
				<div class="btnArea"><button class="btn" type="button" onClick="chkForm()" >로그인</button></div>
			<% Else %>
				<!-- SMS -->
				 <div>
					<p class="inpForm">
						<label for="memId1">아이디</label>
						<input type="text" id="usid" name="usid"  value="<%=vSavedID%>"  AUTOCOMPLETE="off" onKeyPress="if (event.keyCode == 13) frmLogin.uspwd.focus();"/>
					</p>
					<p class="inpForm">
						<label for="memPw1">비밀번호</label>
						<input type="password" id="uspwd" name="uspwd" AUTOCOMPLETE="off" onKeyPress="if (event.keyCode == 13) popSMSAuthNo();"/>
					</p>
					<p class="rt" >
						<input type="checkbox" id="saved_sid" name="saveds_sid" value="o" <%=chkIIF(vSavedID<>"","checked","")%>/>
						<label for="saveId1">아이디저장</label>
					</p>
					<div class="btnArea" id="dvid"><button class="btn" type="button" onClick="popSMSAuthNo()" >인증번호 받기</button> </div>
				</div>
			<% End if %>
				 <!-- 인증번호 입력 -->
				<div  id="dvAuth" style="display:none;">
					<p class="timeLimit">입력유효시간 <strong><input type="text" name="sLimitTime" id="sLimitTime" value="-:--" readonly  style="width:100px;display:inline-block; margin-top:-4px; padding-left:0.5rem; font-size:2.5rem; font-family:arial; vertical-align:middle;border:0;"></strong></p>
					<!--<p><button class="btn btnReapply" type="button" onClick="document.frmLogin.lgnStep.value=1;chkForm();">인증번호 재발송</button></p> -->
					<p class="inpForm tMar20">
						<label for="smsNum">SMS 인증번호 입력</label>
						<input type="text" id="sAuthNo" name="sAuthNo"   value="" AUTOCOMPLETE="off"/>
					</p>
					<div class="btnArea" style="margin-top:0;"><button class="btn" type="button" onClick="chkForm()" >로그인</button></div>
				</div>
			</div>
			<!-- 사번 -->
			<div id="tab2" class="tabCont">
				<p class="inpForm">
					<label for="staffNum">사번</label>
					<input type="text" id="usn" name="usn" value="<%=vSavedEno%>" onKeyPress="if (event.keyCode == 13) frmLogin.unpwd.focus();"/>
				</p>
				<p class="inpForm">
					<label for="memPw3">비밀번호</label>
					<input type="password" id="unpwd" name="unpwd" AUTOCOMPLETE="off" onKeyPress="if (event.keyCode == 13) chkForm();" />
				</p>
				<p class="rt">
					<input type="checkbox" id="saved_eno" name="saved_eno"  value="o" <%=chkIIF(vSavedEno<>"","checked","")%>/>
					<label for="saveId3">사번저장</label>
				</p>
				<div class="btnArea"><button class="btn" type="button" onClick="chkForm()">로그인</button></div>
			</div>
		</div>
	</div>
	</form>
	<ul class="help">
		<li><a href="javascript:popSecLgnInfo('S')">SMS 인증안내</a></li>
	</ul>
	<p class="slogan">YOU ARE ALREADY DIFFERENT <a href="http://www.10x10.co.kr/" target="_blank">10X10.CO.KR</a></p>
	<div class="tenFamily">
		<input type="hidden" name="hidL" id="hidL" value="">
		<dl>
			<dt>서비스 바로가기</dt>
			<dd>
				<ul>
					<li value="http://www.10x10.co.kr/">ONLINE</li>
					<li>OFFLINE</li>
					<li>THE FINGERS</li>
				</ul>
			</dd>
		</dl>
		<button class="btnGo" type="button" onClick="jsGoUrl();">이동</button>
	</div>
	<iframe id="hidFrm" name="hidFrm" src="about:blank" frameborder="0" width="0" height="0"></iframe>
</body>
</html>
