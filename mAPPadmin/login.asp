<%@ codepage="65001" language="VBScript" %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control", "no-cache"
Response.AddHeader "Expires", "0"
Response.AddHeader "Pragma", "no-cache"
%>
<!-- #include virtual="/mAppadmin/inc/incCommon.asp" -->
<!-- #include virtual="/mAppadmin/inc/incHeader.asp" -->
<%
dim cflag, backpath
cflag      = request("cflag")
backpath   = request("backpath")
%>
<script type="text/javascript">
$(function() {
	$("#btn-login").bind("click", function(event, ui) {
		var frm = $("#loginFrm");
		var uid = $("#uid");
		var upwd = $("#upwd");

		if ( !uid.val().length ) {
			alert("아이디를 입력하세요");
			uid.focus();
			return;
		}

		if ( !upwd.val().length ) {
			alert("패스워드를 입력하세요");
			$("#btn-login").removeClass($.mobile.activeBtnClass);
			upwd.focus();
			return;
		}
		$(frm).find("[name='devicekey']").val(getDeviceKey());
		frm.submit();
	});

	$("#btn-goto-dev-server").bind("click", function(event, ui) {
		document.location.href="http://testm.10x10.co.kr/mAppadmin/login.asp";
	});

	// SMS로그인 인증번호 발송
	$("#btn-get-auth-no").click(function() {
		var uid = $("#uid");

		if ( !uid.val().length ) {
			alert("아이디를 입력해주세요.");
			uid.focus();
			return;
		}

		hidFrm.location.href = "iframe_adminLogin_SendSMS.asp?lstp=C&uid=" + uid.val();
	});

	$("#btn-send-push").bind("click", function(event, ui) {
		var frm = document.pushFrm;

		if (confirm("푸시를 전송하시겠습니까?") == true) {
			frm.receiverId.value = "*";
			frm.submit();
		}
	});

	window.getDeviceKey = function(){
		return pushNotification.regid;
	}


	// SMS입력 카운터 작동(3분간:180초)
	var iSecond = 180;
	var timerchecker = null;
	window.startLimitCounter = function(cflg) {
		var frm = $('#loginFrm');
		var sLimitTime = frm.find('[name="sLimitTime"]');

		if ( cflg == "new" ) {
			if ( timerchecker != null ) {
				alert("이미 인증번호를 발송하였습니다.\n휴대폰의 SMS를 확인해주세요.");
				return;
			}
			iSecond = 180;
		}
		rMinute = parseInt(iSecond / 60);
		rSecond = iSecond % 60;
		if ( rSecond < 10 ) { rSecond = "0" + rSecond };
		if ( iSecond > 0 )
		{
			sLimitTime.val(rMinute + ":" + rSecond);
			iSecond--;
			timerchecker = setTimeout("startLimitCounter()", 1000); // 1초 간격으로 체크
		}
		else
		{
			clearTimeout(timerchecker);
			sLimitTime.val("0:00");
			timerchecker = null;
			alert("인증번호 입력 시간이 종료되었습니다.\n\nSMS를 받지 못했다면 다시 번호를 받아주세요.");
		}
	}

	window.PopChgHPNum = function() {
		alert('본인확인을 아직 받지 않은 아이디입니다.\nSCM에서 본인확인후 이용가능합니다.');
	}
});
</script>
</head>
<body>

<h1>텐바이텐 어드민</h1>

<form id="loginFrm" name="loginFrm" method="post" action="/mAppadmin/login/doMobileAppLogin.asp" data-ajax="false">
<input type="hidden" name="devicekey">
<input type="hidden" name="appid" value="<%= getAppIDByUSERAGENT %>">
<input type="hidden" name="cflag" value="<%=cflag%>">
<input type="hidden" name="backpath" value="<%=backpath%>">

ID : <input type="text" name="uid" id="uid" data-clear-btn="true" placeholder="ID" value=""><br>

PASS : <input type="password" name="upwd" id="upwd" data-clear-btn="true" placeholder="Password" value=""><br>

<% if ( cflag = "1" ) then %>
	인증번호 : <input type="text" name="sAuthNo" id="sAuthNo" maxlength="6" data-clear-btn="true" placeholder="인증번호" value="">

	<input type="button" value="인증번호 받기" id="btn-get-auth-no">
	입력 유효시간 : <input type=text name="sLimitTime" value="-:--" readolny><br>
<% end if %>

<br><br>
<input type="submit" value="로그인" id="btn-login" data-role="button" rel="external" /> <input type="button" value="unregister" id="btn-unregister" data-role="button" rel="external" /> <input type="button" value="푸시전송(*)" id="btn-send-push" data-role="button" rel="external" />
<br><br>

</form>
<iframe id="hidFrm" name="hidFrm" src="about:blank" frameborder="0" width="200" height="200"></iframe>


<a href="#panel-log" data-inline="true" data-icon="alert" class="ui-btn-right">로그</a>

<a href="#" data-role="button" data-rel="close">닫기</a>

<input type="button" value="새로고침" id="btn-reload" data-role="button" rel="external" />

<input type="button" value="개발서버이동" id="btn-goto-dev-server" data-role="button" rel="external" />

<form name="pushFrm" action="/mAPPadmin/common/doPush.asp" method="post" onSubmit="return false;">
<input type="hidden" name="mode" value="sendOnePush">
<input type="hidden" name="receiverId" value="">
<!--
<input type="hidden" name="msg" value="test d테a스b트c  jjjj">
-->
<input type="hidden" name="msg" value="test <%=now()%>">
</form>

</body>
</html>
