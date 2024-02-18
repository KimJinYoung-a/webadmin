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

dim userid, searchID
dim manager_name, manager_hp ,jungsan_name, jungsan_hp, deliver_name, deliver_hp, manager_shp, jungsan_shp,deliver_shp
dim sql,reFAddr
dim   recentqcount
dim searchType
dim company_no, ceoname

userid  = requestCheckVar(trim(request("uid")),32)
reFAddr = request.ServerVariables("REMOTE_ADDR")
searchType= requestCheckVar(request("rdoAType"),1)
company_no = requestCheckVar(trim(request("BNo")),12)
ceoname = requestCheckVar(trim(request("Cnm")),32)
	
If userid <> "" Then
	'초기화
		manager_name= ""
	 	manager_hp	=""
	 	jungsan_name=""
	 	jungsan_hp 	=""
	 	deliver_name=""
	 	deliver_hp 	=""
	
	'아이디조회 로그 등록
	sql = "exec db_partner.dbo.sp_Ten_partner_searchPWD_log '"&userid&"','"&Left(reFAddr,16)&"'"
  dbget.Execute sql
	 	 
	'10분 동안 10회 이상 검색시 차단
	recentqcount = 0 	 
	sql = "select count(idx) as cnt "
	sql = sql & " from db_partner.dbo.tbl_partner_searchPWD_log  "
	sql = sql & " where refip='" + Left(reFAddr,16) + "' "
	sql = sql & " and datediff(n,regdate,getdate())<=10" 
	rsget.Open sql, dbget, 1
	if not rsget.eof then
		recentqcount = rsget("cnt")
	end if
	rsget.close

	if recentqcount>=10 then
		response.write "<script type='text/javascript'>alert('단시간 내에 연속으로 여러번 접속하였습니다.\n잠시 후 다시 시도해주세요.');history.back();</script>"
	    dbget.close() : Response.end
	else

	sql =" select id, manager_name, manager_hp ,jungsan_name, jungsan_hp, deliver_name, deliver_hp from db_partner.dbo.tbl_partner where id ='"&userid&"'"
	if searchType ="1" then
		sql = sql & " and left(replace(ceoname,' ',''),3) =left('"&ceoname&"',3) "
	else
		sql = sql & " and company_no ='"&company_no&"' "
	end if
	 
	rsget.Open sql,dbget,1
  if  not rsget.EOF  then
  	  searchID			=rsget("id")
   		manager_name 	=rsget("manager_name")
   		if manager_name <> "" then manager_name= left(manager_name,1)
   		manager_hp =rsget("manager_hp") 
   		if manager_hp <> "" then manager_shp= left(manager_hp,4)&"****"&right(manager_hp,5)
   		jungsan_name 	=rsget("jungsan_name")
   		if jungsan_name <> "" then jungsan_name=left(jungsan_name,1)
   		jungsan_hp =rsget("jungsan_hp")
   		if jungsan_hp <> "" then jungsan_shp= left(jungsan_hp,4)&"****"&right(jungsan_hp,5)
   		deliver_name 	=rsget("deliver_name")
   		if deliver_name <> "" then deliver_name=left(deliver_name,1)
   		deliver_hp =rsget("deliver_hp") 
   		if deliver_hp <> "" then deliver_shp= left(deliver_hp,4)&"****"&right(deliver_hp,5)  
  Else
		if searchType ="1" then
		response.write "<script type='text/javascript'>alert('등록된 정보와 일치하지 않습니다.\n대표자명, 아이디를 다시 확인해주세요.');history.back();</script>"
		Else
		response.write "<script type='text/javascript'>alert('등록된 정보와 일치하지 않습니다.\n사업자 등록번호, 아이디를 다시 확인해주세요.');history.back();</script>"
		End If
		dbget.close() : Response.end
  end if
  rsget.close
	end if
end if
%>
<script>
// SMS입력 카운터 작동(2분30초:150초)
var iSecond=150;
var timerchecker = null;

function startLimitCounter(cflg) {  
	
	if(cflg=="new") {
//		if(timerchecker != null) {
//			alert("이미 인증번호를 발송하였습니다.\n휴대폰의 SMS를 확인해주세요.");
//			return;
//		}
		iSecond=150;	 
	} 
	 
    rMinute = parseInt(iSecond / 60);
    rSecond = iSecond % 60;
    if(rSecond<10) {rSecond="0"+rSecond};

    if(iSecond > 0)
    {
        document.frmAuth.sLimitTime.value  = rMinute+":"+rSecond;
		document.getElementById('timer').innerHTML = rMinute+":"+rSecond;

        iSecond--;
        timerchecker = setTimeout("startLimitCounter()", 1000); // 1초 간격으로 체크
    }
    else
    {
        clearTimeout(timerchecker);
        document.frmAuth.sLimitTime.value = "0:00";
        timerchecker = null;
        alert("인증번호 입력 시간이 종료되었습니다.\n\nSMS를 받지 못했다면 다시 번호를 받아주세요.");
    }
}


$(document).ready(function() {
	// memberFrm폼에 submit이벤트가 일어날때 반응
	// jquery 해당폼 해당이벤트 이런식으로 함수 작성
	$("form#frmAuth").bind("submit", function () {
	if ($("input#sAuthNo").val().length < 6) {
		alert("휴대폰으로 받으신 인증번호를 입력해주세요.");
		$("input#sAuthNo").focus();
		return false;
	}
	if ($("input#sLimitTime").val() == "0:01") {
		alert("인증번호 입력 시간이 종료되었습니다.\n\nSMS를 받지 못했다면 다시 번호를 받아주세요.");
		return false;
	}
	return true;
	});

});

function jsSMSSend(shp){  
	document.frmSMS.sHp.value = shp;
	document.frmSMS.submit();
}

</script>
</head>
<body class="bgWht">
<div class="wrap">
	<div class="container">
		<!-- content -->
		<div class="content">
			<div class="pwrSearch">
				<h1 class="hidden">휴대폰 번호 인증</h1>
				<p class="tit">담당자 정보 확인 후 <br />인증번호받기 버튼을 선택해 주세요.</p>

				<form name="frmSMS" id="frmSMS" method="post" target="hidFrm" action="/apps/academy/login/searchPwd_sendSMS.asp">
				<input type="hidden" name="uid" value="<%=userid%>">
				<input type="hidden" name="sHp" value="">
				<input type="hidden" name="sKey" value="<%=md5(userid&"TPUSMS")%>">
				<div class="certifyNumList">
					<ul>
						<li>
							<div>
								<b>영업담당자</b>
								<p><strong><%=manager_name%>**</strong> / <%=manager_shp%></p>
							</div>
							<div class="btnCertify"><button class="btnS1 btnWht" onClick="jsSMSSend('M');">인증번호 받기</button></div>
						</li>
						<li>
							<div>
								<b>정산담당자</b>
								<p><strong><%=jungsan_name%>**</strong> / <%=jungsan_shp%></p>
							</div>
							<div class="btnCertify"><button class="btnS1 btnWht" onClick="jsSMSSend('J');">인증번호 받기</button></div>
						</li>
						<li>
							<div>
								<b>배송담당자</b>
								<p><strong><%=deliver_name%>**</strong> / <%=deliver_shp%></p>
							</div>
							<div class="btnCertify"><button class="btnS1 btnWht" onClick="jsSMSSend('D');">인증번호 받기</button></div>
						</li>
					</ul>
				</div>
				</form>
				<div class="certifyNumInput" style="display:none;" id="dvAuth" >
				<form id="frmAuth" name="frmAuth" method="post" target="hidFrm" action="/apps/academy/login/searchPwdProc.asp">
				<input type="hidden" name="hidM" value="A">
				<input type="hidden" name="uid" value="<%=userid%>">
				<input type="hidden" name="sKey" value="<%=md5(userid&"TPUAUTH")%>">
				<input type="hidden" name="sLimitTime" value="A">
					<div class="textForm2"><label>인증번호 입력</label><input type="number" id="sAuthNo" name="sAuthNo" placeholder="인증번호를 입력해주세요" style="width:75%;" /><span class="timer" id="timer">02:59</span></div>
					<div class="btnCertify"><button class="btnB1 btnGrn">확 인</button></div>
				</form>
				</div>
			<iframe id="hidFrm" name="hidFrm" src="about:blank" frameborder="0" width="0" height="0"></iframe>
			</div>
		</div>
		<!--// content -->
		<div id="layerMask" class="layerMask"></div>
	</div>
</div>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->