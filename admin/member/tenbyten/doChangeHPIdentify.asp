<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  본인확인을 사용한 휴대폰번호 변경 처리
' History : 2011.05.30 허진원 생성
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<!-- #include virtual="/common/member/oneself/nice.nuguya.oivs.asp"-->
<%
dim cMember
dim C_dumiKey 
C_dumiKey = session.sessionid 

dim mode, userid, username, juminno, dumiKey, strSql
dim MobileNo0, MobileNo1, MobileNo2, MobileNo3, MobileNo
dim manageUrl
dim NiceId, KeyString, ReturnURL, ConfirmMsg, strProcessType, strSendInfo, strOrderNo, SIKey

'// 변수 할당
mode = requestCheckVar(request.form("mode"),5)
userid = requestCheckVar(request.form("userid"),32)
dumiKey = request.form("dumiKey")
'----------
MobileNo0 = requestCheckVar(Request.form("hpNum0"),3)	' 휴대폰번호0
MobileNo1 = requestCheckVar(Request.form("hpNum1"),3)	' 휴대폰번호1
MobileNo2 = requestCheckVar(Request.form("hpNum2"),4)	' 휴대폰번호2
MobileNo3 = requestCheckVar(Request.form("hpNum3"),4)	' 휴대폰번호3

'// 세션값 확인
if (dumiKey<>C_dumiKey) then 
    Call Alert_close("세션정보가 올바르지 않습니다.")
    response.end
end if

'// 직원 기본정보 접수
Set cMember = new CTenByTenMember
	cMember.Fuserid = userid
	cMember.fnGetScmMyInfo

	username      	= cMember.Fusername
	juminno			= Replace(Trim(cMember.FJuminno),"-","")

Set cMember = Nothing

if username="" or isNull(username) then
    Call Alert_close("직원정보가 존재하지 않습니다.")
    response.end
end if

IF application("Svr_Info")="Dev" THEN
 	manageUrl 	    = "http://testwebadmin.10x10.co.kr"
 ELSE
 	manageUrl 	    = "http://webadmin.10x10.co.kr"
 END IF	

'// 모드별 분기
Select Case mode
	Case "chgHP"
		MobileNo = MobileNo1 + "-" + MobileNo2 + "-" + MobileNo3

		strSql = "Update db_partner.dbo.tbl_user_tenbyten " &_
				" Set usercell='" & MobileNo & "'" &_
				"	, isIdentify='Y' " &_
				" Where userid='" & userid & "'"
		dbget.Execute(strSql)
	%>
		<script language="javascript">
		alert('본인확인 및 입력하신 휴대폰번호로 적용되었습니다.');
		parent.opener.history.go(0);
		parent.close();
		</script>
	<%

	Case "ActH"
		MobileNo = MobileNo0 + "-" + MobileNo1 + "-" + MobileNo2 + "-" + MobileNo3
		'=======================================================================================================
		'=====	▣ 계약시에 발급 받은 데이터 설정 : 계약시에 발급된 회원사 ID 및 KeyString값을 설정하십시오. ▣
		'=======================================================================================================
		NiceId= "Ntenxten1"	' 한국신용정보로 부터 전달 받은 회원사 ID ("Nxxx~")
		KeyString = "r6cA3YS9s8WTktrzgfNSOqQXsKf6GnNNpVEnn4DeDuwzgXhICcDpFhTefoTvFUbsux9EvPsbadplISwb" ' 키스트링(80자리)를 넣어주세요.
	
		'========================================================================================
		'=====	▣ 서비스이용시 필요한 데이터 설정 ▣
		'========================================================================================
		' 응답결과를 받아서 처리할 URL을 설정해주세요.
		ReturnURL = manageUrl & "/admin/member/tenbyten/actChangeHPIdentify.asp" ' 본인인증 결과를 리턴 받을 POPUP URL
		' 휴대폰인증 시 인증번호를 직접 지정하고 싶을 때 설정할 수 있습니다.
		' 특수한 경우에 사용되는 사항이니 설정하지 않고, 사용하시면 자동으로 전송됩니다.
		ConfirmMsg = ""	' 전송할 인증번호 (6자리 숫자로 입력해주세요.)
		'========================================================================================
	
		oivsObject.AthKeyStr = KeyString
	
		strProcessType = "5" '//서비스코드 수정하지 마세요.
		strSendInfo = makeSendInfo(NiceId, juminno, SIKey, ReturnURL, MobileNo, ConfirmMsg) '//인증요청시 필요한 암호화 데이터 수정금지
	
		randomize(time())     
		strOrderNo = Replace(date, "-", "")  & round(rnd*(999999999999-100000000000)+100000000000) '//주문번호.. 매 요청마다 중복되지 않도록 유의
		
		'// 해킹방지를 위해 요청정보를 세션에 저장
		session("niceRsdNo") = juminno
		session("niceOrderNo") = strOrderNo
	%>
	<form name="resFrom" method="post" action="https://secure.nuguya.com/nuguya/NiceCert.do">
		<input type="text" name="SendInfo" value="<%=strSendInfo%>">
		<input type="hidden" name="ProcessType" value="<%=strProcessType%>">
		<input type="hidden" name="OrderNo" value="<%=strOrderNo%>">
		<input type="hidden" name="CertMethod" value="CM">	
		</form>		
		<script language="javascript">		
		<!--		
			var w="433";
			var h="540";
		    var x=window.screenLeft;
		    var y=window.screenTop;
		    var l=x+((document.body.offsetWidth-w)/2);
		    var t=y+((document.body.offsetHeight-h)/2);
	
			var frm = document.resFrom;
			var certWin = window.open("","niceCert","toolbars=0,resizable=0,scrolling=0,width="+w+",height="+h+",statusbar=1,top="+t+",left="+l);
			frm.target = "niceCert";
			frm.submit();
			certWin.focus();
			self.location.href ="about:blank";		
		//-->
		</script>
	<%
End Select
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->