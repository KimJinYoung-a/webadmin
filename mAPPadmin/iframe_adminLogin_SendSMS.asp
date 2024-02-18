<%@ codepage="65001" language="VBScript" %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control", "no-cache"
Response.AddHeader "Expires", "0"
Response.AddHeader "Pragma", "no-cache"
Response.CharSet = "UTF-8"
%>
<%
'###########################################################
' Description :  로그인 인증문자 발송
' History : 2011.06.13 허진원 생성
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<!-- #include virtual="/lib/email/smslib.asp" -->
<%

dim cMember
dim userid, lstp
dim sempno, susername, susercell, isIdentify
dim strSql, chkWait
dim authNo

userid=requestCheckVar(Request("uid"),32)
lstp  =requestCheckVar(Request("lstp"),10)  ''C 인경우 mobileApp 인증 /2013/06/19 추가
if (lstp<>"C") then lstp="S"                '' 기본값 S (기존 SMS 인증 로그인)

'// 직원 기본정보 접수
Set cMember = new CTenByTenMember
	cMember.Fuserid = userid
	cMember.fnGetScmMyInfo

	sempno   		= cMember.Fempno
	susername      	= cMember.Fusername
	susercell      	= cMember.Fusercell
	isIdentify		= cMember.FisIdentify

Set cMember = Nothing

'// 본인확인인증을 받았는지 확인 (안받았으면 본인확인 휴대폰번호 변경 팝업 실행)
if isIdentify<>"Y" then
	Response.Write "<script language=javascript>parent.PopChgHPNum();</script>"
	Response.End
end if

'// 휴대폰번호 여부 확인
if susercell="" or isNull(susercell) then
	Call Alert_Move("회원 정보에 휴대폰 번호가 없습니다.\nUSB키를 사용하여 로그인 후 휴대폰정보를 입력해주세요.","about:blank")
	Response.End
end if

'// 인증번호 발송여부 확인
strSql = "select count(idx) " &_
		" from db_log.dbo.tbl_partner_login_log " &_
		" where userid='" & userid & "' " &_
		" 	and loginSuccess='"&lstp&"' " &_
		" 	and datediff(ss,regdate,getdate()) between 0 and 180"
rsget.Open strSql,dbget,1
	chkWait = rsget(0)>0
rsget.Close

if chkWait then
	Call Alert_Move("이미 인증번호를 발송하였습니다.\n휴대폰의 SMS를 확인해주세요.","about:blank")
	Response.End
end if

'// 인증번호 발생 및 DB저장 후 SMS발송
Randomize()
authNo = int(Rnd()*1000000)		'6자리 난수
authNo = Num2Str(authNo,6,"0","R")

'#문자 전송
Call SendNormalSMS(susercell,"","[텐바이텐SCM] " & susername & "님 인증번호는 ["&authNo&"]입니다.")
'#로그 저장
Call AddLoginLog (userid,lstp,authNo)

'//발송 안내 및 카운터 시작
IF application("Svr_Info")="Dev" THEN
	'// TEST서버이면 그냥 Alert처리
	Response.Write "<script language=javascript>" &_
			"	alert('" & susername & "님 인증번호는 [" & authNo & "]입니다.');" &_
			"</script>"
else
	Response.Write "<script language=javascript>" &_
			"	parent.startLimitCounter('new');" &_
			"	alert('휴대폰으로 인증번호를 발송했습니다.\nSMS를 확인 후 로그인해주세요.');" &_
			"</script>"
end if

'-----------------------------------------------------------
'// 웹어드민 접속 로그 저장 함수
Sub AddLoginLog(param1,param2,param3)
    dim sqlStr, reFAddr
    reFAddr = request.ServerVariables("REMOTE_ADDR")

    sqlStr = " insert into [db_log].[dbo].tbl_partner_login_log" + VbCrlf
	sqlStr = sqlStr + " (userid,refip,loginSuccess,USBTokenSn)" + VbCrlf
	sqlStr = sqlStr + " values(" + VbCrlf
	sqlStr = sqlStr + " '" + param1 + "'," + VbCrlf
	sqlStr = sqlStr + " '" + Left(reFAddr,16) + "'," + VbCrlf
	sqlStr = sqlStr + " '" + param2 + "'," + VbCrlf
	sqlStr = sqlStr + " '" + param3 + "'" + VbCrlf
	sqlStr = sqlStr + " )" + VbCrlf

    dbget.Execute sqlStr
end Sub
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
