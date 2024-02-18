<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  어드민 휴대폰 변경 인증문자 발송
' History : 2013.02.18 허진원 생성
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<!-- #include virtual="/lib/email/smslib.asp"-->
<%
dim cMember
dim empNo, chgHp
dim sempno, susername
dim strSql, chkWait
dim authNo

empNo=requestCheckVar(Request("eno"),15)
chgHp=requestCheckVar(Request("chp"),16)

if (empNo="") then
	Call Alert_Move("사원번호가 없습니다.","about:blank")
	Response.End
end if

if (chgHp="" or chgHp="--") then
	Call Alert_Move("발송할 휴대폰번호가 없습니다.","about:blank")
	Response.End
end if

'// 직원 기본정보 접수
Set cMember = new CTenByTenMember
	cMember.Fempno = empNo
	cMember.fnGetMemberData

	sempno   		= cMember.Fempno
	susername      	= cMember.Fusername

Set cMember = Nothing

if (sempno="" or isNull(sempno)) then
	Call Alert_Move("잘못된 사원번호입니다.\n시스템팀에 문의해주세요.","about:blank")
	Response.End
end if

'// 인증번호 발송여부 확인
strSql = "select count(idx) " &_
		" from db_log.dbo.tbl_partner_login_log " &_
		" where userid='" & sempno & "' " &_
		" 	and loginSuccess='S' " &_
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
'Call SendNormalSMS(chgHp,"","[텐바이텐어드민] " & susername & "님 인증번호는 ["&authNo&"]입니다.")
Call SendNormalSMS_LINK(chgHp,"","[텐바이텐어드민] " & susername & "님 인증번호는 ["&authNo&"]입니다.")
'#로그 저장
Call AddLoginLog (sempno,"S",authNo)

'//발송 안내 및 카운터 시작
IF application("Svr_Info")="Dev" THEN
	'// TEST서버이면 그냥 Alert처리
	Response.Write "<script language=javascript>" &_
			"	parent.startLimitCounter('new');" &_
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
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->