<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/md5.asp" -->
<%
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

	'로그인 확인
	if session("ssBctSn")="" or isNull(session("ssBctSn")) then
		Call Alert_Return("잘못된 접속입니다.")
		dbget.close()	:	response.End
	end if

	'// 변수 선언 및 전송값 접수
	dim empno, userpass, userpass2, sql
	empno  = requestCheckVar(trim(request.Form("uid")),32)
	userpass = requestCheckVar(trim(request.Form("upwd")),32)
	userpass2 = requestCheckVar(trim(request.Form("upwd2")),32)

    if (LCASE(empno)<>LCASE(session("ssBctSn"))) then
        Call Alert_Return("잘못된 접속입니다...")
		dbget.close()	:	response.End
    end if

	'패스워드 확인
	if userpass<>userpass2 then
		response.write "<script language='javascript'>" &vbCrLf &_
						"	alert('비밀번호 확인이 틀립니다.\n정확한 비밀번호를 입력해주세요.');" &vbCrLf &_
						"</script>"
		dbget.close()	:	response.End
	end if
	if chkPasswordComplex(empno,userpass)<>"" then
		response.write "<script language='javascript'>" &vbCrLf &_
						"	alert('" & chkPasswordComplex(empno,userpass) & "\n다른 비밀번호를 입력해주세요.');" &vbCrLf &_
						"	parent.document.forms[0].upwd.value='';" &vbCrLf &_
						"	parent.document.forms[0].upwd2.value='';" &vbCrLf &_
						"	parent.document.forms[0].upwd.focus();" &vbCrLf &_
						"</script>"
		dbget.close()	:	response.End
	end if

	'// 패스워드 변경
	dbget.beginTrans

	on Error Resume Next
	sql = "Update [db_partner].[dbo].tbl_user_tenbyten " + vbCrlf
	sql = sql + " set Enc_emppass64='" & SHA256(MD5(userpass)) & "' " + vbCrlf
	sql = sql + " where empno = '" + empno + "'" + vbCrlf
	dbget.Execute(sql)

	'오류검사 및 반영
	If Err.Number = 0 Then
		dbget.CommitTrans				'커밋(정상)
		
		Call AddLoginLog (empno,"R","") ''패스워드 수정 R - flag
		response.write "<script language='javascript'>top.location.replace('/tenmember/index.asp')</script>"

	Else
		dbget.RollBackTrans				'롤백(에러발생시)

		response.write	"<script language='javascript'>" &_
					"	alert('처리중 에러가 발생했습니다.');" &_
					"	history.back();" &_
					"</script>"

	End If
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->