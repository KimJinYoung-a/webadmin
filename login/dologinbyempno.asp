<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 사번로그인
' Hieditor : 서동석 생성
'			 2023.09.07 한용민(사번비밀번호 최종로그인날짜 체크 수정)
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/md5.asp" -->
<!-- #include virtual="/lib/util/tenEncUtil.asp" -->
<!-- #include virtual="/lib/util/base64unicode.asp" -->
<!-- #include virtual="/login/partner_loginCheck_function.asp"-->
<%
'// 웹어드민 접속 로그 저장 함수
''Sub AddLoginLog(param1,param2,param3)
''    dim sqlStr, reFAddr
''    reFAddr = request.ServerVariables("REMOTE_ADDR")
''
''    sqlStr = " insert into [db_log].[dbo].tbl_partner_login_log" + VbCrlf
''	sqlStr = sqlStr + " (userid,refip,loginSuccess,USBTokenSn)" + VbCrlf
''	sqlStr = sqlStr + " values(" + VbCrlf
''	sqlStr = sqlStr + " '" + param1 + "'," + VbCrlf
''	sqlStr = sqlStr + " '" + Left(reFAddr,16) + "'," + VbCrlf
''	sqlStr = sqlStr + " '" + param2 + "'," + VbCrlf
''	sqlStr = sqlStr + " '" + param3 + "'" + VbCrlf
''	sqlStr = sqlStr + " )" + VbCrlf
''
''    dbget.Execute sqlStr
''end Sub 
 
'// 변수 선언 및 전송값 접수
dim empno, userpass, backurl
dim saved_eno
dim IsLoginSuccess
empno  = requestCheckVar(trim(request.Form("usn")),32)			'// 사번
userpass = requestCheckVar(trim(request.Form("unpwd")),32) 
saved_eno= requestCheckVar(trim(request.Form("saved_eno")),1)
	
dim dbpassword
dim sql
dim errMsg

dim lockTerm, failNo
failNo = 5			'// 로그인 실패 허용수
lockTerm = 15		'// 계정 장금 시간 설정(분)

'// ============================================================================
'### 전송값 확인
if (empno = "" or userpass = "") then
    response.write("<script>window.alert('사번 또는 비밀번호가 입력되지 않았습니다.');</script>")
    response.write("<script>window.location.href ='/index.asp?lgnMethod=N'</script>")
    dbget.close()	:	response.End
end if

''// 2017/06/19 추가============================================================
dim GeoIpCCD : GeoIpCCD = getGeoIpCountryCode()
if (GeoIpCCD="--") and (application("Svr_Info")="Dev") then GeoIpCCD="KR" 
    
dim iref : iref = Request.ServerVariables("HTTP_REFERER")
dim irefIP : irefIP = request.ServerVariables("REMOTE_ADDR")

'내부 접속일때 처리
if (GeoIpCCD="--") and (left(irefIP,8)="192.168.") then GeoIpCCD="KR" 

''SCM 통해서만 가능.
' if (Instr(iref,"webadmin.10x10.co.kr")<1) then 
'     Call fn_plogin_AddIISLOG("addlog=plogin&sub=noref&empno="&empno) 
'     response.write("<script>window.alert('더이상 사용할 수 없는 페이지입니다.');</script>")
'     response.write("<script>history.go(-1);</script>")
'     dbget.close()	:	response.End
' end if

''해외IP 불가.
if (GeoIpCCD<>"KR") then
    Call fn_plogin_AddIISLOG("addlog=plogin&sub=geoipfail&empno="&empno&"&geoipccd="&GeoIpCCD) 
    response.write("<script>window.alert('접속이 불가 합니다.');</script>")
    response.write("<script>history.go(-1);</script>")
    dbget.close()	:	response.End
end if

'' 사번로그인은 특정IP만 가능 > 로그인 허용 IP DB 검색
if (NOT (application("Svr_Info")="Dev")) then
if NOT(fncheckAllowIPWithByDB("Y", "", "")) then
    Call fn_plogin_AddIISLOG("addlog=plogin&sub=invalidip&empno="&empno&"&refip="&irefIP) 
    response.write("<script>window.alert('접속이 불가 합니다. 관리자 문의 요망');</script>")
    response.write("<script>history.go(-1);</script>")
    dbget.close()	:	response.End
end if
end if
'// ============================================================================
'### 계정 로그 확인
dim isFirstOrreqChgPwd : isFirstOrreqChgPwd = false ''차후 강제 변경이 필요할경우 같이 사용.
dim isRequirePwdUp : isRequirePwdUp = false '' 패스워드 변경 필요

sql = "select  isNull(max(regdate),getdate()) as regdate " &VbCRLF
sql = sql + "	,isNull(sum(Case loginSuccess When 'N' Then 1 end),0) as FailCnt " &VbCRLF
sql = sql + "	,(select top 1 regdate from [db_log].[dbo].tbl_partner_login_log " &VbCRLF
sql = sql + "		where userid='"&empno&"' " &VbCRLF
sql = sql + "		and loginSuccess in ('Y','R')" &VbCRLF  '' R 패스워드 변경.
sql = sql + "		order by idx desc) as lastloginSuccDt " &VbCRLF  ''최초 로그인 추가
sql = sql + " from (select top " & failNo & " regdate, loginSuccess " &VbCRLF
sql = sql + "	from [db_log].[dbo].tbl_partner_login_log " &VbCRLF
sql = sql + "	where userid='" & empno & "' " &VbCRLF
sql = sql + "	order by idx desc) as pLog " &VbCRLF
rsget.CursorLocation = adUseClient
rsget.Open sql,dbget,adOpenForwardOnly, adLockReadOnly
    isFirstOrreqChgPwd = IsNull(rsget("lastloginSuccDt")) ''2014/05/19
	'// 연속 로그인 실패 후 지정시간 동안 계정 잠금
	if (datediff("n",rsget("regdate"),now)<lockTerm) and (rsget("FailCnt")>=failNo) then
	    response.write("<script>window.alert('비밀번호를 연속으로 " & failNo & "번 틀려 아이디가 잠겼습니다.\n" & (lockTerm-datediff("n",rsget("regdate"),now)) & "분 후 다시 로그인을 해주세요.');</script>")
	     response.write("<script>window.location.href ='/?lgnMethod=N'</script>")
	    dbget.close()	:	response.End
	end if
rsget.Close

'// ============================================================================
'### 유저정보 접수
Dim i_part_sn, i_username, i_level_sn, i_posit_sn
Dim i_LastEmpnoPassWordChangeDate , i_LastLoginOrRegDiff

sql = "SELECT TOP 1 "
sql = sql + "	B.Enc_emppass64 "
sql = sql + "	,B.part_sn "
sql = sql + "	,B.job_sn "
sql = sql + "	,B.username "
sql = sql + "	,B.direct070 "
sql = sql + "	,B.usermail "
sql = sql + "	,B.posit_sn "

'// TODO : 어드민 정보 무시하고 사번 로그인시 개인정보 조회권한으로 지정
if (application("Svr_Info")="Dev") then
	'// sql = sql + "	,IsNull(A.level_sn, 10) as level_sn "
	sql = sql + "	,10 as level_sn "
else
	'// sql = sql + "	,IsNull(A.level_sn, 9) as level_sn "
	sql = sql + "	,9 as level_sn "
end if

sql = sql + "	,isNULL(b.lastEmpnoPwChgDT,'2001-01-01') as lastEmpnoPwChgDT "
sql = sql + "	,datediff(d,isnull((CASE WHEN isNULL(A.lastlogindt,'2001-01-01')>isNULL(b.lastEmpnoPwChgDT,'2001-01-01') THEN A.lastlogindt ELSE b.lastEmpnoPwChgDT END),A.regdate),getdate()) as lastloginOrRegDiff " 
sql = sql & " FROM db_partner.dbo.tbl_user_tenbyten B"				'// 사번 로그인
sql = sql & " LEFT JOIN [db_partner].[dbo].tbl_partner AS A"
sql = sql & "	ON A.id = B.userid"
sql = sql + " WHERE "
sql = sql + "	1 = 1 "
sql = sql + "	AND B.isUsing = 1 "
sql = sql + "	AND B.empno = '" + CStr(empno) + "' "
'sql = sql + "	AND B.statediv = 'Y' "
'sql = sql + "	AND IsNull(A.isusing, 'Y') = 'Y' "

'퇴사예정인 사람까지 로그인이 안되서, 퇴사예정일 까지 고려해서 처리. '/2017.02.23 한용민
sql = sql & " and (B.statediv='Y' or (" + vbCrlf
sql = sql & " 		(B.statediv='N' or IsNull(A.isusing, 'N') = 'N') and dateDiff(d,convert(varchar(10),getdate(),121),B.retireday)>=0 and B.retireday is not null" + vbCrlf
sql = sql & " 	)" + vbCrlf
sql = sql & " )" + vbCrlf

'response.write sql
rsget.CursorLocation = adUseClient
rsget.Open sql,dbget,adOpenForwardOnly, adLockReadOnly

IsLoginSuccess = False
if not rsget.EOF then
	if rtrim(LCase(rsget("Enc_emppass64"))) = trim(LCase(sha256(md5(userpass)))) then
	    i_part_sn = rsget("part_sn")
	    i_username = rsget("username")
	    i_level_sn = rsget("level_sn")
		i_posit_sn = rsget("posit_sn")
		i_LastEmpnoPassWordChangeDate = rsget("lastEmpnoPwChgDT")
		i_LastLoginOrRegDiff = rsget("lastloginOrRegDiff")  ''최종 로그인 OR 등록일
		IsLoginSuccess = True
	end if
end if
rsget.close

if (IsLoginSuccess = True) then
    '// 아래 세션만 생성한다.
	session("ssBctSn") 		= empno
	session("ssBctDiv")		= "5000"		'// 기준 권한 : 개인정보조회
	session("ssAdminPsn") 	= i_part_sn
	session("ssAdminPOSITsn") = i_posit_sn		'직급 번호
	session("ssBctCname") 	= i_username

	'등급 번호
	session("ssAdminLsn") 	= i_level_sn

	if (i_LastLoginOrRegDiff>91) then
		Call fn_plogin_AddIISLOG("addlog=plogin&sub=logntimenosee&empno="&empno)
		response.write("<script>window.alert('장기간 사용하지 않아 계정이 잠겼습니다.');</script>")
		response.write("<script>history.go(-1);</script>")
		dbget.close()	:	response.End
	END IF
    
    ''// 최초 로그인 성공시 패스워드 변경(2014/05/19 서동석)
    if (isFirstOrreqChgPwd) then
        response.write "<script language='javascript'>" &vbCrLf &_
						"	alert('최초 로그인시 비밀번호를 변경하셔야 합니다. \n비밀번호 변경페이지로 이동합니다.');" &vbCrLf &_
						"	self.location='/login/modifyPassword_empno.asp';" &vbCrLf &_
						"</script>"
		dbget.close()	:	response.End

		isRequirePwdUp = true
	else
		isRequirePwdUp = (datediff("d",i_LastEmpnoPassWordChangeDate,now())>91)  ''2017/04/10
    end if

	''// N(3)개월간 비밀번호 변경안한경우 변경페이지로 이동
	if (isRequirePwdUp) then
		response.write "<script language='javascript'>" &vbCrLf &_
						"	alert('3개월 이상 비밀번호를 변경하지 않으셨습니다. \n비밀번호 변경페이지로 이동합니다.');" &vbCrLf &_
						"	self.location.href='/login/modifyPassword_empno.asp';" &vbCrLf &_
						"</script>"
		dbget.close()	:	response.End
	end if
    
    '// 비밀번호 강화 정책 시행(2008.12.12; 허진원)
	if chkPasswordComplex(empno,userpass)<>"" then
		response.write "<script language='javascript'>" &vbCrLf &_
						"	alert('" & chkPasswordComplex(empno,userpass) & "\n비밀번호 변경페이지로 이동합니다.');" &vbCrLf &_
						"	self.location='/login/modifyPassword_empno.asp';" &vbCrLf &_
						"</script>"
		dbget.close()	:	response.End
	end if
	
	'아이디저장
	response.Cookies("ScmSave").domain = "10x10.co.kr"
	response.cookies("ScmSave").Expires = Date + 30	'1개월간 쿠키 저장
    If saved_eno = "o" Then
    	response.cookies("ScmSave")("SAVED_Eno") = tenEnc(CStr(empno))
    Else
    	response.cookies("ScmSave")("SAVED_Eno") = ""
    End If	
end if

if (IsLoginSuccess) then
	''Call AddLoginLog (empno,"Y","")
	Call AddPartnerLoginLogWithGeoIpCode (empno,"Y","",GeoIpCCD)
else
	''Call AddLoginLog (empno,"N","")
	Call AddPartnerLoginLogWithGeoIpCode (empno,"N","",GeoIpCCD)

	response.write("<script>window.alert('아이디 또는 비밀번호가 틀렸습니다.');</script>")
	response.write("<script>window.location.href ='/?lgnMethod=N'</script>")
	dbget.close()	:	response.End
end if


'// ============================================================================
response.write "<script language='javascript'>location.replace('/tenmember/index.asp')</script>"
dbget.close()	:	response.End

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
