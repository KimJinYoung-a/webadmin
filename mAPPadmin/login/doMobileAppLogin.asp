<%@ codepage="65001" language="VBScript" %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control", "no-cache"
Response.AddHeader "Expires", "0"
Response.AddHeader "Pragma", "no-cache"
%>
<!-- #include virtual="/mAppadmin/inc/incUTF8.asp" -->
<!-- #include virtual="/mAppadmin/inc/incCommon.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAppNotiopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/md5.asp" -->
<!-- #include virtual="/lib/util/tenEncUtil.asp" -->
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

    ''USBTokenSn 길이제약 확인
    ''dbget.Execute sqlStr
end Sub

'// 변수 선언 및 전송값 접수
dim sqlStr, AssignedRow
dim userid, userpass, backurl, devicekey, appid, AuthNo, lgnMethod, cflag, backpath
userid  = requestCheckVar(trim(request.Form("uid")),32)
userpass = requestCheckVar(trim(request.Form("upwd")),32)
devicekey = requestCheckVar(trim(request.Form("devicekey")),512)
appid = requestCheckVar(trim(request.Form("appid")),10)
AuthNo = requestCheckVar(trim(request.Form("sAuthNo")),6)
cflag  = requestCheckVar(trim(request.Form("cflag")),10)
backpath = request.Form("backpath")

''lgnMethod = requestCheckVar(trim(request.Form("lgnMethod")),1)

''디바이스 타입 설정
if (appid<>C_APPID_IOS) and (appid<>C_APPID_AND) then
    response.write("<script>window.alert('ERR-001');</script>")
    response.write("<script>history.go(-1);</script>")
    dbAppNotiget.close() : dbget.close()	:	response.End
end if


dim lockTerm, failNo
failNo = 5			'// 로그인 실패 허용수
lockTerm = 15		'// 계정 장금 시간 설정(분)

'### 전송값 확인
if ( userid = "" or userpass = "") then
    response.write("<script>window.alert('아이디 또는 비밀번호가 입력되지 않았습니다.ERR-002');</script>")
    response.write("<script>history.go(-1);</script>")
    dbAppNotiget.close() : dbget.close()	:	response.End
end if

'### 계정 로그 확인
sqlStr = "select  isNull(max(regdate),getdate()) as regdate " &VbCRLF
sqlStr = sqlStr& "	,isNull(sum(Case loginSuccess When 'N' Then 1 end),0) as FailCnt " &VbCRLF
sqlStr = sqlStr& "from (select top " & failNo & " regdate, loginSuccess " &VbCRLF
sqlStr = sqlStr& "	from [db_log].[dbo].tbl_partner_login_log " &VbCRLF
sqlStr = sqlStr& "	where userid='" & userid & "' " &VbCRLF
sqlStr = sqlStr& "	order by idx desc) as pLog "
rsget.Open sqlStr,dbget,1
	'// 연속 로그인 실패 후 지정시간 동안 계정 잠금
	if (datediff("n",rsget("regdate"),now)<lockTerm) and (rsget("FailCnt")>=failNo) then
	    response.write("<script>window.alert('비밀번호를 연속으로 " & failNo & "번 틀려 아이디가 잠겼습니다.\n" & (lockTerm-datediff("n",rsget("regdate"),now)) & "분 후 다시 로그인을 해주세요. ERR-003');</script>")
	    response.write("<script>history.go(-1);</script>")
	    dbAppNotiget.close() : dbget.close()	:	response.End
	end if
rsget.Close

''ID pwd check
Dim UserPassOK : UserPassOK =false
sqlStr = "select top 1 A.id, A.userdiv,A.password, A.company_name" + vbCrlf
sqlStr = sqlStr + " from [db_partner].[dbo].tbl_partner as A " + vbCrlf
sqlStr = sqlStr & " join db_partner.dbo.tbl_user_tenbyten as B" & vbcrlf
sqlStr = sqlStr & "	    ON A.id = B.userid AND B.isUsing = 1" & vbcrlf

' 퇴사예정자 처리	' 2018.10.16 한용민
sqlStr = sqlStr & "	    and (b.statediv ='Y' or (b.statediv ='N' and datediff(dd,b.retireday,getdate())<=0))" & vbcrlf
sqlStr = sqlStr + " where A.id = '" + userid + "'" + vbCrlf
sqlStr = sqlStr + " and A.isusing='Y'" + vbCrlf
sqlStr = sqlStr + " and (A.userdiv<9) " '' 직원만 로그인 가능.
rsget.Open sqlStr,dbget,1
if not rsget.EOF  then
    UserPassOK = rtrim(LCase(rsget("password")))=trim(LCase(userpass))
end if
rsget.Close

if (Not UserPassOK) then
    response.write("<script>window.alert('아이디 또는 비밀번호가 틀렸습니다. ERR-004');</script>")
    response.write("<script>history.go(-1);</script>")
    dbAppNotiget.close() : dbget.close()	:	response.End
end if

''인증
Dim IsConFirmProcOK : IsConFirmProcOK = false
if (cflag="1") and (AuthNo<>"") and (devicekey<>"") and (appid<>"") then
    sqlStr = "select USBTokenSn " + vbCrlf
	sqlStr = sqlStr + " from db_log.dbo.tbl_partner_login_log " + vbCrlf
	sqlStr = sqlStr + " where userid='" & userid & "' " + vbCrlf
	sqlStr = sqlStr + " 	and loginSuccess='C' " + vbCrlf                                 ''방식 추가 C
	sqlStr = sqlStr + " 	and datediff(ss,regdate,getdate()) between 0 and 180"
	rsget.Open sqlStr,dbget,1
	if rsget.EOF or rsget.BOF  then
		response.write("<script>window.alert('입력 제한시간이 초과되었습니다.\n다시 인증번호를 발급받아 입력해주세요. ERR-005');top.location = '/mAPPadmin/login.asp?cflag=1';</script>")
		session.Abandon
		dbAppNotiget.close() : dbget.close()	:	response.End
	else
		if trim(rsget("USBTokenSn"))<>trim(AuthNo) then
			response.write("<script>window.alert('휴대폰으로 발송된 인증번호값이 아닙니다.\n정확히 입력해주세요. ERR-006');top.location = '/mAPPadmin/login.asp?cflag=1';</script>")
			session.Abandon
			dbAppNotiget.close() : dbget.close()	:	response.End
		else
			'' 인증 proc
			sqlStr = " exec db_AppNoti.dbo.sp_Ten_checkAddConfirmUser '"&userid&"',"&appid&",'"&devicekey&"'"
			dbAppNotiget.Execute sqlStr,AssignedRow
			IsConFirmProcOK = (AssignedRow>0)
		end if
	end if
	rsget.Close
end if

'------------------------------------------------------------------------------------------------------------------------------------------------------------------
'' 인증된 사용자인지 Check
Dim isConfirmedUser : isConfirmedUser =false
sqlStr = " select top 1 * from db_AppNoti.dbo.tbl_tbtpns_register"
sqlStr = sqlStr & " where userid='"&userid&"'"
sqlStr = sqlStr & " and appid="&appid&""
sqlStr = sqlStr & " and regkey='"&devicekey&"'"
sqlStr = sqlStr & " and AuthDate is Not NULL"
rsAppNotiget.Open sqlStr,dbAppNotiget,1
if not rsAppNotiget.EOF  then
    isConfirmedUser = true
end if
rsAppNotiget.close

if (Not isConfirmedUser) then
	response.write("<script>window.alert('"&cflag&"|"&AuthNo&"|"&devicekey&"|"&appid&"');</script>")
    response.write("<script>window.alert('인증되지 않은 DEVICE 입니다. 인증 후 사용하시기 바랍니다. ERR-007');</script>")
    response.write("<script>location.href='/mAPPadmin/login.asp?cflag=1';</script>")
    dbAppNotiget.close() : dbget.close()	:	response.End
end if


Dim tstp : tstp = getTimeStampFormat

''rw "Login OK"
''rw "userid:"&userid
''rw "devicekey:"&devicekey
''rw "AuthNo:"&AuthNo
''rw "tstp:"&tstp
''rw "encuid:"&TENENC(userid)
''rw "endkey:"&md5(devicekey&userid&tstp)

'### 유저정보 접수 '//2011-03-9 한용민(정윤정) 수정 - 오푸샵 아이디 추가
sqlStr = "select top 1 A.id, A.company_name, A.userdiv, A.password " + vbCrlf
sqlStr = sqlStr + "	, B.part_sn, A.level_sn, B.job_sn, B.username, B.posit_sn, IsNull(B.empno, '') as empno " + vbCrlf
sqlStr = sqlStr + " from [db_partner].[dbo].tbl_partner as A " + vbCrlf
sqlStr = sqlStr & " join db_partner.dbo.tbl_user_tenbyten as B" & vbcrlf
sqlStr = sqlStr & "	    ON A.id = B.userid AND B.isUsing = 1" & vbcrlf

' 퇴사예정자 처리	' 2018.10.16 한용민
sqlStr = sqlStr & "	    and (b.statediv ='Y' or (b.statediv ='N' and datediff(dd,b.retireday,getdate())<=0))" & vbcrlf
sqlStr = sqlStr + " where A.id = '" + userid + "'" + vbCrlf
sqlStr = sqlStr + " and A.isusing='Y'"
sqlStr = sqlStr + " and (A.userdiv<9)" '' 직원만 로그인 가능.

rsget.Open sqlStr,dbget,1

if  not rsget.EOF  then
	'// 로그인 정보 확인
	if rtrim(LCase(rsget("password")))=trim(LCase(userpass)) then

        Response.Cookies("mAppADM").Domain      = manageDomain

        Response.Cookies("mAppADM")("encuid")   = TENENC(userid)
        Response.Cookies("mAppADM")("tstp")     = tstp
        Response.Cookies("mAppADM")("dvkey")    = devicekey
        Response.Cookies("mAppADM")("appid")    = appid
        Response.Cookies("mAppADM")("enckey")   = MD5(devicekey&userid&tstp)
        Response.cookies("mAppADM").expires     = Dateadd("h" , glb_cookie_time, Now())


        session("mAppBctId")    = userid
        session("mAppBctTstp")  = tstp
        session("mAppBctDiv")   = rsget("userdiv")
		session("mAppBctSn")    = rsget("empno")
        session("mAppBctCname") = db2html(rsget("company_name"))

		session("mAppAdminPsn")  = rsget("part_sn")		    '부서 번호
		session("mAppAdminLsn")  = rsget("level_sn")		    '등급 번호
		session("mAppAdminPOsn") = rsget("job_sn")		    '직책 번호
		session("mAppAdminPOSITsn") = rsget("posit_sn")		'직급 번호


		''로그저장(성공)
	    rsget.close
	    if AuthNo<>"" then
	    	Call AddLoginLog (userid,"Y",AuthNo)
	    else
	    	Call AddLoginLog (userid,"Y",devicekey)
	    end if

        if (backpath<>"") then
            response.redirect backpath
        else
	        response.redirect "/mAPPadmin/main.asp"
	    end if
	else
	    ''로그저장(실패)
	    rsget.close
	    if AuthNo<>"" then
	    	Call AddLoginLog (userid,"N",AuthNo)
	    else
	    	Call AddLoginLog (userid,"N",devicekey)
	    end if

        response.write("<script>window.alert('아이디 또는 비밀번호가 틀렸습니다. ERR-008');</script>")
        response.write("<script>history.go(-1);</script>")
        dbAppNotiget.close() : dbget.close()	:	response.End
	end if
else
	'// 계정없음
    response.write("<script>window.alert('아이디 또는 비밀번호가 틀렸습니다. ERR-009');</script>")
    response.write("<script>history.go(-1);</script>")
    dbAppNotiget.close() : dbget.close()	:	response.End
end if


%>

<!-- #include virtual="/lib/db/dbAppNoticlose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
