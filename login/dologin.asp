<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/md5.asp" -->
<!-- #include virtual="/lib/util/tenEncUtil.asp" -->
<!-- #include virtual="/lib/util/base64unicode.asp" -->
<!-- #include virtual="/lib/NoUSBAllowIpList.asp"-->
<!-- #include virtual="/lib/checkUSBAllowIpList.asp"-->
<!-- #include virtual="/login/partner_loginCheck_function.asp"-->
<!-- #include virtual="/admin/incTenRedisSession.asp"-->
<%
dim manageUrl
IF (application("Svr_Info")	= "Dev") Then
	manageUrl = "http://"&request.ServerVariables("HTTP_HOST")
Else
	manageUrl = "https://"&request.ServerVariables("HTTP_HOST")
End If

'// 변수 선언 및 전송값 접수
dim userid, userpass, Enc_userpass, Enc_userpass64, backurl, tokenSn, lgnMethod, AuthNo
dim saved_id

lgnMethod = requestCheckVar(trim(request.Form("lgnMethod")),1)
if lgnMethod="S" THEN
	userid  = requestCheckVar(trim(request.Form("usid")),32)
	userpass = requestCheckVar(trim(request.Form("uspwd")),32)
	saved_id= requestCheckVar(trim(request.Form("saved_sid")),1)
else
	userid  = requestCheckVar(trim(request.Form("uid")),32)
	userpass = requestCheckVar(trim(request.Form("upwd")),32)
	saved_id= requestCheckVar(trim(request.Form("saved_id")),1)
end if

''USB 인증없이 로그인 체크
'Dim NoUsbValidIP : NoUsbValidIP = fnIsNoUsbAllowIp
Dim NoUsbValidIP
'if (NoUsbValidIP = False) then
	NoUsbValidIP = fncheckAllowIPWithByDB("Y", "", "")
'end if

''2017/04/20 REFERER 없는것 추가로 막음.----------------------, 업체로그인은 더이상 이곳을 못탐..
dim iref : iref = Request.ServerVariables("HTTP_REFERER")
dim irefIP : irefIP = request.ServerVariables("REMOTE_ADDR")

if (Instr(iref,"webadmin.10x10.co.kr")<1) then
	if NOT G_IsLocalDev then
		Call fn_plogin_AddIISLOG("addlog=plogin&sub=noref&uid="&userid)
		response.write("<script>window.alert('더이상 사용할 수 없는 페이지입니다.');</script>")
		response.write("<script>history.go(-1);</script>")
		dbget.close()	:	response.End
	end if
end if
''-----------------------------------------------------------


Enc_userpass = md5(userpass)
Enc_userpass64 = SHA256(md5(userpass))
tokenSn = requestCheckVar(trim(request.Form("tokenSn")),26)

AuthNo = requestCheckVar(trim(request.Form("sAuthNo")),6)

dim dbpassword
dim sql
dim errMsg
dim frontId

dim lockTerm, failNo
failNo = 5			'// 로그인 실패 허용수
lockTerm = 15		'// 계정 장금 시간 설정(분)

'### 전송값 확인
if ( userid = "" or userpass = "") then
    response.write("<script>window.alert('아이디 또는 비밀번호가 입력되지 않았습니다.');</script>")
    response.write("<script>history.go(-1);</script>")
    dbget.close()	:	response.End
end if

'### 계정 로그 확인
dim lastlogindt,lastpwchgdt,lastInfoChgDT, lastloginOrRegDiff, partnerUserdiv, Enc_2password64, regdiffDt
dim isFirstConnect : isFirstConnect = false '' 아이디 발급 이후 최초접속
dim isRequirePwdUp : isRequirePwdUp = false '' 패스워드 변경 필요
dim isRequireInfoUp : isRequireInfoUp = false '' 담당자정보 변경 필요
dim isChangedIp : isChangedIp = false 			'' 접속IP 변동여부

sql = "select  isNull(max(regdate),getdate()) as regdate " &VbCRLF
sql = sql + "	,isNull(sum(Case loginSuccess When 'N' Then 1 end),0) as FailCnt " &VbCRLF
sql = sql + "	,max(Case When rowNum=1 and loginSuccess='Y' Then refip end) as lastIP " &VbCRLF
sql = sql + "from (select top " & failNo & " regdate, loginSuccess, refip, ROW_NUMBER() over(partition by loginSuccess order by idx desc) as rowNum " &VbCRLF
sql = sql + "	from [db_log].[dbo].tbl_partner_login_log with (nolock)" &VbCRLF
sql = sql + "	where userid='" & userid & "' " &VbCRLF
sql = sql + "		and loginSuccess in ('Y','S') " &VbCRLF
sql = sql + "	order by idx desc) as pLog " &VbCRLF
rsget.CursorLocation = adUseClient
rsget.Open sql,dbget,adOpenForwardOnly, adLockReadOnly
	'// 연속 로그인 실패 후 지정시간 동안 계정 잠금
	if (datediff("n",rsget("regdate"),now)<lockTerm) and (rsget("FailCnt")>=failNo) then
	    Call fn_plogin_AddIISLOG("addlog=plogin&sub=lock&uid="&userid)
	    response.write("<script>window.alert('비밀번호를 연속으로 " & failNo & "번 틀려 아이디가 잠겼습니다.\n" & (lockTerm-datediff("n",rsget("regdate"),now)) & "분 후 다시 로그인을 해주세요.');</script>")
	    response.write("<script>history.go(-1);</script>")
	    dbget.close()	:	response.End
	end if
	if rsget("lastIP")<>irefIP and left(irefIP,8)<>"192.168." and left(irefIP,9)<>"172.16.1." and irefIP <> "::1" then isChangedIp = true
rsget.Close


''## IP  확인 2017/04/11
if (IspartnerLoginRejectIP()) then
    Call fn_plogin_AddIISLOG("addlog=plogin&sub=rjtip&uid="&userid)
    response.write("<script>window.alert('계정이 잠기었습니다. ');</script>")
    response.write("<script>history.go(-1);</script>")
	dbget.close()	:	response.End
end if

dim GeoIpCCD : GeoIpCCD = getGeoIpCountryCode()
dim RefCode : RefCode = getConSVCByUagentOrRefer()
''dim AuthReqIP : AuthReqIP = IsPartnerAuthRequireIP(userid,GeoIpCCD,FALSE)

if (GeoIpCCD="--") and (application("Svr_Info")="Dev") then GeoIpCCD="KR"
RefCode = "P:"&RefCode  ''(구)로그인 구분하기위한값


'### 유저정보 접수 '//2011-03-9 한용민(정윤정) 수정 - 오푸샵 아이디 추가
sql = "select top 1 A.id, A.company_name, A.tel, A.fax, A.url, A.email, A.userdiv, A.Enc_password, A.Enc_password64, A.groupid " + vbCrlf
sql = sql + "	, B.part_sn, A.level_sn, B.job_sn, B.username,  B.direct070, B.usermail, B.posit_sn, IsNull(B.empno, '') as empno, B.frontid " + vbCrlf
sql = sql + "	, A.lastlogindt, isNULL(A.lastpwchgdt,'2001-01-01') as lastpwchgdt " + vbCrlf
sql = sql + "	, isNULL(A.lastInfoChgDT,A.regdate) as lastInfoChgDT, IsNull(B.criticinfouser,0) as criticinfouser " + vbCrlf
sql = sql + " , b.lv1customerYN, b.lv2partnerYN, b.lv3InternalYN" + vbCrlf
sql = sql + " ,datediff(d,isnull((CASE WHEN isNULL(A.lastlogindt,'2001-01-01')>isNULL(A.lastPwChgDT,'2001-01-01') THEN A.lastlogindt ELSE A.lastPwChgDT END),A.regdate),getdate()) as lastloginOrRegDiff " + vbCrlf ''최종로그인or등록일 지난기간month.  2017/04/10 추가
sql = sql + " ,datediff(d,A.regdate,getdate()) as regdiffDt " + vbCrlf
sql = sql + " , isNULL(A.Enc_2password64,'') as Enc_2password64" + vbCrlf
sql = sql + " ,(select top 1 shopid" + vbCrlf
sql = sql + " 	from db_partner.dbo.tbl_partner_shopuser with (nolock)" + vbCrlf
sql = sql + " 	where b.empno=empno and firstisusing='Y') as firstshopid" + vbCrlf
sql = sql + " from [db_partner].[dbo].tbl_partner as A with (nolock)" + vbCrlf
sql = sql + " 	left join db_partner.dbo.tbl_user_tenbyten as B with (nolock) ON A.id = B.userid AND B.isUsing = 1" + vbCrlf		'AND B.statediv = 'Y'
sql = sql + " where A.id = '" + userid + "'" + vbCrlf
'sql = sql + " and A.isusing='Y'"

'퇴사예정인 사람까지 로그인이 안되서, 퇴사예정일 까지 고려해서 처리. '/2017.02.23 한용민
sql = sql & " and (A.isusing='Y' or (" + vbCrlf
sql = sql & " 		(A.isusing='N' or B.statediv = 'N') and dateDiff(d,convert(varchar(10),getdate(),121),B.retireday)>=0 and B.retireday is not null" + vbCrlf
sql = sql & " 	)" + vbCrlf
sql = sql & " )" + vbCrlf

rsget.CursorLocation = adUseClient
rsget.Open sql,dbget,adOpenForwardOnly, adLockReadOnly

if  not rsget.EOF  then
	'// 로그인 정보 확인
	partnerUserdiv  = rsget("userdiv")
	RefCode=replace(RefCode,"P:",partnerUserdiv&":")

	if rtrim(UCase(rsget("Enc_password64")))=trim(UCase(Enc_userpass64)) then


    	dbpassword  = rsget("Enc_password64")
    	lastlogindt = rsget("lastlogindt")  ''최종 접속 성공일
    	lastpwchgdt = rsget("lastpwchgdt")  ''최종 패스워드 변경일
    	lastInfoChgDT = rsget("lastInfoChgDT")  ''최종 담당자정보  변경일
		frontId = rsget("frontid") '' 프론트 ID

    	isFirstConnect = isNULL(lastlogindt)
        lastloginOrRegDiff  = rsget("lastloginOrRegDiff")  ''최종 로그인 OR 등록일
        Enc_2password64 = rsget("Enc_2password64")
        regdiffDt = rsget("regdiffDt")  ''등록일 이후 day

    	if (isFirstConnect) then
    	    isRequirePwdUp = true
    	    isRequireInfoUp = true
    	else
    	    ''  isRequirePwdUp = (datediff("d",lastpwchgdt,now())>91)
    	    ''''isRequirePwdUp = (datediff("d",lastpwchgdt,now())>91) and (datediff("d",lastlogindt,now())>0) '' 패스워드 최종변경일을 2014/07/15 부터 넣었으므로.. 우선 lastlogindt 조건넣음.
    	    ''''if (CLNG(rsget("userdiv"))<10) then ''일단 직원
    	    ''''    isRequirePwdUp = (datediff("d",lastpwchgdt,now())>91)
    	    ''''end if

    	    isRequirePwdUp = (datediff("d",lastpwchgdt,now())>91)  ''2017/04/10
    	    isRequireInfoUp=  (datediff("d",lastInfoChgDT,now())>91)
        end if

        ''일단 막음.. // 아예 막음.(2017/04/20) ,CLNG(partnerUserdiv)>=10)(2017/04/21)
        if (CLNG(partnerUserdiv)>=10) then
            Call fn_plogin_AddIISLOG("addlog=plogin&sub=noscmpartner&uid="&userid)
            response.write("<script>window.alert('더이상 사용할 수 없는 페이지입니다.');</script>")
            response.write("<script>history.go(-1);</script>")
            dbget.close()	:	response.End
        end if


        ''2017/04/10 추가 3개월 이상 접속 없는경우 / 2차패스워드 없는경우
        ''/login/partner_loginCheck_function.asp 에 있으나.. 중간에 쿼리 할수 없음.. 좀 뜯어 고쳐야..
        if (lastloginOrRegDiff>91) then
            Call fn_plogin_AddIISLOG("addlog=plogin&sub=logntimenosee&uid="&userid)
            response.write("<script>window.alert('장기간 사용하지 않아 계정이 잠겼습니다.');</script>")
            response.write("<script>history.go(-1);</script>")
            dbget.close()	:	response.End
        elseif ((partnerUserdiv="9999") and (isNULL(Enc_2password64) or Enc_2password64="")) then
            Call fn_plogin_AddIISLOG("addlog=plogin&sub=2ndpassnull&uid="&userid)
            response.write("<script>window.alert('2차 패스워드 등록이 필요합니다.');</script>")
            response.write("<script>history.go(-1);</script>")
            dbget.close()	:	response.End
        end if

        ''2017/04/12 해외IP => SMS 인증을 통과 해야함.

        if (GeoIpCCD<>"KR") and Not(NoUsbValidIP) then
            if (AuthNo="") then
                Call fn_plogin_AddIISLOG("addlog=plogin&sub=geoipfail&uid="&userid&"&geoipccd="&GeoIpCCD)
                response.write("<script>window.alert('SMS 인증이 필요합니다..');</script>")
                response.write("<script>history.go(-1);</script>")
                dbget.close()	:	response.End
            end if
        end if

		'2022/08/01 마지막 접속IP와 다른 IP인경우 SMS인증으로
		if isChangedIp and lgnMethod="U" then
			response.write("<script>window.alert('새로운 환경에서 접속하셨습니다.\nSMS 인증을 진행해주세요.');</script>")
			response.write("<script>location.replace(""/index.asp?lgnMethod=S"");</script>")
			dbget.close()	:	response.End
		end if



        ''------------------------------------------------------------------------------------

        session("ssBctId") = rsget("id")
        session("ssBctDiv") = rsget("userdiv")
        session("ssBctBigo") = rsget("firstshopid")
		session("ssBctSn") = rsget("empno")
        IF session("ssBctDiv") <= 9 THEN
        	 session("ssBctCname") = rsget("username")
        	 session("ssBctEmail") = db2html(rsget("usermail"))
        ELSE
        	if isnull(rsget("company_name")) then
        		session("ssBctCname") = rsget("username")
        	else
        		session("ssBctCname") = db2html(rsget("company_name"))
        	end if

        	session("ssBctEmail") = db2html(rsget("email"))
    	END IF

		session("ssGroupid") = rsget("groupid")
		session("ssAdminPsn") = rsget("part_sn")		'부서 번호
		session("ssAdminLsn") = rsget("level_sn")		'등급 번호
		session("ssAdminPOsn") = rsget("job_sn")		'직책 번호
		session("ssAdminPOSITsn") = rsget("posit_sn")		'직급 번호
		session("ssAdminCLsn") = rsget("criticinfouser")	'개인정보 취급권한
		session("ssAdminlv1customerYN") = rsget("lv1customerYN")
		session("ssAdminlv2partnerYN") = rsget("lv2partnerYN")
		session("ssAdminlv3InternalYN") = rsget("lv3InternalYN")

		'3PL SSO 용 쿠키생성(유저아이디 + 접속아이피 + 접속일자) 및 암호화
		'로그인 후 아이피가 변경되면(스마트폰 접속 등) 로그인이 실패한다.
		'코딩 간소화를 위해 비밀번호는 쿠키로 생성하지 않는다. 차후 변경필요.(비번 단방향 암호화 및 쿠키저장)
		Response.Cookies("ThreePL").Domain				= "10x10.co.kr"
		Response.Cookies("ThreePL")("UserID")			= TBTEncrypt(CStr(rsget("id")) & "," & Request.ServerVariables("REMOTE_HOST") & "," & Left(now(), 10))

        '2014-12-17 김진영 // API서버용 쿠키생성
		Response.Cookies("wapi").Domain				= "10x10.co.kr"
		Response.Cookies("wapi")("UserID")			= TBTEncrypt(CStr(rsget("id")) & "," & Request.ServerVariables("REMOTE_HOST") & "," & Left(now(), 10))
		If isnull(rsget("part_sn")) OR rsget("part_sn") = "" Then
		Else
			Response.Cookies("wapi")("PartSN") 		= TBTEncrypt(rsget("part_sn") & "," & Request.ServerVariables("REMOTE_HOST") & "," & Left(now(), 10))
		End If

		'// FrontAPI용 쿠키생성
		If frontId <> "" Then
			Dim ssnlogindt, retSsnHash, cookieDomain
			ssnlogindt = fnDateTimeToLongTime(now())
			session("ssnlogindt") = ssnlogindt
			retSsnHash = fnDBSessionCreateV2(frontId)

			If Application("Svr_Info") = "Dev" And InStr(Request.ServerVariables("HTTP_REFERER"), "localhost") > 0 Then
				cookieDomain = "localhost"
			Else
				cookieDomain = "10x10.co.kr"
			End If

			Response.Cookies("pinfo").domain = cookieDomain
			Response.Cookies("pinfo")("ssndt") = ssnlogindt
			Response.Cookies("pinfo")("ssnhash") = retSsnHash
		End If

		'아이디저장
		response.Cookies("ScmSave").domain = "10x10.co.kr"
    	response.cookies("ScmSave").Expires = Date + 30	'1개월간 쿠키 저장
	    If saved_id = "o" Then
	    	response.cookies("ScmSave")("SAVED_ID") = tenEnc(CStr(rsget("id")))
	    Else
	    	response.cookies("ScmSave")("SAVED_ID") = ""
	    End If

		'세션 쿠키 전역으로 심기
		Dim cookieData, scResult, lp
		cookieData = Request.ServerVariables("HTTP_COOKIE")

		if instr(cookieData,"ASPSESSIONID")>0 then
			cookieData = Split(cookieData,";")
			for lp=0 to ubound(cookieData)
				if instr(Split(cookieData(lp),"=")(0),"ASPSESSIONID")>0 then
					scResult = scResult & cookieData(lp) & ";"
				end if
			next
		end if

		response.Cookies("TENSSID").domain = "10x10.co.kr"
		response.Cookies("TENSSID") = Base64EncodeUnicode(scResult)

		''로그저장(성공)
	    rsget.close

	    if (isFirstConnect) then
	        ''최초접속인경우 성공으로 안봄 비번 변경으로 // 비번 변경후 최초로그인 일자 Update , 3개월단위 강제로 할경우 isRequirePwdUp 조건추가.
	    else
    	    if AuthNo<>"" then
    	    	Call AddPartnerLoginLogWithGeoIpCode (userid,"Y",AuthNo,GeoIpCCD)
    	    elseif tokenSn<>"" then
    	    	Call AddPartnerLoginLogWithGeoIpCode (userid,"Y",tokenSn,GeoIpCCD)
    	    else
    	        Call AddPartnerLoginLogWithGeoIpCode (userid,"Y",RefCode,GeoIpCCD)
    	    end if
    	end if
    	Call fn_plogin_AddIISLOG("addlog=plogin&sub=loginsucc&uid="&userid&"&userdiv="&partnerUserdiv)
	else
	    ''로그저장(실패)
	    rsget.close
	    if AuthNo<>"" then
	    	Call AddPartnerLoginLogWithGeoIpCode (userid,"N",AuthNo,GeoIpCCD)
	    elseif tokenSn<>"" then
	        Call AddPartnerLoginLogWithGeoIpCode (userid,"N",tokenSn,GeoIpCCD)
	    else
	    	Call AddPartnerLoginLogWithGeoIpCode (userid,"N",RefCode,GeoIpCCD)
	    end if

        Call fn_plogin_AddIISLOG("addlog=plogin&sub=faillogin&uid="&userid&"&userdiv="&partnerUserdiv)
        response.write("<script>window.alert('아이디 또는 비밀번호가 틀렸습니다.\n비밀번호 대소문자를 확인해주세요. ');</script>")
        response.write("<script>history.go(-1);</script>")
        dbget.close()	:	response.End
	end if
else
	'' 로그추가.2017/04/12 F
	if AuthNo<>"" then
    	Call AddPartnerLoginLogWithGeoIpCode (userid,"F",AuthNo,GeoIpCCD)
    elseif tokenSn<>"" then
        Call AddPartnerLoginLogWithGeoIpCode (userid,"F",tokenSn,GeoIpCCD)
    else
    	Call AddPartnerLoginLogWithGeoIpCode (userid,"F",RefCode,GeoIpCCD)
    end if

    '// 계정없음 , 사용중지
	Call fn_plogin_AddIISLOG("addlog=plogin&sub=nouserid&uid="&userid)
    response.write("<script>window.alert('계정이 활성화 되지 않았거나, 아이디 또는 비밀번호가 틀렸습니다.');</script>")
    response.write("<script>history.go(-1);</script>")
    dbget.close()	:	response.End
end if


'### 로그인 계정 구분별 처리 ------------------------------------------------------

''강사임시 => 강사 이곳으로 로그인 불가. 2017/04/21


dim ssnTmpUID
dim isNoComplexPwTxt : isNoComplexPwTxt = chkPasswordComplex_Len6Ver(userid,userpass)

if trim(UCase(dbpassword))=trim(UCase(Enc_userpass64)) then

    ''// 최초 로그인 성공시 패스워드 변경(2014/05/19 서동석)
    if ((isFirstConnect) or (isRequirePwdUp) or (isNoComplexPwTxt<>"")) then

        ssnTmpUID = session("ssBctId")

        '' session.Abandon is Async ?
        '' http://stackoverflow.com/questions/1470445/what-is-the-difference-between-session-abandon-and-session-clear
        Session.Contents.RemoveAll()

        CAll fnCookieExpire()

        session("ssnTmpUID")= ssnTmpUID   ''비번 변경시 사용 세션값

        if (isFirstConnect) then
            response.write "<script language='javascript'>" &vbCrLf &_
    						"	alert('최초 로그인시 비밀번호를 변경하셔야 합니다. \n비밀번호 변경페이지로 이동합니다.');" &vbCrLf &_
    						"	self.location='/login/modifyPassword.asp';" &vbCrLf &_
    						"</script>"
    		dbget.close()	:	response.End
        end if

        ''// N(3)개월간 비밀번호 변경안한경우 변경페이지로 이동
        if (isRequirePwdUp) then
            response.write "<script language='javascript'>" &vbCrLf &_
    						"	alert('3개월 이상 비밀번호를 변경하지 않으셨습니다. \n비밀번호 변경페이지로 이동합니다.');" &vbCrLf &_
    						"	self.location.href='/login/modifyPassword.asp';" &vbCrLf &_
    						"</script>"
    		dbget.close()	:	response.End
        end if

        '// 비밀번호 강화 정책 시행(2008.12.12; 허진원)  'chkPasswordComplex => chkPasswordComplex_Len6Ver 2016/09/20
    	if (isNoComplexPwTxt)<>"" then
    		response.write "<script language='javascript'>" &vbCrLf &_
    						"	alert('" & chkPasswordComplex_Len6Ver(userid,userpass) & "\n비밀번호 변경페이지로 이동합니다.');" &vbCrLf &_
    						"	self.location='/login/modifyPassword.asp';" &vbCrLf &_
    						"</script>"
    		dbget.close()	:	response.End
    	end if

    end if



    response.Cookies("partner").domain = "10x10.co.kr"
    response.Cookies("partner")("userid") = session("ssBctId")
    response.Cookies("partner")("userdiv") = session("ssBctDiv")



    ''직원인경우 프런트Pw와 어드민Pw가 같을경우 errMsg 및 USB토큰 확인
    if (session("ssBctDiv")<=9) then

        '''20120621 추가//서동석 - 권한설정을 빼먹는경우가 있음...
        if (session("ssAdminLsn")<1) then
            session.Abandon
            CAll fnCookieExpire()
            response.write("<script>window.alert('권한이 설정되 있지 않습니다. 관리자 문의요망.');top.location = '/';</script>")
			dbget.close()	:	response.End
        end if

        sql = "select top 1 * from "
        sql = sql + " [db_user].[dbo].tbl_logindata u with (nolock)"
        sql = sql + " where u.userid='" & session("ssBctId") & "'" &VbCRLF
        sql = sql + " and u.Enc_userpass64='" & Enc_userpass64 & "'" &VbCRLF		'2014.06.25 SHA256변경
        ''sql = sql + " and u.Enc_userpass='" & Enc_userpass & "'" &VbCRLF

        rsget.CursorLocation = adUseClient
        rsget.Open sql,dbget,adOpenForwardOnly, adLockReadOnly
    	if  not rsget.EOF  then
    		errMsg = "프런트 와 어드민 비밀번호를 동일하게 사용하고 있습니다. \n\nMyInfo에서 어드민 비밀번호를 변경하여 사용하세요."
    	end if
    	rsget.close


		'// 보안 로그인방법 추가(2011.06.14; 허진원)
		if lgnMethod="" then
		    session.Abandon
		    CAll fnCookieExpire()
		    Call fn_plogin_AddIISLOG("addlog=plogin&sub=lologinmtd&uid="&userid)
		    response.write("<script>window.alert('텐바이텐 직원용 로그인페이지가 아닙니다.\n직원용 페이지로 이동합니다.\n\n※웹어드민의 로그인페이지가 변경되었습니다.\n기존페이지로 들어오신 분은 즐겨찾기를 변경해주세요.');top.location = '"&getSCMURL&"/';</script>")
		    dbget.close()	:	response.End
		else

			if lgnMethod="U" then
				'// USB토큰 확인(2008.06.19; 허진원) //
				if (tokenSn="") then
				    if (NoUsbValidIP) then ''2014/10/29 추가, 2018-05-31 수정, skyer9
				        session("sslgnMethod") = "S"
				    else
				        session.Abandon
				        CAll fnCookieExpire()
				        Call fn_plogin_AddIISLOG("addlog=plogin&sub=notokensn&uid="&userid)
    				    response.write("<script>window.alert('USB키가 없습니다.\n\n텐바이텐 USB키가 제대로 설치되어있는지 확인 후 다시 로그인해주세요.');top.location = '/';</script>")
    				    dbget.close()	:	response.End
    				end if
				else
					'### 유효번호 처리(db_partner.dbo.tbl_admin_key에서 유효번호 확인) ###
					'Token 일련번호 확인(DB)
					sql = "select count(key_idx) " & vbCRLF
					sql = sql & " from db_partner.dbo.tbl_admin_key with (nolock)" & vbCRLF
					sql = sql & " where key_idx='" & tokenSn & "' and del_isusing='Y'"
					rsget.CursorLocation = adUseClient
                    rsget.Open sql,dbget,adOpenForwardOnly, adLockReadOnly

					if rsget(0)<=0 then
					    session.Abandon
					    CAll fnCookieExpire()
					    Call fn_plogin_AddIISLOG("addlog=plogin&sub=invalidkensn&uid="&userid)
						response.write("<script>window.alert('유효한 USB키가 아닙니다.\n관리자에게 문의해주세요.');top.location = '/';</script>")
						dbget.close()	:	response.End
					end if
					rsget.Close

                    ''2017/06/20 추가. USB인증 허용IP 검증.
                    if (Not IsUsbLoginAlowIp) then
                        session.Abandon
                        CAll fnCookieExpire()
					    Call fn_plogin_AddIISLOG("addlog=plogin&sub=usbnoip&uid="&userid)
						response.write("<script>window.alert('접근 가능한 경로가 아닙니다. SMS 인증을 사용하세요.');top.location = '/';</script>")
						dbget.close()	:	response.End
                    end if
				end if
				'// USB토큰확인 끝 //

			elseif lgnMethod="S" then

				'// SMS인증 로그인
				if AuthNo="" then
				    session.Abandon
				    CAll fnCookieExpire()
				    Call fn_plogin_AddIISLOG("addlog=plogin&sub=noauthno&uid="&userid)
				    response.write("<script>window.alert('인증번호가 없습니다.\n휴대폰으로 전송된 인증번호를 정확히 입력해주세요.');top.location = '/?lgnMethod="&lgnMethod&"';</script>")
				    dbget.close()	:	response.End
				else
					'유효한 시간내의 인증번호 확인
					sql = "select USBTokenSn " & vbCRLF
					sql = sql & " from db_log.dbo.tbl_partner_login_log with (nolock)" & vbCRLF
					sql = sql & " where userid='" & userid & "' " & vbCRLF
					sql = sql & " 	and loginSuccess='S' " & vbCRLF
					sql = sql & " 	and datediff(ss,regdate,getdate()) between 0 and 180"
					rsget.CursorLocation = adUseClient
                    rsget.Open sql,dbget,adOpenForwardOnly, adLockReadOnly

					if rsget.EOF or rsget.BOF  then
					    session.Abandon
					    CAll fnCookieExpire()
					    Call fn_plogin_AddIISLOG("addlog=plogin&sub=expireauthno&uid="&userid)
						response.write("<script>window.alert('입력 제한시간이 초과되었습니다.\n다시 인증번호를 발급받아 입력해주세요.');top.location = '/?lgnMethod="&lgnMethod&"';</script>")
						dbget.close()	:	response.End
					else
						if trim(rsget("USBTokenSn"))<>trim(AuthNo) then
						    session.Abandon
						    CAll fnCookieExpire()
						    Call fn_plogin_AddIISLOG("addlog=plogin&sub=nomatchauthno&uid="&userid)
							response.write("<script>window.alert('휴대폰으로 발송된 인증번호값이 아닙니다.\n정확히 입력해주세요.');top.location = '/?lgnMethod="&lgnMethod&"';</script>")
							dbget.close()	:	response.End
						else
							'// adminbodyhead.asp의 USB체크를 피하려면 세션에 SMS인증여부 저장
							session("sslgnMethod") = "S"
						end if
					end if
					rsget.Close
				end if
            else
                ''2017/06/20 추가
                session.Abandon
    		    CAll fnCookieExpire()
    		    Call fn_plogin_AddIISLOG("addlog=plogin&sub=xloginmtd&uid="&userid)
    		    response.write("<script>window.alert('텐바이텐 직원용 로그인페이지가 아닙니다.\n직원용 페이지로 이동합니다.\n\n※웹어드민의 로그인페이지가 변경되었습니다.\n기존페이지로 들어오신 분은 즐겨찾기를 변경해주세요.');top.location = '"&getSCMURL&"/';</script>")
    		    dbget.close()	:	response.End
			end if

		end if
    end if


    if (session("ssBctId")="10x10") then
        ''사용안함.
        session.Abandon
        CAll fnCookieExpire()
        Call fn_plogin_AddIISLOG("addlog=plogin&sub=notenid&uid="&userid)
        dbget.close()	:	response.End

    ''직원Level
    elseif (session("ssBctDiv")<=9) then

    	if (errMsg<>"") then
            response.write "<script language='javascript'>alert('" & errMsg & "');</script>"
        end if

		''2018/12/18 incTenRedisSession
		Call fn_RDS_SSN_SET()

    	response.write "<script language='javascript'>location.replace('" & manageUrl & "/admin/index.asp')</script>"
        dbget.close()	:	response.End

    else
        session.Abandon
        CAll fnCookieExpire()
        Call fn_plogin_AddIISLOG("addlog=plogin&sub=notauth&uid="&userid)
        response.write "<script language='javascript'>alert('권한이없습니다.');location.replace('/')</script>"
        dbget.close()	:	response.End
    end if
end if

function fnCookieExpire()
    Response.Cookies("partner").domain = "10x10.co.kr"
    Response.Cookies("partner") = ""
    Response.Cookies("partner").Expires = Date - 1

    Response.Cookies("ThreePL").Domain	= "10x10.co.kr"
    Response.Cookies("ThreePL") = ""
    Response.Cookies("ThreePL").Expires = Date - 1

    Response.Cookies("wapi").Domain	= "10x10.co.kr"
    Response.Cookies("wapi") = ""
    Response.Cookies("wapi").Expires = Date - 1

	response.Cookies("TENSSID").domain = "10x10.co.kr"
	response.Cookies("TENSSID") = ""
    Response.Cookies("TENSSID").Expires = Date - 1

	'' require /admin/incTenRedisSession.asp
    response.Cookies(GG_RDS_COOKIE_KEYNAME).domain = fn_RDS_getCookieDomain()
    response.Cookies(GG_RDS_COOKIE_KEYNAME) = ""
	Response.Cookies(GG_RDS_COOKIE_KEYNAME).Expires = Date - 1
end function

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
