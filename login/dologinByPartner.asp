<%@ language=vbscript %>
<% option explicit %>
<%
Response.Expires = 0
 Response.AddHeader "Pragma","no-cache"
 Response.AddHeader "Cache-Control","no-cache,must-revalidate"
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/md5.asp" -->
<!-- #include virtual="/lib/util/tenEncUtil.asp" -->
<!-- #include virtual="/lib/util/base64unicode.asp" -->
<!-- #include virtual="/lib/NoUSBAllowIpList.asp"-->
<!-- #include virtual="/login/partner_loginCheck_function.asp"-->
<%

function FnAddIISLOG(iAddLogs)
    ''addLog 추가 로그 //2016/12/29
    if (request.ServerVariables("QUERY_STRING")<>"") then iAddLogs="&"&iAddLogs
    response.AppendToLog iAddLogs
end function

dim manageUrl
IF application("Svr_Info")="Dev" THEN
	'manageUrl 	 = "http://testwebadmin.10x10.co.kr"
	manageUrl 	 = getSCMURL
ELSE
	manageUrl 	 = "http://webadmin.10x10.co.kr"
END IF

'// 변수 선언 및 전송값 접수
dim userid, userpass, Enc_userpass, Enc_userpass64, backurl, tokenSn, lgnMethod, AuthNo
dim loginNo,vIsSec
dim userpassSec1,userpassSec2,userpassSec, Enc_2userpass64
dim saved_id

loginNo = requestCheckVar(trim(request.Form("loginNo")),1)
vIsSec	= requestCheckVar(trim(request.Form("hidSec")),1)
saved_id= requestCheckVar(trim(request.Form("saved_id")),1)
if loginNo ="" then loginNo ="1"
if loginNo = "2" then
	userid  =  session("tmpUID")
	userpass = session("tmpUPWD")
else
	userid  = requestCheckVar(trim(request.Form("uid")),32)
	userpass = requestCheckVar(trim(request.Form("upwd")),32)
end if

Enc_userpass = md5(userpass)
Enc_userpass64 = SHA256(md5(userpass))
userpassSec1 = requestCheckVar(trim(request.Form("upwdS1")),32)
userpassSec2 = requestCheckVar(trim(request.Form("upwdS2")),32)
userpassSec = requestCheckVar(trim(request.Form("upwdS")),32)

tokenSn = requestCheckVar(trim(request.Form("tokenSn")),26)
lgnMethod = requestCheckVar(trim(request.Form("lgnMethod")),1)
AuthNo = requestCheckVar(trim(request.Form("sAuthNo")),6)



dim dbpassword,isdbpassword_sec,dbpassword2
dim db_id, db_userdiv,db_company_name,db_email,db_groupid

dim sql
dim errMsg

dim lockTerm, failNo
failNo = 5			'// 로그인 실패 허용수
lockTerm = 15		'// 계정 장금 시간 설정(분)

'### 전송값 확인
if ( userid = "" or userpass = "") then
    Call FnAddIISLOG("addlog=plogin&sub=no1step&loginNo="&loginNo&"&uid="&userid) ''2016/12/29
    response.write("<script>window.alert('아이디 또는 비밀번호가 입력되지 않았습니다.');</script>")
    response.write("<script>history.go(-1);</script>")
    dbget.close()	:	response.End
end if

'### 계정 로그 확인
dim lastlogindt,lastpwchgdt,lastInfoChgDT
dim isFirstConnect : isFirstConnect = false '' 아이디 발급 이후 최초접속
dim isRequirePwdUp : isRequirePwdUp = false '' 패스워드 변경 필요
dim isRequireInfoUp : isRequireInfoUp = false '' 담당자정보 변경 필요

sql = "select  isNull(max(regdate),getdate()) as regdate " &VbCRLF
sql = sql & "	,isNull(sum(Case loginSuccess When 'N' Then 1 end),0) as FailCnt " &VbCRLF
sql = sql & "from (select top " & failNo & " regdate, loginSuccess " &VbCRLF
sql = sql & "	from [db_log].[dbo].tbl_partner_login_log " &VbCRLF
sql = sql & "	where userid='" & userid & "' " &VbCRLF
sql = sql & "	order by idx desc) as pLog " &VbCRLF
rsget.CursorLocation = adUseClient
rsget.Open sql,dbget,adOpenForwardOnly, adLockReadOnly
	'// 연속 로그인 실패 후 지정시간 동안 계정 잠금
	if (datediff("n",rsget("regdate"),now)<lockTerm) and (rsget("FailCnt")>=failNo) then
	    Call FnAddIISLOG("addlog=plogin&sub=lock&uid="&userid) ''2016/12/29
	    response.write("<script>window.alert('비밀번호를 연속으로 " & failNo & "번 틀려 아이디가 잠겼습니다.\n" & (lockTerm-datediff("n",rsget("regdate"),now)) & "분 후 다시 로그인을 해주세요.');</script>")
	    response.write("<script>history.go(-1);</script>")
	    dbget.close()	:	response.End
	end if

rsget.Close

''## IP  확인 2017/04/11
if (IspartnerLoginRejectIP()) then
    Call FnAddIISLOG("addlog=plogin&sub=rjtip&uid="&userid) ''2016/12/29
    response.write("<script>window.alert('계정이 잠기었습니다. ');</script>")
    response.write("<script>history.go(-1);</script>")
	dbget.close()	:	response.End
end if

dim GeoIpCCD : GeoIpCCD = getGeoIpCountryCode()
dim RefCode : RefCode = getConSVCByUagentOrRefer()
dim AuthReqIP
dim db_id_Exists, db_Enc_password64, db_Enc_2password64
db_id_Exists = FALSE

if (GeoIpCCD="--") and (application("Svr_Info")="Dev") then GeoIpCCD="KR"

if loginNo = "2" then
	if vIsSec = "N" and userpassSec1 ="" then  ''구방식.
	    Call FnAddIISLOG("addlog=plogin&sub=no2step&uid="&userid) ''2016/12/29
		response.write("<script>window.alert('등록된 2차 비밀번호가 없습니다.새로 설정해주세요');</script>")
        response.write("<script>history.go(-1);</script>")
        dbget.close()	:	response.End
	end if

	AuthReqIP = IsPartnerAuthRequireIP(userid,GeoIpCCD,TRUE)

	'### 유저정보 접수 업체 전용으로 변경. 2017/04/24
	sql = "select top 1 A.id, A.company_name, A.tel, A.fax, A.url, A.email, A.userdiv, A.Enc_password, A.Enc_password64, A.groupid " & vbCrlf
	sql = sql & "	, A.level_sn " & vbCrlf
	sql = sql & "	, A.lastlogindt, isNULL(A.lastpwchgdt,'2001-01-01') as lastpwchgdt " & vbCrlf
	sql = sql & "	, isNULL(A.lastInfoChgDT,A.regdate) as lastInfoChgDT" & vbCrlf
	sql = sql & " 	, isNULL(A.Enc_2password64,'') as Enc_2password64"   ''수정 2017/04/10
	sql = sql & " from [db_partner].[dbo].tbl_partner as A " & vbCrlf
	sql = sql & " where A.id = '" & userid & "'" & vbCrlf
	sql = sql & " and A.isusing='Y'"
	rsget.CursorLocation = adUseClient
    rsget.Open sql,dbget,adOpenForwardOnly, adLockReadOnly

	if  not rsget.EOF  then
	    db_id_Exists       = TRUE
	    dbpassword  = rsget("Enc_password64")
    	dbpassword2  = rsget("Enc_2password64")
    	lastlogindt = rsget("lastlogindt")  ''최종 접속 성공일
    	lastpwchgdt = rsget("lastpwchgdt")  ''최종 패스워드 변경일
    	lastInfoChgDT = rsget("lastInfoChgDT")  ''최종 담당자정보  변경일

    	db_id = rsget("id")
    	db_userdiv = rsget("userdiv")
    	db_company_name = db2html(rsget("company_name"))
	    db_email = db2html(rsget("email"))
		db_groupid = rsget("groupid")

    end if
    rsget.close

    if (db_id_Exists) then
		'// 로그인 정보 확인
		if trim(UCase(dbpassword))=trim(UCase(Enc_userpass64)) then
			if vIsSec = "Y" then
				if userpassSec ="" then
					Call FnAddIISLOG("addlog=plogin&sub=no2nd&uid="&userid&"&ruid="&request("uid")) ''2016/12/29
					response.write("<script>window.alert('2차 비밀번호 값이 없습니다.확인해주세요');</script>")
			        response.write("<script>history.go(-1);</script>")
			        dbget.close()	:	response.End
				end if


				Enc_2userpass64 = SHA256(md5(userpassSec))


				if trim(UCase(dbpassword2))<>trim(UCase(Enc_2userpass64)) then

    			    if AuthNo<>"" then
    			    	Call AddPartnerLoginLogWithGeoIpCode (userid,"N",AuthNo,GeoIpCCD)
    			    elseif tokenSn<>"" then
    			        Call AddPartnerLoginLogWithGeoIpCode (userid,"N",tokenSn,GeoIpCCD)
    			    else
    			    	Call AddPartnerLoginLogWithGeoIpCode (userid,"N",RefCode,GeoIpCCD)
    			    end if
    		        Call FnAddIISLOG("addlog=plogin&sub=fail2nd&uid="&userid) ''2016/12/29
    		        response.write("<script>window.alert('2차 비밀번호가 틀렸습니다.확인후 다시 시도해주세요.');</script>")
    		        response.write("<script>history.go(-1);</script>")
    		        dbget.close()	:	response.End
				end if

				if (AuthReqIP) then
                    Call fn_plogin_AddIISLOG("addlog=plogin&sub=reqauth&uid="&userid)
                    Session.Contents.Remove("tmpUID")
  			        Session.Contents.Remove("tmpUPWD")

                    session("reauthUID") =  userid
                    response.write("<script>window.alert('기존 접속 환경과 다른 환경에서 로그인 하셨습니다. 인증 페이지로 이동합니다.');</script>")
                    response.write("<script>location.href='/login/reconfirmip.asp'</script>")
                    dbget.close() : response.End
              end if
			else
				''2차비번생성 더이상 이곳에서 않함. 2017/04/24
				Call FnAddIISLOG("addlog=plogin&sub=2ndfail2&uid="&userid) ''2016/12/29
				response.write("<script>window.alert('유입경로에 문제가 있습니다.');</script>")
		        response.write("<script>history.go(-1);</script>")
		        dbget.close()	:	response.End

''				Enc_2userpass64 = SHA256(md5(userpassSec1))
''
''				''2차비번 생성시 2017/04/10 1,2차 동일할 수 없음. by eastone
''				if (trim(UCase(Enc_userpass64))=trim(UCase(Enc_2userpass64))) then
''				    Call FnAddIISLOG("addlog=plogin&sub=dupp2nd&uid="&userid) ''2016/12/29
''    		        response.write("<script>window.alert('1,2차 비밀번호를 동일하게 설정할 수 없습니다.확인후 다시 시도해주세요.');</script>")
''    		        response.write("<script>history.go(-1);</script>")
''    		        dbget.close()	:	response.End
''				end if
''
''				dim objCmd,returnValue
''					Set objCmd = Server.CreateObject("ADODB.COMMAND")
''						With objCmd
''							.ActiveConnection = dbget
''							.CommandType = adCmdText
''							.CommandText = "{?= call db_partner.[dbo].[sp_Ten_partner_SetSecondPassWord]('"&userid&"',  '"&Enc_2userpass64&"' )}"
''							.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
''							.Execute, , adExecuteNoRecords
''							End With
''						    returnValue = objCmd(0).Value
''					Set objCmd = nothing
''
''				 if returnValue =-1 then
''				 	Call FnAddIISLOG("addlog=plogin&sub=exists2nd&uid="&userid) ''2016/12/29
''					response.write("<script>window.alert('2차 비밀번호가 이미 등록되어있습니다.');</script>")
''			        response.write("<script>history.go(-1);</script>")
''			        dbget.close()	:	response.End
''				 elseif returnValue =0 then
''				 	Call FnAddIISLOG("addlog=plogin&sub=2ndfail&uid="&userid) ''2016/12/29
''					response.write("<script>window.alert('2차 비밀번호 등록에 실패했습니다. 확인 후 다시 등록해주세요');</script>")
''			        response.write("<script>history.go(-1);</script>")
''			        dbget.close()	:	response.End
''			     else
''			        Call FnAddIISLOG("addlog=plogin&sub=2ndok&uid="&userid) ''2016/12/29
''			   	    response.write("<script>window.alert('2차 비밀번호가 등록되었습니다. ');</script>")
''				 end if
			end if

	    	isFirstConnect = isNULL(lastlogindt)

	    	if (isFirstConnect) then
	    	    isRequirePwdUp = true
	    	    isRequireInfoUp = true
	    	else
	    	    '' isRequirePwdUp = (datediff("d",lastpwchgdt,now())>91)
	    	    ''isRequirePwdUp = (datediff("d",lastpwchgdt,now())>91) and (datediff("d",lastlogindt,now())>0) '' 패스워드 최종변경일을 2014/07/15 부터 넣었으므로.. 우선 lastlogindt 조건넣음.
	    	    ''if (CLNG(db_userdiv)<10) then ''일단 직원
	    	    ''    isRequirePwdUp = (datediff("d",lastpwchgdt,now())>91)
	    	    ''end if

	    	    isRequirePwdUp = (datediff("d",lastpwchgdt,now())>91)  ''2017/04/13
	    	    isRequireInfoUp=  (datediff("d",lastInfoChgDT,now())>91)
	        end if

			Session.Contents.Remove("tmpUID")
  			Session.Contents.Remove("tmpUPWD")

	        session("ssBctId") = db_id
	        session("ssBctDiv") = db_userdiv

        	if isnull(db_company_name) then
        		session("ssBctCname") = "..."
        	else
        		session("ssBctCname") = db_company_name
        	end if

        	session("ssBctEmail") = db_email
			session("ssGroupid") = db_groupid

			'아이디저장
    		response.Cookies("PASave").domain = "10x10.co.kr"
        	response.cookies("PASave").Expires = Date + 30	'1개월간 쿠키 저장
    	    If saved_id = "o" Then
    	    	response.cookies("PASave")("SAVED_ID") = tenEnc(CStr(db_id))
    	    Else
    	    	response.cookies("PASave")("SAVED_ID") = ""
    	    End If


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

	    	Call FnAddIISLOG("addlog=plogin&sub=pass2nd&uid="&userid) ''2016/12/29
		else
		    ''로그저장(실패)
		    if AuthNo<>"" then
		    	Call AddPartnerLoginLogWithGeoIpCode (userid,"N",AuthNo,GeoIpCCD)
		    elseif tokenSn<>"" then
		        Call AddPartnerLoginLogWithGeoIpCode (userid,"N",tokenSn,GeoIpCCD)
		    else
		    	Call AddPartnerLoginLogWithGeoIpCode (userid,"N",RefCode,GeoIpCCD)
		    end if

	        Call FnAddIISLOG("addlog=plogin&sub=faillogin&uid="&userid) ''2016/12/29
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

		'// 계정없음
		Call FnAddIISLOG("addlog=plogin&sub=nouserid&uid="&userid) ''2016/12/29
	    response.write("<script>window.alert('계정이 활성화 되지 않았거나, 아이디 또는 비밀번호가 틀렸습니다.');</script>")
	    response.write("<script>history.go(-1);</script>")
	    dbget.close()	:	response.End
	end if


	'### 로그인 계정 구분별 처리 ------------------------------------------------------

	dim ssnTmpUIDPartner
    dim isNoComplexPwTxt : isNoComplexPwTxt = chkPasswordComplex_Len6Ver(userid,userpass)

	''강사임시
	dim cuseridv
	sql = "select top 1 * "
	sql = sql + " from [db_user].[dbo].tbl_user_c"
	sql = sql + " where userid = '" + userid + "'" + vbCrlf

	rsget.CursorLocation = adUseClient
    rsget.Open sql,dbget,adOpenForwardOnly, adLockReadOnly
	if  not rsget.EOF  then
		cuseridv = rsget("userdiv")
	end if
	rsget.close

	if (trim(UCase(dbpassword))=trim(UCase(Enc_userpass64))) AND (trim(UCase(dbpassword2))=trim(UCase(Enc_2userpass64)) ) then
	    if ((isFirstConnect) or (isRequirePwdUp) or (isNoComplexPwTxt<>"")) then

	        ssnTmpUIDPartner = session("ssBctId")

	        Session.Contents.RemoveAll()

	        Response.Cookies("partner").domain = "10x10.co.kr"
            Response.Cookies("partner") = ""
            Response.Cookies("partner").Expires = Date - 1

            session("ssnTmpUIDPartner")= ssnTmpUIDPartner   ''비번 변경시 사용 세션값 (partner)

    	    ''// 최초 로그인 성공시 패스워드 변경(2014/05/19 서동석)
    	    if (isFirstConnect) then
    	        response.write "<script language='javascript'>" &vbCrLf &_
    							"	alert('최초 로그인시 비밀번호를 변경하셔야 합니다. \n비밀번호 변경페이지로 이동합니다.');" &vbCrLf &_
    							"	self.location='/login/modifyPassword_partner.asp';" &vbCrLf &_
    							"</script>"
    			dbget.close()	:	response.End
    	    end if

    	    ''// N(3)개월간 비밀번호 변경안한경우 변경페이지로 이동
    	    if   (isRequirePwdUp) then
    	        response.write "<script language='javascript'>" &vbCrLf &_
    							"	alert('3개월 이상 비밀번호를 변경하지 않으셨습니다. \n비밀번호 변경페이지로 이동합니다.');" &vbCrLf &_
    							"	self.location='/login/modifyPassword_partner.asp';" &vbCrLf &_
    							"</script>"
    			dbget.close()	:	response.End
    	    end if


    		'// 비밀번호 강화 정책 시행(2008.12.12; 허진원)  'chkPasswordComplex => chkPasswordComplex_Len6Ver 2016/09/20
    		if chkPasswordComplex_Len6Ver(userid,userpass)<>"" then
    			response.write "<script language='javascript'>" &vbCrLf &_
    							"	alert('" & chkPasswordComplex_Len6Ver(userid,userpass) & "\n비밀번호 변경페이지로 이동합니다.');" &vbCrLf &_
    							"	self.location='/login/modifyPassword_partner.asp';" &vbCrLf &_
    							"</script>"
    			dbget.close()	:	response.End
    		end if
	    end if

	    response.Cookies("partner").domain = "10x10.co.kr"
	    response.Cookies("partner")("userid") = session("ssBctId")
	    response.Cookies("partner")("userdiv") = session("ssBctDiv")

		''강사 임시 (cuseridv="15") 도 이동 2016/06/23
		if (cuseridv="14") then
		        session("ssUserCDiv")=cuseridv ''2016/08/11
				response.write "<script language='javascript'>location.replace('"&manageUrl&"/lectureadmin/index.asp')</script>"
	        	dbget.close()	:	response.End
		end if


	    if (session("ssBctId")="10x10") then
	        ''사용안함.
	        session.Abandon
	        dbget.close()	:	response.End

	    ''직원Level
	    elseif (session("ssBctDiv")<=9) then
	        ''사용안함.
	        session.Abandon
	        response.write "<script language='javascript'>alert('사용할수 없는페이지입니다.');location.replace('/')</script>"
	        dbget.close()	:	response.End

	    	if (errMsg<>"") then
	            response.write "<script language='javascript'>alert('" & errMsg & "');</script>"
	        end if

	    	response.write "<script language='javascript'>location.replace('" & manageUrl & "/admin/index.asp')</script>"
	        dbget.close()	:	response.End
	    elseif (session("ssBctDiv")=999) then
	    	''제휴 업체 (yahoo, empas..)
	        response.write "<script language='javascript'>location.replace('" & manageUrl & "/company/index.asp')</script>"
	        dbget.close()	:	response.End
	    elseif (session("ssBctDiv")=9999) then
	    	''브랜드 업체

		    ''// N(3)개월간담당자 정보 변경안한경우 변경페이지로 이동
		    if (isRequireInfoUp) then
		   %>
	 		<script language='javascript'>
			 	alert('<%if datediff("d","2014-12-04",date())>90 then%>3개월 이상 담당자정보를 변경하지 않으셨습니다.<%else%>2015년 새해를 맞이하여 담당자 정보 업데이트를 요청 드립니다.<%end if%> \n담당자 정보 변경페이지로 이동합니다.');
			 	self.location='/login/modifyManagerInfo.asp'
			 </script>
			 <%
				dbget.close()	:	response.End
		    end if
	    	response.write "<script language='javascript'>location.replace('" & manageUrl & "/partner/index.asp')</script>"
	        dbget.close()	:	response.End
	    elseif (session("ssBctDiv")=9000) then
	    	''강사 업체
	    	response.write "<script language='javascript'>location.replace('" & manageUrl & "/lectureradmin/index.asp')</script>"
	        dbget.close()	:	response.End
	    elseif (session("ssBctDiv")=501) or (session("ssBctDiv")=502) or (session("ssBctDiv")=503) or (session("ssBctDiv")=509) then

	    	response.write "<script language='javascript'>location.replace('" & manageUrl & "/offshop/index.asp')</script>"
	        dbget.close()	:	response.End
	    elseif (session("ssBctDiv")=101) or (session("ssBctDiv")=111) or (session("ssBctDiv")=112) or (session("ssBctDiv")=201) or (session("ssBctDiv")=301) then

	    	response.write "<script language='javascript'>location.replace('" & manageUrl & "/admin/index.asp')</script>"
	        dbget.close()	:	response.End
	    else
	        session.Abandon
	        response.write "<script language='javascript'>alert('권한이없습니다.');location.replace('/')</script>"
	        dbget.close()	:	response.End
	    end if
	else
	    session.Abandon
        response.write "<script language='javascript'>alert('권한이없습니다.-2');location.replace('/')</script>"
        dbget.close()	:	response.End
	end if
else


	'### 1차 로그인 확인
	sql = "select top 1 A.id,   A.Enc_password, A.Enc_password64, A.groupid,  A.Enc_2password64 " & vbCrlf
	sql = sql & " from [db_partner].[dbo].tbl_partner as A " & vbCrlf
	sql = sql & " where A.id = '" & userid & "'" & vbCrlf
	sql = sql & " and A.isusing='Y'"
	sql = sql & " and A.userdiv>10"  ''2017/04/21 추가 직원은 이쪽으로 불가..

	rsget.CursorLocation = adUseClient
    rsget.Open sql,dbget,adOpenForwardOnly, adLockReadOnly
	if  (not rsget.EOF)  then
	    db_id_Exists       = TRUE
	    db_Enc_password64  = rsget("Enc_password64")
	    db_Enc_2password64 = rsget("Enc_2password64")
	end if
	rsget.Close

	if (db_id_Exists) then
		'// 로그인 정보 확인
        if (rtrim(UCase(db_Enc_password64))=trim(UCase(Enc_userpass64))) then
            ''dbpassword  = db_Enc_password64  '''?

            if isNull(db_Enc_2password64) or (db_Enc_2password64="") then
            	isdbpassword_sec= "N"
            else
            	isdbpassword_sec= "Y"
            end if

		  ''2차패스워드가 빈값인경우 로그인 못함. 2017/04/11
		    if (isdbpassword_sec="N") then
    		    if (Is2ndPwdNotExistsReject(userid)) then
                    Call FnAddIISLOG("addlog=plogin&sub=2ndpassnull&uid="&userid)
                    response.write("<script>window.alert('2차비밀번호 설정후 로그인해주시기 바랍니다.');</script>")
    		        response.write("<script>history.go(-1);</script>")
    		        dbget.close() : response.End
    		    end if
		    end if

		  ''장기간 로그인 안한 경우
		    if (IsLongTimeNotLoginUserid(userid)) then
		        Call FnAddIISLOG("addlog=plogin&sub=logntimenosee&uid="&userid)
                response.write("<script>window.alert('장기간 사용하지 않아 계정이 잠겼습니다.\n비밀번호 찾기를 통해 인증번호 수신후 계정을 활성화 시켜 주시기 바랍니다.');</script>")
		        response.write("<script>history.go(-1);</script>")
		        dbget.close() : response.End
		    end if


		    '임시 세션 생성
		    session("tmpUID") =  userid
		    session("tmpUPWD") = userpass

		%>
			<form name="frmLogin" method="post" action="<%=getSCMSSLURL%>/login/loginS.asp">
				<input type="hidden" name="chkAuth" value="Y">
				<input type="hidden" name="hidSec" value="<%=isdbpassword_sec%>">
				<input type="hidden" name="saved_id" value="<%=saved_id%>">
			</form>
		<%
		      Call FnAddIISLOG("addlog=plogin&sub=pass1st&uid="&userid) ''2016/12/29
		      response.write("<script>document.frmLogin.submit();</script>")
		      dbget.close()	:	response.End
        else
            ''로그저장(실패)
		    if AuthNo<>"" then
		    	Call AddPartnerLoginLogWithGeoIpCode (userid,"N",AuthNo,GeoIpCCD)
		    elseif tokenSn<>"" then
		        Call AddPartnerLoginLogWithGeoIpCode (userid,"N",tokenSn,GeoIpCCD)
		    else
		    	Call AddPartnerLoginLogWithGeoIpCode (userid,"N",RefCode,GeoIpCCD)
		    end if

	        Call FnAddIISLOG("addlog=plogin&sub=faillogin1st&uid="&userid) ''2016/12/29
	        response.write("<script>window.alert('계정이 활성화 되지 않았거나, 아이디 또는 비밀번호가 틀렸습니다.');</script>")
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

		'// 계정없음
		Call FnAddIISLOG("addlog=plogin&sub=nouserid1st&uid="&userid) ''2016/12/29
	    response.write("<script>window.alert('계정이 활성화 되지 않았거나, 아이디 또는 비밀번호가 틀렸습니다.');</script>")
	    response.write("<script>history.go(-1);</script>")
	    dbget.close()	:	response.End
	end if
	response.end
end if
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
