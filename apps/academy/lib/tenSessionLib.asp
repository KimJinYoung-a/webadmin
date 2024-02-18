<%
''session library 2016/11
dim C_MaxSessionTimedOUT : C_MaxSessionTimedOUT = 60*60*12    ''2Hour 이상 적당.  세션이 날라갔을경우 쿠키로 세션을 복구할 시간 (웹서버 세션시간보다 커야함..)
dim C_ssnUpdateReCycleTime : C_ssnUpdateReCycleTime = 60*10  ''10~20분사이 적당할듯. 디비 체크 및 세션 업데이트 체크 주기 C_MaxSessionTimedOUT 보다 작아야.
Dim GG_ACA_APP_CON_NAME : GG_ACA_APP_CON_NAME = "db_academy"

IF (application("Svr_Info")="Dev") then   ''TEST
    C_MaxSessionTimedOUT = 60*60*1
    C_ssnUpdateReCycleTime = 60 ''60*20
    
    '' response.write C_MaxSessionTimedOUT
    '' response.write fnDBSessionExpire("icommang","20161115211750")
    '' response.write fnCheckDBsessionUpdate("icommang","20161115215614","20161115225614",3600)
end if

if NOT (LCASE(request.servervariables("SCRIPT_NAME"))="/apps/protov1/loginproc.asp") then ''로그인페이지는 검사안함.
    CALL fnChkDBSessionUpdate()  ''세션 검사후 업데이트
end if


function fnDateTimeToLongTime(icookieLoginDt)
    dim iorginDt : iorginDt = icookieLoginDt
    iorginDt = CDate(iorginDt)
    
    fnDateTimeToLongTime = Year(iorginDt)&Right("00"&Month(iorginDt),2)&Right("00"&Day(iorginDt),2)&Right("00"&Hour(iorginDt),2)&Right("00"&Minute(iorginDt),2)&Right("00"&Second(iorginDt),2)
end function

function fnLongTimeToDateTime(ilongTime)
    dim iorgDt : iorgDt= ilongTime
    if LEN(ilongTime)<>14 then Exit function
        
    fnLongTimeToDateTime = CDate(LEFT(ilongTime,4)&"-"&MID(ilongTime,5,2)&"-"&MID(ilongTime,7,2)&" "&MID(ilongTime,9,2)&":"&MID(ilongTime,11,2)&":"&MID(ilongTime,13,2))
end function

''디비 세션 생성 log-on
function fnDBSessionCreate(ilgnchannel)
    dim ssnuserid  : ssnuserid =  session("ssnuserid")
    dim ssnlogindt : ssnlogindt = session("ssnlogindt")
    
    if (ssnuserid="") or (ssnlogindt="") then Exit function
        
    dim iSsnCon : set iSsnCon = CreateObject("ADODB.Connection")
    Dim cmd : set cmd = server.CreateObject("ADODB.Command")
    
    dim sqlStr
    sqlStr = "db_academy.[dbo].[sp_ACA_SSN_CREATE]"
    
    iSsnCon.Open Application(GG_ACA_APP_CON_NAME) ''커넥션 스트링.
    cmd.ActiveConnection = iSsnCon
    cmd.CommandText = sqlStr
    cmd.CommandType = adCmdStoredProc
    
    cmd.Parameters.Append cmd.CreateParameter("@ssnuserid", adVarchar, adParamInput, 32, ssnuserid)
    cmd.Parameters.Append cmd.CreateParameter("@ssnlogindt", adVarchar, adParamInput, 14, ssnlogindt)
    cmd.Parameters.Append cmd.CreateParameter("@lgnchannel", adVarchar, adParamInput, 1, ilgnchannel)
   
    cmd.Execute 
    set cmd = Nothing
    iSsnCon.Close
    SET iSsnCon = Nothing
    
end function

''디비세션 날림 log-off시
function fnDBSessionExpire()
    dim ssnuserid  : ssnuserid =  session("ssnuserid")
    dim ssnlogindt : ssnlogindt = session("ssnlogindt")
    
    if (ssnuserid="") or (ssnlogindt="") then Exit function
    
    dim iSsnCon : set iSsnCon = CreateObject("ADODB.Connection")
    Dim cmd : set cmd = server.CreateObject("ADODB.Command")
    dim intResult
    
    dim sqlStr
    sqlStr = "db_academy.[dbo].[sp_ACA_SSN_EXPIRE]"
    
    iSsnCon.Open Application(GG_ACA_APP_CON_NAME) ''커넥션 스트링.
    cmd.ActiveConnection = iSsnCon
    cmd.CommandText = sqlStr
    cmd.CommandType = adCmdStoredProc
    
    cmd.Parameters.Append cmd.CreateParameter("returnValue", adInteger, adParamReturnValue)
    cmd.Parameters.Append cmd.CreateParameter("@ssnuserid", adVarchar, adParamInput, 32, ssnuserid)
    cmd.Parameters.Append cmd.CreateParameter("@ssnlogindt", adVarchar, adParamInput, 14, ssnlogindt)
    cmd.Execute 
    
    intResult = cmd.Parameters("returnValue").Value
    set cmd = Nothing
    iSsnCon.Close
    SET iSsnCon = Nothing
    
    fnDBSessionExpire = (intResult>0)
end function

function fnCheckDBsessionUpdate(icookieUserID,icookieSsnDt,inowSsnDt,iMaxSessionTimedOUT,byRef idbssnlogindt)
    fnCheckDBsessionUpdate = false
    if (icookieUserID="") or (icookieSsnDt="") or (inowSsnDt="") then Exit function
    
    dim iSsnCon : set iSsnCon = CreateObject("ADODB.Connection")
    Dim cmd : set cmd = server.CreateObject("ADODB.Command")
    dim intResult
    
    dim sqlStr
    sqlStr = "db_academy.[dbo].[sp_ACA_SSN_CHECKNUPDATE]"
    
    iSsnCon.Open Application(GG_ACA_APP_CON_NAME) ''커넥션 스트링.
    cmd.ActiveConnection = iSsnCon
    cmd.CommandText = sqlStr
    cmd.CommandType = adCmdStoredProc
    
    cmd.Parameters.Append cmd.CreateParameter("returnValue", adInteger, adParamReturnValue)
    cmd.Parameters.Append cmd.CreateParameter("@ssnuserid", adVarchar, adParamInput, 32, icookieUserID)
    cmd.Parameters.Append cmd.CreateParameter("@ssnlogindt", adVarchar, adParamInput, 14, icookieSsnDt)
    cmd.Parameters.Append cmd.CreateParameter("@updatelogindt", adVarchar, adParamInput, 14, inowSsnDt)
    cmd.Parameters.Append cmd.CreateParameter("@ssntimeoutScond", adInteger, adParamInput, , iMaxSessionTimedOUT)
    cmd.Parameters.Append cmd.CreateParameter("@retdbssnlogindt", adVarchar, adParamOutput, 14, "")
    cmd.Execute
    
    intResult = cmd.Parameters("returnValue").Value
    idbssnlogindt = cmd.Parameters("@retdbssnlogindt").Value
    
    set cmd = Nothing
    iSsnCon.Close
    SET iSsnCon = Nothing
    
    fnCheckDBsessionUpdate = (intResult>0)
end function

function fnChkDBSessionUpdate()
    dim cookieUserID    : cookieUserID = request.cookies("partner")("userid")         ''로그인시 값과 동일해야함.
    dim cookieSsnDt     : cookieSsnDt = request.Cookies("partner")("ssndt")             ''로그인시 값과 동일해야함.  partner 로변경.
    dim isReqSSnUp      : isReqSSnUp = false
    dim isDbssnExists   : isDbssnExists = false
    
    ''if (cookieUserID="") or (cookieSsnDt="") then Exit function
    '' cookieSsnDt 없으면 expired  2016/12/16
    if (cookieUserID="") then Exit function 
    
    dim nowDateTime     : nowDateTime=now()
    dim cookieDateTime  : cookieDateTime=fnLongTimeToDateTime(cookieSsnDt)
    
    dim ssnuserid  : ssnuserid  = session("ssnuserid")
    dim ssnlogindt : ssnlogindt = session("ssnlogindt")
    dim ssnlastcheckdt : ssnlastcheckdt = session("ssnlastcheckdt")
    dim ssnlastcheckDateTime : ssnlastcheckDateTime=fnLongTimeToDateTime(ssnlastcheckdt)
    dim nowSsnDt, dbssnlogindt
    dim isSessionExists : isSessionExists=FALSE
    

    ''세션이존재하고 최종업데이트 시간이 C_ssnUpdateReCycleTime 보다 크면 업데이트. (너무 자주업데이트 하지 않도록)
    if (cookieUserID<>"") and (LCASE(cookieUserID)=LCase(ssnuserid)) then
        if (ssnlogindt=cookieSsnDt) then
            isSessionExists = true
            isReqSSnUp = datediff("s",ssnlastcheckDateTime,nowDateTime)>C_ssnUpdateReCycleTime
        else    ''cookieSsnDt 없는경우 등. 2016/12/15 수정.
            isReqSSnUp = TRUE                   
        end if
    end if
    
    ''비정상적으로 세션이 날라갔을경우.  C_MaxSessionTimedOUT 보다 작은경우에 한해 DB에서 체크 후 세션 업데이트함.
    ''수정 세션이 없으면 무조건 체크.
    if (cookieUserID<>"") and (ssnuserid="") then
        ''isReqSSnUp = datediff("s",cookieDateTime,nowDateTime)<C_MaxSessionTimedOUT
        isReqSSnUp = true
    elseif (cookieUserID<>"") and (LCASE(cookieUserID)<>LCase(ssnuserid)) then   ''2017/05/18 수정 세션이 다르면 Expire **
        Call CookieSessionExpire()
        Exit function
    end if

    if (isReqSSnUp) then
        ''DB에 값이 있는지 체크.

        nowSsnDt = fnDateTimeToLongTime(nowDateTime)
        isDbssnExists = fnCheckDBsessionUpdate(cookieUserID,cookieSsnDt,nowSsnDt,C_MaxSessionTimedOUT,dbssnlogindt)
		    'Response.write nowSsnDt & "<br>"
	'Response.write isDbssnExists
	'Response.end
		
        if (isDbssnExists) then ''세션 업데이트.
            if (NOT isSessionExists) then
                session("ssnuserid") = cookieUserID
                session("ssnlogindt") = dbssnlogindt    ''기존 세션에 있는 
            END IF
            session("ssnlastcheckdt") = nowSsnDt    ''다시체크를 위해업데이트
        else
            ''쿠키 /세션 날림.
            Call CookieSessionExpire()
        end if
    end if
end function

function fnIsSessionCookieValid()
    ''쿠키가 있는경우만. 비로그인은 제외
    if (request.cookies("partner")("userid")<>"") then
        fnIsSessionCookieValid = (LCASE(request.cookies("partner")("userid"))=LCASE(session("ssnuserid")))
    else
        fnIsSessionCookieValid = true    
    end if
end function

function CookieSessionExpire()
    ''log-out
    response.Cookies("partner").domain = "10x10.co.kr"
    response.Cookies("partner") = ""
    response.Cookies("partner").Expires = Date - 1
    
    'response.Cookies("etc").domain = "10x10.co.kr"
    'response.Cookies("etc") = ""
    'response.Cookies("etc").Expires = Date - 1

    session.abandon
    
    ''addLog 추가 로그 //2016/12/16
    dim iAddLogs
    iAddLogs = "r=snexpire"
    if (request.ServerVariables("QUERY_STRING")<>"") then iAddLogs="&"&iAddLogs
    response.AppendToLog iAddLogs

end function
%>