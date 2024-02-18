<!-- #include virtual="/lib/util/aspJSON1.17.asp" -->
<%

''로그인후 호출
'Call fn_RDS_SSN_SET()

''세션검증 필요한페이지에서 호출
'Call fn_RDS_CHK_SSN_RESTORE()

''로그아웃페이지에서 호출
'Call fn_RDS_SSN_Expire()

'' 쿠키값 두개가 필요하다
'' 1.redis 키값용,  2.userid 검증 해시용, 쿠키ex: SCMSSN:XXXXXXXXXX:YYYYYYYY

DIM GG_RDS_APIURL : GG_RDS_APIURL = "https://dapi.10x10.co.kr"  ''"http://52.79.73.177"  
DIM GG_RDS_SSNEXPIRETIME : GG_RDS_SSNEXPIRETIME = 3600 '' default 3600 sec
DIM GG_RDS_AUTHKEY : GG_RDS_AUTHKEY = "key=lEejpMoNDt1GYzODrwlcwMEDqidUkHdskioU7Tl3bdVeMXNFS13xJimboxKx"
DIM GG_RDS_SSNKEYPREFIX : GG_RDS_SSNKEYPREFIX = "SCMSSN"
DIM GG_RDS_COOKIE_KEYNAME : GG_RDS_COOKIE_KEYNAME = "SCMSSNRDS"
DIM GG_RDS_SSN_SALT : GG_RDS_SSN_SALT = "TBT"
DIM GG_RDS_SESSION_CHECKTIMENAME : GG_RDS_SESSION_CHECKTIMENAME = "ssPreChecktime"
DIM GG_RDS_DEBUG : GG_RDS_DEBUG = FALSE

if (application("Svr_Info")="Dev") then 
    GG_RDS_APIURL = "http://52.78.37.37"
    GG_RDS_AUTHKEY = "key=apikey3"
    GG_RDS_COOKIE_KEYNAME = GG_RDS_COOKIE_KEYNAME&"DEV"
    GG_RDS_SSNEXPIRETIME = 3600
    GG_RDS_DEBUG = TRUE
end if

function fn_RDS_debugWrite(istr)
    if (GG_RDS_DEBUG) then
        response.write istr&"<br>"
    end if
end function

function fn_RDS_CHK_getExpTimeDiff()
    dim iexpTimeDiff : iexpTimeDiff = CLNG(GG_RDS_SSNEXPIRETIME / 10)
    if (iexpTimeDiff=0) then iexpTimeDiff = 3600/10
    if (application("Svr_Info")="Dev") then iexpTimeDiff=60

    fn_RDS_CHK_getExpTimeDiff = iexpTimeDiff
end function

''-- SH256 ----------------------------------------------------
function fn_GG_RDS_sha256hashBytes(aBytes)
    Dim sha256
    set sha256 = CreateObject("System.Security.Cryptography.SHA256Managed")

    sha256.Initialize()
    'Note you MUST use computehash_2 to get the correct version of this method, and the bytes MUST be double wrapped in brackets to ensure they get passed in correctly.
    fn_GG_RDS_sha256hashBytes = sha256.ComputeHash_2( (aBytes) )

    set sha256 = Nothing
end function

function fn_GG_RDS_stringToUTFBytes(aString)
    Dim UTF8
    Set UTF8 = CreateObject("System.Text.UTF8Encoding")
    fn_GG_RDS_stringToUTFBytes = UTF8.GetBytes_4(aString)
    SET UTF8 = Nothing
end function

function fn_GG_RDS_bytesToHex(aBytes)
    dim hexStr, x
    for x=1 to lenb(aBytes)
        hexStr= hex(ascb(midb( (aBytes),x,1)))
        if len(hexStr)=1 then hexStr="0" & hexStr
        fn_GG_RDS_bytesToHex=fn_GG_RDS_bytesToHex & hexStr
    next
end function

''SHA256
function fn_GG_RDS_SHA256(ostr)
    fn_GG_RDS_SHA256 = fn_GG_RDS_bytesToHex(fn_GG_RDS_sha256hashBytes(fn_GG_RDS_stringToUTFBytes(ostr)))
end function
''-- SH256 ----------------------------------------------------

function fn_RDS_getCookieDomain()
    Dim iCookieDomain : iCookieDomain = "10x10.co.kr"
    IF application("Svr_Info")="Dev" THEN
        if (request.ServerVariables("LOCAL_ADDR")="::1") or (request.ServerVariables("LOCAL_ADDR")="127.0.0.1") then
            iCookieDomain = "localhost"
        end if
    End if
    fn_RDS_getCookieDomain = iCookieDomain
end function

function fn_RDS_SSN_MakeKey()
    fn_RDS_SSN_MakeKey = ""
    
    dim iuserid : iuserid = LCASE(session("ssBctId"))

    if LEN(iuserid)<1 then Exit function

    Randomize()
    Dim rndVal : rndVal = CLNG((Rnd * 9000000) + 1)

    Dim icookieKey 
    icookieKey = fn_GG_RDS_SHA256(iuserid&rndVal)&":"&fn_GG_RDS_SHA256(iuserid&GG_RDS_SSN_SALT)

    Dim icookieDomain : icookieDomain = fn_RDS_getCookieDomain()
    response.Cookies(GG_RDS_COOKIE_KEYNAME).domain = iCookieDomain
    response.Cookies(GG_RDS_COOKIE_KEYNAME) = icookieKey

    dim retVal 
    retVal = GG_RDS_SSNKEYPREFIX&":"&icookieKey

    fn_RDS_SSN_MakeKey = retVal
end function


function fn_RDS_SSN_KeyGet()
    fn_RDS_SSN_KeyGet = ""

    Dim icookieKey : icookieKey = request.Cookies(GG_RDS_COOKIE_KEYNAME)
    if (LEN(icookieKey)<1)  then exit function

    dim retVal 
    retVal = GG_RDS_SSNKEYPREFIX&":"&icookieKey

    fn_RDS_SSN_KeyGet = retVal
end function


function fn_RDS_SSN_MakeBodyJson()
    Dim retBody

    Dim iexprTime : iexprTime = GG_RDS_SSNEXPIRETIME
    Dim issnKey : issnKey = fn_RDS_SSN_MakeKey()
    Dim issnVal : issnVal = fn_RDS_SSN_serializeSessionToJson() ''fn_RDS_SSN_serializeSessionToStr()

    fn_RDS_SSN_MakeBodyJson = ""
    if (LEN(issnKey)<1 or LEN(issnVal)<1) then Exit function

    retBody = "{"
    retBody = retBody & """key"": """&issnKey&""","
    retBody = retBody & """value"": """&issnVal&""","
    retBody = retBody & """expirationTime"": "&iexprTime
    retBody = retBody & "}"

    fn_RDS_SSN_MakeBodyJson = retBody
end function



'' Async 잘 안됨..
' function X_fn_RDS_SSN_SET_ASync()
'     Dim objXML
'     Dim iBody 
'     iBody = fn_RDS_SSN_MakeBodyJson()
'     if LEN(iBody)<1 then Exit function ''쿠키 값이 없거나 

'     Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")

'     objXML.open "POST", "" & GG_RDS_APIURL&"/api/RedisValues/", True ''Ascync true
' 	objXML.setRequestHeader "Authorization", GG_RDS_AUTHKEY
' 	objXML.setRequestHeader "CONTENT-TYPE", "application/json"
' 	objXML.send(iBody)

'     ''objXML.WaitForResponse()  ''이방식은 Sync와 같음.
'     '' Async 가 되었다안되었다함. script가 끝나기전에 POST가 완료되었는지 알수 없음.
' end function

'' Sync방식
function fn_RDS_SSN_SET()
    Dim objXML
    Dim iBody 
    iBody = fn_RDS_SSN_MakeBodyJson()
    if LEN(iBody)<1 then Exit function ''쿠키 값이 없거나 

    Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")

    objXML.open "POST", "" & GG_RDS_APIURL&"/api/RedisValues/", False
    objXML.SetTimeouts 10*1000, 10*1000, 10*1000, 10*1000  ''조금 짧게 잡자.
	objXML.setRequestHeader "Authorization", GG_RDS_AUTHKEY
	objXML.setRequestHeader "CONTENT-TYPE", "application/json"
    On Error Resume Next
	objXML.send(iBody)
    On Error Goto 0
    session(GG_RDS_SESSION_CHECKTIMENAME) = Now()
    SET objXML = Nothing 
end function

function fn_RDS_SSN_serializeSessionToJson()
    ''타입 구분을 잘해야 String인지 Long 인지.
    ' response.write "ssBctDiv:"&TypeName(session("ssBctDiv"))
    ' response.write "ssAdminPsn:"&TypeName(session("ssAdminPsn"))
    ' response.write "ssAdminLsn:"&TypeName(session("ssAdminLsn"))
    ' response.write "ssAdminPOsn:"&TypeName(session("ssAdminPOsn"))
    ' response.write "ssAdminCLsn:"&TypeName(session("ssAdminCLsn"))
    ' response.end

    Dim retData
    retData = "{"
    retData = retData & "\""ssBctId\"":\"""&Trim(session("ssBctId"))&"\"""        '로그인 아이디
    retData = retData & ",\""ssBctDiv\"":"&Trim(session("ssBctDiv"))&""      '회원구분
    retData = retData & ",\""ssBctBigo\"":\"""&Trim(session("ssBctBigo"))&"\"""    '매장 추가 정보
    retData = retData & ",\""ssBctSn\"":\"""&Trim(session("ssBctSn"))&"\"""        '직원번호
    retData = retData & ",\""ssBctCname\"":\"""&Trim(session("ssBctCname"))&"\"""  '직원 이름
	retData = retData & ",\""ssBctEmail\"":\"""&Trim(session("ssBctEmail"))&"\"""  '직원 이메일
    retData = retData & ",\""ssGroupid\"":\"""&Trim(session("ssGroupid"))&"\"""    '그룹 코드
    retData = retData & ",\""ssAdminPsn\"":"&Trim(session("ssAdminPsn"))&""   '부서 번호
    retData = retData & ",\""ssAdminLsn\"":"&Trim(session("ssAdminLsn"))&""   '등급 번호
    retData = retData & ",\""ssAdminPOsn\"":"&Trim(session("ssAdminPOsn"))&""   '직책 번호
    retData = retData & ",\""ssAdminPOSITsn\"":"&Trim(session("ssAdminPOSITsn"))&""   '직급 번호
    retData = retData & ",\""ssAdminCLsn\"":"&Trim(session("ssAdminCLsn"))&""   '개인정보 취급권한
    retData = retData & ",\""sslgnMethod\"":\"""&Trim(session("sslgnMethod"))&"\"""   'SMS인증여부
    retData = retData & ",\""ssAdminlv1customerYN\"":\"""&Trim(session("ssAdminlv1customerYN"))&"\"""   ' 개인정보취급권한(고객정보)
    retData = retData & ",\""ssAdminlv2partnerYN\"":\"""&Trim(session("ssAdminlv2partnerYN"))&"\"""   ' 개인정보취급권한(파트너정보)
    retData = retData & ",\""ssAdminlv3InternalYN\"":\"""&Trim(session("ssAdminlv3InternalYN"))&"\"""   ' 개인정보취급권한(내부정보)
    retData = retData & ",\""WEHAGO_access_token\"":\"""&Trim(session("WEHAGO_access_token"))&"\"""     ' 처음 위하고 접속시 토큰값을 저장하고 그 토큰으로 8시간 통신을 한다. 토큰이 없을때만 로그인해서 토큰 받아옴.
    retData = retData & ",\""WEHAGO_state\"":\"""&Trim(session("WEHAGO_state"))&"\"""
    retData = retData & ",\""WEHAGO_wehago_id\"":\"""&Trim(session("WEHAGO_wehago_id"))&"\"""
    retData = retData & ",\""WEHAGO_time\"":\"""&Trim(session("WEHAGO_time"))&"\"""
    retData = retData & "}"

    fn_RDS_SSN_serializeSessionToJson = retData
end function


function fn_RDS_SSN_deSerializeJsonToSession(irediskey, ival)
    if Len(irediskey)<1 then Exit function
    if Len(ival)<1 then Exit function
    if LCASE(ival)="null" then Exit function

    Dim i
    Dim Is_SsBctId_OK : Is_SsBctId_OK = FALSE
    Dim mayBctId, mayBctIdHash, mayBctIdHashArr
    mayBctIdHashArr = split(irediskey,":")
    if isArray(mayBctIdHashArr) then
        if UBOUND(mayBctIdHashArr)>=2 then mayBctIdHash = mayBctIdHashArr(2)
    end if

    Dim jsonObj, jsonItemKey
    set jsonObj = New aspJson
        jsonObj.loadJSON(ival)


        for each jsonItemKey in jsonObj.data
            if (jsonItemKey="ssBctId") then
                mayBctId = jsonObj.data(jsonItemKey)
                Is_SsBctId_OK = (fn_GG_RDS_SHA256(LCase(mayBctId)&GG_RDS_SSN_SALT)=mayBctIdHash)

                if NOT (Is_SsBctId_OK) then
                    Call fn_RDS_debugWrite("session diff:"&mayBctId&":"&mayBctIdHash)
                end if
                ''세션id 값이 다르다.
            end if
        Next

        if (Is_SsBctId_OK) then 
            for each jsonItemKey in jsonObj.data
                'Call fn_RDS_debugWrite(jsonItemKey&":"&jsonObj.data(jsonItemKey)&":"&TypeName(jsonObj.data(jsonItemKey)))
                if (TypeName(jsonObj.data(jsonItemKey))="Double") then
                    session(jsonItemKey) = CLNG(jsonObj.data(jsonItemKey))
                else
                    session(jsonItemKey) = jsonObj.data(jsonItemKey)

                    if isNULL(session(jsonItemKey)) then session(jsonItemKey)=""
                end if
            Next
        end if
    Set jsonObj = Nothing

end function


function fn_RDS_SSN_GET()
    Dim objXML
    Dim retJson, maykey, mayval
    Dim jsonObj, oJSONoutput

    Dim iredisKey : iredisKey = fn_RDS_SSN_KeyGet()
    if LEN(iredisKey)<1 then Exit function
    Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
    ''setTimeouts (long resolveTimeout, long connectTimeout, long sendTimeout, long receiveTimeout)
    objXML.open "GET", "" & GG_RDS_APIURL&"/api/RedisValues/"&iredisKey, False 
    objXML.SetTimeouts 10*1000, 10*1000, 10*1000, 10*1000
    objXML.setRequestHeader "Authorization", GG_RDS_AUTHKEY
	objXML.setRequestHeader "CONTENT-TYPE", "application/json"

    On Error Resume Next
    objXML.send()
    If (Err) then Exit function
    On Error Goto 0
 
    if (objXML.Status = "200") then
        retJson = TRIM(objXML.responseText)

        set jsonObj = New aspJSON
            jsonObj.loadJSON(retJson)

            maykey = jsonObj.data("key")
            mayval = jsonObj.data("value")

            if (iredisKey=maykey) and Not isNULL(mayval) then
                Call fn_RDS_SSN_deSerializeJsonToSession(maykey,mayval)
            end if
        Set jsonObj = Nothing
    else
        '' ERR - nothing do '' 장애시 내부 세션으로 쓰자.
    end if
    SET objXML = Nothing
end function


function fn_RDS_SSN_Expire()
    Dim iredisKey : iredisKey = fn_RDS_SSN_KeyGet()
    Dim objXML
    if LEN(iredisKey)>0 then 
        Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")

        objXML.open "DELETE", "" & GG_RDS_APIURL&"/api/RedisValues/"&iredisKey, False  ''Ascync true
        objXML.SetTimeouts 10*1000, 10*1000, 10*1000, 10*1000
        objXML.setRequestHeader "Authorization", GG_RDS_AUTHKEY
        objXML.setRequestHeader "CONTENT-TYPE", "application/json"

        On Error Resume Next
        objXML.send()
        On Error Goto 0
    end if

    Dim icookieDomain : icookieDomain = fn_RDS_getCookieDomain()
    response.Cookies(GG_RDS_COOKIE_KEYNAME).domain = iCookieDomain
    response.Cookies(GG_RDS_COOKIE_KEYNAME).Expires = Date - 1

    ''session.abandon  ''호출 페이지에서 처리하자
    SET objXML = Nothing
end function


function fn_RDS_CHK_SSN_RESTORE()
    Dim iredisKey : iredisKey = fn_RDS_SSN_KeyGet()
    if LEN(iredisKey)<1 then Exit function  ''쿠키가 없으면 SKIP

    Dim preCheckTime : preCheckTime = session(GG_RDS_SESSION_CHECKTIMENAME)
    Dim isReqCheck : isReqCheck = False
    Dim expCheckTimeDiff : expCheckTimeDiff = fn_RDS_CHK_getExpTimeDiff()

    if (session("ssBctId")="") then
        isReqCheck = True
    elseif (preCheckTime="") then
        isReqCheck = True
    else
        if NOT isDate(preCheckTime) then
            isReqCheck = True
        else
            isReqCheck = (datediff("s",preCheckTime,now())>expCheckTimeDiff) 
        end if
    end if

    if (isReqCheck) then
        Call fn_RDS_debugWrite("diffTime:"&datediff("s",preCheckTime,now()))
        Call fn_RDS_SSN_GET()
        session(GG_RDS_SESSION_CHECKTIMENAME) = Now() ''최종체크시각
    end if

    '' id세션만 있고 레벨SN 등이 없는경우.
    if (session("ssBctId")<>"") then
        if (isNULL(session("ssBctId")) or isNULL(session("ssBctDiv")) or (session("ssBctDiv")="") or isNULL(session("ssAdminLsn")) or (session("ssAdminLsn")="") or isNULL(session("ssAdminPsn")) or (session("ssAdminPsn")="")) then
            session("ssBctId")=""
            session.abandon
            Call fn_RDS_SSN_Expire()
        end if
    end if
end function

%>