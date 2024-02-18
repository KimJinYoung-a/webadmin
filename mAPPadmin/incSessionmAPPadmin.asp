<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%

''로그인 체크 :
'' 1. 세션으로 로그인 체크하지 않음 (안드로이드 세션 살아 있음?)
'' 2. 쿠키 expire 시간의 1/2 지난경우 쿠키 재생성
'' 3. 쿠키 검증은  session("mAppBctTstp")<>glb_tstp

Dim glb_isLogined : glb_isLogined = false
Dim glb_encuid, glb_userid, glb_tstp, glb_enckey, glb_appid, glb_devicekey


function checkLoginCookie()
    checkLoginCookie = false
    glb_encuid      = request.Cookies("mAppADM")("encuid")   ''= userid
    glb_tstp        = request.Cookies("mAppADM")("tstp")     ''= tstp
    glb_devicekey   = request.Cookies("mAppADM")("dvkey")    ''= devicekey
    glb_appid       = request.Cookies("mAppADM")("appid")    ''= appid
    glb_enckey      = request.Cookies("mAppADM")("enckey")   ''= devicekey&userid&tstp

    if (glb_encuid="") or (glb_devicekey="") or (glb_tstp="") or (glb_appid="") or (glb_enckey="") then Exit function

    ''검증

    ''쿠키 Expire 타임 1/2 이후 로그인 했으면 다시 생성
    if (datediff("h",request.Cookies("mAppADM")("tstp"),now())>glb_cookie_time) then             ''쿠키 expire 보다 지난경우나 쿠키 살아 있는경우
        exit function
    elseif (datediff("n",request.Cookies("mAppADM")("tstp"),now())>glb_cookie_time*60/2) or (session("mAppBctTstp")="") then  ''쿠키로 세션 재생성/ 타임스탬프 재생성
        Call chkLoginFromCookie
    end if

    '''세션 검증
    if (session("mAppBctTstp")<>glb_tstp) then Exit function

    checkLoginCookie = true
end function

if (Not checkLoginCookie) then      ''쿠키 없으면 무조건 로그오프
    glb_isLogined = false
    session.Abandon                 ''안드로이드 세션 살아 있음
    Response.Cookies("mAppADM").Domain      = manageDomain
    Response.cookies("mAppADM").expires     = date -1
else
    glb_isLogined = true
end if


dim ibackpathAll
dim strBackPath, strGetData

If (Not glb_isLogined) then
	strBackPath 	= request.ServerVariables("URL")
	strGetData  	= request.ServerVariables("QUERY_STRING")
	ibackpathAll = "backpath="+ server.URLEncode(strBackPath)
	if (strGetData<>"") then
	    ibackpathAll = ibackpathAll&server.URLEncode("?"&strGetData)
	end if

	Response.CharSet = "UTF-8"
 %>
    <script>
    alert("세션이 종료되었습니다. \n재로그인후 사용하실수 있습니다.<%=glb_isLogined%>");
    <% if (ibackpathAll<>"backpath=") then %>
    top.location = "/mAPPadmin/login.asp?<%=ibackpathAll%>";
    <% else %>
    top.location = "/mAPPadmin/login.asp";
    <% end if %>
    </script>
    <%
    response.End
End if

%>

<%
function chkLoginFromCookie()
    chkLoginFromCookie = false
    if (glb_encuid="") or (glb_tstp="") or (glb_devicekey="") or (glb_enckey="") then Exit function

    dim objXML, xmlDOM, retval, params
    params = "glb_encuid="&server.URLEncode(glb_encuid)&"&glb_tstp=" & server.URLEncode(glb_tstp)&"&glb_devicekey=" & server.URLEncode(glb_devicekey)&"&glb_appid=" & server.URLEncode(glb_appid)&"&glb_enckey=" & server.URLEncode(glb_enckey)

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.Open "POST", manageUrl & "/mAppadmin/login/actMAppCookieLoginData.asp", false
	objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	objXML.Send(params)

	If objXML.Status = "200" Then
	    Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
		xmlDOM.async = False
		'DOM 객체에 XML을 담는다.(바이너리 데이터로 받아서 UTF-8로 변환(한글문제))
		xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "UTF-8")
		''xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "EUC-KR")

        retval = xmlDOM.getElementsByTagName("retval").item(0).text
        if (retval="Y") then

            Response.Cookies("mAppADM").Domain      = manageDomain

            glb_encuid = xmlDOM.getElementsByTagName("mAppencuid").item(0).text
            glb_tstp   = xmlDOM.getElementsByTagName("mApptstp").item(0).text
            glb_enckey = xmlDOM.getElementsByTagName("mAppenckey").item(0).text

            Response.Cookies("mAppADM")("encuid")   = glb_encuid
            Response.Cookies("mAppADM")("tstp")     = glb_tstp
            Response.Cookies("mAppADM")("dvkey")    = glb_devicekey
            Response.Cookies("mAppADM")("appid")    = glb_appid
            Response.Cookies("mAppADM")("enckey")   = glb_enckey
            Response.cookies("mAppADM").expires     = Dateadd("h" , glb_cookie_time, Now())

            session("mAppBctId")    = xmlDOM.getElementsByTagName("mAppBctId").item(0).text
            session("mAppBctTstp")  = glb_tstp
            session("mAppBctDiv") = xmlDOM.getElementsByTagName("mAppBctDiv").item(0).text
    		session("mAppBctSn")  = xmlDOM.getElementsByTagName("mAppBctSn").item(0).text
            session("mAppBctCname") = xmlDOM.getElementsByTagName("mAppBctCname").item(0).text

    		session("mAppAdminPsn") = xmlDOM.getElementsByTagName("mAppAdminPsn").item(0).text		    '부서 번호
    		session("mAppAdminLsn") = xmlDOM.getElementsByTagName("mAppAdminLsn").item(0).text		    '등급 번호
    		session("mAppAdminPOsn") = xmlDOM.getElementsByTagName("mAppAdminPOsn").item(0).text		    '직책 번호
    		session("mAppAdminPOSITsn") = xmlDOM.getElementsByTagName("mAppAdminPOSITsn").item(0).text		'직급 번호

            chkLoginFromCookie=true
        end if
	    Set xmlDOM = Nothing
	ELSE
	    rw objXML.Status
	end if

	Set objXML= Nothing
end function


'if (session("mAppBctId")="") then
'    glb_encuid = request.Cookies("mAppADM")("encuid")   ''= userid
'    glb_tstp   = request.Cookies("mAppADM")("tstp")     ''= tstp
'    glb_devicekey = request.Cookies("mAppADM")("dvkey") ''= devicekey
'    glb_appid  = request.Cookies("mAppADM")("appid")    ''= appid
'    glb_enckey = request.Cookies("mAppADM")("enckey")   ''= devicekey&userid&tstp
'
'    if chkLoginFromCookie() then
'        glb_isLogined=(session("mAppBctId")<>"")
'    end if
'
''    glb_userid = TENDEC(glb_encuid)
''    rw glb_userid
''    rw glb_tstp
''    rw glb_enckey
''
''    if (glb_userid<>"") then
''        if (MD5(glb_userid&glb_tstp)=glb_enckey) then
''            ''fillSession
''
''        end if
''    end if
'
'else
'    glb_isLogined=true
'
''    rw session("mAppBctId")
'end if
%>
