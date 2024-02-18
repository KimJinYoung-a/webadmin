<%
CONST C_APPID_IOS = "1"
CONST C_APPID_AND = "2"

CONST glb_cookie_time = 1 ''hour


Dim manageDomain
DIM staticImgUrl,webImgUrl,manageUrl,wwwUrl,uploadImgUrl,staticUploadUrl

IF application("Svr_Info")="Dev" THEN
    manageDomain    = "testm.10x10.co.kr"                           '' using cookie write
    staticImgUrl    = "http://testimgstatic.10x10.co.kr"	        '테스트
 	webImgUrl		= "http://testwebimage.10x10.co.kr"				'웹이미지

 	manageUrl 	    = "http://"&manageDomain
    wwwUrl 		    = "http://2012www.10x10.co.kr"

 	uploadImgUrl    = "http://testupload.10x10.co.kr"
 	staticUploadUrl	= "http://testimgstatic.10x10.co.kr"
else
    manageDomain    = "webadmin.10x10.co.kr"
    staticImgUrl    = "http://imgstatic.10x10.co.kr"
 	webImgUrl		= "http://webimage.10x10.co.kr"				'웹이미지

    manageUrl 	    = "http://"&manageDomain
    wwwUrl 		    = "http://www.10x10.co.kr"

    uploadImgUrl    = "http://upload.10x10.co.kr"          '' upload.10x10.co.kr 통해서 Nas Server로 업로드
	staticUploadUrl = "http://oimgstatic.10x10.co.kr"
end if


''------------------------------------------------------------------------------
'' checkDevice 통합
	Dim uAgent, flgDevice
	'///// 접속기종 및 브라우져 종류 가져오기(모바일에서만 ipad 추가;2012.05.23) /////
	uAgent = Lcase(Request.ServerVariables("HTTP_USER_AGENT"))
	if instr(uAgent,"windows ce")>0 or instr(uAgent,"lgtelecom")>0 or instr(uAgent,"midp")>0 or instr(uAgent,"wipi")>0 or instr(uAgent,"android")>0 or instr(uAgent,"ipod")>0 or instr(uAgent,"iphone")>0 or instr(uAgent,"ipad")>0 or instr(uAgent,"playstation") or instr(uAgent,"blackberry") then
		'휴대기기
		if instr(uAgent,"ppc")>0 or instr(uAgent,"iemobile")>0 then
			flgDevice = "P"	'PDA
		elseif instr(uAgent,"ipod")>0 or instr(uAgent,"iphone")>0 or instr(uAgent,"ipad")>0 then
			flgDevice = "I" 'iPhone,iPod,iPad
		elseif instr(uAgent,"android")>0 then
			flgDevice = "A" 'Android
		else
			flgDevice = "M"	'Mobile
		end if
	else
		'일반
		flgDevice = "W"
	end If

''------------------------------------------------------------------------------
function getAppIDByUSERAGENT()
    getAppIDByUSERAGENT = -1        ''unknown

    if flgDevice="A" then
	    getAppIDByUSERAGENT = 2
    elseif flgDevice="I" then
		getAppIDByUSERAGENT = 1
    end if
end function

function getTimeStampFormat()
    getTimeStampFormat = FormatDateTime(now(),2)&" "&FormatDateTime(now(),4)&Right(FormatDateTime(now(),3),3)
end function

function getMAppLoginUserID()
    getMAppLoginUserID = LCASE(session("mAppBctId"))
end function
%>
