<%
''OnErrResumeNext 사용시 에러통보
function OnErrNoti()
    Const lngMaxFormBytes = 800
    dim strServerIP
    dim errDescription, errSource
    dim strMethod,datNow

    errDescription = ERR.Description
    errSource =  "["&ERR.Number&"]"&ERR.Source

	strServerIP = Request.ServerVariables("LOCAL_ADDR")

    dim strMsg : strMsg=""

    strMsg = strMsg & "errDescription: "&errDescription&"<br>"
    strMsg = strMsg & "errSource: "&errSource&"<br>"&"<br>"

    strMsg = strMsg & "<li>서버:<br>"
	strMsg = strMsg & application("Svr_Info") & ":"&strServerIP
	strMsg = strMsg & "<br><br></li>"

	'// 접속자 브라우저 정보
	strMsg = strMsg & "<li>브라우저 종류:<br>"
	strMsg = strMsg & Server.HTMLEncode(Request.ServerVariables("HTTP_USER_AGENT"))
	strMsg = strMsg & "<br><br></li>"

	strMsg = strMsg & "<li>접속자 IP:<br>"
	strMsg = strMsg & Server.HTMLEncode(Request.ServerVariables("REMOTE_ADDR"))
	strMsg = strMsg & "<br><br></li>"

	strMsg = strMsg & "<li>경유페이지:<br>"
	strMsg = strMsg & request.ServerVariables("HTTP_REFERER")
	strMsg = strMsg & "<br><br></li>"

	'// 오류 페이지 정보
	strMsg = strMsg & "<li>페이지:<br>"
	strMethod = Request.ServerVariables("REQUEST_METHOD")
	strMsg = strMsg & "HOST : " & Request.ServerVariables("HTTP_HOST") & "<BR>"
	strMsg = strMsg & strMethod & " : "

	If strMethod = "POST" Then
		strMsg = strMsg & Request.TotalBytes & " bytes to "
	End If

	strMsg = strMsg & Request.ServerVariables("SCRIPT_NAME")
	strMsg = strMsg & "</li>"

	If strMethod = "POST" Then
		strMsg = strMsg & "<br><li>POST Data:<br>"

		'실행에 관련된 에러를 출력합니다.
		If Request.TotalBytes > lngMaxFormBytes Then
			strMsg = strMsg & Server.HTMLEncode(Left(Request.Form, lngMaxFormBytes)) & " . . ."'
		Else
			strMsg = strMsg & Server.HTMLEncode(Request.Form)
		End If
		strMsg = strMsg & "</li>"
	elseif strMethod = "GET" then
		strMsg = strMsg & "<br><li>GET Data:<br>"
		strMsg = strMsg & Request.QueryString
	End If
	strMsg = strMsg & "<br><br></li>"

	'// 오류 발생시간 정보
	strMsg = strMsg & "<li>시간:<br>"
	datNow = Now()
	strMsg = strMsg & Server.HTMLEncode(FormatDateTime(datNow, 1) & ", " & FormatDateTime(datNow, 3))
	on error resume next
		Session.Codepage = bakCodepage
	on error goto 0
	strMsg = strMsg & "<br><br></li>"


    '### 시스템팀 구성원에게 오류 발생 내용 발송 ###
	dim cdoMessage,cdoConfig
	Set cdoConfig = CreateObject("CDO.Configuration")

    '-> 서버 접근방법을 설정합니다
	cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 '1 - (cdoSendUsingPickUp)  2 - (cdoSendUsingPort)
	'-> 서버 주소를 설정합니다
	If application("Svr_Info")="Dev" Then
	    cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver")="61.252.133.2" ''"110.93.128.94"
	else
	    cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver")="110.93.128.94" ''"110.93.128.94"
    end if
	'-> 접근할 포트번호를 설정합니다
	cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
	'-> 접속시도할 제한시간을 설정합니다
	cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 30
	'-> SMTP 접속 인증방법을 설정합니다
	cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
	'-> SMTP 서버에 인증할 ID를 입력합니다
	cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "MailSendUser"
	'-> SMTP 서버에 인증할 암호를 입력합니다
	cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "wjddlswjddls"
	cdoConfig.Fields.Update

	Set cdoMessage = CreateObject("CDO.Message")
	Set cdoMessage.Configuration = cdoConfig

	'개발 중 막음(corpse2)
	cdoMessage.To 		= "kobula@10x10.co.kr;corpse2@10x10.co.kr"
	cdoMessage.From 	= "webserver@10x10.co.kr"
	cdoMessage.SubJect 	= "["&date()&"] Fingers App 페이지 오류 발생"
	cdoMessage.HTMLBody	= strMsg

	cdoMessage.BodyPart.Charset="ks_c_5601-1987"         '/// 한글을 위해선 꼭 넣어 주어야 합니다.
    cdoMessage.HTMLBodyPart.Charset="ks_c_5601-1987"     '/// 한글을 위해선 꼭 넣어 주어야 합니다.
	cdoMessage.Send

	Set cdoMessage = nothing
	Set cdoConfig = nothing

end function

'// 앱 Master Key 접수
function getWishAppKey(sOs)
	Select Case uCase(sOs)
		Case "IOS", "IPHONE", "IPAD"
			getWishAppKey = "7"
		Case "ANDROID"
			getWishAppKey = "8"
	end Select
end Function

function getWishAppKey2(sOs)
	Select Case uCase(sOs)
		Case "IOS", "IPHONE", "IPAD"
			getWishAppKey2 = "9"
		Case "ANDROID"
			getWishAppKey2 = "10"
	end Select
end Function

'// 오류 메시지 출력
function getErrMsg(sCd,byRef fDesc)
	Select Case sCd
		Case "1000"
			getErrMsg = true
		Case "2102"
			getErrMsg = "fail_password"
			fDesc = "아이디 또는 비밀번호가 맞지 않습니다."
		Case "2103"
			getErrMsg = "user_not_exist"
			fDesc = "아이디 또는 비밀번호가 맞지 않습니다."
		Case "2201"
			getErrMsg = "user_not_allow"
			fDesc = "텐바이텐 이용를 허용하지 않으셨습니다."
		Case "2202"
			getErrMsg = "user_not_use"
			fDesc = "죄송합니다. 사용이 불가한 회원입니다."
		Case "2301"
			getErrMsg = "user_not_auth"
			fDesc = "아직 인증을 받지 않은 회원입니다."
		Case "9000"
			getErrMsg = "no_login"
			fDesc = "로그인이 필요합니다."
		Case "9100"
			getErrMsg = "low_version"
			fDesc = "업그래이드가 필요합니다."
		Case "9999"
			getErrMsg = false
		Case Else
			getErrMsg = false
	End Select
end function

'// 디바이스 등록/ 버전업 로그
function addDeviceLog(appKey,deviceid,userid,appVer,lastact)
   Exit function  ''더이상 로그쌓지 않음. 2014/12/02
    dim sqlStr
  
    sqlStr = " insert into db_academy.dbo.tbl_app_regInfo_log" & vbCrLf
    sqlStr = sqlStr & "	(appKey,deviceid,userid,appVer,lastact)" & vbCrLf
    sqlStr = sqlStr & "	values("&appKey&"" & vbCrLf
    sqlStr = sqlStr & "	,'"&deviceid&"'" & vbCrLf
    sqlStr = sqlStr & "	,'"&userid&"'" & vbCrLf
    sqlStr = sqlStr & "	,'"&appVer&"'" & vbCrLf
    sqlStr = sqlStr & "	,'"&lastact&"'" & vbCrLf
    sqlStr = sqlStr & "	)" & vbCrLf

    dbACADEMYget.Execute(sqlStr)
end function

'// uuid 정보 저장
function addUUIDInfo(appKey,deviceid,uuID)
    dim sqlStr
    if (uuID="") then Exit function
    if (appKey="") then Exit function
        
    sqlStr = "exec db_academy.dbo.sp_Academy_save_App_UUID_INFO "&appKey&",'"&deviceid&"','"&uuID&"'"
    dbACADEMYget.Execute(sqlStr)
end function

'// uuid 및 nid 정보 저장 2015/07/23
Function addUUIDNidInfo(appKey,deviceid,uuID,nid)
    dim sqlStr
    if (uuID="") and (nid="") then Exit function
        
    sqlStr = "exec db_academy.dbo.sp_Academy_save_App_UUID_NID_INFO "&appKey&",'"&deviceid&"','"&uuID&"','"&nid&"'"
    dbACADEMYget.Execute(sqlStr)
End Function

























'// 필터 파징
Sub getParseFilter(sJson, byRef dspTp, byRef prcMin, byRef prcMax, byRef ColrCd, byRef catCd, byRef keyWd, byRef mkrid)
	Dim oFlt, oColr, i
	Set oFlt = sJson

	dspTp = oFlt.displaytype
	prcMin = getNumeric(requestCheckVar(oFlt.pricelimitlow,8))
	prcMax = chkIIF(oFlt.pricelimithigh="-1","-1",getNumeric(requestCheckVar(oFlt.pricelimithigh,8)))
	catCd = ReplaceRequestSpecialChar(oFlt.categoryid)
	keyWd = requestCheckVar(oFlt.keyword,100)
	mkrid = ReplaceRequestSpecialChar(oFlt.brandid)

	Set oColr = oFlt.color
	if oColr.length>0 then
		for i=0 to oColr.length-1
			if i>0 then ColrCd=ColrCd & ","
			ColrCd = ColrCd & getTenColorCd(oColr.get(i).colorindex)
		next
	end if
	Set oColr = Nothing

	Set oFlt = Nothing
end Sub





'// URLEncode
function b64encode(str)
	str = base64encode(str)
	str = replace(str,"+","-")
	str = replace(str,"/","_")
	b64encode = str
end function





'// 필터 파징 (v2프로토콜)
Sub getParseFilterV2(sJson, byRef dspTp, byRef prcMin, byRef prcMax, byRef ColrCd, byRef catCd, byRef keyWd, byRef mkrid, byRef delitp)
	Dim oFlt, oColr, i
	Set oFlt = sJson

	dspTp = oFlt.displaytype
	prcMin = getNumeric(requestCheckVar(oFlt.pricelimitlow,8))
	prcMax = chkIIF(oFlt.pricelimithigh="-1","-1",getNumeric(requestCheckVar(oFlt.pricelimithigh,8)))
	catCd = ReplaceRequestSpecialChar(oFlt.categoryid)
	keyWd = requestCheckVar(oFlt.keyword,100)
	mkrid = ReplaceRequestSpecialChar(oFlt.brandid)
    on Error resume Next
    delitp = requestCheckVar(oFlt.delitp,2)  ''V2에서만 있음.
    if Err then on Error Goto 0
        
	Set oColr = oFlt.color
	if oColr.length>0 then
		for i=0 to oColr.length-1
			if i>0 then ColrCd=ColrCd & ","
			ColrCd = ColrCd & (oColr.get(i).colorindex)  ''V2 환경에서는 getTenColorCd 안씀
		next
	end if
	Set oColr = Nothing

	Set oFlt = Nothing
end Sub

'// 필터 파징 (v3프로토콜)
Sub getParseFilterV3(sJson, byRef dspTp, byRef prcMin, byRef prcMax, byRef ColrCd, byRef catCd, byRef keyWd, byRef mkrid, byRef delitp, byRef rstxt)
	Dim oFlt, oColr, i
	Set oFlt = sJson

	dspTp = oFlt.displaytype
	prcMin = getNumeric(requestCheckVar(oFlt.pricelimitlow,8))
	prcMax = chkIIF(oFlt.pricelimithigh="-1","-1",getNumeric(requestCheckVar(oFlt.pricelimithigh,8)))
	catCd = ReplaceRequestSpecialChar(oFlt.categoryid)
	keyWd = requestCheckVar(oFlt.keyword,100)
	mkrid = ReplaceRequestSpecialChar(oFlt.brandid)
    on Error resume Next
    delitp = requestCheckVar(oFlt.delitp,2)  ''V2에서만 있음.
    rstxt = requestCheckVar(oFlt.rstxt,100)  ''V3에서만 있음.
    if Err then on Error Goto 0
        
	Set oColr = oFlt.color
	if oColr.length>0 then
		for i=0 to oColr.length-1
			if i>0 then ColrCd=ColrCd & ","
			ColrCd = ColrCd & (oColr.get(i).colorindex)  ''V2 환경에서는 getTenColorCd 안씀
		next
	end if
	Set oColr = Nothing

	Set oFlt = Nothing
end Sub

'// 필터 파징 (v3.1프로토콜)
Sub getParseFilterV31(sJson, byRef dspTp, byRef prcMin, byRef prcMax, byRef ColrCd, byRef catCd, byRef keyWd, byRef mkrid, byRef delitp, byRef rstxt,byRef sflag,byRef sscp)
	Dim oFlt, oColr, i
	Set oFlt = sJson

	dspTp = oFlt.displaytype
	prcMin = getNumeric(requestCheckVar(oFlt.pricelimitlow,8))
	prcMax = chkIIF(oFlt.pricelimithigh="-1","-1",getNumeric(requestCheckVar(oFlt.pricelimithigh,8)))
	catCd = ReplaceRequestSpecialChar(oFlt.categoryid)
	keyWd = requestCheckVar(oFlt.keyword,100)
	mkrid = ReplaceRequestSpecialChar(oFlt.brandid)
    on Error resume Next
    delitp = requestCheckVar(oFlt.delitp,2)  ''V2에서만 있음.
    rstxt = requestCheckVar(oFlt.rstxt,100)  ''V3에서만 있음.
    
    sflag = requestCheckVar(oFlt.sflag,10)  ''V31에서만 있음.
    sscp = requestCheckVar(oFlt.sscp,10)  ''V31에서만 있음.
    
    if Err then on Error Goto 0
        
	Set oColr = oFlt.color
	if oColr.length>0 then
		for i=0 to oColr.length-1
			if i>0 then ColrCd=ColrCd & ","
			ColrCd = ColrCd & (oColr.get(i).colorindex)  ''V2 환경에서는 getTenColorCd 안씀
		next
	end if
	Set oColr = Nothing

	Set oFlt = Nothing
end Sub

'' popFilterApp.asp 에서 쓰임
Sub getParseFilterPop(sJson, byRef prcMin, byRef prcMax, byRef ColrCd, byRef delitp)
	Dim oFlt, oColr, i
	Set oFlt = sJson

	prcMin = getNumeric(requestCheckVar(oFlt.pricelimitlow,8))
	prcMax = chkIIF(oFlt.pricelimithigh="-1","-1",getNumeric(requestCheckVar(oFlt.pricelimithigh,8)))
    delitp = requestCheckVar(oFlt.delitp,2)  ''V2에서만 있음.
        
	Set oColr = oFlt.color
	if oColr.length>0 then
		for i=0 to oColr.length-1
			if i>0 then ColrCd=ColrCd & ","
			ColrCd = ColrCd & (oColr.get(i).colorindex)  ''V2 환경에서는 getTenColorCd 안씀
		next
	end if
	Set oColr = Nothing

	Set oFlt = Nothing
end Sub


'// 컬러코드 매칭 > 반환 (설정 필요)
Function getTenColorCd(appColrCd)
	Select Case cStr(appColrCd)
		Case "1"	'엘로우
			getTenColorCd = "003,010,021"
		Case "2"	'오렌지
			getTenColorCd = "002"
		Case "3"	'레드
			getTenColorCd = "001,023"
		Case "4"	'핑크
			getTenColorCd = "009,017"
		Case "5"	'바이올렛
			getTenColorCd = "008,018"
		Case "6"	'블루
			getTenColorCd = "006,007,020"
		Case "7"	'그린
			getTenColorCd = "005,016,019"
		Case "8"	'화이트
			getTenColorCd = "004,011,024"
		Case "9"	'그레이
			getTenColorCd = "012,022"
		Case "10"	'블랙
			getTenColorCd = "013"
		Case "11"	'골드
			getTenColorCd = "015"
		Case "12"	'실버
			getTenColorCd = "014"
		Case "13"	'스트라이프
			getTenColorCd = "026"
		Case "14"	'체크
			getTenColorCd = "025"
		Case "15"	'도트
			getTenColorCd = "027"
		Case "16"	'플라워
			getTenColorCd = "028"
		Case "17"	'드로잉
			getTenColorCd = "029"
		Case "18"	'애니멀
			getTenColorCd = "030"
		Case "19"	'지오매틱
			getTenColorCd = "031"
	end Select
end Function

''// 팔로잉 푸시메시지 발송
Sub sendFollowingPushMsg(muid,fuid)
	Dim sqlStr
	sqlStr = "exec [db_contents].[dbo].[sp_Ten_sendPushMsg_Follow] '" & muid & "','" & fuid & "' "
	dbget.Execute(sqlStr)
end Sub
%>