<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<% response.Charset="UTF-8" %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/apps/academy/lib/commlib.asp" -->
<!-- #include virtual="/lib/util/md5.asp"-->
<!-- #include virtual="/lib/util/base64New.asp"-->
<!-- #include virtual="/lib/util/tenEncUtil.asp"-->
<!-- #include virtual="/apps/common/appFunction.asp"-->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<!-- #include virtual="/apps/protoV1/protoFunction.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
'###############################################
' PageName : /apps/appCom/wish/startupProc.asp
' Discription : Wish APP 최초구동시 정보 처리
' Request : json > type, pushid, OS, versioncode, versionname, verserion
' Response : response > 결과
' History : 2016.10.18 허진원 : 신규 생성
'###############################################

'//헤더 출력
Response.ContentType = "text/html"

'---------------------------
'@@ 서버 점검시 아래 주석을 풀어주세요
Set oJson = jsObject()
oJson("response") = getErrMsg("9999",sFDesc)
oJson("faildesc") = "핑거스 서비스가 텐바이텐으로 이전되고 종료 되었습니다. 텐바이텐으로 문의 바랍니다. 감사합니다."
oJson.flush
Set oJson = Nothing
Response.End
'---------------------------

Dim sFDesc
Dim sType, sDeviceId
Dim sOS, sVerCd, sAppKey, sMinUpVer, sCurrVer, sAppId, sUUID, sPushyn, snID
Dim sData : sData = Request("json")
Dim oJson, sForcedUpdate, sVerCheck

'// 전송결과 파징
on Error Resume Next
dim oResult
set oResult = JSON.parse(sData)
	sType = oResult.type
	sDeviceId = requestCheckVar(oResult.pushid,256)

	sOS = requestCheckVar(oResult.OS,10)
	sVerCd = requestCheckVar(oResult.versionname,20)   ''' versioncode => versionname ''2016/12/19
	sVerCheck = requestCheckVar(oResult.version,6)         ''' 2016/12/19 추가 api 버전 확인
	
	sAppKey = getWishAppKey(sOS)

	if Not ERR THEN
		sUUID = requestCheckVar(oResult.uuid,40)
		if ERR THEN Err.Clear ''uuid 프로토콜 없음
	END IF

	if Not ERR THEN
		snID = requestCheckVar(oResult.idfa,40)
		if ERR THEN Err.Clear ''uuid 프로토콜 없음
	END IF



set oResult = Nothing

'// json객체 선언
Set oJson = jsObject()

If (Err) Then
	oJson("response") = getErrMsg("9999",sFDesc)
	oJson("faildesc") = "처리중 오류가 발생했습니다."

ElseIf sType<>"firstconnection" Then
	'// 잘못된 콜싸인 아님
	oJson("response") = getErrMsg("9999",sFDesc)
	oJson("faildesc") = "잘못된 접근입니다."

ElseIf sAppKey="" Then

	'// 잘못된 접근
	oJson("response") = getErrMsg("9999",sFDesc)
	oJson("faildesc") = "파라메터 정보가 없습니다."

Else
	Dim sqlStr

	'// 위시 앱버전 확인
	sqlStr = "Select minuUpVer, currVer, appId from db_academy.dbo.tbl_app_master Where appKey=" & sAppKey
	rsACADEMYget.CursorLocation = adUseClient
	rsACADEMYget.Open sqlStr,dbACADEMYget,adOpenForwardOnly, adLockReadOnly
	If Not(rsACADEMYget.EOF Or rsACADEMYget.BOF) Then
		sMinUpVer = rsACADEMYget("minuUpVer")			'크리티컬한 최소 구동 버전   '' 막음.
		sCurrVer = rsACADEMYget("currVer")				'최신 APP버전                '' 권장업데이트시
		sAppId = rsACADEMYget("appId")			        '앱스토어 AppId
	Else
		sMinUpVer = "2"
		sCurrVer = "2"
		sAppId = "1181771109"
	End If
	rsACADEMYget.Close


'' IOS
' 테스트 1 - 낮은버전(일반 업데이트)
'    if (sAppKey="5") then
'        sMinUpVer = "0"
'		sCurrVer = "1.7"
'		sAppId = "864817011"
'    end if

'' 테스트 2 - 강제업데이트
'    if (sAppKey="5") then
'        sMinUpVer = "1.7"
'		sCurrVer = "1.7"
'		sAppId = "864817011"
'    end if

'' Android
'' 테스트 1 - 낮은버전(일반 업데이트)
'    if (sAppKey="6") then
'       sMinUpVer = "32"
'		sCurrVer = "42"
'		sAppId = "kr.tenbyten.shoping"
'    end if

'' 테스트 2 - 강제업데이트
'    if (sAppKey="6") then
'        sMinUpVer = "42"
'		sCurrVer = "42"
'		sAppId = "kr.tenbyten.shoping"
'    end if

    ''강제업데이트시 
    if (cStr(sVerCd)<cStr(sMinUpVer)) then
        sForcedUpdate = true
    else
        sForcedUpdate = false
    end if

	if sDeviceId<>"" then
		'// 접속 기기 정보 저장 //필요 없을듯 deviceProc.asp 추가됨
		sqlStr = "IF NOT EXISTS(select regidx from db_academy.dbo.tbl_app_regInfo where appkey=" & sAppKey & " and deviceid='" & sDeviceId & "') " & vbCrLf
		sqlStr = sqlStr & " begin " & vbCrLf
		sqlStr = sqlStr & "	insert into db_academy.dbo.tbl_app_regInfo " & vbCrLf
		sqlStr = sqlStr & "		(appKey,deviceid,regdate,appVer,lastact,isAlarm01,isAlarm02,isAlarm03,isAlarm04,isAlarm05,regrefip) values " & vbCrLf
		sqlStr = sqlStr & "	(" & sAppKey			'앱고유Key
		sqlStr = sqlStr & ",'" & sDeviceId & "'"	'접속기기 DeviceID
		sqlStr = sqlStr & ",getdate()"				'최초접속 일시
		sqlStr = sqlStr & ",'"&sVerCd&"'"			'버전                       ''/2014/03/21
		sqlStr = sqlStr & ",'str'"                  '' 최종액션구분
		sqlStr = sqlStr & ",'Y'"					'공지사항 알림 여부
		sqlStr = sqlStr & ",'Y'"					'작품 등록 승인 알림 여부
		sqlStr = sqlStr & ",'Y'"					'작품 주문 알림 여부
		sqlStr = sqlStr & ",'Y'"					'작품 Q&A 알림 여부
		sqlStr = sqlStr & ",'N','"&Request.ServerVariables("REMOTE_ADDR")&"') " & vbCrLf
		sqlStr = sqlStr & " end" & vbCrLf
		sqlStr = sqlStr & " ELSE" & vbCrLf
		sqlStr = sqlStr & " begin " & vbCrLf
		sqlStr = sqlStr & " update db_academy.dbo.tbl_app_regInfo" & vbCrLf
	    sqlStr = sqlStr & "	set lastact='stU'" & vbCrLf
	    sqlStr = sqlStr & "	,appVer='"&sVerCd&"'" & vbCrLf
	    sqlStr = sqlStr & "	,isusing='Y'" & vbCrLf
	    sqlStr = sqlStr & "	,lastUpdate=getdate()" & vbCrLf
	    sqlStr = sqlStr & "	where appkey=" & sAppKey & " and deviceid='" & sDeviceId & "'" & vbCrLf
		sqlStr = sqlStr & " end" & vbCrLf
		dbACADEMYget.Execute(sqlStr)

		''call addDeviceLog(sAppKey,sDeviceId,"",sVerCd,"str")
	else
	    ''call addDeviceLog(sAppKey,sDeviceId,"",sVerCd,"ttt")
	end if

    '' uuid 추가 2014/06/25 --------------------------------
    ''call addUUIDInfo(sAppKey,sDeviceId,sUUID)
    '' nid 추가 2015/07/23 --------------------------------
    call addUUIDNidInfo(sAppKey,sDeviceId,sUUID,snid)
    ''------------------------------------------------------

	oJson("response") = getErrMsg("1000",sFDesc)
	oJson("lastversionname") = cStr(sCurrVer)		'현재App 버전 == 버전명

	dim strAppWVUrl
	IF application("Svr_Info")="Dev" THEN
		strAppWVUrl = "http://testwebadmin.10x10.co.kr/apps/academy/"
	else
		strAppWVUrl = "https://webadmin.10x10.co.kr/apps/academy/"
	end if

	dim currenttime
	currenttime = now()


	'// 추후 업데이트시 이동될 AppID : 필요없을듯.
	oJson("appid") = sAppId
    
	if (sVerCheck =1) then
		Set oJson("topmenu") = getTopMenuJSon_2015
	else
		Set oJson("topmenu") = getTopMenuJSon_TEST
		Set oJson("submenu") = getSubMenuJSon_TEST
	end if
    
    oJson("forcedupdate") = sForcedUpdate  
    

end if

if ERR then Call OnErrNoti()
On Error Goto 0

'Json 출력(JSON)
oJson.flush
Set oJson = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->