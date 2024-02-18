<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<% response.Charset="UTF-8" %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/apps/academy/lib/commlib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/md5.asp"-->
<!-- #include virtual="/lib/util/base64New.asp"-->
<!-- #include virtual="/lib/util/tenEncUtil.asp"-->
<!-- #include virtual="/apps/common/appFunction.asp"-->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<!-- #include virtual="/apps/protoV1/protoFunction.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
'###############################################
' PageName : /apps/protoV1/deviceProc.asp
' Discription : 푸쉬 알림 정보 등록
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
Dim sType, sDeviceId, sVerCheck
Dim sOS, sVerCd, sAppKey, sUUID, snID
Dim sData : sData = Request("json")
Dim oJson

'// 전송결과 파징
on Error Resume Next

dim oResult
set oResult = JSON.parse(sData)
	sType = oResult.type
	sDeviceId = requestCheckVar(oResult.pushid,256)
	sOS = requestCheckVar(oResult.OS,10)
	sVerCd = requestCheckVar(oResult.versionname,6)         ''' versioncode => versionname ''2016/12/19
	sVerCheck = requestCheckVar(oResult.version,6)         ''' 2016/12/19 추가 api 버전 확인
	if Not ERR THEN
		sUUID = requestCheckVar(oResult.uuid,40)
		If ERR THEN Err.Clear ''uuid 프로토콜 없음
	End If

	if Not ERR THEN
		snID = requestCheckVar(oResult.idfa,40)
		If ERR THEN Err.Clear ''uuid 프로토콜 없음
	End If

set oResult = Nothing

	sAppKey = getWishAppKey(sOS)

'// json객체 선언
Set oJson = jsObject()

IF (Err) then
	oJson("response") = getErrMsg("9999",sFDesc)
	oJson("faildesc") = "처리중 오류가 발생했습니다."
ElseIf (sType<>"reg") and (sType<>"rmv") then
	'// 잘못된 콜싸인 아님
	oJson("response") = getErrMsg("9999",sFDesc)
	oJson("faildesc") = "잘못된 접근입니다."
ElseIf (sDeviceId="") then
	'// 잘못된 sDeviceId
	oJson("response") = getErrMsg("9999",sFDesc)
	oJson("faildesc") = "잘못된 접근입니다."
ElseIf sAppKey="" then
	'// 잘못된 접근
	oJson("response") = getErrMsg("9999",sFDesc)
	oJson("faildesc") = "파라메터 정보가 없습니다."
else
	dim sqlStr

	if sDeviceId<>"" Then
		'// 접속 기기 정보 저장
		if (sType="reg") then  ''구글서버에서 어씽크로 받아옴
			sqlStr = "IF NOT EXISTS(select regidx from db_academy.dbo.tbl_app_regInfo where appkey=" & sAppKey & " and deviceid='" & sDeviceId & "') " & vbCrLf
    		sqlStr = sqlStr & "begin " & vbCrLf
    		sqlStr = sqlStr & "	insert into db_academy.dbo.tbl_app_regInfo " & vbCrLf
    		sqlStr = sqlStr & "		(appKey,deviceid,regdate,appVer,lastact,isAlarm01,isAlarm02,isAlarm03,isAlarm04,isAlarm05,regrefip) values " & vbCrLf
    		sqlStr = sqlStr & "	(" & sAppKey			'앱고유Key
    		sqlStr = sqlStr & ",'" & sDeviceId & "'"	'접속기기 DeviceID
    		sqlStr = sqlStr & ",getdate()"				'최초접속 일시
    		sqlStr = sqlStr & ",'"&sVerCd&"'"			'버전                       ''/2014/03/21
    		sqlStr = sqlStr & ",'reg'"                  '' 최종액션구분
    		sqlStr = sqlStr & ",'Y'"					'공지사항 알림 여부
			sqlStr = sqlStr & ",'Y'"					'작품 등록 승인 알림 여부
			sqlStr = sqlStr & ",'Y'"					'작품 주문 알림 여부
			sqlStr = sqlStr & ",'Y'"					'작품 Q&A 알림 여부
			sqlStr = sqlStr & ",'N','"&Request.ServerVariables("REMOTE_ADDR")&"') " & vbCrLf
    		sqlStr = sqlStr & " end"& vbCrLf

    		sqlStr = sqlStr & " ELSE"& vbCrLf
    		sqlStr = sqlStr & " begin " & vbCrLf
    		sqlStr = sqlStr & " update db_academy.dbo.tbl_app_regInfo" & vbCrLf
    	    sqlStr = sqlStr & "	set lastact='rrg'" & vbCrLf                         ''기기삭제후 재등록 등
    	    sqlStr = sqlStr & "	,appVer='"&sVerCd&"'" & vbCrLf
    	    sqlStr = sqlStr & "	,isusing='Y'" & vbCrLf
    	    sqlStr = sqlStr & "	,lastUpdate=getdate()" & vbCrLf
    	    sqlStr = sqlStr & "	where appkey=" & sAppKey & " and deviceid='" & sDeviceId & "'" & vbCrLf
    		sqlStr = sqlStr & " end"& vbCrLf
			dbACADEMYget.Execute(sqlStr)

    	ElseIf (sType="rmv") then ''버전업이 된경우 삭제
    	    sqlStr = "update db_academy.dbo.tbl_app_regInfo" & vbCrLf
    	    sqlStr = sqlStr & "	set isusing='N'" & vbCrLf
    	    sqlStr = sqlStr & "	,lastact='rmv'" & vbCrLf
    	    sqlStr = sqlStr & "	,lastUpdate=getdate()" & vbCrLf
    	    sqlStr = sqlStr & "	where appkey=" & sAppKey & " and deviceid='" & sDeviceId & "'" & vbCrLf
    	    dbACADEMYget.Execute(sqlStr)
    	End If

        ''변경로그 작성
    	Call addDeviceLog(sAppKey,sDeviceId,"",sVerCd,sType)
	End If
    
    If (sDeviceId<>"") And (sType<>"rmv") Then
    'Response.write sUUID & "<br>"
	'Response.write snid
	'Response.end
        '' uuid 추가 2014/06/25 --------------------------------
        ''call addUUIDInfo(sAppKey,sDeviceId,sUUID)
        '' nid 추가 2015/07/23 --------------------------------
        Call addUUIDNidInfo(sAppKey,sDeviceId,sUUID,snid)
        ''------------------------------------------------------
    End If

	oJson("response") = getErrMsg("1000",sFDesc)

End If

if ERR then Call OnErrNoti()
On Error Goto 0
'Json 출력(JSON)
oJson.flush

Set oJson = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->