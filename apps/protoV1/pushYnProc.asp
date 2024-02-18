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
' PageName : /apps/protoV1/pushYnProc.asp
' Discription : 푸쉬 알림 켜고 끄기 설정 등록
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

Function pushYProc(sPermit,sAppKey,sDeviceId,sUUID,snid,sActTp,userid)
    Dim sqlStr, ret 
    sqlStr = " exec db_academy.[dbo].[sp_Academy_App_Push_Proc] '"&sPermit&"','"&sAppKey&"','"&sDeviceId&"','"&sUUID&"','"&snid&"','"&Request.ServerVariables("REMOTE_ADDR")&"',"&sActTp&",'"&userid&"'"
    dbACADEMYget.Execute sqlStr
    
    If (Err) Then
        pushYProc = False
    Else
        pushYProc = True
    End If
End Function


Function pushYProc2(NotiPushYN,ItemPushYN,OrderPushYN,QnAPushYN,sAppKey,sDeviceId,sUUID,snid,sActTp,userid)
    Dim sqlStr, ret 
    sqlStr = " exec db_academy.[dbo].[sp_Academy_App_Push_Proc_New] '"&NotiPushYN&"','"&ItemPushYN&"','"&OrderPushYN&"','"&QnAPushYN&"','"&sAppKey&"','"&sDeviceId&"','"&sUUID&"','"&snid&"','"&Request.ServerVariables("REMOTE_ADDR")&"',"&sActTp&",'"&userid&"'"
    dbACADEMYget.Execute sqlStr
    
    If (Err) Then
        pushYProc2 = False
    Else
        pushYProc2 = True
    End If
End Function

Dim sFDesc
Dim sType, sDeviceId, sPermit, sActTp
Dim sOS, sVerCd, sAppKey, sUUID, snID
Dim sData : sData = Request("json")
Dim oJson, retMsg, sVerCheck
Dim sPermitMsg 

Dim userid : userid = GetLoginUserID
Dim NotiPushYN, ItemPushYN, OrderPushYN, QnAPushYN

'// 전송결과 파징
On Error Resume Next

Dim oResult
Set oResult = JSON.parse(sData)
	sType = oResult.type
	sDeviceId = requestCheckVar(oResult.pushid,256)
	sOS = requestCheckVar(oResult.OS,10)
	sVerCd = requestCheckVar(oResult.versionname,6)     ''' versioncode => versionname ''2016/12/19
	sVerCheck = requestCheckVar(oResult.version,6)         ''' 2016/12/19 추가 api 버전 확인
	sAppKey = getWishAppKey(sOS)
	if Not ERR THEN
		sUUID = requestCheckVar(oResult.uuid,40)
		If ERR THEN Err.Clear ''uuid 프로토콜 없음
	END IF

	if Not ERR THEN
		snID = requestCheckVar(oResult.idfa,40)
		If ERR THEN Err.Clear ''uuid 프로토콜 없음
	END IF
    
'Response.write sPermit
	If sVerCheck = 2 Then
		NotiPushYN = requestCheckVar(oResult.notipush,1)
		ItemPushYN = requestCheckVar(oResult.itempush,1)
		OrderPushYN = requestCheckVar(oResult.orderpush,1)
		QnAPushYN = requestCheckVar(oResult.qnapush,1)
	Else
		sPermit = requestCheckVar(oResult.notiyn,1)
	End If

    sActTp="0" ''기본값
Set oResult = Nothing

'// json객체 선언
Set oJson = jsObject()

If (Err) Then
    oJson("response") = getErrMsg("9999",sFDesc)
    oJson("faildesc") = "처리중 오류가 발생했습니다."
ElseIf (LCASE(sType)<>"adpush") and (LCASE(sType)<>"adpush") Then
    '// 잘못된 콜싸인 아님
    oJson("response") = getErrMsg("9999",sFDesc)
    oJson("faildesc") = "잘못된 접근입니다."
'ElseIf (sPermit="") Then
'    oJson("response") = getErrMsg("9999",sFDesc)
'    oJson("faildesc") = "잘못된 접근입니다."
''ElseIf (sDeviceId="") Then                                ''IOS는 sDeviceId 없을 수 있음.
''    '// 잘못된 sDeviceId
''    oJson("response") = getErrMsg("9999",sFDesc)
''    oJson("faildesc") = "처리중 오류가 발생했습니다."
ElseIf sAppKey="" Then
    '// 잘못된 접근
    oJson("response") = getErrMsg("9999",sFDesc)
    oJson("faildesc") = "파라메터 정보가 없습니다."
Else
	If sVerCheck <> 2 Then 
		''A=모두받기 | N=모두끄기 | C=공지만받기 | P=상품공지만 받기
		sPermitMsg = sPermit
		if (sPermit="A") then
			sPermitMsg = "수신으"
		elseif (sPermit="N") then
			sPermitMsg = "수신거부"
		elseif (sPermit="C") then
			sPermitMsg = "공지수신으"
		elseif (sPermit="P") then
			sPermitMsg = "상품공지수신으"
		end if
		
		If (pushYProc(sPermit,sAppKey,sDeviceId,sUUID,snid,sActTp,userid)) Then
			oJson("response") = getErrMsg("1000",sFDesc)
			retMsg = "더핑거스 아티스트에서 보내는"&VBCRLF
			retMsg = retMsg &"PUSH(알림) 수신여부가 "&sPermitMsg&"로 변경되었습니다."&VBCRLF&VBCRLF
			retMsg = retMsg &LEFT(NOW(),4)&"년 "&MID(NOW(),6,2)&"월 "&MID(NOW(),9,2)&"일"&VBCRLF
			oJson("resultmsg") = retMsg
		Else
			oJson("response") = getErrMsg("9999",sFDesc)
			oJson("faildesc") = "처리중 오류가 발생했습니다."
		End If
	Else
        If (pushYProc2(NotiPushYN,ItemPushYN,OrderPushYN,QnAPushYN,sAppKey,sDeviceId,sUUID,snid,sActTp,userid)) Then
            oJson("response") = getErrMsg("1000",sFDesc)
            retMsg = "더핑거스 아티스트에서 보내는"&VBCRLF
            retMsg = retMsg &"PUSH(알림) 수신여부가 변경되었습니다."&VBCRLF&VBCRLF
            retMsg = retMsg &LEFT(NOW(),4)&"년 "&MID(NOW(),6,2)&"월 "&MID(NOW(),9,2)&"일"&VBCRLF
            oJson("resultmsg") = retMsg
        Else
            oJson("response") = getErrMsg("9999",sFDesc)
            oJson("faildesc") = "처리중 오류가 발생했습니다."
        End If
	End If
End If

If Err Then Call OnErrNoti()
On Error Goto 0
'Json 출력(JSON)
oJson.flush

Set oJson = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->