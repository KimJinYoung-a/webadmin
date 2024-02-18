<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<% response.Charset="UTF-8" %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
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
' PageName : /apps/protoV1/logoutProc.asp
' Discription : 로그 아웃 처리
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
Dim sOS, sVerCd
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
set oResult = Nothing

'// json객체 선언
Set oJson = jsObject()

IF (Err) then
	oJson("response") = getErrMsg("9999",sFDesc)
	oJson("faildesc") = "처리중 오류가 발생했습니다."
ElseIf sType<>"logout" then
	'// 로그아웃
	oJson("response") = getErrMsg("9999",sFDesc)
	oJson("faildesc") = "잘못된 접근입니다."
ElseIf sType="" then
	'// 잘못된 접근
	oJson("response") = getErrMsg("9999",sFDesc)
	oJson("faildesc") = "파라메터 정보가 없습니다."
Else
	response.Cookies("partner").domain = "10x10.co.kr"
	response.Cookies("partner") = ""
	response.Cookies("partner").Expires = Date - 1

	response.Cookies("PASave").domain = "10x10.co.kr"
	response.Cookies("PASave") = ""
	response.Cookies("PASave").Expires = Date - 1
	session.abandon
	
	oJson("response") = getErrMsg("1000",sFDesc)
End If

if ERR then Call OnErrNoti()
On Error Goto 0
'Json 출력(JSON)
oJson.flush

Set oJson = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->