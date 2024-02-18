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
' PageName : /apps/protoV1/set2ndPassdel.asp
' Discription : 2단계 비밀번호 삭제
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
Dim sType
Dim sOS, sVerCd, sUid
Dim sData : sData = Request("json")
Dim oJson

'// 전송결과 파징
on Error Resume Next

dim oResult
set oResult = JSON.parse(sData)
	sType = oResult.type
	sUid = requestCheckVar(oResult.id,32)
	sOS = requestCheckVar(oResult.OS,10)
	sVerCd = requestCheckVar(oResult.versionname,6)         ''' versioncode => versionname ''2016/12/19
set oResult = Nothing

'// json객체 선언
Set oJson = jsObject()

dim lockTerm, failNo, loginlock, errmsg, sql
failNo = 5			'// 로그인 실패 허용수
lockTerm = 15		'// 계정 장금 시간 설정(분)

IF (Err) then
	oJson("response") = getErrMsg("9999",sFDesc)
	oJson("faildesc") = "처리중 오류가 발생했습니다."
ElseIf sType<>"set2ndpassdel" then
	'// 2차 비번 설정 타입 아님
	oJson("response") = getErrMsg("9999",sFDesc)
	oJson("faildesc") = "잘못된 접근입니다."
ElseIf sUid="" then
	'// 잘못된 접근
	oJson("response") = getErrMsg("9999",sFDesc)
	oJson("faildesc") = "파라메터 정보가 없습니다."
Else
	Dim objCmd,returnValue 
	Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_partner_SetSecondPassWordDel]('"&sUid&"')}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
			returnValue = objCmd(0).Value
	Set objCmd = Nothing

	If returnValue =0 Then'2
		rsget.close
		oJson("response") = getErrMsg("9999",sFDesc)
		oJson("faildesc") = "2차 비밀번호 등록에 실패했습니다. 확인 후 다시 등록해주세요."
	Else'등록 성공'2
		oJson("response") = getErrMsg("1000",sFDesc)
	End If
End If

If ERR Then Call OnErrNoti()
On Error Goto 0
'Json 출력(JSON)
oJson.flush

Set oJson = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->