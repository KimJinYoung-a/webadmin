<%@ language="VBScript" %>
<%


  Option Explicit

  Const lngMaxFormBytes = 800

  Dim objASPError, blnErrorWritten, strServername, strServerIP, strRemoteIP
  Dim strMethod, lngPos, datNow, strQueryString, strURL

	If Response.Buffer Then
		Response.Clear
		Response.Status = "500 Internal Server Error"
		Response.ContentType = "text/html"
		''Response.ChaRset = "euc-kr"
		Response.Expires = 0
	End If

	'// 오류객체 선언
	Set objASPError = Server.GetLastError

	'### 오류 메시지 작성 ###
	Dim bakCodepage, strMsg

	'// 오류 유형 정보
	strMsg = "<li>오류 유형:<br>"

	on error resume next
		bakCodepage = Session.Codepage
		Session.Codepage = 1252
		on error goto 0

		strMsg = strMsg & Server.HTMLEncode(objASPError.Category)

		If objASPError.ASPCode > "" Then strMsg = strMsg & Server.HTMLEncode(", " & objASPError.ASPCode)
			strMsg = strMsg &  Server.HTMLEncode(" (0x" & Hex(objASPError.Number) & ")" ) & "<br>"
		If objASPError.ASPDescription > "" Then
			strMsg = strMsg & Server.HTMLEncode(objASPError.ASPDescription) & "<br>"
		elseIf (objASPError.Description > "") Then
			strMsg = strMsg & Server.HTMLEncode(objASPError.Description) & "<br>"
		end if

		blnErrorWritten = False

		'IIS에서 발생한 오류 코드를 출력합니다.
		If objASPError.Source > "" Then
			strServername = LCase(Request.ServerVariables("SERVER_NAME"))
			strServerIP = Request.ServerVariables("LOCAL_ADDR")
			strRemoteIP =  Request.ServerVariables("REMOTE_ADDR")

			If (strServerIP = strRemoteIP) And objASPError.File <> "?" Then
				strMsg = strMsg & Server.HTMLEncode(objASPError.File)

				If objASPError.Line > 0 Then strMsg = strMsg & ", line " & objASPError.Line
				If objASPError.Column > 0 Then strMsg = strMsg & ", column " & objASPError.Column
				strMsg = strMsg & "<br>"
				strMsg = strMsg & "<font style=""COLOR:000000; FONT: 8pt/11pt courier new""><b>"
				strMsg = strMsg & Server.HTMLEncode(objASPError.Source) & "<br>"
				If objASPError.Column > 0 Then strMsg = strMsg & String((objASPError.Column - 1), "-") & "^<br>"
				strMsg = strMsg & "</b></font>"
				blnErrorWritten = True
			End If
		End If

		If Not blnErrorWritten And objASPError.File <> "?" Then
			strMsg = strMsg & "<b>" & Server.HTMLEncode(  objASPError.File)
			If objASPError.Line > 0 Then strMsg = strMsg & Server.HTMLEncode(", line " & objASPError.Line)
			If objASPError.Column > 0 Then strMsg = strMsg & ", column " & objASPError.Column
			strMsg = strMsg & "</b><br>"
		End If
		strMsg = strMsg & "</li>"

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
		On Error Resume Next
		If Request.TotalBytes > lngMaxFormBytes Then
			strMsg = strMsg & Server.HTMLEncode(Left(Request.Form, lngMaxFormBytes)) & " . . ."'
		Else
			strMsg = strMsg & Server.HTMLEncode(Request.Form)
		End If
		On Error Goto 0
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

	'// 사용자 정보
	strMsg = strMsg & "<br><br><li>접속User:<br>"
	strMsg = strMsg & session("ssBctID")

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
	cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver")= "110.93.128.95"

	'-> 접근할 포트번호를 설정합니다
	cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25

	'-> 접속시도할 제한시간을 설정합니다
	cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 15

	'-> SMTP 접속 인증방법을 설정합니다
	cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1

	'-> SMTP 서버에 인증할 ID를 입력합니다
	cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "MailSendUser"

	'-> SMTP 서버에 인증할 암호를 입력합니다
	cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "wjddlswjddls"

	cdoConfig.Fields.Update

	Set cdoMessage = CreateObject("CDO.Message")
	Set cdoMessage.Configuration = cdoConfig

'	cdoMessage.To 		= "kobula@10x10.co.kr;tozzinet@10x10.co.kr;kjy8517@10x10.co.kr;errmail@10x10.co.kr;thensi7@10x10.co.kr;corpse2@10x10.co.kr;"
	cdoMessage.To 		= "errmail@10x10.co.kr"
	cdoMessage.From 	= "webserver@10x10.co.kr"
	cdoMessage.SubJect 	= "["&date()&"] WebAdmin 페이지 오류 발생"
	cdoMessage.HTMLBody	= strMsg
	cdoMessage.Send

	Set cdoMessage = nothing
	Set cdoConfig = nothing
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="ko" xml:lang="ko">
<head>
<title>텐바이텐 10X10 = 감성채널 감성에너지..</title>
</head>
<body>
<table width="500" border="0" cellspacing="0" cellpadding="0" align="center">
<tr>
	<td><img src="http://fiximage.10x10.co.kr/web2008/cscenter/sorry_top2.gif" width="500" height="278" /></td>
</tr>
<tr>
	<td height="66" bgcolor="#f7f7f7">
		<table width="100%" border="0" cellspacing="0" cellpadding="0">
		<tr>
			<td align="right"><a href="javascript:history.back()" onFocus="this.blur();"><img src="http://fiximage.10x10.co.kr/web2008/cscenter/btn_back.gif" width="94" height="35" border="0" /></a></td>
		</tr>
		</table>
	</td>
</tr>
</table>
</body>
</html>