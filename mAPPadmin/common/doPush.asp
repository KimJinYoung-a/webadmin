<%@ codepage="65001" language="VBScript" %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control", "no-cache"
Response.AddHeader "Expires", "0"
Response.AddHeader "Pragma", "no-cache"
%>
<!-- #include virtual="/mAppadmin/inc/incUTF8.asp" -->
<!-- #include virtual="/mAppadmin/inc/incCommon.asp" -->
<!-- #i//n//cl//ude vi//rt//ual="/mAPPadmin/incSessionmAPPadmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAppNotiopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/email/pushServiceLib.asp" -->
<%

dim refer
refer = request.ServerVariables("HTTP_REFERER")

dim mode, receiverId, msg
dim appkey, receiver_userid, receiver_role, message, url

mode = requestCheckVar(html2db(request("mode")),32)
receiverId = requestCheckVar(html2db(request("receiverId")),32)
msg = requestCheckVar(html2db(request("msg")),320)


if (mode = "sendOnePush") then
	appkey				= "admin_app"
	receiver_userid		= receiverId
	receiver_role		= "20"
	message				= msg
	url					= "http://webadmin.10x10.co.kr/mAppadmin/TTT.asp?ridx=111"


	Call sendPushMessage(appkey, receiver_userid, receiver_role, message, url)

	response.write "<script>alert('전송요청되었습니다.');</script>"
	response.write "<script>location.replace('" + CStr(refer) + "');</script>"
else
	'
end if

%>
<!-- #include virtual="/lib/db/dbAppNoticlose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
