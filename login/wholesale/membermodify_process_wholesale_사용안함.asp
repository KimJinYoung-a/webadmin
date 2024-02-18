<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<% response.charset = "utf-8" %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : 텐바이텐 대량구매 사이트
' Hieditor : 2013.07.15 한용민 생성
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib_utf8.asp" -->
<!-- #include virtual="/lib/util/md5.asp" -->
<!-- #include virtual="/lib/util/tenEncUtil.asp" -->
<!-- #include virtual="/lib/util/base64unicode.asp" -->
<!-- #include virtual="/lib/classes/member/userloginclass_wholesale.asp" -->
<%
dim wwwwholesale, wwwwSSLholesale
IF application("Svr_Info")="Dev" THEN
	wwwwholesale  	= "http://testwholesale.10x10.co.kr"	
	wwwwSSLholesale	= "https://testwholesale.10x10.co.kr"	
ELSE
	wwwwholesale  	= "http://wholesale.10x10.co.kr"
	wwwwSSLholesale	= "https://wholesale.10x10.co.kr"	
END IF

Dim vQuery, vGubun, vUserID, vManName, vManPhone, vManEmail, vOldPass, vNewPass1, vNewPass2, vResult, vdeliver_name, vdeliver_phone
dim vshopzipcode, vshopaddr1, vshopaddr2, vreturn_zipcode, vreturn_address, vreturn_address2
	vUserID 	= tenDec(request.Cookies("winfo")("userid"))
	vGubun		= Request.Form("gubun")
	vManName	= requestcheckvar(Request.Form("manname"),32)
	vManPhone	= requestcheckvar(Request.Form("manphone"),16)
	vManEmail	= requestcheckvar(Request.Form("manemail"),128)
	vOldPass	= Request.Form("oldpass")
	vNewPass1	= Request.Form("newpass1")
	vNewPass2	= Request.Form("newpass2")
	vdeliver_name	= requestcheckvar(Request.Form("deliver_name"),32)
	vdeliver_phone	= requestcheckvar(Request.Form("deliver_phone"),32)
	'vshopzipcode = requestcheckvar(Request.Form("shopzipcode"),7)
	vshopaddr1 = requestcheckvar(Request.Form("shopaddr1"),128)
	vshopaddr2 = requestcheckvar(Request.Form("shopaddr2"),128)
	'vreturn_zipcode = requestcheckvar(Request.Form("return_zipcode"),7)
	vreturn_address	= requestcheckvar(Request.Form("return_address"),128)
	vreturn_address2 = requestcheckvar(Request.Form("return_address2"),128)
%>
<form name="loginForm" action="<%=wwwwholesale%>/login/SSLreload.asp" method="post" style="margin:0px;">
	<input type="hidden" name="mode" />
</form>
<%
If vGubun = "info" Then
	response.Cookies("winfo")("manemail") = vManEmail

	vQuery = "EXECUTE [db_shop].[dbo].[sp_Ten_UserInfo_proc_wholesale]" & vbcrlf
	vQuery = vQuery & " '" & trim(vGubun) & "', '" & trim(vUserID) & "', '" & html2db(trim(vManName)) & "', '" & trim(vManPhone) & "', '" & html2db(trim(vManEmail)) & "', ''" & vbcrlf
	vQuery = vQuery & " , '" & html2db(trim(vdeliver_name)) & "', '" & trim(vdeliver_phone) & "', '', '" & html2db(trim(vshopaddr1)) & "', '" & html2db(trim(vshopaddr2)) & "'" & vbcrlf
	vQuery = vQuery & " , '', '" & html2db(trim(vreturn_address)) & "', '" & html2db(trim(vreturn_address2)) & "'" & vbcrlf

	'response.write vQuery & "<Br>"
	dbget.execute vQuery

	Response.Write "<script type='text/javascript'>document.loginForm.mode.value='l11'; document.loginForm.submit();</script>"
	dbget.close() : Response.End

ElseIf vGubun = "changepw" Then
	'### 비번 체크.
	vResult = "x"
	vQuery = "EXECUTE [db_shop].[dbo].[sp_Ten_UserInfo_oldpassCheck_wholesale] '" & trim(vUserID) & "', '" & trim(md5(vOldPass)) & "'"

	'response.write vQuery & "<Br>"	
	rsget.Open vQuery,dbget,1
	If Not rsget.Eof Then
		vResult = rsget(0)
	End If
	rsget.close()

	If vResult = "ok" Then
		vQuery = "EXECUTE [db_shop].[dbo].[sp_Ten_UserInfo_proc_wholesale] '" & trim(vGubun) & "', '" & trim(vUserID) & "', '', '', '', '" & trim(md5(vNewPass1)) & "','','','','','','','',''"
		
		'response.write vQuery & "<Br>"			
		dbget.execute vQuery
		Response.Write "<script type='text/javascript'>document.loginForm.mode.value='l11'; document.loginForm.submit();</script>"
		dbget.close() : Response.End
	ElseIf vResult = "x" Then
		Response.Write "<script type='text/javascript'>document.loginForm.mode.value='l12'; document.loginForm.submit();</script>"
		dbget.close() : Response.End
	End If
End If
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->