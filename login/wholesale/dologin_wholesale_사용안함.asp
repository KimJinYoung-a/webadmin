<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<% response.charset = "utf-8" %>
<%
'###########################################################
' Description : 텐바이텐 대량구매 사이트
' Hieditor : 2013.07.15 한용민 생성
'###########################################################
%>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/md5.asp" -->
<!-- #include virtual="/lib/util/tenEncUtil.asp" -->
<!-- #include virtual="/lib/util/base64unicode.asp" -->
<!-- #include virtual="/lib/classes/member/userloginclass_wholesale.asp" -->

<%
dim wwwwholesale
IF application("Svr_Info")="Dev" THEN
	wwwwholesale  	= "http://testwholesale.10x10.co.kr"	
ELSE
	wwwwholesale  	= "http://wholesale.10x10.co.kr"
END IF

'Response.Write "<script type='text/javascript'>alert('dologin');</script>"

dim vmode, vBackPath, vStrGetData
	vBackPath 	= ReplaceRequestSpecialChar(Request.Form("backpath"))
	vStrGetData	= ReplaceRequestSpecialChar(request.Form("strGD"))

If vBackPath = "" Then
	IF application("Svr_Info")="Dev" THEN
		vBackPath = wwwwholesale & "/" & "_index.asp"
	else
		vBackPath = wwwwholesale
	end if
Else

	'/로그인페이지에서 타고 들어온 경우라면
	if vBackPath="/login.asp" then
		IF application("Svr_Info")="Dev" THEN
			vBackPath = wwwwholesale & "/" & "_index.asp"
		else
			vBackPath = wwwwholesale
		end if
	else
		vBackPath = vBackPath & "?" & vStrGetData
	end if
End IF	
%>
<form name="loginForm" action="<%=wwwwholesale%>/login/SSLreload.asp" method="post" style="margin:0px;">
	<input type="hidden" name="strPath" value="<%= vBackPath %>" />
	<input type="hidden" name="mode" value="<%= vmode %>" />
</form>
<%
Dim cLogin, vUserID, vUserPW,vEnc_UserPW, vResult, sqlStr
	vUserID 	= Trim(Request.Form("userid"))
	vUserPW 	= Trim(Request.Form("userpw"))
	vEnc_UserPW = Md5(vUserPW)
	
	If vUserID = "" Then
		Response.Write "<script type='text/javascript'>document.loginForm.mode.value='l1'; document.loginForm.submit();</script>"
		dbget.close() : Response.End
	End IF
	
	If vUserPW = "" Then
		Response.Write "<script type='text/javascript'>document.loginForm.mode.value='l2'; document.loginForm.submit();</script>"
		dbget.close() : Response.End
	End IF
	
SET cLogin = New CwholesaleLogin
	cLogin.FRectUserID = vUserID
	cLogin.FRectUserPW = vEnc_UserPW
	cLogin.GetLoginData
	
	vResult = cLogin.FOneItem.FResult

	'/로그인 로그 저장
	sqlStr = "exec [db_shop].[dbo].[usp_UserLogin_multiSite_AddLog] 'WSLWEB','"& vUserID &"','"& Request.ServerVariables("REMOTE_ADDR") &"',"& vResult &""
	
	'response.write sqlStr & "<Br>"
	dbget.execute sqlStr

	If vResult = "0" Then
		IF application("Svr_Info")="Dev" THEN
			'response.Cookies("winfo").domain = "testwholesale.10x10.co.kr"
			response.Cookies("winfo").domain = "10x10.co.kr"
		else
			'response.Cookies("winfo").domain = "wholesale.10x10.co.kr"
			response.Cookies("winfo").domain = "10x10.co.kr"
		end if
		response.Cookies("winfo")("userid") = tenEnc(cLogin.FOneItem.FUserID)
		response.Cookies("winfo")("shix") = HashTenID(cLogin.FOneItem.FUserID)
		response.Cookies("winfo")("shopname") = cLogin.FOneItem.FShopName
		response.Cookies("winfo")("currencyunit") = cLogin.FOneItem.FcurrencyUnit
		response.Cookies("winfo")("currencychar") = cLogin.FOneItem.FcurrencyChar
		response.Cookies("winfo")("countrylangcd") = cLogin.FOneItem.fcountrylangcd
		response.Cookies("winfo")("manemail") = cLogin.FOneItem.Fmanemail
		response.Cookies("winfo")("GetLogincountrylangcd") = cLogin.FOneItem.fcountrylangcd
		
		'####### 일단은 위에만 저장. 필요시 아래 주석만 풀어주면 됨.
		'response.Cookies("winfo")("shopdiv") = cLogin.FOneItem.FShopdiv
		'response.Cookies("winfo")("groupid") = cLogin.FOneItem.Fgroupid
		'response.Cookies("winfo")("countrynamekr") = cLogin.FOneItem.FcountryNamekr
		'response.Cookies("winfo")("ismobileusing") = cLogin.FOneItem.Fismobileusing
		'response.Cookies("winfo")("shopcountrycode") = cLogin.FOneItem.FshopCountryCode
		
		Response.Write "<script type='text/javascript'>document.loginForm.mode.value='l10'; document.loginForm.submit();</script>"
		dbget.close() : Response.End

	ElseIf vResult = "1" Then	'### 아이디 틀릴때.
		Response.Write "<script type='text/javascript'>document.loginForm.mode.value='l3'; document.loginForm.submit();</script>"
		dbget.close() : Response.End

	ElseIf vResult = "2" Then	'### 비번 틀릴때.
		Response.Write "<script type='text/javascript'>document.loginForm.mode.value='l4'; document.loginForm.submit();</script>"
		dbget.close() : Response.End

	End If
SET cLogin = Nothing
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->