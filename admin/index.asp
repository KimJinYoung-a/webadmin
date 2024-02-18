<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<%
'/https 로 안들어 오면 전부 리다이렉트시킴.
if request.servervariables("HTTPS")="off" THEN
	IF application("Svr_Info")="Dev" THEN
		if G_IsLocalDev then
		else
			response.redirect(getSCMSSLURL & "/admin/index.asp")
		end if
	else
		response.redirect(getSCMSSLURL & "/admin/index.asp")
	end if
end if
%>
<html>
<head>
<title>[10x10] Business Comunication</title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link REL="SHORTCUT ICON" href="http://fiximage.10x10.co.kr/icons/10x10SCM.ico">
</head>

<frameset rows="60,*" frameborder="NO" border="0" framespacing="0" cols="*">
    <frame name="header" scrolling="NO" noresize src="/admin/lib/frameheader.asp" >
    <frameset name="menuset" cols="180,*" frameborder="NO" border="0" framespacing="0">
        <frame name="menu" src="/admin/menu/left_menu.asp" scrolling="AUTO">
        <frame name="contents" src="/admin/scmmain.asp">
    </frameset>
</frameset>
<noframes>
<body bgcolor="#FFFFFF" text="#000000">
</body></noframes>
</html>
