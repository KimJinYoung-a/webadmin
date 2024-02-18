<%@ codepage="65001" language="VBScript" %>
<% option explicit %>
<!-- #include virtual="/mAppadmin/inc/incUTF8.asp" -->
<!-- #include virtual="/mAppadmin/inc/incCommon.asp" -->
<%
    session.abandon
    Response.Cookies("mAppADM").Domain      = manageDomain
    Response.cookies("mAppADM").expires     = date -1

    response.redirect "/mAPPadmin/login.asp"
%>
