<%
dim btcid 
btcid= session("ssBctID")
if (btcid="") then response.End
%>
<html>
<head>
<title>[10x10] Business Comunication</title>
<link rel="stylesheet" href="/css/scm.css" type="text/css">
</head>
<body bgcolor="#FFFFFF" text="#000000" topmargin="0" leftmargin="0">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr bgcolor="#666699"> 
    <td><a href="/company/" target="_top"><img src="/images/bst_title.gif" width="500" height="38" border="0"></a></td>
  </tr>
</table>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr bgcolor="#DDDDDD"> 
    <td width="65%" bgcolor="#DDDDDD">
    &nbsp;&nbsp;<font size="2"><a href="/login/dologout.asp" target="_top">[로그아웃x]</a></font>
     </td>
    <td align="right" width="35%" valign="middle">
      <b><font size="2"><%=session("ssBctId")%></font></b>
      <font face="Verdana, Arial, Helvetica, sans-serif" size="2"> (<%=session("ssBctCname")%>) 님이 로그인 하셨습니다.&nbsp;&nbsp;</font></font>
    </td>
  </tr>
</table>
</body>
</html>