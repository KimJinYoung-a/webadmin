<!-- #include virtual="/lib/classes/admin/OLDmenucls.asp"-->
<%
dim menupos, imenupos, menuposStr
menupos = request("menupos")
if menupos ="" then menupos=1
set imenupos = new CMenu
menuposStr = imenupos.getMenuPos(menupos)
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<script language="JavaScript" src="/js/xl.js"></script>
<script language="JavaScript" src="/js/common.js"></script>
<script language="JavaScript" src="/js/report.js"></script>
<link rel="stylesheet" href="/css/scm.css" type="text/css">
</head>
<body>
<table width="700" border="0" class="a">
	<tr>
		<td><%= menuposStr %></td>
	</tr>
</table>
<%
set imenupos = Nothing
%>