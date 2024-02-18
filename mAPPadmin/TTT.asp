<%@ codepage="65001" language="VBScript" %>
<% option explicit %>
<!-- #include virtual="/mAppadmin/inc/incCommon.asp" -->
<!-- #include virtual="/mAPPadmin/incSessionmAPPadmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAppNotiopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/mAppadmin/inc/incHeader.asp" -->
<script type="text/javascript">

</script>
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
안녕하세요!! <%= session("mAppBctId") %> 님
<table>
<tr>
<td>한글 한글11</td>
</tr>
</table>

<br><br><br>

<a href="/mAppadmin/test.asp" data-ajax="false">업무협조</a>


<input type="button" value="새로고침" id="btn-reload" data-role="button" rel="external" />

<input type="button" value="로그아웃" id="btn-logout" data-role="button" rel="external" />

</body>
</html>
<!-- #include virtual="/lib/db/dbAppNoticlose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
