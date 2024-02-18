<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/tenmember/incSessionTenMember.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%

dim login_username, login_empno

login_empno		= session("ssBctSn")
login_username	= session("ssBctCname")

%>

<html>
<head>
<title>[10x10] Business Comunication</title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link rel="stylesheet" href="/css/scm.css" type="text/css">
<script language='javascript'>
function WindowMinSize(){
	parent.document.all('menuset').cols = "20,*";
	document.all.WINSIZE[0].style.display = "none";
	document.all.WINSIZE[1].style.display = "";
}

function WindowMaxSize(){
	parent.document.all('menuset').cols = "180,*";
	document.all.WINSIZE[0].style.display = "";
	document.all.WINSIZE[1].style.display = "none";
}
</script>
</head>
<body bgcolor="#FFFFFF" text="#000000" topmargin="0" leftmargin="0">

<% if (application("Svr_Info")="Dev") then %>
	<center><b><font color="red">This is <%= Year(now) %> Test Server...</font></b></center>
<% else %>
	<!-- 상단 여백 -->
	<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
		<tr height="5">
			<td></td>
		</tr>
	</table>
	<!-- 상단 여백 -->
<% end if %>

<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
   	<tr height="10" valign="bottom" bgcolor="F4F4F4">
	        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top" bgcolor="F4F4F4">
	        <td background="/images/tbl_blue_round_04.gif"></td>
	        <td valign="center" bgcolor="F4F4F4">
	        	<img src="/images/admin_logo_10x10.jpg" width="90" height="25" align="absbottom">
	        	<b>10x10 Business Communication Tool</b>
	        </td>
	        <td valign="center" align="right" bgcolor="F4F4F4">
				<!-- 여기에 상단 메뉴 추가 -->
            </td>
	        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
<!-- 표 상단바 끝-->

<!-- 표 중간바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
    <tr bgcolor="#CCCCCC" height="20">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td width="170" align="right">
			<div id=WINSIZE style="display:">창 확대하기
				<input type=button class="button" value="☜" onClick="javascript:WindowMinSize()">
			</div>
			<div id=WINSIZE style="display:none">창 축소하기
				<input type=button class="button" value="☞" onClick="javascript:WindowMaxSize()">
			</div>
		</td>
        <td align="right">
	        <b><%= login_username %>(<%= login_empno %>)</b> 님이 로그인 하셨습니다.
	    	&nbsp;
	    	<a href="/login/dologout.asp" target="_top"><img src="/images/icon_logout.gif" width="64" height="17" border="0" align="absbottom"></a>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- 표 중간바 끝-->

</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->