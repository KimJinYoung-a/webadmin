<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<%
dim btcid
	btcid= session("ssBctID")
%>
<html>
<head>
<title>[10x10] Business Comunication</title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link rel="stylesheet" href="/css/scm.css" type="text/css">
<script language="JavaScript" src="/js/common.js"></script>
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

function pop_editcompany(){
	var popwin = window.open('/designer/company/editcompany3.asp?menupos=53' ,'op1','width=750,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function pop_10x10_person(){
	var popwin = window.open('/common/pop_10x10_person.asp','op2','width=450,height=450,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function pop_10x10_map(){
	var popwin = window.open('/common/pop_10x10_map.asp','op3','width=650,height=800,scrollbars=yes,resizable=yes')
	popwin.focus();
}

</script>
</head>
<body bgcolor="#FFFFFF" text="#000000" topmargin="0" leftmargin="0">
<!-- 상단 여백 -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
<tr height="5">
	<td align="center"></td>
</tr>
</table>
<!-- 상단 여백 -->
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
        	<a href="/common/offshop/member/offlinemember_list.asp" target="_blank" onMouseOver="this.style.color = 'red'; this.style.fontWeight = 'bold'" onMouseOut="this.style.color = 'black'; this.style.fontWeight = 'normal'">비상연락망</a> |
	        <a href="javascript:pop_10x10_map();" onMouseOver="this.style.color = 'red'; this.style.fontWeight = 'bold'" onMouseOut="this.style.color = 'black'; this.style.fontWeight = 'normal'">텐바이텐 약도</a> |
	        <% 
	        '//가맹점,해외매장일 경우 '/2012.03.02 용만 추가
	        if (session("ssBctDiv")="502") or (session("ssBctDiv")="503") then
	        %>
        		<a href="<%= manageUrl %>/admin/offshop/board/offshop_board.asp" target="_blank" onMouseOver="this.style.color = 'red'; this.style.fontWeight = 'bold'" onMouseOut="this.style.color = 'black'; this.style.fontWeight = 'normal'">매장통합게시판</a>
	        	|
	        	<a href="#" onclick="printbarcode_on_off_multi(); return false;" onMouseOver="this.style.color = 'red'; this.style.fontWeight = 'bold'" onMouseOut="this.style.color = 'black'; this.style.fontWeight = 'normal'" >바코드출력</a>
	    	<% end if %>
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
			<input type=button value="☜" onClick="javascript:WindowMinSize()">
		</div>
		<div id=WINSIZE style="display:none">창 축소하기
			<input type=button value="☞" onClick="javascript:WindowMaxSize()">
		</div>
	</td>
    <td align="right">
        <b><%=session("ssBctID")%>(<%=session("ssBctCname")%>)</b> 님이 로그인 하셨습니다.
    	&nbsp;
    	<a href="/login/dologout.asp" target="_top"><img src="/images/icon_logout.gif" width="64" height="17" border="0" align="absbottom"></a>
    </td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<!-- 표 중간바 끝-->

</body>
</html>