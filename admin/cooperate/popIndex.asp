<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cooperate/chk_auth.asp"-->

<%
	Dim vMenu, vMenuLink
	vMenu = NullFillWith(requestCheckVar(Request("mn"),10),"C00")
	
	SELECT CASE vMenu
		CASE "C00" : vMenuLink = "/admin/cooperate/main.asp"						'### 업무협조 메인
		CASE "C10" : vMenuLink = "/admin/cooperate/my_cooperate.asp"				'### 보낸업무협조 전체 리스트
		CASE "C11" : vMenuLink = "/admin/cooperate/my_cooperate.asp?doc_status=1"	'### 보낸업무협조 기안 리스트
		CASE "C12" : vMenuLink = "/admin/cooperate/my_cooperate.asp?doc_status=2"	'### 보낸업무협조 작업중 리스트
		CASE "C13" : vMenuLink = "/admin/cooperate/my_cooperate.asp?doc_status=3"	'### 보낸업무협조 작업완료 리스트
		CASE "C14" : vMenuLink = "/admin/cooperate/my_cooperate.asp?doc_status=4"	'### 보낸업무협조 반려 리스트
		CASE "C15" : vMenuLink = "/admin/cooperate/my_cooperate.asp?doc_status=5"	'### 보낸업무협조 반려 후 최종완료리스트
		CASE "C16" : vMenuLink = "/admin/cooperate/my_cooperate.asp?doc_status=6"	'### 보낸업무협조 참조 리스트
		CASE "C20" : vMenuLink = "/admin/cooperate/"								'### 받은업무협조 전체 리스트
		CASE "C21" : vMenuLink = "/admin/cooperate/?doc_status=1"					'### 받은업무협조 기안 리스트
		CASE "C22" : vMenuLink = "/admin/cooperate/?doc_status=2"					'### 받은업무협조 작업중 리스트
		CASE "C23" : vMenuLink = "/admin/cooperate/?doc_status=3"					'### 받은업무협조 작업완료 리스트
		CASE "C24" : vMenuLink = "/admin/cooperate/?doc_status=4"					'### 받은업무협조 반려 리스트
		CASE "C25" : vMenuLink = "/admin/cooperate/?doc_status=5"					'### 받은업무협조 반려 후 최종완료리스트
		CASE "C26" : vMenuLink = "/admin/cooperate/?doc_status=6"					'### 받은업무협조 참조 리스트
	END SELECT
%>

<html> 
<head>
<title>[10x10] SCM 업무협조</title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr"> 
</head> 
<frameset id="coopmenuset" cols="180,*" frameborder="NO" border="0" framespacing="0">
<% If g_VertiHoriz = "v" OR g_VertiHoriz = "" Then %>
	<frame id="coopmenu" src="/admin/cooperate/lib/leftmenu.asp?mn=<%=vMenu%>" scrolling="No" style="border-right:1px solid gray;" />
	<% If vMenu = "C00" Then %>
		<frame name="coopcontents" id="coopcontents" src="<%=vMenuLink%>" scrolling="auto" />
	<% Else %>
		<frameset id="coopsubmenuset" cols="600,*" frameborder="NO" border="0" framespacing="0">
			<frame name="coopcontents" id="coopcontents" src="<%=vMenuLink%>" scrolling="no" style="border-right:1px solid gray;" />
			<frame name="coopDetail" id="coopDetail" src="about:blank" scrolling="auto" />
		</frameset> 
	<% End If %>
<% Else %>
	<frame id="coopmenu" src="/admin/cooperate/lib/leftmenu.asp?mn=<%=vMenu%>" scrolling="No" style="border-right:1px solid gray;" />
	<frameset id="coopsubmenuset" rows="500,*" frameborder="NO" border="0" framespacing="0">
		<% If vMenu = "C00" Then %>
			<frame name="coopcontents" id="coopcontents" src="<%=vMenuLink%>" scrolling="auto" />
		<% Else %>
			<frame name="coopcontents" id="coopcontents" src="<%=vMenuLink%>" scrolling="auto" style="border-bottom:2px solid gray;" />
			<frame name="coopDetail" id="coopDetail" src="about:blank" scrolling="auto" />
		<% End If %>
	</frameset>
<% End If %>
</frameset>
<noframes>
<body bgcolor="#FFFFFF" text="#000000">
</body>
</noframes>
</html>