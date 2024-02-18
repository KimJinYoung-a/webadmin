 <%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 전자결재 팝업 index
' History : 2011.03.14 정윤정  생성
'###########################################################
%>
<!-- #include virtual="/lib/util/htmllib.asp"-->  
<!-- #include virtual="/admin/incSessionAdmin.asp" -->   
<%
Dim sDirectUrl, iRectMenu, sListURL
Dim ireportidx,ipayrequestidx
sDirectUrl = requestCheckvar(Request("sDUrl"),100)
iRectMenu =	requestCheckvar(Request("iRM"),10)
ireportidx =  requestCheckvar(Request("iridx"),10)
ipayrequestidx=  requestCheckvar(Request("ipridx"),10)
IF iRectMenu = "" THEN iRectMenu = "F100"
IF ireportidx = "" THEN ireportidx = 0
IF ipayrequestidx = "" THEN ipayrequestidx = 0
sListURL = "about:blank"
 
IF iRectMenu ="F110" THEN
 sDirectUrl = "/admin/approval/eapp/payrequestreceive.asp?iPRS=1&iAS=1&blnL=0&igbn=1&ipridx="&ipayrequestidx
 IF ipayrequestidx > 0 THEN sListURL = "/admin/approval/eapp/confirmpayrequest.asp?iridx="&ireportidx&"&ipridx="&ipayrequestidx&"&iAS=1&igbn=1" 
ELSEIF iRectMenu ="F711" THEN
 sDirectUrl = "/admin/approval/eapp/payrequestreceive.asp?iPRS=7&iAS=1&igbn=1&ipridx="&ipayrequestidx
 	IF ipayrequestidx > 0 THEN sListURL = "/admin/approval/eapp/confirmpayrequest.asp?iridx="&ireportidx&"&ipridx="&ipayrequestidx&"&iAS=1&igbn=1"
ELSEIF iRectMenu ="F971" THEN
 sDirectUrl = "/admin/approval/eapp/payrequestreceive.asp?iPRS=9&iAS=7&igbn=1&ipridx="&ipayrequestidx
 	IF ipayrequestidx > 0 THEN sListURL = "/admin/approval/eapp/confirmpayrequest.asp?iridx="&ireportidx&"&ipridx="&ipayrequestidx&"&iAS=7&igbn=1"
ELSEIF iRectMenu ="F551" THEN
 sDirectUrl = "/admin/approval/eapp/payrequestreceive.asp?iPRS=5&iAS=5&igbn=1&ipridx="&ipayrequestidx    
 	IF ipayrequestidx > 0 THEN sListURL = "/admin/approval/eapp/confirmpayrequest.asp?iridx="&ireportidx&"&ipridx="&ipayrequestidx&"&iAS=5&igbn=1"  
ELSE
	 sDirectUrl = "/admin/approval/eapp/payrequestreceive.asp?iPRS=1&iAS=0&blnL=0&igbn=1&ipridx="&ipayrequestidx
	 	IF ipayrequestidx > 0 THEN sListURL = "/admin/approval/eapp/confirmpayrequest.asp?iridx="&ireportidx&"&ipridx="&ipayrequestidx&"&iAS=0&igbn=1" 
END IF	 
%>
<html> 
<head>
<title>[10x10] Business Comunication</title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr"> 
</head> 
<frameset id="eappmenuset" cols="180,*" frameborder="NO" border="0" framespacing="0">
	<frame id="eappmenu" src="/admin/approval/eapp/ft_leftmenu.asp?iRM=<%=iRectMenu%>" scrolling="No" style="border-right:1px solid gray;"> 
  <frameset id="eappsubmenuset" cols="400,*" frameborder="NO" border="0" framespacing="0">
		<frame id="eappcontents" src="<%=sDirectUrl%>" scrolling="No" style="border-right:1px solid gray;">
	 	<frame id="eappDetail" src="<%=sListURL%>" scrolling="auto">
	</frameset>  
</frameset> 
<noframes>
<body bgcolor="#FFFFFF" text="#000000">
</body></noframes>
</html>
 	