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
IF iRectMenu = "" THEN iRectMenu = "M999"
IF ireportidx = "" THEN ireportidx = 0
IF ipayrequestidx = "" THEN ipayrequestidx = 0
sListURL = "about:blank"
		
 IF iRectMenu ="M010" THEN
	sDirectUrl = "/admin/approval/eapp/eappsend.asp?iRS=0&iridx="&ireportidx&"&iRM="&iRectMenu 
	IF ireportidx > 0 THEN sListURL = "/admin/approval/eapp/modeapp.asp?iridx="&ireportidx&"&iRM="&iRectMenu 
ELSEIF iRectMenu ="T010" THEN
	sDirectUrl = "/admin/approval/eapp/eappsend.asp?iRS=-1&iridx="&ireportidx&"&iRM="&iRectMenu   
ELSEIF iRectMenu ="T020" THEN
	sDirectUrl = "/admin/approval/eapp/payrequestsend.asp?iprs=-1&iridx="&ireportidx&"&ipridx="&ipayrequestidx&"&iRM="&iRectMenu   
ELSEIF iRectMenu ="T011" THEN
	sDirectUrl = "/admin/approval/eapp/eappreceive.asp?iRS=-1&iridx="&ireportidx&"&iRM="&iRectMenu   
ELSEIF iRectMenu ="T021" THEN
	sDirectUrl = "/admin/approval/eapp/payrequestreceive.asp?iprs=-1&iridx="&ireportidx&"&ipridx="&ipayrequestidx&"&iRM="&iRectMenu   
ELSEIF iRectMenu ="M011" THEN
	sDirectUrl = "/admin/approval/eapp/eappsend.asp?iRS=1&iridx="&ireportidx&"&iRM="&iRectMenu
	IF ireportidx > 0 THEN sListURL = "/admin/approval/eapp/modeapp.asp?iridx="&ireportidx&"&iRM="&iRectMenu   
ELSEIF iRectMenu ="M017" THEN
	sDirectUrl = "/admin/approval/eapp/eappsend.asp?iRS=7&iridx="&ireportidx&"&iRM="&iRectMenu  
	IF ireportidx > 0 THEN sListURL = "/admin/approval/eapp/modeapp.asp?iridx="&ireportidx&"&iRM="&iRectMenu 
ELSEIF iRectMenu ="M013" THEN			
	sDirectUrl = "/admin/approval/eapp/eappsend.asp?iRS=3&iridx="&ireportidx&"&iRM="&iRectMenu  
	IF ireportidx > 0 THEN sListURL = "/admin/approval/eapp/modeapp.asp?iridx="&ireportidx&"&iRM="&iRectMenu 
ELSEIF iRectMenu ="M015" THEN
	sDirectUrl = "/admin/approval/eapp/eappsend.asp?iRS=5&iridx="&ireportidx&"&iRM="&iRectMenu  
	IF ireportidx > 0 THEN sListURL = "/admin/approval/eapp/modeapp.asp?iridx="&ireportidx&"&iRM="&iRectMenu 
ELSEIF iRectMenu ="M020" THEN
sDirectUrl = "/admin/approval/eapp/payrequestsend.asp?iprs=0&iridx="&ireportidx&"&ipridx="&ipayrequestidx&"&iRM="&iRectMenu  
	IF ipayrequestidx > 0 THEN sListURL = "/admin/approval/eapp/regpayrequest.asp?iridx="&ireportidx&"&ipridx="&ipayrequestidx&"&iRM="&iRectMenu   
ELSEIF iRectMenu ="M021" THEN
	sDirectUrl = "/admin/approval/eapp/payrequestsend.asp?iprs=1&iridx="&ireportidx&"&ipridx="&ipayrequestidx&"&iRM="&iRectMenu  
ELSEIF iRectMenu ="M027" THEN
	sDirectUrl = "/admin/approval/eapp/payrequestsend.asp?iprs=7&iridx="&ireportidx&"&ipridx="&ipayrequestidx&"&iRM="&iRectMenu    
ELSEIF iRectMenu ="M029" THEN	
	sDirectUrl = "/admin/approval/eapp/payrequestsend.asp?iprs=9&iridx="&ireportidx&"&ipridx="&ipayrequestidx&"&iRM="&iRectMenu   
ELSEIF iRectMenu ="M025" THEN
	sDirectUrl = "/admin/approval/eapp/payrequestsend.asp?iprs=5&iridx="&ireportidx&"&ipridx="&ipayrequestidx&"&iRM="&iRectMenu   
ELSEIF iRectMenu ="M028" THEN
	sDirectUrl = "/admin/approval/eapp/payrequestdoc.asp?iridx="&ireportidx&"&ipridx="&ipayrequestidx&"&iRM="&iRectMenu   	
ELSEIF iRectMenu ="M110" THEN
	sDirectUrl = "/admin/approval/eapp/eappreceive.asp?iRS=1&iAS=0&iridx="&ireportidx&"&iRM="&iRectMenu 
		IF ireportidx > 0 THEN sListURL = "/admin/approval/eapp/confirmeapp.asp?iridx="&ireportidx&"&iRM="&iRectMenu  
ELSEIF iRectMenu ="M111" THEN
	sDirectUrl = "/admin/approval/eapp/eappreceive.asp?iRS=1&iAS=1&iridx="&ireportidx&"&iRM="&iRectMenu  
ELSEIF iRectMenu ="M171" THEN	
	sDirectUrl = "/admin/approval/eapp/eappreceive.asp?iRS=7&iAS=1&iridx="&ireportidx&"&iRM="&iRectMenu  
ELSEIF iRectMenu ="M113" OR iRectMenu ="M133" THEN
	sDirectUrl = "/admin/approval/eapp/eappreceive.asp?iRS=1&iAS=3&iridx="&ireportidx&"&iRM="&iRectMenu  
ELSEIF iRectMenu ="M115"OR iRectMenu ="M155"  THEN	
	sDirectUrl = "/admin/approval/eapp/eappreceive.asp?iRS=1&iAS=5&iridx="&ireportidx&"&iRM="&iRectMenu  
ELSEIF iRectMenu ="M112" THEN
	sDirectUrl = "/admin/approval/eapp/eapprefer.asp?iridx="&ireportidx&"&iRM="&iRectMenu  
ELSEIF iRectMenu ="M120" THEN
 sDirectUrl = "/admin/approval/eapp/payrequestview.asp?iridx="&ireportidx&"&iRM="&iRectMenu  
ELSEIF iRectMenu ="F100" THEN
 sDirectUrl = "/admin/approval/eapp/payrequestreceive.asp?iPRS=1&iAS=0&blnL=0&iRM="&iRectMenu
ELSEIF iRectMenu ="F110" THEN
 sDirectUrl = "/admin/approval/eapp/payrequestreceive.asp?iPRS=1&iAS=1&blnL=0&iRM="&iRectMenu
ELSEIF iRectMenu ="F711" THEN
 sDirectUrl = "/admin/approval/eapp/payrequestreceive.asp?iPRS=7&iAS=1&iRM="&iRectMenu
ELSEIF iRectMenu ="F971" THEN
 sDirectUrl = "/admin/approval/eapp/payrequestreceive.asp?iPRS=9&iAS=7&iRM="&iRectMenu
ELSEIF iRectMenu ="F551" THEN
 sDirectUrl = "/admin/approval/eapp/payrequestreceive.asp?iPRS=5&iAS=5&iRM="&iRectMenu    
ELSE
	sDirectUrl = "/admin/approval/eapp/main.asp"
END IF	 
%>
<html> 
<head>
<title>[10x10] Business Comunication</title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr"> 
</head> 
<frameset id="eappmenuset" name="eappmenuset" cols="180,*" frameborder="NO" border="0" framespacing="0">
	<frame id="eappmenu" name="eappmenu" src="/admin/approval/eapp/leftmenu.asp?iRM=<%=iRectMenu%>" scrolling="No" style="border-right:1px solid gray;">
	<%IF iRectMenu = "M999"   THEN%>
	<frame id="eappcontents" name="eappcontents" src="<%=sDirectUrl%>" scrolling="auto"> 
	<%ELSE%>
  <frameset id="eappsubmenuset" name="eappsubmenuset" cols="500,*" frameborder="NO" border="0" framespacing="0">
		<frame id="eappcontents" name="eappcontents" src="<%=sDirectUrl%>" scrolling="No" style="border-right:1px solid gray;">
	 	<frame id="eappDetail" name="eappDetail" src="<%=sListURL%>" scrolling="auto">
	</frameset> 
  <%END IF%>
</frameset> 
<noframes>
<body bgcolor="#FFFFFF" text="#000000">
</body></noframes>
</html>
 	