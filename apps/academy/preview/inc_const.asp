<%
'-----------------------------------------------------------------------
' 전역변수 선언
'-----------------------------------------------------------------------
DIM CAddDetailSpliter : CAddDetailSpliter= CHR(3)&CHR(4)

 Dim staticImgUrl,uploadUrl,wwwUrl,SSLUrl, g_AdminURL,fingersImgUrl,mailzinUrl , webImgUrl, www1Url, SSL1Url
 Dim googleANAL_EXTSCRIPT ''Google analytics 관련 변수 선언 신규 GA관련
 Dim IsDevSever : IsDevSever = false
 
 Dim DocSvrAddr, DocSvrPort, DocAuthCode
 IF application("Svr_Info")="Dev" THEN
 	staticImgUrl	= "http://testimgstatic.10x10.co.kr"	'테스트
 	uploadUrl		= "http://testimage.thefingers.co.kr"
 	fingersImgUrl	= "http://testimage.thefingers.co.kr"
	webImgUrl		= "http://testwebimage.10x10.co.kr"
 	wwwUrl			= "http://testm.thefingers.co.kr"
 	SSLUrl			= "http://testm.thefingers.co.kr"
 	www1Url         = "http://testm.thefingers.co.kr"
	SSL1Url         = "http://testm.thefingers.co.kr"
	
 	g_AdminURL		= "http://testwebadmin.10x10.co.kr"
	mailzinUrl			= "http://testmailzine.10x10.co.kr"
 	
 	IsDevSever      = True
 	
 	'DocSvrAddr      = "61.252.133.4"
 	'DocSvrPort      = "6167"
 ELSE
 	staticImgUrl	= "http://imgstatic.10x10.co.kr"
 	uploadUrl		= "http://oimage.thefingers.co.kr"
 	fingersImgUrl	= "http://image.thefingers.co.kr"
 	webImgUrl		= "http://webimage.10x10.co.kr"
 	wwwUrl			= "http://m.thefingers.co.kr"
	SSLUrl			= "https://m.thefingers.co.kr"
	www1Url         = "http://m1.thefingers.co.kr"
	SSL1Url         = "https://m1.thefingers.co.kr"
	
	g_AdminURL		= "http://webadmin.10x10.co.kr"

	mailzinUrl			= "http://mailzine.10x10.co.kr"
	
 	'DocSvrAddr      = ""	'110.93.128.107
 	'DocSvrPort      = ""
 END IF

%>
