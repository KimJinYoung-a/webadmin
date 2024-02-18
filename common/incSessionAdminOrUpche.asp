<%

dim C_IS_SSL_ENABLED : C_IS_SSL_ENABLED = (Request.ServerVariables("HTTPS") = "on")

dim C_ADMIN_AUTH
dim C_OFF_AUTH, C_OFF_part
dim C_MngPart               '' 경영지원팀 인지.
dim C_InspectorUser			''감사
dim C_CSUser, C_CSPowerUser, C_CSOutsourcingUser, C_CSOutsourcingPowerUser, C_CSpermanentUser		'cs팀

C_ADMIN_AUTH = (session("ssBctId") = "icommang") or (session("ssBctId") = "coolhas") or (session("ssBctId") = "kobula") or (session("ssBctId") = "tozzinet") or (session("ssBctId") = "leesjun25") or (session("ssBctId") = "kjy8517") or (session("ssBctId") = "hrkang97") or (session("ssBctId") = "motions") or (session("ssBctId") = "thensi7")

'오프라인본사이거나 직영점이고 파트선임 이상. 온라인MD(11)도 열어줌
C_OFF_AUTH = (session("ssAdminPsn") = "13" or session("ssAdminPsn") = "24" or session("ssAdminPsn") = "11") and session("ssAdminLsn")<="3"
C_OFF_AUTH = C_OFF_AUTH or session("ssBctId") = "hrkang97" or session("ssBctId") = "tozzinet"
'C_OFF_AUTH = (session("ssBctId") = "hrkang97") or (session("ssBctId") = "nownhere21") or (session("ssBctId") = "bleuciel05") or (session("ssBctId") = "hsc85") or (session("ssBctId") = "john6136") or (session("ssBctId") = "kjyfer") or (session("ssBctId") = "eccrose") or (session("ssBctId") = "endandand") or (session("ssBctId") = "chhs2000")
	C_OFF_part = (session("ssAdminPsn") = "13" or session("ssAdminPsn") = "24")

	'cs 파트선임권한		' 2019.05.28 한용민
	C_CSPowerUser   = (session("ssAdminPsn")="10" and session("ssAdminLsn")<="3") or session("ssBctId") = "boom15"
    C_CSPowerUser	= C_CSPowerUser or session("ssBctId") = "rabbit1693" or session("ssBctId") = "heendoongi" or session("ssBctId") = "seokmi1221"
	' cs팀	' 2019.11.20 한용민
	C_CSUser   = session("ssAdminPsn")="10"
	' cs팀 위탁업체		' 2021.04.09 한용민
	C_CSOutsourcingUser = C_CSUser and session("ssAdminLsn")="5"
	' cs팀 위탁업체 파트선임 이상		' 2021.09.13 한용민
	C_CSOutsourcingPowerUser = C_CSOutsourcingUser and session("ssBctId") = "kangkong83"
	if session("ssAdminPOSITsn")<>"" and not(isnull(session("ssAdminPOSITsn"))) then
		C_CSpermanentUser = session("ssAdminPsn")="10" and session("ssAdminPOSITsn")<=11
	end if
C_MngPart = (session("ssAdminPsn")="8")
C_InspectorUser = session("ssBctId") = "aimcta1"

'' 공급업체
dim C_IS_Maker_Upche, C_ADMIN_USER

C_IS_Maker_Upche = (session("ssBctDiv") = "9999")
C_ADMIN_USER     = (session("ssBctDiv") < 10) or (session("ssBctDiv")="301")


If (session("ssBctId") = "") or ((Not C_IS_Maker_Upche) and (Not C_ADMIN_USER)) then
    %><html>
    <script>
    alert("세션이 종료되었습니다. \n재로그인후 사용하실수 있습니다.<%= session("ssBctDiv") %>");
    top.location = "/index.asp";
    </script>
    </html><%
    response.End
End if

'-----------------------------------------------------------------------
' 이벤트 전역변수 선언 (2007.02.07; 정윤정)
'-----------------------------------------------------------------------
Dim staticImgUrl,uploadUrl,manageUrl,wwwUrl,uploadImgUrl, ItemUploadUrl, staticUploadUrl, imgFingers ,wwwFingers, partnerUrl, fixImgUrl
dim webImgUrl, fingersImgUrl, UploadImgFingers
dim staticImgUrlForMAIL, webImgUrlForMAIL, fixImgUrlForMAIL		'// 메일 발송시 SSL 고려안함

IF application("Svr_Info")="Dev" THEN

	staticImgUrlForMAIL		= "http://testimgstatic.10x10.co.kr"		'// 메일 발송시 SSL 고려안함
	webImgUrlForMAIL		= "http://testwebimage.10x10.co.kr"
	fixImgUrlForMAIL		= "http://fiximage.10x10.co.kr"

 	staticImgUrl 	= "http://testimgstatic.10x10.co.kr"	'테스트
 	imgFingers 		= "http://testimage.thefingers.co.kr"
 	uploadUrl		= "http://testimgstatic.10x10.co.kr"
  	fingersImgUrl	= "http://testimage.thefingers.co.kr"
 	manageUrl 	 	= "http://testwebadmin.10x10.co.kr"
 	wwwUrl		 	= "http://test.10x10.co.kr"
 	webImgUrl		= "http://testwebimage.10x10.co.kr"				'웹이미지
 	uploadImgUrl 	= "http://testupload.10x10.co.kr"
 	ItemUploadUrl	= "http://testupload.10x10.co.kr"
 	wwwFingers		= "http://test.thefingers.co.kr"
    UploadImgFingers = "http://testimage.thefingers.co.kr"
    
 	partnerUrl		= "http://testwebimage.10x10.co.kr/partner"		'임시상품이미지(파트너)
ELSE
	if (C_IS_SSL_ENABLED = True) then
 		staticImgUrl    = "/imgstatic"
 		webImgUrl		= "/webimage"							'웹이미지
		fixImgUrl		= "/fiximage"

 		staticImgUrlForMAIL		= "http://imgstatic.10x10.co.kr"		'// 메일 발송시 SSL 고려안함
 		webImgUrlForMAIL		= "http://webimage.10x10.co.kr"
		fixImgUrlForMAIL		= "http://fiximage.10x10.co.kr"

		wwwFingers		= "http://www.thefingers.co.kr"
 		imgFingers 		= "http://image.thefingers.co.kr"
 		uploadUrl		= "http://oimgstatic.10x10.co.kr"
  		fingersImgUrl	= "http://image.thefingers.co.kr"
 		wwwUrl 		 	= "http://www1.10x10.co.kr"
 		manageUrl 	 	= "http://webadmin.10x10.co.kr"

 		'' 131 to Nas Svr
 		uploadImgUrl 	= "https://upload.10x10.co.kr"          '' upload.10x10.co.kr 통해서 Nas Server로 업로드
 		ItemUploadUrl	= "https://upload.10x10.co.kr"

 		partnerUrl		= "http://partner.10x10.co.kr"				'임시상품이미지(파트너)
 		UploadImgFingers = "http://oimage.thefingers.co.kr"
	else
 		staticImgUrl 	= "http://imgstatic.10x10.co.kr"
		webImgUrl		= "http://webimage.10x10.co.kr"				'웹이미지
		fixImgUrl		= "http://fiximage.10x10.co.kr"

 		staticImgUrlForMAIL		= "http://imgstatic.10x10.co.kr"		'// 메일 발송시 SSL 고려안함
 		webImgUrlForMAIL		= "http://webimage.10x10.co.kr"
		fixImgUrlForMAIL		= "http://fiximage.10x10.co.kr"

		wwwFingers		= "http://www.thefingers.co.kr"
 		imgFingers 		= "http://image.thefingers.co.kr"
 		uploadUrl		= "http://oimgstatic.10x10.co.kr"
  		fingersImgUrl	= "http://image.thefingers.co.kr"
 		wwwUrl 		 	= "http://www1.10x10.co.kr"
 		manageUrl 	 	= "http://webadmin.10x10.co.kr"

 		'' 131 to Nas Svr
 		uploadImgUrl 	= "http://upload.10x10.co.kr"          '' upload.10x10.co.kr 통해서 Nas Server로 업로드
 		ItemUploadUrl	= "http://upload.10x10.co.kr"

 		partnerUrl		= "http://partner.10x10.co.kr"				'임시상품이미지(파트너)
 		UploadImgFingers = "http://oimage.thefingers.co.kr"
	end if
END IF

%>
