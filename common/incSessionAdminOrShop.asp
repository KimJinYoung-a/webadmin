<!-- #include virtual="/admin/incTenRedisSession.asp"-->
<%
'###########################################################
' Description : 본사 & 매장 공용 incsessionadmin
' History : 2009.04.07 서동석 생성
'			2011.02.17 한용민 수정
'###########################################################
Call fn_RDS_CHK_SSN_RESTORE()

dim C_IS_SSL_ENABLED : C_IS_SSL_ENABLED = (Request.ServerVariables("HTTPS") = "on")

dim C_ADMIN_AUTH
dim C_OFF_AUTH, C_OFF_part
dim C_MngPowerUser          ''경영지원팀 주요 인물
dim C_MngPart               '' 경영지원팀 인지.
dim C_ManagerUpJob          '' 점장 //현재 레벨 기준 점장이 계약직 보다 낮음.. ==> 정리 필요;;
dim C_PSMngPart,C_PSMngPartPower							''인사파트
dim C_HideCouponInfo		''not used
dim C_InspectorUser			''감사
dim C_OP, C_OP_AUTH               ''운영개발팀
dim C_ManagerPartTimeMember	'' 시급계약직 관리자
dim C_ERP_VERSION			'' ERP 버전
dim C_SYSTEM_Part		' 개발팀
dim C_logics_Part,C_logicsPowerUser		' 물류팀
dim C_CSUser, C_CSPowerUser, C_CSOutsourcingUser, C_CSpermanentUser		'cs팀
dim C_AUTH, C_Relationship_Part
dim C_MD_AUTH, C_MD, C_CriticInfoUserLV1, C_CriticInfoUserLV2, C_CriticInfoUserLV3, C_privacyadminuser
	C_ADMIN_AUTH = (session("ssBctId") = "icommang") or (session("ssBctId") = "coolhas") or (session("ssBctId") = "kobula") or (session("ssBctId") = "tozzinet") or (session("ssBctId") = "leesjun25") or (session("ssBctId") = "kjy8517") or (session("ssBctId") = "hrkang97") or (session("ssBctId") = "motions") or (session("ssBctId") = "thensi7")
	C_HideCouponInfo = (session("ssBctId") = "tozzinet1")
    C_InspectorUser = session("ssBctId") = "aimcta1"

	'C_CriticInfoUserLV1 = (session("ssAdminCLsn")=500)		' 개인정보취급권한(개인정보)		' 2019.09.06 한용민
	'C_CriticInfoUserLV2 = (session("ssAdminCLsn")=100)		' 개인정보취급권한(배송정보)		' 2019.09.06 한용민
	'C_CriticInfoUserLV3 = (session("ssAdminCLsn")=1)		' 개인정보취급권한(주문정보)		' 2019.09.06 한용민
	'C_CriticInfoUserLV4 = (session("ssAdminCLsn")=200)		' 개인정보취급권한(인사정보)		' 2019.09.06 한용민
	C_CriticInfoUserLV1 = (session("ssAdminlv1customerYN")="Y")		' 개인정보취급권한(고객정보)		' 2022.05.11 한용민 생성
	C_CriticInfoUserLV2 = (session("ssAdminlv2partnerYN")="Y")		' 개인정보취급권한(파트너정보)		' 2022.05.11 한용민 생성
	C_CriticInfoUserLV3 = (session("ssAdminlv3InternalYN")="Y")		' 개인정보취급권한(내부정보)		' 2022.05.11 한용민 생성
	C_privacyadminuser = (session("ssBctId") = "kobula") or (session("ssBctId") = "tozzinet") or (session("ssBctId") = "coolhas")		' 2019.09.06 한용민

	'오프라인본사이거나 직영점이고 파트선임 이상
	C_OFF_AUTH = (session("ssAdminPsn") = "13" or session("ssAdminPsn") = "24") and session("ssAdminLsn")<="3"
	C_OFF_AUTH = C_OFF_AUTH or C_MD_AUTH
	C_OFF_part = (session("ssAdminPsn") = "13" or session("ssAdminPsn") = "24")

	' 온라인MD운영(11), 온라인MD수입(21) 파트선임 이상		' 2019.07.09 한용민
	C_MD_AUTH = (session("ssAdminPsn") = "11" or session("ssAdminPsn") = "21") and session("ssAdminLsn")<="3"

	' 온라인MD운영(11), 온라인MD수입(21)		' 2019.07.09 한용민
	C_MD = (session("ssAdminPsn") = "11" or session("ssAdminPsn") = "21")

    C_SYSTEM_Part = session("ssAdminPsn")="7" or session("ssAdminPsn")="30" or session("ssAdminPsn")="31"

	' 물류팀  파트선임 이상	' 2020.02.12 한용민
	C_logicsPowerUser = session("ssAdminPsn")="9" and session("ssAdminLsn")<="3"
	' 물류팀	' 2020.02.12 한용민
	C_logics_Part = (session("ssAdminPsn")="9")

	'cs 파트선임권한		' 2019.05.28 한용민
	C_CSPowerUser   = (session("ssAdminPsn")="10" and session("ssAdminLsn")<="3") or session("ssBctId") = "boom15" or session("ssBctId") = "yj22354"
    C_CSPowerUser	= C_CSPowerUser or session("ssBctId") = "rabbit1693" or session("ssBctId") = "heendoongi" or session("ssBctId") = "seokmi1221"
	' cs팀	' 2019.11.20 한용민
	C_CSUser   = session("ssAdminPsn")="10"
	' cs팀 위탁업체		' 2021.04.09 한용민
	C_CSOutsourcingUser = session("ssAdminLsn")="5"
	' cs팀 정규직
	if session("ssAdminPOSITsn")<>"" and not(isnull(session("ssAdminPOSITsn"))) then
		C_CSpermanentUser = session("ssAdminPsn")="10" and session("ssAdminPOSITsn")<=11
	end if

	'경영지원 파트선임권한		' 2019.07.09 한용민
	C_MngPowerUser = session("ssAdminPsn")="8" and session("ssAdminLsn")<="3"
    C_MngPart = (session("ssAdminPsn")="8")
    ' 인사총무(20)
	C_PSMngPartPower = (session("ssAdminPsn")="20" and session("ssAdminLsn")<="3") or session("ssBctId") = "aimcta"
	C_PSMngPart= session("ssAdminPsn")="20" or session("ssBctId") = "aimcta"

    C_ManagerUpJob = (Not IsNULL(session("ssAdminPOsn")) and (session("ssAdminPOsn")>"0") and (session("ssAdminPOsn")<"7") )
    C_ManagerUpJob = ((C_ManagerUpJob) or (session("ssAdminLsn")<="3") or (session("ssAdminLsn")="6"))
    C_ManagerUpJob = ((C_ManagerUpJob) or (session("ssBctId") = "kei0329"))
	C_ManagerUpJob = ((C_ManagerUpJob) or (session("ssBctId") = "azure0502"))

	C_OP_AUTH   = session("ssAdminPsn")="30" and session("ssAdminLsn")<="3"  '운영기획 파트선임권한 추가(2015.05.04 정윤정)
	' 운영기획팀 권한	' 2021.01.12 한용민
	C_OP   = session("ssAdminPsn")="30"
	C_AUTH = session("ssAdminLsn")<="3" '일반 파트 선임 이상

    C_ManagerPartTimeMember = C_MngPart or (session("ssBctId") = "josin222") or (session("ssBctId") = "john6136") or (session("ssBctId") = "nownhere21")
	C_ManagerPartTimeMember = C_ManagerPartTimeMember or (session("ssBctId") = "bleuciel05") or (session("ssBctId") = "mussin83") or (session("ssBctId") = "bseo")
	C_ManagerPartTimeMember = C_ManagerPartTimeMember or (session("ssBctId") = "boyishP")

	' 관계사 팀		' 2020.06.02 한용민
	C_Relationship_Part = session("ssAdminPsn")="17" 

	'// 프로시저명에 사용가능한 변수이어야 한다.(예, [db_SCM_LINK].[dbo].[sp_BA_CUST_ContsInsert_2015])
	C_ERP_VERSION = ""

'/오프라인 매장과 업체의 공통 부분으로 사용 하는 부분에서 session("ssAdminPsn") null 들어 갈경우 권한이 틀려짐..
if isnull(session("ssAdminPsn")) then session("ssAdminPsn") = 0

'' 공급업체
dim C_IS_Maker_Upche

'' 직영점
dim C_IS_OWN_SHOP

'' 가맹점
dim C_IS_FRN_SHOP

'' 직영 또는 가맹점
dim C_IS_SHOP

'' 매장 아이디
dim C_STREETSHOPID

''직원
dim C_ADMIN_USER

C_IS_Maker_Upche = (session("ssBctDiv") = "9999")

''session("ssAdminLsn")=10 , 11 매장직원권한 2011-01-14 eastone추가
'C_IS_OWN_SHOP = C_IS_OWN_SHOP or ( ((session("ssAdminLsn")="10") or (session("ssAdminLsn")="11")) and (session("ssBctDiv") <> "503"))
'핑거스
'C_IS_OWN_SHOP = C_IS_OWN_SHOP or (session("ssBctDiv")="301" or session("ssAdminPsn")="16")

' 매장 자체 아이디의 경우 권한 설정을 따로 안함. userdiv 체크해야 해서 추가		' 2019.05.14 한용민
C_IS_OWN_SHOP = (session("ssBctDiv") = "501") or (session("ssBctDiv") = "101") or (session("ssBctDiv") = "111") or (session("ssBctDiv") = "112")
C_IS_OWN_SHOP = C_IS_OWN_SHOP or NOT(isNULL(session("ssBctBigo")) or session("ssBctBigo")="")

C_IS_FRN_SHOP = (session("ssBctDiv") = "502") or (session("ssBctDiv") = "503") or (session("ssBctId")="streetshop102")
C_IS_SHOP = (C_IS_OWN_SHOP or C_IS_FRN_SHOP)
C_ADMIN_USER     = (session("ssBctDiv") < 10)

if C_IS_FRN_SHOP then
	C_STREETSHOPID = session("ssBctId")
elseif C_IS_OWN_SHOP then
	if (session("ssBctDiv") = "501") or (session("ssBctDiv") = "502") then
		C_STREETSHOPID = session("ssBctid")
	'elseif (session("ssBctDiv")="301" or session("ssAdminPsn")="16") then       ''아카데미
	'    C_STREETSHOPID = "cafe003"
	else
		C_STREETSHOPID = session("ssBctBigo")
	end if
else
    C_STREETSHOPID = session("ssBctBigo")
end if

If (session("ssBctId") = "") then
    %><html>
    <script>
    alert("세션이 종료되었습니다. \n재로그인후 사용하실수 있습니다.");
    top.location = "/index.asp";
    </script>
    </html><%
    response.End
End if

'-----------------------------------------------------------------------
' 이벤트 전역변수 선언 (2007.02.07; 정윤정)
'-----------------------------------------------------------------------
Dim staticImgUrl,uploadUrl,manageUrl,wwwUrl, uploadImgUrl,othermall,mailzine,www2009url, ItemUploadUrl, staticUploadUrl, webImgUrl, mobileUrl, fixImgUrl
Dim wwwFingers, imgFingers
dim webImgSSLUrl, stsAdmURL
''검색엔진 관련
Dim DocSvrAddr, DocSvrPort, DocAuthCode

IF application("Svr_Info")="Dev" THEN
 	staticImgUrl = "http://testimgstatic.10x10.co.kr"	'테스트
 	webImgUrl		= "http://testwebimage.10x10.co.kr"				'웹이미지
 	webImgSSLUrl	= "http://testwebimage.10x10.co.kr"
 	manageUrl 	    = "http://testwebadmin.10x10.co.kr"
 	wwwUrl		    = "http://2010www.10x10.co.kr"            ''차후 정리요망
 	othermall       = "http://othermall.10x10.co.kr"
	mailzine        = "http://testmailzine.10x10.co.kr"
	www2009url      = "http://2009www.10x10.co.kr"
	mobileUrl	    = "http://61.252.133.2"

	wwwFingers		= "http://test.thefingers.co.kr"
	imgFingers		= "http://testimage.thefingers.co.kr"

	''** Upload 구분.;;
	uploadUrl	    = "http://testimgstatic.10x10.co.kr"   ''차후 정리요망
	uploadImgUrl    = "http://testupload.10x10.co.kr"
	ItemUploadUrl	= "http://testupload.10x10.co.kr"
	stsAdmURL		= "http://testwebadmin.10x10.co.kr"
	if (request.ServerVariables("LOCAL_ADDR")="::1" or request.ServerVariables("LOCAL_ADDR")="127.0.0.1") then
		manageUrl = ""
		stsAdmURL = ""
	end if
ELSE
	if (C_IS_SSL_ENABLED = True) then
 		staticImgUrl    = "/imgstatic"
 		webImgUrl		= "/webimage"							'웹이미지
		webImgSSLUrl	= "https://webimage.10x10.co.kr"
		fixImgUrl		= "/fiximage"

 		wwwUrl 		    = "http://www1.10x10.co.kr"
 		manageUrl 	    = "http://webadmin.10x10.co.kr"
 		othermall       = "http://gseshop.10x10.co.kr"
		mailzine        = "http://mailzine.10x10.co.kr"
		www2009url      = "http://www.10x10.co.kr"
		mobileUrl	    = "http://m.10x10.co.kr"

		wwwFingers		= "http://www.thefingers.co.kr"
		imgFingers		= "http://image.thefingers.co.kr"

		''** Upload 구분.;;
		uploadUrl	    = "https://oimgstatic.10x10.co.kr"
		uploadImgUrl    = "https://upload.10x10.co.kr"          '' upload.10x10.co.kr 통해서 Nas Server로 업로드
 		ItemUploadUrl	= "https://upload.10x10.co.kr"

		staticUploadUrl = "http://oimgstatic.10x10.co.kr"
		stsAdmURL		= "https://stscm.10x10.co.kr"
	else
 		staticImgUrl    = "http://imgstatic.10x10.co.kr"
 		webImgUrl		= "http://webimage.10x10.co.kr"				'웹이미지
 		webImgSSLUrl	= "http://webimage.10x10.co.kr"
		fixImgUrl		= "http://fiximage.10x10.co.kr"

 		wwwUrl 		    = "http://www1.10x10.co.kr"
 		manageUrl 	    = "http://webadmin.10x10.co.kr"
 		othermall       = "http://gseshop.10x10.co.kr"
		mailzine        = "http://mailzine.10x10.co.kr"
		www2009url      = "http://www.10x10.co.kr"
		mobileUrl	    = "http://m.10x10.co.kr"

		wwwFingers		= "http://www.thefingers.co.kr"
		imgFingers		= "http://image.thefingers.co.kr"

		''** Upload 구분.;;
		uploadUrl	    = "http://oimgstatic.10x10.co.kr"
		uploadImgUrl    = "http://upload.10x10.co.kr"          '' upload.10x10.co.kr 통해서 Nas Server로 업로드
 		ItemUploadUrl	= "http://upload.10x10.co.kr"

		staticUploadUrl = "http://oimgstatic.10x10.co.kr"
		stsAdmURL		= "http://stscm.10x10.co.kr"
	end if
END IF

%>
