<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%

'// ============================================================================
dim C_IS_LOGIN_OK
C_IS_LOGIN_OK = True

if (C_IS_LOGIN_OK = False) or IsNull(session("ssBctSn")) or IsNull(session("ssBctDiv"))  then
	C_IS_LOGIN_OK = False
end if

if (C_IS_LOGIN_OK = False) or (session("ssBctSn") = "") or (session("ssBctDiv") = "")  then
	C_IS_LOGIN_OK = False
end if

if (C_IS_LOGIN_OK = False) or ((session("ssBctDiv") > "9") and (session("ssBctDiv") <> "5000")) then
	C_IS_LOGIN_OK = False
end if


If (Not C_IS_LOGIN_OK) then
 %>
    <script>
    alert("세션이 종료되었습니다. \n재로그인후 사용하실수 있습니다.<%= session("ssBctSn") %>");
    top.location = "/";
    </script>
    <%
    response.End
End if


'// ============================================================================
dim C_IS_ADMIN_LOGIN
C_IS_ADMIN_LOGIN = False


'// ============================================================================
dim C_IS_SCM_LOGIN
C_IS_SCM_LOGIN = (session("ssAdminLsn") < 9)


'// ============================================================================
 
DIM CAddDetailSpliter : CAddDetailSpliter= CHR(3)&CHR(4)
dim C_ADMIN_AUTH
dim C_OFF_AUTH              ''오프 주요인물
dim C_MngPowerUser          ''경영지원팀 주요 인물
dim C_MngPart               '' 경영지원팀 인지.
dim C_ManagerUpJob          '' 점장 //현재 레벨 기준 점장이 계약직 보다 낮음.. ==> 정리 필요;;
dim C_PSMngPart,C_PSMngPartPower							''인사파트
dim C_HideCouponInfo		''not used
dim C_InspectorUser			''감사
dim C_OP, C_OP_AUTH               ''운영개발팀
dim C_DataSales_AUTH		''영업 데이터세일즈 주요 인물
dim C_ManagerPartTimeMember	'' 시급계약직 관리자
dim C_ERP_VERSION			'' ERP 버전
dim C_SYSTEM_Part		' 개발팀
dim C_logics_Part,C_logicsPowerUser		' 물류팀
dim C_CSUser, C_CSPowerUser, C_CSOutsourcingUser, C_CSOutsourcingPowerUser, C_CSpermanentUser		'cs팀
dim C_AUTH, C_Relationship_Part, C_MKT_Part
dim C_MD_AUTH, C_MD, C_CriticInfoUserLV1, C_CriticInfoUserLV2, C_CriticInfoUserLV3, C_privacyadminuser
	C_ADMIN_AUTH = (session("ssBctId") = "coolhas") or (session("ssBctId") = "kobula") or (session("ssBctId") = "tozzinet") or (session("ssBctId") = "corpse2") or (session("ssBctId") = "kjy8517") or (session("ssBctId") = "thensi7") or (session("ssBctId") = "qpark99")
	C_HideCouponInfo = (session("ssBctId") = "tozzinet")
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
	C_OFF_AUTH = C_OFF_AUTH or session("ssBctId") = "rabbit1693"

	' 온라인MD운영(11), 온라인MD수입(21) 파트선임 이상		' 2019.07.09 한용민
	C_MD_AUTH = (session("ssAdminPsn") = "11" or session("ssAdminPsn") = "21") and session("ssAdminLsn")<="3"
	C_MD_AUTH = C_MD_AUTH or session("ssBctId")="as2304"

	' 온라인MD운영(11), 온라인MD수입(21)		' 2019.07.09 한용민
	C_MD = (session("ssAdminPsn") = "11" or session("ssAdminPsn") = "21")

    C_SYSTEM_Part = session("ssAdminPsn")="7" or session("ssAdminPsn")="30" or session("ssAdminPsn")="31"

	C_MKT_Part = session("ssAdminPsn")="14"

	C_DataSales_AUTH = C_ADMIN_AUTH or session("ssBctId")="mydrizzle" or session("ssBctId")="fpahsskfk200"	'이슬비, 이민지

	' 물류팀  파트선임 이상	' 2020.02.12 한용민
	C_logicsPowerUser = session("ssAdminPsn")="9" and session("ssAdminLsn")<="3"
	' 물류팀	' 2020.02.12 한용민
	C_logics_Part = (session("ssAdminPsn")="9")

	'cs 파트선임권한		' 2019.05.28 한용민
	C_CSPowerUser   = (session("ssAdminPsn")="10" and session("ssAdminLsn")<="3") or session("ssBctId") = "boom15"
    C_CSPowerUser	= C_CSPowerUser or session("ssBctId") = "rabbit1693" or session("ssBctId") = "heendoongi" or session("ssBctId") = "seokmi1221"
	' cs팀	' 2019.11.20 한용민
	C_CSUser   = session("ssAdminPsn")="10"
	' cs팀 위탁업체		' 2021.04.09 한용민
	C_CSOutsourcingUser = C_CSUser and session("ssAdminLsn")="5"
	' cs팀 위탁업체 파트선임 이상		' 2021.09.13 한용민
	C_CSOutsourcingPowerUser = C_CSOutsourcingUser and session("ssBctId") = "kangkong83"
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

	C_OP_AUTH   = session("ssAdminPsn")="30" and session("ssAdminLsn")<="3"  '운영개발팀 파트선임권한 추가(2015.05.04 정윤정)
	' 운영개발팀 권한	' 2021.01.12 한용민
	C_OP   = session("ssAdminPsn")="30"
	C_AUTH = session("ssAdminLsn")<="3" '일반 파트 선임 이상

    C_ManagerPartTimeMember = C_MngPart or (session("ssBctId") = "josin222")
	C_ManagerPartTimeMember = C_ManagerPartTimeMember or (session("ssBctId") = "boyishP")

	' 관계사 팀		' 2020.06.02 한용민
	C_Relationship_Part = session("ssAdminPsn")="17"

	'// 프로시저명에 사용가능한 변수이어야 한다.(예, [db_SCM_LINK].[dbo].[sp_BA_CUST_ContsInsert_2015])
	C_ERP_VERSION = ""

 Dim staticImgUrl,uploadUrl,manageUrl,wwwUrl, uploadImgUrl,othermall,mailzine,www2009url, ItemUploadUrl, staticUploadUrl, webImgUrl, mobileUrl, partnerUrl
 Dim wwwFingers, imgFingers
  ''검색엔진 관련
 Dim DocSvrAddr, DocSvrPort, DocAuthCode

 IF application("Svr_Info")="Dev" THEN
 	staticImgUrl = "http://testimgstatic.10x10.co.kr"	'테스트
 	webImgUrl		= "http://testwebimage.10x10.co.kr"				'웹이미지

 	manageUrl 	    = "http://testwebadmin.10x10.co.kr"
 	wwwUrl		    = "http://2012www.10x10.co.kr"            ''차후 정리요망
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
	partnerUrl		= "http://testwebimage.10x10.co.kr/partner"		'임시상품이미지(파트너)
 ELSE
 	staticImgUrl    = "https://imgstatic.10x10.co.kr"
 	webImgUrl		= "https://webimage.10x10.co.kr"				'웹이미지

 	wwwUrl 		    = "https://www1.10x10.co.kr"
 	manageUrl 	    = "https://webadmin.10x10.co.kr"
 	othermall       = "https://gseshop.10x10.co.kr"
	mailzine        = "http://mailzine.10x10.co.kr"
	www2009url      = "https://www.10x10.co.kr"
	mobileUrl	    = "https://m.10x10.co.kr"

	wwwFingers		= "http://www.thefingers.co.kr"
	imgFingers		= "http://image.thefingers.co.kr"

	''** Upload 구분.;;
	uploadUrl	    = "https://oimgstatic.10x10.co.kr"
	uploadImgUrl    = "https://upload.10x10.co.kr"          '' upload.10x10.co.kr 통해서 Nas Server로 업로드
 	ItemUploadUrl	= "https://upload.10x10.co.kr"
 	partnerUrl		= "http://partner.10x10.co.kr"				'임시상품이미지(파트너)

	staticUploadUrl = "https://oimgstatic.10x10.co.kr"
 END IF

%>