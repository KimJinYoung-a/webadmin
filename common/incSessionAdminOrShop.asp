<!-- #include virtual="/admin/incTenRedisSession.asp"-->
<%
'###########################################################
' Description : ���� & ���� ���� incsessionadmin
' History : 2009.04.07 ������ ����
'			2011.02.17 �ѿ�� ����
'###########################################################
Call fn_RDS_CHK_SSN_RESTORE()

dim C_IS_SSL_ENABLED : C_IS_SSL_ENABLED = (Request.ServerVariables("HTTPS") = "on")

dim C_ADMIN_AUTH
dim C_OFF_AUTH, C_OFF_part
dim C_MngPowerUser          ''�濵������ �ֿ� �ι�
dim C_MngPart               '' �濵������ ����.
dim C_ManagerUpJob          '' ���� //���� ���� ���� ������ ����� ���� ����.. ==> ���� �ʿ�;;
dim C_PSMngPart,C_PSMngPartPower							''�λ���Ʈ
dim C_HideCouponInfo		''not used
dim C_InspectorUser			''����
dim C_OP, C_OP_AUTH               ''�������
dim C_ManagerPartTimeMember	'' �ñް���� ������
dim C_ERP_VERSION			'' ERP ����
dim C_SYSTEM_Part		' ������
dim C_logics_Part,C_logicsPowerUser		' ������
dim C_CSUser, C_CSPowerUser, C_CSOutsourcingUser, C_CSpermanentUser		'cs��
dim C_AUTH, C_Relationship_Part
dim C_MD_AUTH, C_MD, C_CriticInfoUserLV1, C_CriticInfoUserLV2, C_CriticInfoUserLV3, C_privacyadminuser
	C_ADMIN_AUTH = (session("ssBctId") = "icommang") or (session("ssBctId") = "coolhas") or (session("ssBctId") = "kobula") or (session("ssBctId") = "tozzinet") or (session("ssBctId") = "leesjun25") or (session("ssBctId") = "kjy8517") or (session("ssBctId") = "hrkang97") or (session("ssBctId") = "motions") or (session("ssBctId") = "thensi7")
	C_HideCouponInfo = (session("ssBctId") = "tozzinet1")
    C_InspectorUser = session("ssBctId") = "aimcta1"

	'C_CriticInfoUserLV1 = (session("ssAdminCLsn")=500)		' ����������ޱ���(��������)		' 2019.09.06 �ѿ��
	'C_CriticInfoUserLV2 = (session("ssAdminCLsn")=100)		' ����������ޱ���(�������)		' 2019.09.06 �ѿ��
	'C_CriticInfoUserLV3 = (session("ssAdminCLsn")=1)		' ����������ޱ���(�ֹ�����)		' 2019.09.06 �ѿ��
	'C_CriticInfoUserLV4 = (session("ssAdminCLsn")=200)		' ����������ޱ���(�λ�����)		' 2019.09.06 �ѿ��
	C_CriticInfoUserLV1 = (session("ssAdminlv1customerYN")="Y")		' ����������ޱ���(������)		' 2022.05.11 �ѿ�� ����
	C_CriticInfoUserLV2 = (session("ssAdminlv2partnerYN")="Y")		' ����������ޱ���(��Ʈ������)		' 2022.05.11 �ѿ�� ����
	C_CriticInfoUserLV3 = (session("ssAdminlv3InternalYN")="Y")		' ����������ޱ���(��������)		' 2022.05.11 �ѿ�� ����
	C_privacyadminuser = (session("ssBctId") = "kobula") or (session("ssBctId") = "tozzinet") or (session("ssBctId") = "coolhas")		' 2019.09.06 �ѿ��

	'�������κ����̰ų� �������̰� ��Ʈ���� �̻�
	C_OFF_AUTH = (session("ssAdminPsn") = "13" or session("ssAdminPsn") = "24") and session("ssAdminLsn")<="3"
	C_OFF_AUTH = C_OFF_AUTH or C_MD_AUTH
	C_OFF_part = (session("ssAdminPsn") = "13" or session("ssAdminPsn") = "24")

	' �¶���MD�(11), �¶���MD����(21) ��Ʈ���� �̻�		' 2019.07.09 �ѿ��
	C_MD_AUTH = (session("ssAdminPsn") = "11" or session("ssAdminPsn") = "21") and session("ssAdminLsn")<="3"

	' �¶���MD�(11), �¶���MD����(21)		' 2019.07.09 �ѿ��
	C_MD = (session("ssAdminPsn") = "11" or session("ssAdminPsn") = "21")

    C_SYSTEM_Part = session("ssAdminPsn")="7" or session("ssAdminPsn")="30" or session("ssAdminPsn")="31"

	' ������  ��Ʈ���� �̻�	' 2020.02.12 �ѿ��
	C_logicsPowerUser = session("ssAdminPsn")="9" and session("ssAdminLsn")<="3"
	' ������	' 2020.02.12 �ѿ��
	C_logics_Part = (session("ssAdminPsn")="9")

	'cs ��Ʈ���ӱ���		' 2019.05.28 �ѿ��
	C_CSPowerUser   = (session("ssAdminPsn")="10" and session("ssAdminLsn")<="3") or session("ssBctId") = "boom15" or session("ssBctId") = "yj22354"
    C_CSPowerUser	= C_CSPowerUser or session("ssBctId") = "rabbit1693" or session("ssBctId") = "heendoongi" or session("ssBctId") = "seokmi1221"
	' cs��	' 2019.11.20 �ѿ��
	C_CSUser   = session("ssAdminPsn")="10"
	' cs�� ��Ź��ü		' 2021.04.09 �ѿ��
	C_CSOutsourcingUser = session("ssAdminLsn")="5"
	' cs�� ������
	if session("ssAdminPOSITsn")<>"" and not(isnull(session("ssAdminPOSITsn"))) then
		C_CSpermanentUser = session("ssAdminPsn")="10" and session("ssAdminPOSITsn")<=11
	end if

	'�濵���� ��Ʈ���ӱ���		' 2019.07.09 �ѿ��
	C_MngPowerUser = session("ssAdminPsn")="8" and session("ssAdminLsn")<="3"
    C_MngPart = (session("ssAdminPsn")="8")
    ' �λ��ѹ�(20)
	C_PSMngPartPower = (session("ssAdminPsn")="20" and session("ssAdminLsn")<="3") or session("ssBctId") = "aimcta"
	C_PSMngPart= session("ssAdminPsn")="20" or session("ssBctId") = "aimcta"

    C_ManagerUpJob = (Not IsNULL(session("ssAdminPOsn")) and (session("ssAdminPOsn")>"0") and (session("ssAdminPOsn")<"7") )
    C_ManagerUpJob = ((C_ManagerUpJob) or (session("ssAdminLsn")<="3") or (session("ssAdminLsn")="6"))
    C_ManagerUpJob = ((C_ManagerUpJob) or (session("ssBctId") = "kei0329"))
	C_ManagerUpJob = ((C_ManagerUpJob) or (session("ssBctId") = "azure0502"))

	C_OP_AUTH   = session("ssAdminPsn")="30" and session("ssAdminLsn")<="3"  '���ȹ ��Ʈ���ӱ��� �߰�(2015.05.04 ������)
	' ���ȹ�� ����	' 2021.01.12 �ѿ��
	C_OP   = session("ssAdminPsn")="30"
	C_AUTH = session("ssAdminLsn")<="3" '�Ϲ� ��Ʈ ���� �̻�

    C_ManagerPartTimeMember = C_MngPart or (session("ssBctId") = "josin222") or (session("ssBctId") = "john6136") or (session("ssBctId") = "nownhere21")
	C_ManagerPartTimeMember = C_ManagerPartTimeMember or (session("ssBctId") = "bleuciel05") or (session("ssBctId") = "mussin83") or (session("ssBctId") = "bseo")
	C_ManagerPartTimeMember = C_ManagerPartTimeMember or (session("ssBctId") = "boyishP")

	' ����� ��		' 2020.06.02 �ѿ��
	C_Relationship_Part = session("ssAdminPsn")="17" 

	'// ���ν����� ��밡���� �����̾�� �Ѵ�.(��, [db_SCM_LINK].[dbo].[sp_BA_CUST_ContsInsert_2015])
	C_ERP_VERSION = ""

'/�������� ����� ��ü�� ���� �κ����� ��� �ϴ� �κп��� session("ssAdminPsn") null ��� ����� ������ Ʋ����..
if isnull(session("ssAdminPsn")) then session("ssAdminPsn") = 0

'' ���޾�ü
dim C_IS_Maker_Upche

'' ������
dim C_IS_OWN_SHOP

'' ������
dim C_IS_FRN_SHOP

'' ���� �Ǵ� ������
dim C_IS_SHOP

'' ���� ���̵�
dim C_STREETSHOPID

''����
dim C_ADMIN_USER

C_IS_Maker_Upche = (session("ssBctDiv") = "9999")

''session("ssAdminLsn")=10 , 11 ������������ 2011-01-14 eastone�߰�
'C_IS_OWN_SHOP = C_IS_OWN_SHOP or ( ((session("ssAdminLsn")="10") or (session("ssAdminLsn")="11")) and (session("ssBctDiv") <> "503"))
'�ΰŽ�
'C_IS_OWN_SHOP = C_IS_OWN_SHOP or (session("ssBctDiv")="301" or session("ssAdminPsn")="16")

' ���� ��ü ���̵��� ��� ���� ������ ���� ����. userdiv üũ�ؾ� �ؼ� �߰�		' 2019.05.14 �ѿ��
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
	'elseif (session("ssBctDiv")="301" or session("ssAdminPsn")="16") then       ''��ī����
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
    alert("������ ����Ǿ����ϴ�. \n��α����� ����ϽǼ� �ֽ��ϴ�.");
    top.location = "/index.asp";
    </script>
    </html><%
    response.End
End if

'-----------------------------------------------------------------------
' �̺�Ʈ �������� ���� (2007.02.07; ������)
'-----------------------------------------------------------------------
Dim staticImgUrl,uploadUrl,manageUrl,wwwUrl, uploadImgUrl,othermall,mailzine,www2009url, ItemUploadUrl, staticUploadUrl, webImgUrl, mobileUrl, fixImgUrl
Dim wwwFingers, imgFingers
dim webImgSSLUrl, stsAdmURL
''�˻����� ����
Dim DocSvrAddr, DocSvrPort, DocAuthCode

IF application("Svr_Info")="Dev" THEN
 	staticImgUrl = "http://testimgstatic.10x10.co.kr"	'�׽�Ʈ
 	webImgUrl		= "http://testwebimage.10x10.co.kr"				'���̹���
 	webImgSSLUrl	= "http://testwebimage.10x10.co.kr"
 	manageUrl 	    = "http://testwebadmin.10x10.co.kr"
 	wwwUrl		    = "http://2010www.10x10.co.kr"            ''���� �������
 	othermall       = "http://othermall.10x10.co.kr"
	mailzine        = "http://testmailzine.10x10.co.kr"
	www2009url      = "http://2009www.10x10.co.kr"
	mobileUrl	    = "http://61.252.133.2"

	wwwFingers		= "http://test.thefingers.co.kr"
	imgFingers		= "http://testimage.thefingers.co.kr"

	''** Upload ����.;;
	uploadUrl	    = "http://testimgstatic.10x10.co.kr"   ''���� �������
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
 		webImgUrl		= "/webimage"							'���̹���
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

		''** Upload ����.;;
		uploadUrl	    = "https://oimgstatic.10x10.co.kr"
		uploadImgUrl    = "https://upload.10x10.co.kr"          '' upload.10x10.co.kr ���ؼ� Nas Server�� ���ε�
 		ItemUploadUrl	= "https://upload.10x10.co.kr"

		staticUploadUrl = "http://oimgstatic.10x10.co.kr"
		stsAdmURL		= "https://stscm.10x10.co.kr"
	else
 		staticImgUrl    = "http://imgstatic.10x10.co.kr"
 		webImgUrl		= "http://webimage.10x10.co.kr"				'���̹���
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

		''** Upload ����.;;
		uploadUrl	    = "http://oimgstatic.10x10.co.kr"
		uploadImgUrl    = "http://upload.10x10.co.kr"          '' upload.10x10.co.kr ���ؼ� Nas Server�� ���ε�
 		ItemUploadUrl	= "http://upload.10x10.co.kr"

		staticUploadUrl = "http://oimgstatic.10x10.co.kr"
		stsAdmURL		= "http://stscm.10x10.co.kr"
	end if
END IF

%>
