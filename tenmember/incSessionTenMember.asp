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
    alert("������ ����Ǿ����ϴ�. \n��α����� ����ϽǼ� �ֽ��ϴ�.<%= session("ssBctSn") %>");
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
dim C_OFF_AUTH              ''���� �ֿ��ι�
dim C_MngPowerUser          ''�濵������ �ֿ� �ι�
dim C_MngPart               '' �濵������ ����.
dim C_ManagerUpJob          '' ���� //���� ���� ���� ������ ����� ���� ����.. ==> ���� �ʿ�;;
dim C_PSMngPart,C_PSMngPartPower							''�λ���Ʈ
dim C_HideCouponInfo		''not used
dim C_InspectorUser			''����
dim C_OP, C_OP_AUTH               ''�������
dim C_DataSales_AUTH		''���� �����ͼ����� �ֿ� �ι�
dim C_ManagerPartTimeMember	'' �ñް���� ������
dim C_ERP_VERSION			'' ERP ����
dim C_SYSTEM_Part		' ������
dim C_logics_Part,C_logicsPowerUser		' ������
dim C_CSUser, C_CSPowerUser, C_CSOutsourcingUser, C_CSOutsourcingPowerUser, C_CSpermanentUser		'cs��
dim C_AUTH, C_Relationship_Part, C_MKT_Part
dim C_MD_AUTH, C_MD, C_CriticInfoUserLV1, C_CriticInfoUserLV2, C_CriticInfoUserLV3, C_privacyadminuser
	C_ADMIN_AUTH = (session("ssBctId") = "coolhas") or (session("ssBctId") = "kobula") or (session("ssBctId") = "tozzinet") or (session("ssBctId") = "corpse2") or (session("ssBctId") = "kjy8517") or (session("ssBctId") = "thensi7") or (session("ssBctId") = "qpark99")
	C_HideCouponInfo = (session("ssBctId") = "tozzinet")
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
	C_OFF_AUTH = C_OFF_AUTH or session("ssBctId") = "rabbit1693"

	' �¶���MD�(11), �¶���MD����(21) ��Ʈ���� �̻�		' 2019.07.09 �ѿ��
	C_MD_AUTH = (session("ssAdminPsn") = "11" or session("ssAdminPsn") = "21") and session("ssAdminLsn")<="3"
	C_MD_AUTH = C_MD_AUTH or session("ssBctId")="as2304"

	' �¶���MD�(11), �¶���MD����(21)		' 2019.07.09 �ѿ��
	C_MD = (session("ssAdminPsn") = "11" or session("ssAdminPsn") = "21")

    C_SYSTEM_Part = session("ssAdminPsn")="7" or session("ssAdminPsn")="30" or session("ssAdminPsn")="31"

	C_MKT_Part = session("ssAdminPsn")="14"

	C_DataSales_AUTH = C_ADMIN_AUTH or session("ssBctId")="mydrizzle" or session("ssBctId")="fpahsskfk200"	'�̽���, �̹���

	' ������  ��Ʈ���� �̻�	' 2020.02.12 �ѿ��
	C_logicsPowerUser = session("ssAdminPsn")="9" and session("ssAdminLsn")<="3"
	' ������	' 2020.02.12 �ѿ��
	C_logics_Part = (session("ssAdminPsn")="9")

	'cs ��Ʈ���ӱ���		' 2019.05.28 �ѿ��
	C_CSPowerUser   = (session("ssAdminPsn")="10" and session("ssAdminLsn")<="3") or session("ssBctId") = "boom15"
    C_CSPowerUser	= C_CSPowerUser or session("ssBctId") = "rabbit1693" or session("ssBctId") = "heendoongi" or session("ssBctId") = "seokmi1221"
	' cs��	' 2019.11.20 �ѿ��
	C_CSUser   = session("ssAdminPsn")="10"
	' cs�� ��Ź��ü		' 2021.04.09 �ѿ��
	C_CSOutsourcingUser = C_CSUser and session("ssAdminLsn")="5"
	' cs�� ��Ź��ü ��Ʈ���� �̻�		' 2021.09.13 �ѿ��
	C_CSOutsourcingPowerUser = C_CSOutsourcingUser and session("ssBctId") = "kangkong83"
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

	C_OP_AUTH   = session("ssAdminPsn")="30" and session("ssAdminLsn")<="3"  '������� ��Ʈ���ӱ��� �߰�(2015.05.04 ������)
	' ������� ����	' 2021.01.12 �ѿ��
	C_OP   = session("ssAdminPsn")="30"
	C_AUTH = session("ssAdminLsn")<="3" '�Ϲ� ��Ʈ ���� �̻�

    C_ManagerPartTimeMember = C_MngPart or (session("ssBctId") = "josin222")
	C_ManagerPartTimeMember = C_ManagerPartTimeMember or (session("ssBctId") = "boyishP")

	' ����� ��		' 2020.06.02 �ѿ��
	C_Relationship_Part = session("ssAdminPsn")="17"

	'// ���ν����� ��밡���� �����̾�� �Ѵ�.(��, [db_SCM_LINK].[dbo].[sp_BA_CUST_ContsInsert_2015])
	C_ERP_VERSION = ""

 Dim staticImgUrl,uploadUrl,manageUrl,wwwUrl, uploadImgUrl,othermall,mailzine,www2009url, ItemUploadUrl, staticUploadUrl, webImgUrl, mobileUrl, partnerUrl
 Dim wwwFingers, imgFingers
  ''�˻����� ����
 Dim DocSvrAddr, DocSvrPort, DocAuthCode

 IF application("Svr_Info")="Dev" THEN
 	staticImgUrl = "http://testimgstatic.10x10.co.kr"	'�׽�Ʈ
 	webImgUrl		= "http://testwebimage.10x10.co.kr"				'���̹���

 	manageUrl 	    = "http://testwebadmin.10x10.co.kr"
 	wwwUrl		    = "http://2012www.10x10.co.kr"            ''���� �������
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
	partnerUrl		= "http://testwebimage.10x10.co.kr/partner"		'�ӽû�ǰ�̹���(��Ʈ��)
 ELSE
 	staticImgUrl    = "https://imgstatic.10x10.co.kr"
 	webImgUrl		= "https://webimage.10x10.co.kr"				'���̹���

 	wwwUrl 		    = "https://www1.10x10.co.kr"
 	manageUrl 	    = "https://webadmin.10x10.co.kr"
 	othermall       = "https://gseshop.10x10.co.kr"
	mailzine        = "http://mailzine.10x10.co.kr"
	www2009url      = "https://www.10x10.co.kr"
	mobileUrl	    = "https://m.10x10.co.kr"

	wwwFingers		= "http://www.thefingers.co.kr"
	imgFingers		= "http://image.thefingers.co.kr"

	''** Upload ����.;;
	uploadUrl	    = "https://oimgstatic.10x10.co.kr"
	uploadImgUrl    = "https://upload.10x10.co.kr"          '' upload.10x10.co.kr ���ؼ� Nas Server�� ���ε�
 	ItemUploadUrl	= "https://upload.10x10.co.kr"
 	partnerUrl		= "http://partner.10x10.co.kr"				'�ӽû�ǰ�̹���(��Ʈ��)

	staticUploadUrl = "https://oimgstatic.10x10.co.kr"
 END IF

%>