<%

dim C_IS_SSL_ENABLED : C_IS_SSL_ENABLED = (Request.ServerVariables("HTTPS") = "on")

dim C_ADMIN_AUTH
dim C_OFF_AUTH, C_OFF_part
dim C_MngPart               '' �濵������ ����.
dim C_InspectorUser			''����
dim C_CSUser, C_CSPowerUser, C_CSOutsourcingUser, C_CSOutsourcingPowerUser, C_CSpermanentUser		'cs��

C_ADMIN_AUTH = (session("ssBctId") = "icommang") or (session("ssBctId") = "coolhas") or (session("ssBctId") = "kobula") or (session("ssBctId") = "tozzinet") or (session("ssBctId") = "leesjun25") or (session("ssBctId") = "kjy8517") or (session("ssBctId") = "hrkang97") or (session("ssBctId") = "motions") or (session("ssBctId") = "thensi7")

'�������κ����̰ų� �������̰� ��Ʈ���� �̻�. �¶���MD(11)�� ������
C_OFF_AUTH = (session("ssAdminPsn") = "13" or session("ssAdminPsn") = "24" or session("ssAdminPsn") = "11") and session("ssAdminLsn")<="3"
C_OFF_AUTH = C_OFF_AUTH or session("ssBctId") = "hrkang97" or session("ssBctId") = "tozzinet"
'C_OFF_AUTH = (session("ssBctId") = "hrkang97") or (session("ssBctId") = "nownhere21") or (session("ssBctId") = "bleuciel05") or (session("ssBctId") = "hsc85") or (session("ssBctId") = "john6136") or (session("ssBctId") = "kjyfer") or (session("ssBctId") = "eccrose") or (session("ssBctId") = "endandand") or (session("ssBctId") = "chhs2000")
	C_OFF_part = (session("ssAdminPsn") = "13" or session("ssAdminPsn") = "24")

	'cs ��Ʈ���ӱ���		' 2019.05.28 �ѿ��
	C_CSPowerUser   = (session("ssAdminPsn")="10" and session("ssAdminLsn")<="3") or session("ssBctId") = "boom15"
    C_CSPowerUser	= C_CSPowerUser or session("ssBctId") = "rabbit1693" or session("ssBctId") = "heendoongi" or session("ssBctId") = "seokmi1221"
	' cs��	' 2019.11.20 �ѿ��
	C_CSUser   = session("ssAdminPsn")="10"
	' cs�� ��Ź��ü		' 2021.04.09 �ѿ��
	C_CSOutsourcingUser = C_CSUser and session("ssAdminLsn")="5"
	' cs�� ��Ź��ü ��Ʈ���� �̻�		' 2021.09.13 �ѿ��
	C_CSOutsourcingPowerUser = C_CSOutsourcingUser and session("ssBctId") = "kangkong83"
	if session("ssAdminPOSITsn")<>"" and not(isnull(session("ssAdminPOSITsn"))) then
		C_CSpermanentUser = session("ssAdminPsn")="10" and session("ssAdminPOSITsn")<=11
	end if
C_MngPart = (session("ssAdminPsn")="8")
C_InspectorUser = session("ssBctId") = "aimcta1"

'' ���޾�ü
dim C_IS_Maker_Upche, C_ADMIN_USER

C_IS_Maker_Upche = (session("ssBctDiv") = "9999")
C_ADMIN_USER     = (session("ssBctDiv") < 10) or (session("ssBctDiv")="301")


If (session("ssBctId") = "") or ((Not C_IS_Maker_Upche) and (Not C_ADMIN_USER)) then
    %><html>
    <script>
    alert("������ ����Ǿ����ϴ�. \n��α����� ����ϽǼ� �ֽ��ϴ�.<%= session("ssBctDiv") %>");
    top.location = "/index.asp";
    </script>
    </html><%
    response.End
End if

'-----------------------------------------------------------------------
' �̺�Ʈ �������� ���� (2007.02.07; ������)
'-----------------------------------------------------------------------
Dim staticImgUrl,uploadUrl,manageUrl,wwwUrl,uploadImgUrl, ItemUploadUrl, staticUploadUrl, imgFingers ,wwwFingers, partnerUrl, fixImgUrl
dim webImgUrl, fingersImgUrl, UploadImgFingers
dim staticImgUrlForMAIL, webImgUrlForMAIL, fixImgUrlForMAIL		'// ���� �߼۽� SSL �������

IF application("Svr_Info")="Dev" THEN

	staticImgUrlForMAIL		= "http://testimgstatic.10x10.co.kr"		'// ���� �߼۽� SSL �������
	webImgUrlForMAIL		= "http://testwebimage.10x10.co.kr"
	fixImgUrlForMAIL		= "http://fiximage.10x10.co.kr"

 	staticImgUrl 	= "http://testimgstatic.10x10.co.kr"	'�׽�Ʈ
 	imgFingers 		= "http://testimage.thefingers.co.kr"
 	uploadUrl		= "http://testimgstatic.10x10.co.kr"
  	fingersImgUrl	= "http://testimage.thefingers.co.kr"
 	manageUrl 	 	= "http://testwebadmin.10x10.co.kr"
 	wwwUrl		 	= "http://test.10x10.co.kr"
 	webImgUrl		= "http://testwebimage.10x10.co.kr"				'���̹���
 	uploadImgUrl 	= "http://testupload.10x10.co.kr"
 	ItemUploadUrl	= "http://testupload.10x10.co.kr"
 	wwwFingers		= "http://test.thefingers.co.kr"
    UploadImgFingers = "http://testimage.thefingers.co.kr"
    
 	partnerUrl		= "http://testwebimage.10x10.co.kr/partner"		'�ӽû�ǰ�̹���(��Ʈ��)
ELSE
	if (C_IS_SSL_ENABLED = True) then
 		staticImgUrl    = "/imgstatic"
 		webImgUrl		= "/webimage"							'���̹���
		fixImgUrl		= "/fiximage"

 		staticImgUrlForMAIL		= "http://imgstatic.10x10.co.kr"		'// ���� �߼۽� SSL �������
 		webImgUrlForMAIL		= "http://webimage.10x10.co.kr"
		fixImgUrlForMAIL		= "http://fiximage.10x10.co.kr"

		wwwFingers		= "http://www.thefingers.co.kr"
 		imgFingers 		= "http://image.thefingers.co.kr"
 		uploadUrl		= "http://oimgstatic.10x10.co.kr"
  		fingersImgUrl	= "http://image.thefingers.co.kr"
 		wwwUrl 		 	= "http://www1.10x10.co.kr"
 		manageUrl 	 	= "http://webadmin.10x10.co.kr"

 		'' 131 to Nas Svr
 		uploadImgUrl 	= "https://upload.10x10.co.kr"          '' upload.10x10.co.kr ���ؼ� Nas Server�� ���ε�
 		ItemUploadUrl	= "https://upload.10x10.co.kr"

 		partnerUrl		= "http://partner.10x10.co.kr"				'�ӽû�ǰ�̹���(��Ʈ��)
 		UploadImgFingers = "http://oimage.thefingers.co.kr"
	else
 		staticImgUrl 	= "http://imgstatic.10x10.co.kr"
		webImgUrl		= "http://webimage.10x10.co.kr"				'���̹���
		fixImgUrl		= "http://fiximage.10x10.co.kr"

 		staticImgUrlForMAIL		= "http://imgstatic.10x10.co.kr"		'// ���� �߼۽� SSL �������
 		webImgUrlForMAIL		= "http://webimage.10x10.co.kr"
		fixImgUrlForMAIL		= "http://fiximage.10x10.co.kr"

		wwwFingers		= "http://www.thefingers.co.kr"
 		imgFingers 		= "http://image.thefingers.co.kr"
 		uploadUrl		= "http://oimgstatic.10x10.co.kr"
  		fingersImgUrl	= "http://image.thefingers.co.kr"
 		wwwUrl 		 	= "http://www1.10x10.co.kr"
 		manageUrl 	 	= "http://webadmin.10x10.co.kr"

 		'' 131 to Nas Svr
 		uploadImgUrl 	= "http://upload.10x10.co.kr"          '' upload.10x10.co.kr ���ؼ� Nas Server�� ���ε�
 		ItemUploadUrl	= "http://upload.10x10.co.kr"

 		partnerUrl		= "http://partner.10x10.co.kr"				'�ӽû�ǰ�̹���(��Ʈ��)
 		UploadImgFingers = "http://oimage.thefingers.co.kr"
	end if
END IF

%>
