<%

dim C_IS_SSL_ENABLED : C_IS_SSL_ENABLED = (Request.ServerVariables("HTTPS") = "on")
dim C_Relationship_Part
dim C_ADMIN_AUTH
dim C_OFF_AUTH, C_OFF_part
dim C_MngPart               '' �濵������ ����.
dim C_InspectorUser			''����

C_ADMIN_AUTH = (session("ssBctId") = "coolhas") or (session("ssBctId") = "kobula") or (session("ssBctId") = "tozzinet") or (session("ssBctId") = "kjy8517") or (session("ssBctId") = "thensi7")

'�������κ����̰ų� �������̰� ��Ʈ���� �̻�. �¶���MD(11)�� ������
C_OFF_AUTH = (session("ssAdminPsn") = "13" or session("ssAdminPsn") = "24" or session("ssAdminPsn") = "11") and session("ssAdminLsn")<="3"
C_OFF_AUTH = C_OFF_AUTH or session("ssBctId") = "tozzinet"
	C_OFF_part = (session("ssAdminPsn") = "13" or session("ssAdminPsn") = "24")

C_MngPart = (session("ssAdminPsn")="8")
C_InspectorUser = session("ssBctId") = "aimcta1"

	' ����� ��		' 2020.06.02 �ѿ��
	C_Relationship_Part = session("ssAdminPsn")="17" 


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

C_IS_FRN_SHOP = (session("ssBctDiv") = "502") or (session("ssBctDiv") = "503")
C_IS_SHOP = (C_IS_OWN_SHOP or C_IS_FRN_SHOP)
C_ADMIN_USER     = (session("ssBctDiv") < 10)

if C_IS_FRN_SHOP then
	C_STREETSHOPID = session("ssBctId")
elseif C_IS_OWN_SHOP then
	if (session("ssBctDiv") = "501") or (session("ssBctDiv") = "502") then
		C_STREETSHOPID = session("ssBctid")
	''elseif (session("ssBctDiv")="201" or session("ssAdminPsn")="6") then        ''��ȭ��
	''    C_STREETSHOPID = "cafe002"
	''elseif (session("ssBctDiv")="301" or session("ssAdminPsn")="16") then       ''��ī����
	''    C_STREETSHOPID = "cafe003"
	else
		C_STREETSHOPID = session("ssBctBigo")
	end if
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
''�˻����� ����
Dim DocSvrAddr, DocSvrPort, DocAuthCode

IF application("Svr_Info")="Dev" THEN
 	staticImgUrl = "http://testimgstatic.10x10.co.kr"	'�׽�Ʈ
 	webImgUrl		= "http://testwebimage.10x10.co.kr"				'���̹���
	fixImgUrl			= "http://fiximage.10x10.co.kr"

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
ELSE
	if (C_IS_SSL_ENABLED = True) then
 		staticImgUrl    = "/imgstatic"
 		webImgUrl		= "/webimage"							'���̹���
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
		uploadUrl	    = "http://oimgstatic.10x10.co.kr"
		uploadImgUrl    = "https://upload.10x10.co.kr"          '' upload.10x10.co.kr ���ؼ� Nas Server�� ���ε�
 		ItemUploadUrl	= "https://upload.10x10.co.kr"

		staticUploadUrl = "http://oimgstatic.10x10.co.kr"
	else
 		staticImgUrl    = "http://imgstatic.10x10.co.kr"
 		webImgUrl		= "http://webimage.10x10.co.kr"				'���̹���
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
	end if
END IF

%>
