<%
'-----------------------------------------------------------------------
' �̺�Ʈ �������� ����
'-----------------------------------------------------------------------
 Dim staticImgUrl,uploadUrl,manageUrl,wwwUrl, uploadImgUrl,othermall,mailzine,www2009url, ItemUploadUrl, staticUploadUrl, webImgUrl, mobileUrl, partnerUrl
 Dim wwwFingers, imgFingers
  ''�˻����� ����
 Dim DocSvrAddr, DocSvrPort, DocAuthCode, menupos

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
 	staticImgUrl    = "http://imgstatic.10x10.co.kr"
 	webImgUrl		= "http://webimage.10x10.co.kr"				'���̹���

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
 	partnerUrl		= "http://partner.10x10.co.kr"				'�ӽû�ǰ�̹���(��Ʈ��)

	staticUploadUrl = "http://oimgstatic.10x10.co.kr"
 END IF
%>