<%

'// ���Ȼ� ����ϴ�. �߿����� ���� �����̰�, ���� �ٸ� ���굵���� ������ ���� �뵵 �̿ܿ��� ����ϸ� �ʵȴ�.

'' �Ʒ��� ���� �δܰ�� �ؾ� referrer �� �νĵȴ�.
''function PopProductOffline(barcode) {
''	var popwin = window.open('','PopProductOffline','width=1100,height=600,resizabled=yes,scrollbars=yes');
''	popwin.location.href = "http://webadmin.10x10.co.kr/common/offshop/item/pop_itemview_off_view.asp?barcode=" + barcode;
''	popwin.focus();
''}

Function CheckReferrerOnly()
	dim refURL
	dim allowURL, arrAllowURL, tmpAllowURL
	dim i

	allowURL = "http://testlogics.10x10.co.kr/,http://logics.10x10.co.kr/,http://testwebadmin.10x10.co.kr/,http://webadmin.10x10.co.kr/"
	CheckReferrerOnly = False
	refURL = Request.ServerVariables("HTTP_REFERER")
	arrAllowURL = Split(allowURL, ",")

	for i = 0 to UBound(arrAllowURL) - 1
		tmpAllowURL = arrAllowURL(i)
		if (tmpAllowURL = Left(refURL, Len(tmpAllowURL))) then
			CheckReferrerOnly = True
		end if
	next
End Function

If CheckReferrerOnly() = False Then
	Response.Write "�߸��� �����Դϴ�."
	Response.End
End If

'-----------------------------------------------------------------------
' �̺�Ʈ �������� ���� (2007.02.07; ������)
'-----------------------------------------------------------------------
 Dim staticImgUrl,uploadUrl,manageUrl,wwwUrl, uploadImgUrl,othermall,mailzine,www2009url, ItemUploadUrl, staticUploadUrl, webImgUrl, mobileUrl
 Dim wwwFingers, imgFingers
  ''�˻����� ����
 Dim DocSvrAddr, DocSvrPort, DocAuthCode

 IF application("Svr_Info")="Dev" THEN
 	staticImgUrl = "http://testimgstatic.10x10.co.kr"	'�׽�Ʈ
 	webImgUrl		= "http://testwebimage.10x10.co.kr"				'���̹���

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

	staticUploadUrl = "http://oimgstatic.10x10.co.kr"
 END IF

%>