<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%

dim C_IS_SSL_ENABLED : C_IS_SSL_ENABLED = (Request.ServerVariables("HTTPS") = "on")

DIM CAddDetailSpliter : CAddDetailSpliter= CHR(3)&CHR(4)

If (session("ssBctId") = "") or (session("ssBctDiv") <> "9999" and session("ssBctDiv") > "9") then
    %><html>
    <script>
    alert("60���� ����Ǿ� �α׾ƿ��Ǿ����ϴ�. \n�ٽ� �α��� �� ����ϽǼ� �ֽ��ϴ�.");
    top.location = "/index.asp";
    </script>
    </html><%
    response.End
End if

'-----------------------------------------------------------------------
' �̺�Ʈ �������� ���� (2007.02.07; ������)
'-----------------------------------------------------------------------
Dim staticImgUrl,uploadUrl,manageUrl,wwwUrl,uploadImgUrl, ItemUploadUrl, staticUploadUrl, imgFingers, wwwFingers, partnerUrl, fixImgUrl, UploadImgFingers
dim webImgUrl
dim staticImgUrlForMAIL, webImgUrlForMAIL, fixImgUrlForMAIL		'// ���� �߼۽� SSL �������

IF application("Svr_Info")="Dev" THEN
 	staticImgUrl 		= "http://testimgstatic.10x10.co.kr"	'�׽�Ʈ
	webImgUrl			= "http://testwebimage.10x10.co.kr"				'���̹���
	fixImgUrl			= "http://fiximage.10x10.co.kr"

	staticImgUrlForMAIL		= "http://testimgstatic.10x10.co.kr"		'// ���� �߼۽� SSL �������
	webImgUrlForMAIL		= "http://testwebimage.10x10.co.kr"
	fixImgUrlForMAIL		= "http://fiximage.10x10.co.kr"

 	imgFingers 			= "http://testimage.thefingers.co.kr"
 	wwwFingers			= "http://test.thefingers.co.kr"
 	uploadUrl			= "http://testimgstatic.10x10.co.kr"
 	manageUrl 	 		= "http://testwebadmin.10x10.co.kr"
 	wwwUrl		 		= "http://test.10x10.co.kr"

 	uploadImgUrl 		= "http://testupload.10x10.co.kr"
 	ItemUploadUrl		= "http://testupload.10x10.co.kr"
 	partnerUrl 			= "http://testwebimage.10x10.co.kr/partner"		'�ӽû�ǰ�̹���(��Ʈ��)
 	UploadImgFingers    = "http://testimage.thefingers.co.kr"
ELSE
	if (C_IS_SSL_ENABLED = True) then
		staticImgUrl    	= "/imgstatic"
 		webImgUrl			= "/webimage"							'���̹���
		fixImgUrl			= "/fiximage"

 		staticImgUrlForMAIL		= "http://imgstatic.10x10.co.kr"		'// ���� �߼۽� SSL �������
 		webImgUrlForMAIL		= "http://webimage.10x10.co.kr"
		fixImgUrlForMAIL		= "http://fiximage.10x10.co.kr"

 		imgFingers 			= "http://image.thefingers.co.kr"
 		wwwFingers			= "http://www.thefingers.co.kr"
 		uploadUrl			= "http://oimgstatic.10x10.co.kr"
 		wwwUrl 		 		= "http://www1.10x10.co.kr"
 		manageUrl 	 		= "http://webadmin.10x10.co.kr"

 		'' 131 to Nas Svr
 		uploadImgUrl 		= "https://upload.10x10.co.kr"          '' upload.10x10.co.kr ���ؼ� Nas Server�� ���ε�
 		ItemUploadUrl		= "https://upload.10x10.co.kr"
 		partnerUrl 			= "http://partner.10x10.co.kr"				'�ӽû�ǰ�̹���(��Ʈ��)
 		UploadImgFingers    = "http://oimage.thefingers.co.kr"
	else
 		staticImgUrl 		= "http://imgstatic.10x10.co.kr"
		webImgUrl			= "http://webimage.10x10.co.kr"				'���̹���
		fixImgUrl			= "http://fiximage.10x10.co.kr"

 		staticImgUrlForMAIL		= "http://imgstatic.10x10.co.kr"		'// ���� �߼۽� SSL �������
 		webImgUrlForMAIL		= "http://webimage.10x10.co.kr"
		fixImgUrlForMAIL		= "http://fiximage.10x10.co.kr"

 		imgFingers 			= "http://image.thefingers.co.kr"
 		wwwFingers			= "http://www.thefingers.co.kr"
 		uploadUrl			= "http://oimgstatic.10x10.co.kr"
 		wwwUrl 		 		= "http://www1.10x10.co.kr"
 		manageUrl 	 		= "http://webadmin.10x10.co.kr"

 		'' 131 to Nas Svr
 		uploadImgUrl 		= "http://upload.10x10.co.kr"          '' upload.10x10.co.kr ���ؼ� Nas Server�� ���ε�
 		ItemUploadUrl		= "http://upload.10x10.co.kr"
 		partnerUrl 			= "http://partner.10x10.co.kr"				'�ӽû�ǰ�̹���(��Ʈ��)
 		UploadImgFingers    = "http://oimage.thefingers.co.kr"
	end if
 END IF
%>
