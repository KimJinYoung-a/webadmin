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
    alert("60분이 경과되어 로그아웃되었습니다. \n다시 로그인 후 사용하실수 있습니다.");
    top.location = "/index.asp";
    </script>
    </html><%
    response.End
End if

'-----------------------------------------------------------------------
' 이벤트 전역변수 선언 (2007.02.07; 정윤정)
'-----------------------------------------------------------------------
Dim staticImgUrl,uploadUrl,manageUrl,wwwUrl,uploadImgUrl, ItemUploadUrl, staticUploadUrl, imgFingers, wwwFingers, partnerUrl, fixImgUrl, UploadImgFingers
dim webImgUrl
dim staticImgUrlForMAIL, webImgUrlForMAIL, fixImgUrlForMAIL		'// 메일 발송시 SSL 고려안함

IF application("Svr_Info")="Dev" THEN
 	staticImgUrl 		= "http://testimgstatic.10x10.co.kr"	'테스트
	webImgUrl			= "http://testwebimage.10x10.co.kr"				'웹이미지
	fixImgUrl			= "http://fiximage.10x10.co.kr"

	staticImgUrlForMAIL		= "http://testimgstatic.10x10.co.kr"		'// 메일 발송시 SSL 고려안함
	webImgUrlForMAIL		= "http://testwebimage.10x10.co.kr"
	fixImgUrlForMAIL		= "http://fiximage.10x10.co.kr"

 	imgFingers 			= "http://testimage.thefingers.co.kr"
 	wwwFingers			= "http://test.thefingers.co.kr"
 	uploadUrl			= "http://testimgstatic.10x10.co.kr"
 	manageUrl 	 		= "http://testwebadmin.10x10.co.kr"
 	wwwUrl		 		= "http://test.10x10.co.kr"

 	uploadImgUrl 		= "http://testupload.10x10.co.kr"
 	ItemUploadUrl		= "http://testupload.10x10.co.kr"
 	partnerUrl 			= "http://testwebimage.10x10.co.kr/partner"		'임시상품이미지(파트너)
 	UploadImgFingers    = "http://testimage.thefingers.co.kr"
ELSE
	if (C_IS_SSL_ENABLED = True) then
		staticImgUrl    	= "/imgstatic"
 		webImgUrl			= "/webimage"							'웹이미지
		fixImgUrl			= "/fiximage"

 		staticImgUrlForMAIL		= "http://imgstatic.10x10.co.kr"		'// 메일 발송시 SSL 고려안함
 		webImgUrlForMAIL		= "http://webimage.10x10.co.kr"
		fixImgUrlForMAIL		= "http://fiximage.10x10.co.kr"

 		imgFingers 			= "http://image.thefingers.co.kr"
 		wwwFingers			= "http://www.thefingers.co.kr"
 		uploadUrl			= "http://oimgstatic.10x10.co.kr"
 		wwwUrl 		 		= "http://www1.10x10.co.kr"
 		manageUrl 	 		= "http://webadmin.10x10.co.kr"

 		'' 131 to Nas Svr
 		uploadImgUrl 		= "https://upload.10x10.co.kr"          '' upload.10x10.co.kr 통해서 Nas Server로 업로드
 		ItemUploadUrl		= "https://upload.10x10.co.kr"
 		partnerUrl 			= "http://partner.10x10.co.kr"				'임시상품이미지(파트너)
 		UploadImgFingers    = "http://oimage.thefingers.co.kr"
	else
 		staticImgUrl 		= "http://imgstatic.10x10.co.kr"
		webImgUrl			= "http://webimage.10x10.co.kr"				'웹이미지
		fixImgUrl			= "http://fiximage.10x10.co.kr"

 		staticImgUrlForMAIL		= "http://imgstatic.10x10.co.kr"		'// 메일 발송시 SSL 고려안함
 		webImgUrlForMAIL		= "http://webimage.10x10.co.kr"
		fixImgUrlForMAIL		= "http://fiximage.10x10.co.kr"

 		imgFingers 			= "http://image.thefingers.co.kr"
 		wwwFingers			= "http://www.thefingers.co.kr"
 		uploadUrl			= "http://oimgstatic.10x10.co.kr"
 		wwwUrl 		 		= "http://www1.10x10.co.kr"
 		manageUrl 	 		= "http://webadmin.10x10.co.kr"

 		'' 131 to Nas Svr
 		uploadImgUrl 		= "http://upload.10x10.co.kr"          '' upload.10x10.co.kr 통해서 Nas Server로 업로드
 		ItemUploadUrl		= "http://upload.10x10.co.kr"
 		partnerUrl 			= "http://partner.10x10.co.kr"				'임시상품이미지(파트너)
 		UploadImgFingers    = "http://oimage.thefingers.co.kr"
	end if
 END IF
%>
