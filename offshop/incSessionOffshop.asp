<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
dim C_ADMIN_AUTH : C_ADMIN_AUTH = FALSE

If (session("ssBctId") = "") or ((session("ssBctDiv") <> "501") and (session("ssBctDiv") <> "502") and (session("ssBctDiv") <> "503") and (session("ssBctDiv") <> "509")) then
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
 Dim staticImgUrl,uploadUrl,manageUrl,wwwUrl, uploadImgUrl
 IF application("Svr_Info")="Dev" THEN
 	staticImgUrl = "http://testimgstatic.10x10.co.kr"	'테스트
 	uploadUrl	 = "http://testimgstatic.10x10.co.kr"   ''차후 정리요망
 	manageUrl 	 = "http://testwebadmin.10x10.co.kr"
 	wwwUrl		 = "http://test.10x10.co.kr"            ''차후 정리요망
 	uploadImgUrl = "http://testupload.10x10.co.kr"
 ELSE
 	staticImgUrl = "http://imgstatic.10x10.co.kr"	
 	uploadUrl	="http://oimgstatic.10x10.co.kr"
 	wwwUrl 		 = "http://www1.10x10.co.kr"
 	manageUrl 	 = "http://webadmin.10x10.co.kr"
 	uploadImgUrl = "http://upload.10x10.co.kr"          '' upload.10x10.co.kr 통해서 Nas Server로 업로드
 END IF	
%>