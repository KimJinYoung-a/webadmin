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
    alert("������ ����Ǿ����ϴ�. \n��α����� ����ϽǼ� �ֽ��ϴ�.");
    top.location = "/index.asp";
    </script>
    </html><%
    response.End
End if

'-----------------------------------------------------------------------
' �̺�Ʈ �������� ���� (2007.02.07; ������)
'-----------------------------------------------------------------------
 Dim staticImgUrl,uploadUrl,manageUrl,wwwUrl, uploadImgUrl
 IF application("Svr_Info")="Dev" THEN
 	staticImgUrl = "http://testimgstatic.10x10.co.kr"	'�׽�Ʈ
 	uploadUrl	 = "http://testimgstatic.10x10.co.kr"   ''���� �������
 	manageUrl 	 = "http://testwebadmin.10x10.co.kr"
 	wwwUrl		 = "http://test.10x10.co.kr"            ''���� �������
 	uploadImgUrl = "http://testupload.10x10.co.kr"
 ELSE
 	staticImgUrl = "http://imgstatic.10x10.co.kr"	
 	uploadUrl	="http://oimgstatic.10x10.co.kr"
 	wwwUrl 		 = "http://www1.10x10.co.kr"
 	manageUrl 	 = "http://webadmin.10x10.co.kr"
 	uploadImgUrl = "http://upload.10x10.co.kr"          '' upload.10x10.co.kr ���ؼ� Nas Server�� ���ε�
 END IF	
%>