<%
'<!-- #include virtual="/admin/incTenRedisSession.asp"-->
'Call fn_RDS_CHK_SSN_RESTORE()
%>
<%
' '<!-- #include virtual="/lib/util/base64unicode.asp" -->
' '<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>

' '' stscm. ���� ���°�� session ���� �缼������
' function G_isReqCheckHost()
'     G_isReqCheckHost = false
'     dim iHostName : iHostName = request.serverVariables("HTTP_HOST")
    
'     IF application("Svr_Info")="Dev" THEN
'         G_isReqCheckHost = (LCASE(iHostName)="teststscm.10x10.co.kr")
'     else
'         G_isReqCheckHost = (LCASE(iHostName)="stscm.10x10.co.kr")
'     end if
    
' end function

' function G_getLoginSessionByScm()
'     G_getLoginSessionByScm = False
    
'     dim ssnChkUrl : ssnChkUrl = "https://webadmin.10x10.co.kr/apps/protoV1/getLoginSessions.asp"
'     dim ibearerTkn : ibearerTkn = "bearer 1L6O9L>N8CAM@CEFH:D<G:N?O:L6NO6e8O>[7F?^?FGO>=FHF=NTN4K9M4U\7R6P6I>N>IFT"
'     dim istrParam : istrParam = ""
    
'     ''iaspsessionIdCookies="ASPSESSIONIDSAQBSCTA=EBLLAKGBILAIJALDLKCKIMDH"
'     ''iaspsessionIdCookies="ASPSESSIONIDSARDRCTB=MEMNGAGBKBEPBBCONGPOOEME"

'     IF application("Svr_Info")="Dev" THEN
'         ssnChkUrl = "https://testwebadmin.10x10.co.kr/apps/protoV1/getLoginSessions.asp"
'     end if
    
'     if (request.cookies("TENSSID")="") then Exit function
        
'     Dim kk, icookieName, icookieValue
'     Dim iTENSSID : iTENSSID = Base64decodeUnicode(request.cookies("TENSSID"))
'     Dim iTENSSIDArr : iTENSSIDArr = split(iTENSSID,";")
'     Dim retSsnJson, objObj
    
'     Dim xmlHttp
'     Set xmlHttp= CreateObject("MSXML2.ServerXMLHTTP.3.0")
' 	    xmlHttp.Open "POST", ssnChkUrl , False
' 		xmlHttp.setRequestHeader "Content-Type", "application/json"
' 		xmlHttp.SetRequestHeader "Authorization", ibearerTkn
' 		If isArray(iTENSSIDArr) then
'             For kk = LBound(iTENSSIDArr) To UBound(iTENSSIDArr)
'                 if InStr(iTENSSIDArr(kk),"=")>0 then
'                     icookieName = Trim(split(iTENSSIDArr(kk),"=")(0))
'                     icookieValue = Trim(split(iTENSSIDArr(kk),"=")(1))
'                     If instr(icookieName,"ASPSESSIONID") then
'                         ''response.write icookieName & "=" & icookieValue&"<br>"
'                         xmlHttp.SetRequestHeader "Cookie", icookieName & "=" & icookieValue
'                         end if
'                  End If
'             Next
'         end if

' 		xmlHttp.Send(istrParam)
' 		If xmlHttp.Status = "200" Then
' 		    ''response.write xmlHttp.ResponseBody
' 			retSsnJson = BinaryToText(xmlHttp.ResponseBody,"utf-8")
			
' 			Set objObj = JSON.parse(retSsnJson)
' 			session("ssBctId") = objObj.ssBctId         '�α��� ���̵�
'             session("ssBctDiv")	= objObj.ssBctDiv  		'ȸ������
'             session("ssBctBigo") = objObj.ssBctBigo		'���� �߰� ����
'             session("ssBctSn")	 = objObj.ssBctSn		'������ȣ
'             session("ssBctCname") = objObj.ssBctCname	'���� �̸�
'             session("ssBctEmail") = objObj.ssBctEmail   '���� �̸���
            
'             session("ssGroupid") = objObj.ssGroupid     '�׷� �ڵ�
'             session("ssAdminPsn") = objObj.ssAdminPsn   '�μ� ��ȣ
'             session("ssAdminLsn") = objObj.ssAdminLsn   '��� ��ȣ
'             session("ssAdminPOsn") = objObj.ssAdminPOsn '��å ��ȣ
'             session("ssAdminPOSITsn") = objObj.ssAdminPOSITsn   '���� ��ȣ
'             session("ssAdminCLsn") = objObj.ssAdminCLsn '�������� ��ޱ���

' 			Set objObj = Nothing
			
' 			G_getLoginSessionByScm = true
' 		else
' 		    'response.write xmlHttp.ResponseBody 'Status
' 		end if
' 	Set xmlHttp= Nothing
' end function

' if (session("ssBctId")="") then
'     if (G_isReqCheckHost) then
'         if NOT (G_getLoginSessionByScm()) then
'             ''������ ����Ǿ����� ������ ����.
'             if application("Svr_Info")="Dev" then 
'                 response.write "<script>alert('oops! session expired');top.location = 'http://testwebadmin.10x10.co.kr/adminIndex.asp';</script>"
'             else
'                 response.write "<script>alert('oops! session expired');top.location = 'http://webadmin.10x10.co.kr/adminIndex.asp';</script>"
'             end if
'             response.end
'         end if
'     end if
' end if
%>