<%@ language="VBScript" %>
<% option explicit %>
 
<%
'###########################################################
' Description : ����  ���
' History : 2011.03.16 ������  ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/email/smslib.asp"-->
<!-- #include virtual="/lib/classes/cooperate/cooperateCls.asp"-->
<%
Dim objCmd, returnValue, sMode, sMode_H
Dim reportIdx,payrequestidx
Dim arapcd,reportName,reportPrice,scmLink,bigo,reportContents,reportState,referId,adminID,authId,authposition,issms,Comment, authId_H, issms_H,authId_L, authId_F
Dim eapppartIdx, partMoney
Dim fileName,referLink, filelink, filetype
Dim i ,blnusing, commentidx, isLast, authstate
Dim iMenuNo
dim arreapppartIdx,lasteapppartIdx
Dim returnUrl
Dim edmsIdx,sscmsubmitlink
Dim reqhp, reqEmail, smstext,susername,sauthname,sUserid, smsurl
Dim iRectMenu
Dim ipayType, sCurrencyType, mCurrencyPrice
Dim arrAuthId
Dim ipos, sRectAuthType 
Dim slastApprovalid
	 
iRectMenu = requestCheckvar(Request("iRM"),10)
sMode		= requestCheckvar(Request("hidM"),2)
sMode_H     = requestCheckvar(Request("hidM_H"),2)
reportIdx	= requestCheckvar(Request("irIdx"),10)
payrequestidx=requestCheckvar(Request("iprIdx"),10)
IF payrequestidx = "" THEN payrequestidx = 0
arapcd	= requestCheckvar(Request("iAIdx"),10)
edmsIdx		=  requestCheckvar(Request("ieidx"),10)
reportName	= requestCheckvar(Request("sRN"),60)
reportPrice	= getNumeric(requestCheckvar(Request("mRP"),20))
scmLink		= requestCheckvar(Request("iSL"),10)
sscmsubmitlink =  requestCheckvar(Request("sSSL"),128)
if scmLink = "" THEN scmLink = 0
bigo		= ReplaceRequestSpecialChar(Request("sB"))
 
reportContents	= ReplaceRequestSpecialChar(Request("editor")) 

reportState	= requestCheckvar(Request("hidRS"),4)
'���缱
referId		= ReplaceRequestSpecialChar(Request("hidRfI")) '����
authId		= ReplaceRequestSpecialChar(Request("hidAI")) '���缱
authId_L	= requestCheckvar(Request("hidALI"),32) '����������
authId_F	= requestCheckvar(Request("hidAHI"),32) '����������
authId_H	= requestCheckvar(Request("hidAI_H"),32) '����
isLast		= requestCheckvar(Request("blnL"),1) '�������� ��Ͽ��� 
issms		= requestCheckvar(Request("chkSms"),1)
issms_H		= requestCheckvar(Request("chkSms_H"),1) 
if issms ="" then issms =0
if issms_H ="" then issms_H =0
	 
authposition= requestCheckvar(Request("iAP"),10)	
sUserid   = requestCheckvar(Request("hidAId"),32)'�����
adminID		= session("ssBctId") 
Comment		= ReplaceRequestSpecialChar(Request("tCmt"))

eapppartIdx= ReplaceRequestSpecialChar(Request("ip"))
partMoney	= ReplaceRequestSpecialChar(Request("mP"))

IF  partMoney = "" THEN partMoney = 0

fileName 	= ReplaceRequestSpecialChar(Request("sFile")) 
referLink	= ReplaceRequestSpecialChar(Request("sL"))
blnusing	= requestCheckvar(Request("blnU"),1)
commentidx = requestCheckvar(Request("iCidx"),10)

susername	= requestCheckvar(Request("hidUN"),30)
authstate	= requestCheckvar(Request("hidAS"),4)
sauthname = replace(requestCheckvar(Request("hidAN"),30),"&nbsp;"," ")
returnUrl   =  requestCheckvar(Request("hidRU"),100)

ipayType	= requestCheckvar(Request("selPT"),4)
sCurrencyType	= requestCheckvar(Request("selCT"),3)
mCurrencyPrice	= requestCheckvar(Request("sCP"),20)
sRectAuthType =requestCheckvar(Request("iRAT"),1)
IF ipayType = "" THEN ipayType = 0
slastApprovalid =  requestCheckvar(Request("iLAID"),10)  

''rw sMode
SELECT CASE sMode
Case "I"
	Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_eAppReport_insert]("&edmsIdx&", "&arapcd&" ,'"&reportName&"', '"&reportPRice&"'"&_
						+","&scmLink&",'"&bigo&"','"&reportContents&"',"&reportState&",'"&referId&"','"&adminID&"',"&ipayType&",'"&sCurrencyType&"','"&mCurrencyPrice&"' ,'"&slastApprovalid&"')}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
	Set objCmd = nothing

	IF returnValue > 0 THEN
		reportIdx = returnValue

		'÷������ ���
		fileName = split(fileName,",")
		For i = 0 To UBound(fileName)
		if (trim(fileName(i)) <> "") then
		Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_eAppAttachFile_insert]( "&reportIdx&" ,0,'"&trim(fileName(i))&"', 1)}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
		Set objCmd = nothing
		end if
		Next

		'��ũ ���
		referLink = split(referLink,",")
		For i = 0 To UBound(referLink)
		if(trim(referLink(i)) <> "") then
		Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_eAppAttachFile_insert]( "&reportIdx&",0 ,'"&trim(referLink(i))&"', 0)}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
		Set objCmd = nothing
		end if
		Next


		'�μ��� �ڱݱ��� ���
		IF eapppartIdx <> "" THEN
		eapppartIdx = split(eapppartIdx,",")
		partMoney = split(partMoney,",")

		For i = 0 To UBound(eapppartIdx)
		Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_eappPartMoney_insert]( "&reportIdx&" ,0,'"&trim(eapppartIdx(i))&"','"&getNumeric(partMoney(i))&"')}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
		Set objCmd = nothing
		Next
		END IF

	 '���缱 ���  
	 ipos = 0
	  IF authId <> "" THEN
	  	arrAuthId = split(authId,",")
	  	For i = 0 To UBound(arrAuthId)
			Set objCmd = Server.CreateObject("ADODB.COMMAND")
			With objCmd
				.ActiveConnection = dbget
				.CommandType = adCmdText
				.CommandText = "{?= call db_partner.[dbo].[sp_Ten_eAppAuthLine_insert]( "&reportIdx&" ,"&i+1&",'"&trim(arrAuthId(i))&"','D','"&isSMS&"')}"
				.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
				.Execute, , adExecuteNoRecords
				End With
			    returnValue = objCmd(0).Value
			Set objCmd = nothing
			Next
			ipos = i
		END IF
 
			'���ǵ��
		 IF authId_H <> "" THEN
		 	ipos = ipos + 1
		 	Set objCmd = Server.CreateObject("ADODB.COMMAND")
			With objCmd
				.ActiveConnection = dbget
				.CommandType = adCmdText
				.CommandText = "{?= call db_partner.[dbo].[sp_Ten_eAppAuthLine_insert]( "&reportIdx&" ,"&ipos&",'"&trim(authId_H)&"','A','"&isSMS_H&"')}"
				.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
				.Execute, , adExecuteNoRecords
				End With
			    returnValue = objCmd(0).Value
			Set objCmd = nothing
		 END IF
		 
		'���������ڵ��
		 IF authId_L <> "" THEN 
		 		ipos = ipos + 1
		 	Set objCmd = Server.CreateObject("ADODB.COMMAND")
			With objCmd
				.ActiveConnection = dbget
				.CommandType = adCmdText
				.CommandText = "{?= call db_partner.[dbo].[sp_Ten_eAppAuthLine_insert]( "&reportIdx&" ,"&ipos&",'"&trim(authId_L)&"','L','"&isSMS&"')}"
				.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
				.Execute, , adExecuteNoRecords
				End With
			    returnValue = objCmd(0).Value
			Set objCmd = nothing
		 END IF 

		'���������ڵ��
		 IF authId_F <> "" THEN 
		 		ipos = ipos + 1
		 	Set objCmd = Server.CreateObject("ADODB.COMMAND")
			With objCmd
				.ActiveConnection = dbget
				.CommandType = adCmdText
				.CommandText = "{?= call db_partner.[dbo].[sp_Ten_eAppAuthLine_insert]( "&reportIdx&" ,"&ipos&",'"&trim(authId_F)&"','F','"&isSMS&"')}"
				.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
				.Execute, , adExecuteNoRecords
				End With
			    returnValue = objCmd(0).Value
			Set objCmd = nothing
		 END IF 

		IF (isSMS_H <> "1") and isSMS = "1" and reportState = 1 and (authid <> "" or authid_L <> "" ) THEN
			if authid <> "" then
				reqhp = fnGetMemberHp(trim(arrAuthId(0)))
				reqEmail = fnGetMemberEmail(trim(arrAuthId(0)))
			else
				reqhp = fnGetMemberHp(authId_L)
				reqEmail = fnGetMemberEmail(authId_L)
			end if	
    		smstext = "["&reportName&"] ������ ["&susername&"]�����κ��� �����û�Ǿ����ϴ�."
			smstext = smstext & vbCrLf & vbCrLf & "> �ٷΰ���"
			smsurl = getSCMSSLURL & "/admin/approval/eapp/confirmeapp.asp?iridx=" & reportIdx
    		
			if reqEmail<>"" then
				smstext = chrbyte(Trim(smstext),1000,"Y")
				Call SendRadioWebHookMessage(reqEmail,"admin","SCM ���ڰ���","�����û",smstext,smsurl)
			else
				Call SendMultiRowsSMS(reqhp,"",smstext,"") 
			end if
    	END IF

		''2013/10/21 �߰�
		IF isSMS_H = "1" and reportState = 1 and authid_H <> "" THEN
    		reqhp = fnGetMemberHp(authid_H)
			reqEmail = fnGetMemberEmail(authid_H)
    		smstext = "["&reportName&"] ������ ["&susername&"]�����κ��� ���ǿ�û�Ǿ����ϴ�."
			smstext = smstext & vbCrLf & vbCrLf & "> �ٷΰ���"
			smsurl = getSCMSSLURL & "/admin/approval/eapp/confirmeapp.asp?iridx=" & reportIdx

			if reqEmail<>"" then
				smstext = chrbyte(Trim(smstext),1000,"Y")
				Call SendRadioWebHookMessage(reqEmail,"admin","SCM ���ڰ���","���ǿ�û",smstext,smsurl)
			else
				Call SendMultiRowsSMS(reqhp,"",smstext,"") 
			end if
    	END IF
 
		IF edmsIdx = 1 THEN
			%>
			<script language="javascript">
			<!--
			 var winEapp = window.open("/admin/approval/eapp/popIndex.asp?iRM=M01<%=reportState%>","popEapp","width="+screen.availWidth+", height="+ screen.availHeight +",resizable=yes, scrollbars=yes");
	 			winEapp.focus();
				opener.self.close();
				self.close();
			//-->
			</script>
			<%
		ELSEIF edmsIdx = 2  or edmsidx = 33 THEN
			%>
			<script language="javascript">
			<!--
				 var winEapp = window.open("/admin/approval/eapp/popIndex.asp?iRM=M01<%=reportState%>","popEapp","width="+screen.availWidth+", height="+ screen.availHeight +",resizable=yes, scrollbars=yes");
	 			winEapp.focus();
				self.close();
			//-->
			</script>
			<%
		ELSE
%>
		<script language="javascript">
			<!--
				alert("��ϵǾ����ϴ�");
				opener.top.location.href = "/admin/approval/eapp/popIndex.asp?iRM=M01<%=reportState%>";
				self.close();
			//-->
			</script>
<%
		END IF
	ELSEIF 	returnValue = -1 THEN
			Call Alert_return ("���������Ͱ� �����մϴ�.Ȯ�� �� �ٽ� �Է����ּ���")
	ELSE
			Call Alert_return ("������ ó���� ������ �߻��Ͽ����ϴ�.1")
	END IF
	 
	response.end
Case "U"
''�ӽ� 2014-10-29
if reportContents = "undefined" and  ReplaceRequestSpecialChar(Request("Ueditor")) <> "" then
	reportContents = ReplaceRequestSpecialChar(Request("Ueditor")) 
end if 
 
	Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_eAppReport_update]( "&reportIdx&",'"&reportName&"', '"&reportPRice&"'"&_
						+","&scmLink&",'"&bigo&"','"&reportContents&"',"&reportState&",'"&referId&"','"&adminID&"',"&ipayType&",'"&sCurrencyType&"','"&mCurrencyPrice&"',"&arapcd&")}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
	Set objCmd = nothing

	IF returnValue > 0 THEN
		 Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_eAppAttachFile_Delete]("&reportIdx&",0)}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
		Set objCmd = nothing

		fileName = split(fileName,",")
		For i = 0 To UBound(fileName)
		if(trim(fileName(i)) <> "") then
		Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_eAppAttachFile_insert]( "&reportIdx&" ,0,'"&trim(fileName(i))&"', 1)}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
		Set objCmd = nothing
		end if
		Next

		referLink = split(referLink,",")
		For i = 0 To UBound(referLink)
		if(trim(referLink(i)) <> "") then
		Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_eAppAttachFile_insert]( "&reportIdx&" ,0,'"&trim(referLink(i))&"', 0)}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
		Set objCmd = nothing
		end if
		Next

	 '�μ��� �ڱݱ��� ���
	 	Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_eappPartMoney_Delete]( "&reportIdx&" ,0)}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
		 	.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
		Set objCmd = nothing


		IF eapppartIdx <> "" THEN
		eapppartIdx = split(eapppartIdx,",")
		partMoney = split(partMoney,",")

		For i = 0 To UBound(eapppartIdx)
		Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_eappPartMoney_insert]( "&reportIdx&" ,0,'"&trim(eapppartIdx(i))&"','"&getNumeric(partMoney(i))&"')}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
		Set objCmd = nothing
		Next
		END IF

		'���缱 ��� 
			'-- �������� authstate = 0 �϶� ���缱 ������ ���
			'-- �������� ���� �� ���� ���
			Set objCmd = Server.CreateObject("ADODB.COMMAND")
			With objCmd
				.ActiveConnection = dbget
				.CommandType = adCmdText
				.CommandText = "{?= call db_partner.[dbo].[sp_Ten_eAppAuthLine_Delete]( "&reportIdx&")}"
				.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
				.Execute, , adExecuteNoRecords
				End With
			    returnValue = objCmd(0).Value
			Set objCmd = nothing 
 
	 ipos = 0	
	  IF authId <> "" THEN
	  	arrAuthId = split(authId,",")
	  	For i = 0 To UBound(arrAuthId)
			Set objCmd = Server.CreateObject("ADODB.COMMAND")
			With objCmd
				.ActiveConnection = dbget
				.CommandType = adCmdText
				.CommandText = "{?= call db_partner.[dbo].[sp_Ten_eAppAuthLine_insert]( "&reportIdx&" ,"&i+1&",'"&trim(arrAuthId(i))&"','D','"&isSMS&"')}"
				.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
				.Execute, , adExecuteNoRecords
				End With
			    returnValue = objCmd(0).Value
			Set objCmd = nothing
			Next
			ipos = i
		END IF
		
		'���ǵ��
		 IF authId_H <> "" THEN
		  ipos = ipos +1
		 	Set objCmd = Server.CreateObject("ADODB.COMMAND")
			With objCmd
				.ActiveConnection = dbget
				.CommandType = adCmdText
				.CommandText = "{?= call db_partner.[dbo].[sp_Ten_eAppAuthLine_insert]( "&reportIdx&" ,"&ipos&",'"&trim(authId_H)&"','A','"&isSMS_H&"')}"
				.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
				.Execute, , adExecuteNoRecords
				End With
			    returnValue = objCmd(0).Value
			Set objCmd = nothing
		 END IF
		 
		'���������ڵ��
		 IF authId_L <> "" THEN 
		 	ipos = ipos +1
		 	Set objCmd = Server.CreateObject("ADODB.COMMAND")
			With objCmd
				.ActiveConnection = dbget
				.CommandType = adCmdText
				.CommandText = "{?= call db_partner.[dbo].[sp_Ten_eAppAuthLine_insert]( "&reportIdx&" ,"&ipos&",'"&trim(authId_L)&"','L','"&isSMS&"')}"
				.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
				.Execute, , adExecuteNoRecords
				End With
			    returnValue = objCmd(0).Value
			Set objCmd = nothing
		 END IF 

		'���������ڵ��
		 IF authId_F <> "" THEN
		 	ipos = ipos + 1
		 	Set objCmd = Server.CreateObject("ADODB.COMMAND")
			With objCmd
				.ActiveConnection = dbget
				.CommandType = adCmdText
				.CommandText = "{?= call db_partner.[dbo].[sp_Ten_eAppAuthLine_insert]( "&reportIdx&" ,"&ipos&",'"&trim(authId_F)&"','F','"&isSMS&"')}"
				.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
				.Execute, , adExecuteNoRecords
				End With
			    returnValue = objCmd(0).Value
			Set objCmd = nothing
		 END IF 
		 
	 IF (isSMS_H <> "1") and isSMS = "1" and reportState = 1 and (authid <> "" or authid_L <> "" ) THEN
		if authid <> "" then
    		reqhp = fnGetMemberHp(trim(arrAuthId(0)))
			reqEmail = fnGetMemberEmail(trim(arrAuthId(0)))
    	else
    		reqhp = fnGetMemberHp(authId_L)
			reqEmail = fnGetMemberEmail(authId_L)
    	end if	
		smstext = "["&reportName&"] ������ ["&susername&"]�����κ��� �����û�Ǿ����ϴ�." 
		smstext = smstext & vbCrLf & vbCrLf & "> �ٷΰ���"
		smsurl = getSCMSSLURL & "/admin/approval/eapp/confirmeapp.asp?iridx=" & reportIdx

		if reqEmail<>"" then
			smstext = chrbyte(Trim(smstext),1000,"Y")
			Call SendRadioWebHookMessage(reqEmail,"admin","SCM ���ڰ���","�����û",smstext,smsurl)
		else
			Call SendMultiRowsSMS(reqhp,"",smstext,"") 
		end if
    END IF

	''2013/10/21 �߰�
	IF isSMS_H = "1" and reportState = 1 and authid_H <> "" THEN
		reqhp = fnGetMemberHp(authid_H)
		reqEmail = fnGetMemberEmail(authid_H)
		smstext = "["&reportName&"] ������ ["&susername&"]�����κ��� ���ǿ�û�Ǿ����ϴ�."
		smstext = smstext & vbCrLf & vbCrLf & "> �ٷΰ���"
		smsurl = getSCMSSLURL & "/admin/approval/eapp/confirmeapp.asp?iridx=" & reportIdx

		if reqEmail<>"" then
			smstext = chrbyte(Trim(smstext),1000,"Y")
			Call SendRadioWebHookMessage(reqEmail,"admin","SCM ���ڰ���","���ǿ�û",smstext,smsurl)
		else
			Call SendMultiRowsSMS(reqhp,"",smstext,"") 
		end if
	END IF
 %>
		<script language="javascript">
			<!--
				alert("��ϵǾ����ϴ�");
				top.location.href = "/admin/approval/eapp/popIndex.asp?iridx=<%=reportIdx%>&iRM=<%=iRectMenu%>";
			//-->
			</script>
<%
	ELSE
			Call Alert_return ("������ ó���� ������ �߻��Ͽ����ϴ�.1")
	END IF
	
	response.end
CASE "CU"
			Set objCmd = Server.CreateObject("ADODB.COMMAND")
				With objCmd
					.ActiveConnection = dbget
					.CommandType = adCmdText
					.CommandText = "{?= call db_partner.[dbo].[sp_Ten_eAppReport_ConfirmUpdate]( "&reportIdx&",'"&referId&"','"&adminID&"')}"
					.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
					.Execute, , adExecuteNoRecords
					End With
				    returnValue = objCmd(0).Value
			Set objCmd = nothing

			IF returnValue = 0 THEN
				Call Alert_return ("������ ó���� ������ �߻��Ͽ����ϴ�.1")
				
			response.end
			END IF

	 '�μ��� �ڱݱ��� ���
	 	Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_eappPartMoney_Delete]( "&reportIdx&" ,0)}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
		 	.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
		Set objCmd = nothing


		IF eapppartIdx <> "" THEN
		eapppartIdx = split(eapppartIdx,",")
		partMoney = split(partMoney,",")

		For i = 0 To UBound(eapppartIdx)
		Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_eappPartMoney_insert]( "&reportIdx&" ,0,'"&trim(eapppartIdx(i))&"','"&getNumeric(partMoney(i))&"')}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
		Set objCmd = nothing
		Next
		END IF
		 %>
		<script language="javascript">
			<!--
				alert("ó���Ǿ����ϴ�");
				top.location.href = "/admin/approval/eapp/popIndex.asp?iRM=<%=iRectMenu%>&iAS=<%=AuthState%>&iridx=<%=reportIdx%>";
			//-->
			</script>
<%

CASE "C" ''����  
	Dim rlm
		if isSMS = "" THEN isSMS = 0
			'�ݷ��� ��� : �ݷ��� ���� 5�� ����, �׿� ���缱 ������ó��
			'�ݷ��� �� ǰ���Ѱ�� : ���� �ݷ��� ���� 6���� ����
			if AuthState = 5 THEN authposition = 0 '�ݷ��� ��쿡�� ���� ����ó�� ���� 
		Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_eAppReport_Confirm]( "&reportidx&" , "&AuthState&","&reportState&",'"&adminId&"',"&authposition&",  "&isSMS&")}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
		Set objCmd = nothing

	IF 	returnValue = 1 THEN

		IF isSMS = "1"   THEN 
			reqhp = fnGetMemberHp(authid)
			reqEmail = fnGetMemberEmail(authid)
			smstext = "["&reportName&"] ������ ["&susername&"]�����κ��� "& "�����û�Ǿ����ϴ�."
			smstext = smstext & vbCrLf & vbCrLf & "> �ٷΰ���"
			smsurl = getSCMSSLURL & "/admin/approval/eapp/confirmeapp.asp?iridx=" & reportIdx

			if reqEmail<>"" then
				smstext = chrbyte(Trim(smstext),1000,"Y")
				Call SendRadioWebHookMessage(reqEmail,"admin","SCM ���ڰ���","������û",smstext,smsurl)
			else
				Call SendMultiRowsSMS(reqhp,"",smstext,"") 
			end if
		END IF
 
        ''2013/10/21 �߰�
		IF isSMS_H = "1" THEN
    		reqhp = fnGetMemberHp(authid_H)
			reqEmail = fnGetMemberEmail(authid_H)
    		smstext = "["&reportName&"] ������ ["&susername&"]�����κ��� ���ǿ�û�Ǿ����ϴ�."
			smstext = smstext & vbCrLf & vbCrLf & "> �ٷΰ���"
			smsurl = getSCMSSLURL & "/admin/approval/eapp/confirmeapp.asp?iridx=" & reportIdx

			if reqEmail<>"" then
				smstext = chrbyte(Trim(smstext),1000,"Y")
				Call SendRadioWebHookMessage(reqEmail,"admin","SCM ���ڰ���","���ǿ�û",smstext,smsurl)
			else
				Call SendMultiRowsSMS(reqhp,"",smstext,"") 
			end if
    	END IF

		IF   AuthState <> 1 THEN
			Dim strStatus
			IF AuthState = "3" THEN
				strStatus = "����"
			ELSEIF	AuthState = "5" THEN
				strStatus = "�ݷ�"
			END IF
			reqhp = fnGetMemberHp(sUserid)
			reqEmail = fnGetMemberEmail(sUserid)
			smstext = "["&reportName&"] ������ ["&sauthname&"]�����κ��� ����"&strStatus&"�Ǿ����ϴ�."
			smstext = smstext & vbCrLf & vbCrLf & "> �ٷΰ���"
			smsurl = getSCMSSLURL & "/admin/approval/eapp/modeapp.asp?iridx=" & reportIdx

			if reqEmail<>"" then
				smstext = chrbyte(Trim(smstext),1000,"Y")
				Call SendRadioWebHookMessage(reqEmail,"admin","SCM ���ڰ���","����"&strStatus,smstext,smsurl)
			else
				Call SendMultiRowsSMS(reqhp,"",smstext,"") 
			end if
		ELSEIF reportState = "7" and AuthState = 1 THEN
			reqhp = fnGetMemberHp(sUserid)
			reqEmail = fnGetMemberEmail(sUserid)
			smstext = "["&reportName&"] ������ ������εǾ����ϴ�."
			smstext = smstext & vbCrLf & vbCrLf & "> �ٷΰ���"
			smsurl = getSCMSSLURL & "/admin/approval/eapp/modeapp.asp?iridx=" & reportIdx

			if reqEmail<>"" then
				smstext = chrbyte(Trim(smstext),1000,"Y")
				Call SendRadioWebHookMessage(reqEmail,"admin","SCM ���ڰ���","�������",smstext,smsurl)
			else
				Call SendMultiRowsSMS(reqhp,"",smstext,"") 
			end if
		END IF

		IF sscmsubmitlink <> "" and  AuthState <> 3 and (( authId <> "" and AuthState<>1) or authId = "") THEN	'���°� �����̰ų� ���������� �ƴѰ�쿡�� ���� scm ��ũ ó���� ���� �ʴ´�.
			%>
			<form name="frmLink" method="post" action="<%=sscmsubmitlink&scmLink%>" target="_top">
			<input type="hidden" name="ias" value="<%=AuthState%>">
			<input type="hidden" name="hidRU" value="/admin/approval/eapp/popIndex.asp?iridx=<%=reportIdx%>&iRM=<%=iRectMenu%>&iAS=<%=AuthState%>&abc">
			</form>
			<script language="javascript">
			<!--
			document.frmLink.submit();
			//-->
			</script>
			<%
			
			response.end
		END IF
		 %>
		<script language="javascript">
			<!--
				alert("ó���Ǿ����ϴ�");
				top.location.href = "/admin/approval/eapp/popIndex.asp?iRM=<%=iRectMenu%>&iAS=<%=AuthState%>&123";
			//-->
			</script>
<%
	ELSE
			Call Alert_return ("������ ó���� ������ �߻��Ͽ����ϴ�.1")
	END IF
	
	response.end
CASE "D"

	Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_eAppReport_Delete]( "&reportidx&",'"&adminId&"')}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
		Set objCmd = nothing
	IF 	returnValue = 1 THEN
	 %>
		<script language="javascript">
			<!--
				alert("�����Ǿ����ϴ�");
				top.location.href = "/admin/approval/eapp/popIndex.asp?iRM=<%=iRectMenu%>";
			//-->
			</script>
<%
	ELSE
			Call Alert_return ("������ ó���� ������ �߻��Ͽ����ϴ�.1")
	END IF
	
	response.end
	CASE "DA" '--�����ڻ��� 
	Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_eAppReport_DeleteAdmin]( "&reportidx&",'"&adminId&"')}" 
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
		Set objCmd = nothing
		
		
	IF 	returnValue = 1 THEN
	 	 %>
		<script language="javascript">
			<!--
				alert("�����Ǿ����ϴ�");
				self.close();
				opener.location.reload();
			//-->
			</script>
<%
	ELSE
			Call Alert_return ("������ ó���� ������ �߻��Ͽ����ϴ�.1")
	END IF
	
	response.end

	 
CASE "A"
		Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_eAppReport_AddUpdate]("&reportidx&", '"&reportPRice&"'"&_
						+" ,'"&adminID&"','"&authId&"',"&authposition&", "&isSMS&" )}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
	Set objCmd = nothing
	IF 	returnValue > 0 THEN
		IF isSMS = "1" and reportState = 1 THEN
			reqhp = fnGetMemberHp(authid)
			reqEmail = fnGetMemberEmail(authid)
			smstext = "["&reportName&"] ������ ["&susername&"]�����κ��� �߰�ǰ�� �����û�Ǿ����ϴ�."
			smstext = smstext & vbCrLf & vbCrLf & "> �ٷΰ���"
			smsurl = getSCMSSLURL & "/admin/approval/eapp/confirmeapp.asp?iridx=" & reportIdx

			if reqEmail<>"" then
				smstext = chrbyte(Trim(smstext),1000,"Y")
				Call SendRadioWebHookMessage(reqEmail,"admin","SCM ���ڰ���","�߰�ǰ�ǿ�û",smstext,smsurl)
			else
				Call SendMultiRowsSMS(reqhp,"",smstext,"") 
			end if
	  	END IF

		 call Alert_close("ǰ�Ǳݾ� �߰�ǰ�ǰ� �����ϵǾ����ϴ�.")
	ELSE
		Call Alert_return ("������ ó���� ������ �߻��Ͽ����ϴ�.1")
	END IF
	
	response.end
CASE ELSE
	Call Alert_return ("������ ó���� ������ �߻��Ͽ����ϴ�.0")
END SELECT
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
 