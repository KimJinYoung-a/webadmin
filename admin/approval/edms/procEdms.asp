<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
Dim objCmd, returnValue, sMode
Dim menupos
Dim iedmsidx, blnusing,page
Dim icateidx1,icateidx2,sserialnum,sedmsname,sedmscode,iviewno,sedmsfile,blnApproval,blnscmapproval,ilastapprovalid,sscmlink,sadminid , blnPay,edmsform, rdoH
Dim sscmsubmitlink
Dim isAgreeNeed, isAgreeNeedTarget
sMode		= requestCheckvar(Request("hidM"),1)

iedmsidx	= requestCheckvar(Request("ieidx"),10)
icateidx1	= requestCheckvar(Request("selC1"),10)
icateidx2	= requestCheckvar(Request("hidC2"),10)
sserialnum	= requestCheckvar(Request("hidSN"),3)
sedmsname	= requestCheckvar(Request("sDN"),64)
sedmscode	= requestCheckvar(Request("sDC"),10)
iviewno		= requestCheckvar(Request("iVN"),10)
sedmsfile	= requestCheckvar(Request("hidAF"),128)
blnApproval	= requestCheckvar(Request("rdoA"),1)

blnscmapproval	= requestCheckvar(Request("rdoEA"),1)
ilastapprovalid	= requestCheckvar(Request("selJN"),4)
sscmlink	= requestCheckvar(Request("sSL"),128)
sscmsubmitlink	= requestCheckvar(Request("sSSL"),128)
sadminid	=   session("ssBctId")
blnPay		= requestCheckvar(Request("rdoP"),1)
blnUsing	= requestCheckvar(Request("rdoU"),1)
rdoH        = requestCheckvar(Request("rdoH"),1)
menupos		= requestCheckvar(Request("menupos"),10)
page 		= requestCheckvar(Request("page"),10)
edmsform	= ReplaceRequestSpecialChar(Request("editor"))
isAgreeNeed = requestCheckvar(Request("isAgreeNeed"),1)
isAgreeNeedTarget = requestCheckvar(Request("sId"),32)

If isAgreeNeed <> "Y" Then
	isAgreeNeedTarget = ""
End If

if (rdoH="") then rdoH=0
IF iviewno = "" THEN iviewno = 0

if (checkNotValidHTML(sedmsname) = true) Then
	response.write "<script>alert('���ڰ��� ǰ�Ǽ� �̸����� HTML�� ����Ͻ� �� �����ϴ�.');history.back();</script>"
	dbget.Close
	response.End
End If

if (checkNotValidHTML(edmsform) = true) Then
	response.write "<script>alert('���ڰ��� ������ Script �Ǵ� Action�� ����Ͻ� �� �����ϴ�.');history.back();</script>"
	dbget.Close
	response.End
End If
	
SELECT CASE sMode
Case "I"

	Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_edms_insert]( "&icateidx1&","&icateidx2&",'"&sserialnum&"' ,'"&sedmsname&"', '"&sedmscode&"' ,"&iviewno&_
					",'"&sedmsfile&"',"&blnApproval&","&blnscmapproval&",'"&ilastapprovalid&"','"&sscmlink&"','"&sscmsubmitlink&"','"&sadminid&"',"&blnPay&","&rdoH&",'"&isAgreeNeed&"','"&isAgreeNeedTarget&"')}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
	Set objCmd = nothing

	IF returnValue = "1" THEN
		call Alert_closenmove("��ϵǾ����ϴ�.","index.asp?selC1="&icateidx1&"&selC2="&icateidx2&"&menupos="&menupos&"&page="&page)
	ELSEIF 	returnValue = "2" THEN
			Call Alert_move ("�Է��Ͻ� �����ڵ尪�� ������ ������Դϴ�.�ٽ� �Է����ּ���","popEdmsConts.asp?menupos="&menupos&"&page="&page)
	ELSE
			Call Alert_return ("������ ó���� ������ �߻��Ͽ����ϴ�.1")
	END IF
	response.end
Case "U"
	Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_edms_update]("&iedmsidx&","&icateidx1&","&icateidx2&",'"&sserialnum&"' ,'"&sedmsname&"', '"&sedmscode&"' ,"&iviewno&_
					",'"&sedmsfile&"',"&blnApproval&","&blnscmapproval&",'"&ilastapprovalid&"','"&sscmlink&"','"&sscmsubmitlink&"',"&blnUsing&",'"&sadminid&"',"&blnPay&","&rdoH&",'"&isAgreeNeed&"','"&isAgreeNeedTarget&"')}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
	Set objCmd = nothing

	IF returnValue = "1" THEN
		IF blnPay = 0 THEN '������û�� �������ΰ� False �̸� ������ �������񳻿� Nulló��
			Set objCmd = Server.CreateObject("ADODB.COMMAND")
				With objCmd
					.ActiveConnection = dbget
					.CommandType = adCmdText
					.CommandText = "{?= call db_partner.[dbo].[sp_Ten_eAappAccount_UpdateEdms]("&iedmsidx&")}"
					.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
					.Execute, , adExecuteNoRecords
					End With
				    returnValue = objCmd(0).Value
			Set objCmd = nothing
		END IF
		Call Alert_closenmove ("�����Ǿ����ϴ�.","index.asp?selC1="&icateidx1&"&selC2="&icateidx2&"&menupos="&menupos&"&page="&page)
	ELSEIF 	returnValue = "2" THEN
		Call Alert_move ("�Է��Ͻ� ī�װ��ڵ尪�� ������ ������Դϴ�.�ٽ� �Է����ּ���","popEdmsConts.asp?ieidx="&iedmsidx&"&menupos="&menupos&"&page="&page)
	ELSE
		Call Alert_return ("������ ó���� ������ �߻��Ͽ����ϴ�.1")
	END IF
	response.end


Case "A"
	Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_edms_updateFile]( "&iedmsidx&",'"&sedmsfile&"','"&sadminid&"')}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
	Set objCmd = nothing

	IF returnValue = "1" THEN
		IF sedmsfile <> "" THEN
			call Alert_closenmove("��ϵǾ����ϴ�.","index.asp?selC1="&icateidx1&"&selC2="&icateidx2&"&menupos="&menupos&"&page="&page)
		ELSE
			Call Alert_move ("�����Ǿ����ϴ�.","index.asp?selC1="&icateidx1&"&selC2="&icateidx2&"&menupos="&menupos&"&page="&page)
		END IF
	ELSE
			Call Alert_return ("������ ó���� ������ �߻��Ͽ����ϴ�.1")
	END IF
	response.end

CASE "F"
			Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_edms_updateForm]( "&iedmsidx&",'"&edmsform&"','"&sadminid&"')}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
	Set objCmd = nothing

	IF returnValue = "1" THEN
			call Alert_closenmove("��ϵǾ����ϴ�.","index.asp?selC1="&icateidx1&"&selC2="&icateidx2&"&menupos="&menupos&"&page="&page)
	ELSE
			Call Alert_return ("������ ó���� ������ �߻��Ͽ����ϴ�.1")
	END IF
	response.end
CASE ELSE
	Call Alert_return ("������ ó���� ������ �߻��Ͽ����ϴ�.0")
END SELECT
%>
