<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
dim sMode,itemid,logmsgcd,logmsg, adminid, currstate, sellCS
dim itemidarr, makerid, itemname, dispCate, cdl, cdm, cds, onlyNotSet, selCtr
dim objCmd, returnValue
dim sReturnURL

sMode = requestCheckvar(request("hidM"),1)
itemid = requestCheckvar(Request("itemid"),16)
logmsg = ReplaceRequestSpecialChar(request("sMsg"))
logmsgcd = requestCheckvar(request("sMsgcd"),50)
currstate = requestCheckvar(request("sCS"),1)	'�����û ����
sReturnURL = ReplaceRequestSpecialChar(request("sRU"))
itemidarr = ReplaceRequestSpecialChar(request("itemidarr"))

'�ϰ� ó����
makerid	= requestCheckvar(Request("makerid"),32)
itemname	= requestCheckvar(Request("itemname"),64)
dispCate = requestCheckvar(request("disp"),16)
cdl = requestCheckvar(request("cdl"),3)
cdm = requestCheckvar(request("cdm"),3)
cds = requestCheckvar(request("cds"),3)
sellCS = requestCheckvar(request("sellCS"),1)	'�˻� ����
onlyNotSet = requestCheckvar(request("onlyNotSet"),1)
selCtr = requestCheckvar(request("selCtr"),1)
if onlyNotSet="Y" then dispCate="n"

adminid = session("ssBctId") 

SELECT CASE sMode
Case "U" '//����ó��
	Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_temp.[dbo].[sp_Ten_wait_item_proc]("&itemid&", '"&currstate&"' ,'"&logmsgcd&"','"&logmsg&"', '"&adminid&"' )}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
	Set objCmd = nothing
	
	IF	returnValue = 0 THEN
			Call Alert_return ("������ ó���� ������ �߻��Ͽ����ϴ�.[E10]")
	ELSE	
			Call Alert_move("ó���Ϸ�Ǿ����ϴ�.", sReturnURL)
	END IF
	response.end
Case "M" '//����ó��
	Dim arrItem, i,   chkReturn0, chkReturnCount, ItemCount
		arrItem = split(itemidarr,",") '��ǰ �迭�� ������
	
	chkReturn0 = ""
	chkReturnCount = 0
	ItemCount =  UBound(arrItem)+1
		For i = 0 To UBound(arrItem)
			Set objCmd = Server.CreateObject("ADODB.COMMAND")
			With objCmd
				.ActiveConnection = dbget
				.CommandType = adCmdText
				.CommandText = "{?= call db_temp.[dbo].[sp_Ten_wait_item_proc]("&trim(arrItem(i))&", '"&currstate&"' ,'"&logmsgcd&"','"&logmsg&"', '"&adminid&"' )}"
				.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
				.Execute, , adExecuteNoRecords
				End With
				returnValue = objCmd(0).Value
		Set objCmd = nothing
			IF returnValue = 0 THEN	
				IF chkReturn0 = "" THEN
					chkReturn0 =   trim(arrItem(i))
				ELSE
					chkReturn0 = chkReturn0 +","+ trim(arrItem(i))
				END IF
			ELSE
				chkReturnCount = chkReturnCount + 1	
			END IF
		Next
	
	Dim strResultMsg
	strResultMsg = "" 
	IF chkReturn0 <> "" THEN
		strResultMsg = strResultMsg & "��ǰ�ӽ��ڵ� ["&chkReturn0 &"] �� ó�� �����߽��ϴ�.\n"
	END IF
		strResultMsg = strResultMsg & "�����Ͻ� [" & ItemCount &"]���� ��ǰ �� "& "["&chkReturnCount&"]���� ���������� ó���Ϸ�Ǿ����ϴ�."
		Call Alert_move(strResultMsg, sReturnURL)
	
	response.end
Case "B"	'//�ϰ�ó��
	if makerid="" or dispCate="" or sellCS="" then
		Call Alert_return ("�ϰ�ó���� �ϱ����ؼ��� �˻������� �ʿ��մϴ�.")
		response.end
	end if

	if itemidarr<>"" then
		itemidarr = replace(itemidarr,chr(10),",")
		itemidarr = replace(itemidarr,chr(13),",")
		itemidarr = replace(itemidarr,",,",",")
	end if

	Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_temp.[dbo].[usp_Ten_wait_item_batch_reject]('"&dispCate&"','"&makerid&"','"&itemname&"','"&currstate&"','"&sellCS&"','"&itemidarr&"','"&cdl&"','"&cdm&"','"&cds&"','"&selCtr&"','"&logmsgcd&"','"&logmsg&"','"&adminid&"')}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
	Set objCmd = nothing
	
	IF	returnValue = 0 THEN
			Call Alert_return ("������ ó���� ������ �߻��Ͽ����ϴ�.[E10]")
	ELSE	
			Call Alert_move("ó���Ϸ�Ǿ����ϴ�.", sReturnURL)
	END IF

CASE ELSE
	Call Alert_return ("������ ó���� ������ �߻��Ͽ����ϴ�.[E00]")
END SELECT
 %>
<!-- #include virtual="/lib/db/dbclose.asp" -->