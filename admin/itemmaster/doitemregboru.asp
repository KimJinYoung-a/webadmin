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
currstate = requestCheckvar(request("sCS"),1)	'변경요청 상태
sReturnURL = ReplaceRequestSpecialChar(request("sRU"))
itemidarr = ReplaceRequestSpecialChar(request("itemidarr"))

'일괄 처리용
makerid	= requestCheckvar(Request("makerid"),32)
itemname	= requestCheckvar(Request("itemname"),64)
dispCate = requestCheckvar(request("disp"),16)
cdl = requestCheckvar(request("cdl"),3)
cdm = requestCheckvar(request("cdm"),3)
cds = requestCheckvar(request("cds"),3)
sellCS = requestCheckvar(request("sellCS"),1)	'검색 상태
onlyNotSet = requestCheckvar(request("onlyNotSet"),1)
selCtr = requestCheckvar(request("selCtr"),1)
if onlyNotSet="Y" then dispCate="n"

adminid = session("ssBctId") 

SELECT CASE sMode
Case "U" '//단일처리
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
			Call Alert_return ("데이터 처리에 문제가 발생하였습니다.[E10]")
	ELSE	
			Call Alert_move("처리완료되었습니다.", sReturnURL)
	END IF
	response.end
Case "M" '//다중처리
	Dim arrItem, i,   chkReturn0, chkReturnCount, ItemCount
		arrItem = split(itemidarr,",") '상품 배열로 나누기
	
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
		strResultMsg = strResultMsg & "상품임시코드 ["&chkReturn0 &"] 는 처리 실패했습니다.\n"
	END IF
		strResultMsg = strResultMsg & "선택하신 [" & ItemCount &"]건의 상품 중 "& "["&chkReturnCount&"]건이 성공적으로 처리완료되었습니다."
		Call Alert_move(strResultMsg, sReturnURL)
	
	response.end
Case "B"	'//일괄처리
	if makerid="" or dispCate="" or sellCS="" then
		Call Alert_return ("일괄처리를 하기위해서는 검색조건이 필요합니다.")
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
			Call Alert_return ("데이터 처리에 문제가 발생하였습니다.[E10]")
	ELSE	
			Call Alert_move("처리완료되었습니다.", sReturnURL)
	END IF

CASE ELSE
	Call Alert_return ("데이터 처리에 문제가 발생하였습니다.[E00]")
END SELECT
 %>
<!-- #include virtual="/lib/db/dbclose.asp" -->