<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
dim sMode,itemid,itemname,sellcash,buycash,adminid ,rejectstr
dim itemidarr,itemnamearr,sellcasharr,buycasharr,rejectstrarr
dim sitemname, ssellcash, sbuycash
dim objCmd, returnValue
dim strResultMsg, itemcount,  chkReturnCount
Dim errID, i
dim menupos
dim rectSort,rectmakerid, rectitemid, rectitemname, rectstartdate, rectenddate, rectreqtype, rectdispCate
dim editidxarr,editidx,edittypearr, edittype

'검색 파라미터-------------------------------------------
rectmakerid     = requestCheckvar(request("rmakerid"),32)
rectitemid  = RequestCheckVar(request("ritemid"),500) 
rectitemname = RequestCheckVar(request("ritemname"),64) 
rectdispCate = requestCheckvar(request("rdispCate"),16) 
rectstartdate  = RequestCheckVar(request("rSD"),10) 
rectenddate  = RequestCheckVar(request("rED"),10) 
rectreqtype = RequestCheckVar(request("rRT"),1)
rectSort= RequestCheckVar(request("rS"),2)
menupos=requestCheckvar(request("menupos"),10)
rejectstr	= RequestCheckVar(Request("rejectstr"),64)


sMode = requestCheckvar(request("hidM"),1) 

itemidarr = ReplaceRequestSpecialChar(Request("itemid"))
itemnamearr = ReplaceRequestSpecialChar(Request("itemname"))
sellcasharr= ReplaceRequestSpecialChar(Request("sellcash"))
buycasharr= ReplaceRequestSpecialChar(Request("buycash"))

editidxarr= ReplaceRequestSpecialChar(Request("editidx"))
edittypearr =ReplaceRequestSpecialChar(Request("edittype"))

itemcount=requestCheckvar(request("itemcount"),10)
adminid= session("ssBctId") 

SELECT CASE sMode 
Case "A" '//승인
	itemid 	= split(itemidarr,",") 
	editidx = split(editidxarr,",") 
	itemname= split(itemnamearr,",")  
	sellcash= split(sellcasharr,",") 
	buycash = split(buycasharr,",") 
	edittype = split(edittypearr,",")
chkReturnCount = 0
 
	For i=0 To UBound(itemid)
		itemid(i) = Left(trim(itemid(i)),16)  
		editidx(i) = left(trim(editidx(i)),16)
		edittype(i) = left(trim(edittype(i)),1)
		 
 		
		Set objCmd = Server.CreateObject("ADODB.COMMAND")
			With objCmd
				.ActiveConnection = dbget
				.CommandType = adCmdText
				.CommandText = "{?= call db_item.[dbo].[sp_Ten_item_UpcheReq_update]("&itemid(i)&", "&editidx(i)&", '"&adminid&"' ,'"&edittype(i)&"')}"
				.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
				.Execute, , adExecuteNoRecords
				End With
			    returnValue = objCmd(0).Value
		Set objCmd = nothing
		 
		IF returnValue <> "1" THEN 
			if errID = "" THEN
				errID = itemid(i)
			else
			errID = errID+","+itemid(i)
			end if
		ELSE
			chkReturnCount = chkReturnCount  + 1
		END IF	
	Next
 
 strResultMsg = "" 
 	IF errID <> "" THEN
 		strResultMsg = strResultMsg & "상품코드 ["&errID &"] 는 처리 실패했습니다.\n"
	END IF
	strResultMsg = strResultMsg & "선택하신 [" & ItemCount &"]건의 상품 중 "& "["&chkReturnCount&"]건의 상품수정요청이 승인되었습니다."
	Call Alert_move(strResultMsg, "/admin/itemmaster/item_modReq_confirm.asp?menupos="&menupos&"&sS="&rectSort&"&makerid="&rectmakerid&"&itemname="&rectitemname&"&itemid="&rectitemid&"&disp="&rectdispCate&"&dSD="&rectstartdate&"&dED="&rectenddate&"&selRT="&rectreqtype)
 
response.end
Case "D" '//반려
	itemid = split(itemidarr,",") 
	editidx = split(editidxarr,",") 
 
chkReturnCount = 0
 
	For i=0 To UBound(itemid)
		itemid(i) = Left(trim(itemid(i)),16)  
		editidx(i) = left(trim(editidx(i)),16)
 	 
   
	Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_item.[dbo].[sp_Ten_item_UpcheReq_Return]("&itemid(i)&", "&editidx(i)&",'"&rejectstr&"', '"&adminid&"' )}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
	Set objCmd = nothing
	
		IF returnValue <> "1" THEN 
			if errID = "" THEN
				errID = itemid(i)
			else
			errID = errID+","+itemid(i)
			end if
		ELSE
			chkReturnCount = chkReturnCount  + 1
		END IF	
	Next
 
 strResultMsg = "" 
 	IF errID <> "" THEN
 		strResultMsg = strResultMsg & "상품코드 ["&errID &"] 는 처리 실패했습니다.\n"
	END IF
	strResultMsg = strResultMsg & "선택하신 [" & ItemCount &"]건의 상품 중 "& "["&chkReturnCount&"]건의 상품수정요청이 반려되었습니다."
	Call Alert_move(strResultMsg, "/admin/itemmaster/item_modReq_confirm.asp?menupos="&menupos&"&sS="&rectSort&"&makerid="&rectmakerid&"&itemname="&rectitemname&"&itemid="&rectitemid&"&disp="&rectdispCate&"&dSD="&rectstartdate&"&dED="&rectenddate&"&selRT="&rectreqtype)
	 
response.end

CASE "C" '//대기승인으로 변경
	Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_item.[dbo].[sp_Ten_item_UpcheReq_Change]("&itemidarr&","&editidxarr&",'"&adminid&"' )}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
	Set objCmd = nothing
	
	IF	returnValue = 0 THEN
			Call Alert_return ("데이터 처리에 문제가 발생하였습니다.1")
	ELSE	
			Call Alert_move("승인대기 상태로 변경되었습니다.", "/admin/itemmaster/item_modReq_confirm.asp?menupos="&menupos&"&sS="&rectSort&"&makerid="&rectmakerid&"&itemname="&rectitemname&"&itemid="&rectitemid&"&disp="&rectdispCate&"&dSD="&rectstartdate&"&dED="&rectenddate&"&selRT="&rectreqtype)
	END IF
response.end

CASE ELSE
	Call Alert_return ("데이터 처리에 문제가 발생하였습니다.0")
END SELECT
 %>
<!-- #include virtual="/lib/db/dbclose.asp" -->