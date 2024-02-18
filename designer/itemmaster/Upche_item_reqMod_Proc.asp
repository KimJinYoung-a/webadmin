<%@ language=vbscript %>
<% option explicit %>
<%
'########################################################### 
' Description : 업체배송 상품명 수정요청 처리
' History : 2014.03.19 정윤정 등록 
'###########################################################
%>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
dim itemidarr, olditemname, itemname, etcstr
dim i,itemid
dim objCmd, returnValue, errID,ItemCount,chkReturnCount
Dim mode, menupos
Dim strResultMsg
Dim oldsellcash, oldbuycash, sellcash, buycash
Dim makerid, edtype
 
mode=requestCheckvar(request("hidM"),1)
menupos=requestCheckvar(request("menupos"),10)
itemidarr = ReplaceRequestSpecialChar(request("itemidarr")) 
olditemname= ReplaceRequestSpecialChar(request("olditemname"))
itemname= ReplaceRequestSpecialChar(request("itemname"))
oldsellcash= ReplaceRequestSpecialChar(request("oldsellcash"))
oldbuycash= ReplaceRequestSpecialChar(request("oldbuycash"))
sellcash= ReplaceRequestSpecialChar(request("sellcash"))
buycash= ReplaceRequestSpecialChar(request("buycash"))
etcstr=  requestCheckvar(request("etcStr"),64)
ItemCount=requestCheckvar(request("itemcount"),10)
makerid = session("ssBctID")
 
SELECT  CASE mode
CASE "N" '--상품명 수정요청
	itemid = split(itemidarr,"|,|")
	olditemname = split(olditemname,"|,|")
	itemname = split(itemname,"|,|")
	chkReturnCount = 0
	For i=0 To UBound(itemid)
		itemid(i) = Left(trim(itemid(i)),16)
		olditemname(i) = Left(trim(olditemname(i)),64)
		itemname(i) = Left(trim(itemname(i)),64)
		 
		if itemname(i) = "" then 
			Call Alert_return ("상품명이 등록되지 않았습니다.")
	 		response.end
		end if
		
		Set objCmd = Server.CreateObject("ADODB.COMMAND")
			With objCmd
				.ActiveConnection = dbget
				.CommandType = adCmdText
				.CommandText = "{?= call db_temp.[dbo].[sp_Ten_upche_itemedit_itemanmeInsert]("&trim(itemid(i))&", '"&trim(olditemname(i))&"' ,'"&trim(itemname(i))&"','"&etcstr&"','"&makerid&"')}"
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
	strResultMsg = strResultMsg & "선택하신 [" & ItemCount &"]건의 상품 중 "& "["&chkReturnCount&"]건의 상품이 수정요청되었습니다.\n 상품수정 요청결과는 업체상품관리>>상품수정처리결과에서 확인하세요"
	Call Alert_move(strResultMsg, "/designer/itemmaster/upche_item_reqMod_itemname.asp?menupos="&menupos)
CASE "P" '--가격 수정요청
	itemid = split(itemidarr,",")
	oldsellcash = split(oldsellcash,",")
	sellcash = split(sellcash,",")
	oldbuycash = split(oldbuycash,",")
	buycash = split(buycash,",")
	chkReturnCount = 0
	For i=0 To UBound(itemid)
		itemid(i) = Left(trim(itemid(i)),16) 
		
	 '등록조건 확인
	 if trim(sellcash(i)) = ""  then
	 	 Call Alert_return ("판매가가 등록되지 않았습니다.")
	 response.end
	 end if
	 
	  if (Clng(trim(sellcash(i))) <=100 or   Clng(trim(buycash(i))) <=100 ) then
	 	 Call Alert_return ("판매가나 공급가는 100원 이상만 가능합니다.")
	 response.end
	 end if
	 
	  if Clng(trim(sellcash(i))) < Clng(trim(buycash(i))) then
	 	 Call Alert_return ("판매가는 공급가보다 큰 가격만 가능합니다." & sellcash(i)&"-" &buycash(i))
	 response.end
	 end if
	
	    ''2015/03/10 추가
		if (UBOUND(itemid)<>UBOUND(oldsellcash)) or (UBOUND(itemid)<>UBOUND(sellcash)) or (UBOUND(itemid)<>UBOUND(oldbuycash)) or (UBOUND(itemid)<>UBOUND(buycash)) then
		    Call Alert_return ("전송 파라메터 오류-관리자문의 요망")
	        response.end
		end if
		
		Set objCmd = Server.CreateObject("ADODB.COMMAND")
			With objCmd
				.ActiveConnection = dbget
				.CommandType = adCmdText
				.CommandText = "{?= call db_temp.[dbo].[sp_Ten_upche_itemedit_sellcashInsert]("&trim(itemid(i))&", '"&trim(oldsellcash(i))&"' ,'"&trim(oldbuycash(i))&"', '"&trim(sellcash(i))&"' ,'"&trim(buycash(i))&"','"&etcstr&"','"&makerid&"')}"
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
	strResultMsg = strResultMsg & "선택하신 [" & ItemCount &"]건의 상품 중 "& "["&chkReturnCount&"]건의 상품이 수정요청되었습니다.\n 상품수정 요청결과는 업체상품관리>>상품수정처리결과에서 확인하세요"
	Call Alert_move(strResultMsg, "/designer/itemmaster/upche_item_reqMod_sellprice.asp?menupos="&menupos)
CASE "C" '수정요청 취소 
itemidarr = left(itemidarr,16)
olditemname = left(olditemname,64)
 
		Set objCmd = Server.CreateObject("ADODB.COMMAND")
			With objCmd
				.ActiveConnection = dbget
				.CommandType = adCmdText
				.CommandText = "{?= call db_temp.[dbo].[sp_Ten_upche_itemedit_cancel]("&itemidarr&",'"&makerid&"','"&olditemname&"','"&oldsellcash&"')}"
				.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
				.Execute, , adExecuteNoRecords
				End With
			    returnValue = objCmd(0).Value
		Set objCmd = nothing
		
		IF returnValue <> 1 THEN
				Call Alert_return ("데이터 처리에 문제가 발생하였습니다.-error: case 'c' returnValue")
 		ELSE
 				Call Alert_move("상품수정요청이 취소되었습니다.", "/designer/itemmaster/upche_item_reqMod_result.asp?menupos="&menupos)
		END IF
CASE ELSE
		Call Alert_return ("데이터 처리에 문제가 발생하였습니다.-error: case else")
END SELECT	
%>