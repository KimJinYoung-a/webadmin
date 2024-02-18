<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<%
Session.CodePage = 65001
Response.Charset = "UTF-8"
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/apps/academy/lib/head.asp" -->
<%
Dim pageTitle
pageTitle="더핑거스 - 가격수정요청 처리"

dim olditemname, itemname, etcstr
dim i,itemid
dim objCmd, returnValue
Dim mode
Dim oldsellcash, oldbuycash, sellcash, buycash
Dim makerid, iidx

mode=requestCheckvar(request("hidM"),1)
itemid = requestCheckvar(request("itemid"),12) 
olditemname= requestCheckvar(request("olditemname"),64)
itemname= requestCheckvar(request("itemname"),64)
oldsellcash= requestCheckvar(request("oldsellcash"),24)
oldbuycash= requestCheckvar(request("oldbuycash"),24)
sellcash= requestCheckvar(request("sellcash"),24)
buycash= requestCheckvar(request("buycash"),24)
etcstr=  requestCheckvar(request("etcStr"),64)
makerid = request.cookies("partner")("userid")
iidx = requestCheckvar(request("idx"),12)

if makerid="" then
 	Call Alert_return ("로그인이 필요합니다.")
 	response.end
end if

SELECT  CASE mode
CASE "N" '--상품명 수정요청

CASE "P" '--가격 수정요청

	 '등록조건 확인
	 if trim(sellcash) = ""  then
	 	 Call Alert_return ("판매가가 등록되지 않았습니다.")
	 	response.end
	 end if
	 
	  if (Clng(trim(sellcash)) <=100 or   Clng(trim(buycash)) <=100 ) then
	 	 Call Alert_return ("판매가나 공급가는 100원 이상만 가능합니다.")
	 response.end
	 end if
	 
	  if Clng(trim(sellcash)) < Clng(trim(buycash)) then
	 	 Call Alert_return ("판매가는 공급가보다 큰 가격만 가능합니다." & sellcash&"-" &buycash)
	 response.end
	 end if
	
	Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbACADEMYget
			.CommandType = adCmdText
			.CommandText = "{?= call db_academy.[dbo].[sp_Fingers_upche_itemedit_sellcashInsert]("&trim(itemid)&", '"&trim(oldsellcash)&"' ,'"&trim(oldbuycash)&"', '"&trim(sellcash)&"' ,'"&trim(buycash)&"','"&etcstr&"','"&makerid&"')}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
	Set objCmd = nothing

	IF returnValue <> 1 THEN
		Call Alert_return("상품가격 수정 처리가 실패했습니다.")
	else
		Response.Write "<script>alert('상품이 수정요청되었습니다.'); fnAPPopenerJsCallClose('fnItemPriceEditEnd(1)');</script>"
	END IF	


CASE "C" '수정요청 취소 

		Set objCmd = Server.CreateObject("ADODB.COMMAND")
			With objCmd
				.ActiveConnection = dbACADEMYget
				.CommandType = adCmdText
				.CommandText = "{?= call db_academy.[dbo].[sp_Fingers_upche_itemedit_cancel]("&itemid&",'"&makerid&"','"&olditemname&"','"&oldsellcash&"')}"
				.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
				.Execute, , adExecuteNoRecords
				End With
			    returnValue = objCmd(0).Value
		Set objCmd = nothing
		
		IF returnValue <> 1 THEN
			Call Alert_return ("데이터 처리에 문제가 발생하였습니다.-error: case 'c' returnValue")
 		ELSE
 			Response.Write "<script>alert('상품수정요청이 취소되었습니다.'); parent.document.location.reload();</script>"
		END IF

CASE "R" '반려 정보 확인

		Set objCmd = Server.CreateObject("ADODB.COMMAND")
			With objCmd
				.ActiveConnection = dbACADEMYget
				.CommandType = adCmdText
				.CommandText = "{?= call db_academy.[dbo].[sp_Fingers_upche_itemedit_readed]("&itemid&",'"&iidx&"')}"
				.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
				.Execute, , adExecuteNoRecords
				End With
			    returnValue = objCmd(0).Value
		Set objCmd = nothing
		
		IF returnValue <> 1 THEN
			Call Alert_return ("데이터 처리에 문제가 발생하였습니다.-error: case 'r' returnValue")
 		ELSE
 			Response.Write "<script>alert('반려가 확인되었습니다.'); parent.document.location.reload();</script>"
		END IF

CASE ELSE
		Call Alert_return ("데이터 처리에 문제가 발생하였습니다.-error: case else")
END SELECT	
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->