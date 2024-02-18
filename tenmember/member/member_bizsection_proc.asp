<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/tenmember/incSessionTenMember.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<%

dim empno
empno = session("ssBctSn")

Dim sMode, sbizsectionCd,iPersent,sadminid
Dim sYear,sMonth,sYYYYMM
Dim dberr,returnValue,intLoop,objCmd
Dim sUserBizCd, sbizsectionUseCd

sMode	= requestCheckVar(Request("hidM"),1)
sYear 	= requestCheckVar(Request("selY"),4)
sMonth 	= requestCheckVar(Request("selM"),2)
sbizsectionCd = split(ReplaceRequestSpecialChar(request("sBCD")),",")
sbizsectionUseCd = split(ReplaceRequestSpecialChar(request("sBUCD")),",")
iPersent= split(ReplaceRequestSpecialChar(request("sPR")),",")
sadminid =  session("ssBctId")
sYYYYMM = sYear&"-"&Format00(2,sMonth)
sUserBizCd = requestCheckVar(Request("hidUBCD"),10)

dberr = 0
SELECT CASE sMode
	CASE "I"
	'IF day(date()) > 10 THEN
		'	Call Alert_return ("10일 이후에는 수정불가능합니다.")
		'response.end
	'END IF

	dbget.beginTrans

	Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
		.ActiveConnection = dbget
		.CommandType = adCmdText
		.CommandText = "{?= call db_partner.[dbo].[sp_Ten_user_Bizsection_Delete]('"& empno &"','"&sYYYYMM&"')}"

		.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
		.Execute, , adExecuteNoRecords
		End With
	    returnValue = objCmd(0).Value
	Set objCmd = nothing
		IF returnValue <> "1" THEN	dberr = dberr + 1

	For intLoop = 0 To UBound(sbizsectionCd)
		IF iPersent(intLoop) > 0 THEN
			Set objCmd = Server.CreateObject("ADODB.COMMAND")
			With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_user_Bizsection_Insert]('"& empno &"','"&sYYYYMM&"', '"&trim(sbizsectionCd(intLoop))&"' ,"&	iPersent(intLoop) &",'"&sadminid&"', '"&trim(sbizsectionUseCd(intLoop))&"')}"

			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
			Set objCmd = nothing
			IF returnValue <> "1" THEN	dberr = dberr + 1
		END IF
	Next

	IF dberr = "0" THEN
			dbget.CommitTrans
			Call Alert_move ("등록되었습니다.","pop_member_bizsection_Reg.asp?sen="& empno &"&sD="&sYYYYMM)
	ELSE
			dbget.RollBackTrans
			Call Alert_return ("데이터 처리에 문제가 발생하였습니다.")
	END IF
CASE "M"
	dbget.beginTrans
		 	For intLoop = 0 To UBound(sbizsectionCd)
		IF iPersent(intLoop) > 0 THEN
			Set objCmd = Server.CreateObject("ADODB.COMMAND")
			With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_userBizsection_avg_manualInsert]('"&sYYYYMM&"','"&sUserBizCd&"', '"&trim(sbizsectionCd(intLoop))&"' ,"&	iPersent(intLoop)&")}"

			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
			Set objCmd = nothing
			IF returnValue <> "1" THEN	dberr = dberr + 1
		END IF
	Next

	IF dberr = "0" THEN
			dbget.CommitTrans
			Call Alert_move ("등록되었습니다.","pop_userBiz_bizsection_Reg.asp?sen="& empno &"&sD="&sYYYYMM)
	ELSE
			dbget.RollBackTrans
			Call Alert_return ("데이터 처리에 문제가 발생하였습니다.")
	END IF
CASE ELSE
		Call Alert_return ("데이터 처리에 문제가 발생하였습니다.")
END SELECT


%>
