<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/cscenter/lib/csAsfunction.asp"-->
<%

dim refer, i
refer = request.ServerVariables("HTTP_REFERER")

dim sqlStr
dim mode
dim addsongjangdiv, addsongjangno

mode = requestCheckVar(request("mode"), 32)
addsongjangno = requestCheckVar(request("addsongjangno"), 32)
addsongjangdiv = requestCheckVar(request("addsongjangdiv"), 10)

dim ChkIxCnt : ChkIxCnt = request("chkix").count
dim chkidx,chgsongjangno,chgsongjangdiv, iodetailidx, iorderserial,isongjangno,isongjangdiv, chgdlvfinishdt

dim exceptmakerid		: exceptmakerid =  requestCheckVar(request("exceptmakerid"), 32)
dim exceptsongjangdiv	: exceptsongjangdiv =  requestCheckVar(request("exceptsongjangdiv"), 10)

dim objCmd, returnValue, retErrText, retText

if (mode = "retry") then
	sqlStr = " exec [db_order].[dbo].[usp_Ten_Delivery_Trace_Queue_ADDByHand] '"&addsongjangno&"',"&addsongjangdiv
	dbget.Execute sqlStr

	response.write	"<script language='javascript'>" &_
					"	alert('저장되었습니다.'); " &_
					"	location.replace('" + CStr(refer) + "'); " &_
					"</script>"
elseif (mode = "chgdftsongjangdiv") then
	sqlStr = " exec [db_partner].[dbo].[usp_Ten_partner_InfoChange_songjangdiv] '"&requestCheckVar(request("makerid"), 32)&"',"&requestCheckVar(request("chgdiv"), 10)
	dbget.Execute sqlStr

	response.write	"<script language='javascript'>" &_
					"	alert('수정되었습니다..'); " &_
					"	opener.location.reload(); window.close();" &_
					"</script>"
elseif (mode="refreshfakesummary") then
	sqlStr = " exec [db_order].[dbo].[usp_Ten_Delivery_Trace_CheckOrderList_MAKE] "
	dbget.Execute sqlStr

	response.write	"<script language='javascript'>" &_
					"	alert('반영되었습니다..'); " &_
					"	opener.location.reload(); window.close();" &_
					"</script>"
elseif (mode="etcdlvfinauto") then
	sqlStr = " exec [db_order].[dbo].[usp_Ten_Delivery_Trace_ETCSongjang_AutoFinish] "
	dbget.Execute sqlStr

	response.write	"<script language='javascript'>" &_
					"	alert('반영되었습니다..'); " &_
					"	opener.location.reload(); window.close();" &_
					"</script>"
elseif (mode="addexceptbrand") or (mode="delexceptbrand") then
	Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdStoredProc
			.CommandText = "db_order.[dbo].[usp_TEN_Delivery_Trace_ExceptBrand_Add]"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Parameters.Append .CreateParameter("@songjangdiv", adInteger, adParamInput, , exceptsongjangdiv)
			.Parameters.Append .CreateParameter("@makerid", adVarchar, adParamInput, 32, exceptmakerid)
			.Parameters.Append .CreateParameter("@reguserid", adVarchar, adParamInput, 32, session("ssBctID"))
			.Parameters.Append .CreateParameter("@actionType", adVarchar, adParamInput, 1, CHKIIF(mode="delexceptbrand","D",""))
			.Parameters.Append .CreateParameter("@retErrText", adVarchar, adParamOutput, 100, retErrText)

			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd.Parameters("RETURN_VALUE").Value
			retErrText  = objCmd.Parameters("@retErrText").Value
	Set objCmd = nothing

	retText = "처리되었습니다."

	if returnValue<0 then
		retText = retErrText
	end if
	%>
	<script language='javascript'>
		alert('<%=retText%>');
		location.replace('<%=CStr(refer)%>');
	</script>
	<%
elseif (mode="chgdtl") then
	if (ChkIxCnt>0) then
		 for i=1 to ChkIxCnt
		 	chkidx			= request("chkix")(i)
			chgsongjangno	= Trim(request("chgsongjangno")(chkidx+1))
			chgsongjangdiv	= Trim(request("chgsongjangdiv")(chkidx+1))
			chgdlvfinishdt	= Trim(request("chgdlvfinishdt")(chkidx+1))

			iodetailidx		= Trim(request("odetailidx")(chkidx+1))
			iorderserial	= Trim(request("orderserial")(chkidx+1))
			isongjangno		= (request("songjangno")(chkidx+1))
			isongjangdiv	= Trim(request("songjangdiv")(chkidx+1))

			if (chgdlvfinishdt<>"") and NOT isDate(chgdlvfinishdt) then
				rw "skip ("&iodetailidx&") 배송완료일오류:"&chgdlvfinishdt
			else
				'sqlStr = " exec [db_order].[dbo].[usp_Ten_Delivery_Trace_ChgOrderDtlEdit] "&iodetailidx&",'"&iorderserial&"','"&isongjangno&"',"&isongjangdiv&",'"&chgsongjangno&"',"&chgsongjangdiv&","&CHKIIF(chgdlvfinishdt="","NULL","'"&chgdlvfinishdt&"'")&",'"&session("ssBctId")&"'"
				'dbget.Execute sqlStr

				Set objCmd = Server.CreateObject("ADODB.COMMAND")
				With objCmd
					.ActiveConnection = dbget
					.CommandType = adCmdStoredProc
					.CommandText = "db_order.[dbo].[usp_Ten_Delivery_Trace_ChgOrderDtlEdit_WITH_RetERR]"
					.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
					.Parameters.Append .CreateParameter("@odetailidx", adBigInt, adParamInput, , iodetailidx)
					.Parameters.Append .CreateParameter("@orderserial", adVarchar, adParamInput, 11, iorderserial)
					.Parameters.Append .CreateParameter("@songjangno", adVarchar, adParamInput, 32, isongjangno)
					.Parameters.Append .CreateParameter("@songjangdiv", adInteger, adParamInput, , isongjangdiv)
					.Parameters.Append .CreateParameter("@chgsongjangno", adVarchar, adParamInput, 32, chgsongjangno)
					.Parameters.Append .CreateParameter("@chgsongjangdiv", adInteger, adParamInput, , chgsongjangdiv)

					.Parameters.Append .CreateParameter("@chgdlvfinishdt", adDBTimeStamp, adParamInput, , CHKIIF(chgdlvfinishdt="",Null,chgdlvfinishdt))

					.Parameters.Append .CreateParameter("@chguserid", adVarchar, adParamInput, 32, session("ssBctID"))
					.Parameters.Append .CreateParameter("@retErrText", adVarchar, adParamOutput, 100, retErrText)

					.Execute, , adExecuteNoRecords
					End With
					returnValue = objCmd.Parameters("RETURN_VALUE").Value
					retErrText  = objCmd.Parameters("@retErrText").Value
				Set objCmd = nothing

				IF (returnValue<1) then
					retText = retText + retErrText &"\r\n"
				end if
			end if
		 next
	else
		if (request("chkix")="") then
			response.write "잘못된 접근입니다.[001]"
			dbget.close() : response.end
		end if

		chkidx = requestCheckvar(request("chkix"),10)
		chgsongjangno = requestCheckvar(request("chgsongjangno"),32)
		chgsongjangdiv = requestCheckvar(request("chgsongjangdiv"),10)
		chgdlvfinishdt = requestCheckvar(request("chgdlvfinishdt"),19)
		iodetailidx  = requestCheckvar(request("odetailidx"),20)
		iorderserial = requestCheckvar(request("orderserial"),20)
		isongjangno = requestCheckvar(request("songjangno"),32)
		isongjangdiv = requestCheckvar(request("songjangdiv"),32)

		''rw chkidx&"|"&iorderserial&"|"&isongjangno&"|"&isongjangdiv&"|"&chgsongjangno&"|"&chgsongjangdiv
		if (chgdlvfinishdt<>"") and NOT isDate(chgdlvfinishdt) then
			rw "skip ("&iodetailidx&") 배송완료일오류:"&chgdlvfinishdt
		else
			'sqlStr = " exec [db_order].[dbo].[usp_Ten_Delivery_Trace_ChgOrderDtlEdit] "&iodetailidx&",'"&iorderserial&"','"&isongjangno&"',"&isongjangdiv&",'"&chgsongjangno&"',"&chgsongjangdiv&","&CHKIIF(chgdlvfinishdt="","NULL","'"&chgdlvfinishdt&"'")&",'"&session("ssBctId")&"'"
			'dbget.Execute sqlStr

			Set objCmd = Server.CreateObject("ADODB.COMMAND")
			With objCmd
				.ActiveConnection = dbget
				.CommandType = adCmdStoredProc
				.CommandText = "db_order.[dbo].[usp_Ten_Delivery_Trace_ChgOrderDtlEdit_WITH_RetERR]"
				.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
				.Parameters.Append .CreateParameter("@odetailidx", adBigInt, adParamInput, , iodetailidx)
				.Parameters.Append .CreateParameter("@orderserial", adVarchar, adParamInput, 11, iorderserial)
				.Parameters.Append .CreateParameter("@songjangno", adVarchar, adParamInput, 32, isongjangno)
				.Parameters.Append .CreateParameter("@songjangdiv", adInteger, adParamInput, , isongjangdiv)
				.Parameters.Append .CreateParameter("@chgsongjangno", adVarchar, adParamInput, 32, chgsongjangno)
				.Parameters.Append .CreateParameter("@chgsongjangdiv", adInteger, adParamInput, , chgsongjangdiv)

				.Parameters.Append .CreateParameter("@chgdlvfinishdt", adDBTimeStamp, adParamInput, , CHKIIF(chgdlvfinishdt="",Null,chgdlvfinishdt))

				.Parameters.Append .CreateParameter("@chguserid", adVarchar, adParamInput, 32, session("ssBctID"))
				.Parameters.Append .CreateParameter("@retErrText", adVarchar, adParamOutput, 100, retErrText)

				.Execute, , adExecuteNoRecords
				End With
				returnValue = objCmd.Parameters("RETURN_VALUE").Value
				retErrText  = objCmd.Parameters("@retErrText").Value
			Set objCmd = nothing

			IF (returnValue<1) then
				retText = retText + retErrText &"\r\n"
			end if

		end if
	end if

	if (retText<>"")  then
		response.write	"<script language='javascript'>" &_
					"	alert('"&retText&"'); " &_
					"	location.replace('" + CStr(refer) + "'); " &_
					"</script>"
	else
		response.write	"<script language='javascript'>" &_
					"	alert('수정 되었습니다.'); " &_
					"	location.replace('" + CStr(refer) + "'); " &_
					"</script>"
	end if
elseif (mode="finetc") then
	if (ChkIxCnt>0) then
		 for i=1 to ChkIxCnt
		 	chkidx			= request("chkix")(i)
			chgsongjangno	= Trim(request("chgsongjangno")(chkidx+1))
			chgsongjangdiv	= Trim(request("chgsongjangdiv")(chkidx+1))

			iodetailidx		= Trim(request("odetailidx")(chkidx+1))
			iorderserial	= Trim(request("orderserial")(chkidx+1))
			isongjangno		= (request("songjangno")(chkidx+1))
			isongjangdiv	= Trim(request("songjangdiv")(chkidx+1))


			sqlStr = " exec [db_order].[dbo].[usp_Ten_Delivery_Trace_ChgForceDlvFinishDt] "&iodetailidx&",'"&iorderserial&"','"&isongjangno&"',"&isongjangdiv&",'"&chgsongjangno&"',"&chgsongjangdiv&",'"&session("ssBctId")&"'"
			dbget.Execute sqlStr

			call AddCsMemo(iorderserial,"1","",session("ssBctId"), "기타내역 배송완료처리(" & iodetailidx & ")")
		 next
	else
		if (request("chkix")="") then
			response.write "잘못된 접근입니다.[002]"
			dbget.close() : response.end
		end if

		chkidx = requestCheckvar(request("chkix"),10)
		chgsongjangno = requestCheckvar(request("chgsongjangno"),32)
		chgsongjangdiv = requestCheckvar(request("chgsongjangdiv"),10)
		iodetailidx  = requestCheckvar(request("odetailidx"),20)
		iorderserial = requestCheckvar(request("orderserial"),20)
		isongjangno = requestCheckvar(request("songjangno"),32)
		isongjangdiv = requestCheckvar(request("songjangdiv"),32)

		''rw chkidx&"|"&iorderserial&"|"&isongjangno&"|"&isongjangdiv&"|"&chgsongjangno&"|"&chgsongjangdiv
		sqlStr = " exec [db_order].[dbo].[usp_Ten_Delivery_Trace_ChgForceDlvFinishDt] "&iodetailidx&",'"&iorderserial&"','"&isongjangno&"',"&isongjangdiv&",'"&chgsongjangno&"',"&chgsongjangdiv&",'"&session("ssBctId")&"'"
		dbget.Execute sqlStr

		call AddCsMemo(iorderserial,"1","",session("ssBctId"), "기타내역 배송완료처리(" & iodetailidx & ")")
	end if

	response.write	"<script language='javascript'>" &_
					"	alert('수정 되었습니다.'); " &_
					"	location.replace('" + CStr(refer) + "'); " &_
					"</script>"

elseif (mode="chgsongjang") then

	if (ChkIxCnt>0) then
		 for i=1 to ChkIxCnt
		 	chkidx			= request("chkix")(i)
			chgsongjangno	= Trim(request("chgsongjangno")(chkidx+1))
			chgsongjangdiv	= Trim(request("chgsongjangdiv")(chkidx+1))

			iodetailidx		= Trim(request("odetailidx")(chkidx+1))
			iorderserial	= Trim(request("orderserial")(chkidx+1))
			isongjangno		= (request("songjangno")(chkidx+1))
			isongjangdiv	= Trim(request("songjangdiv")(chkidx+1))


			''rw chkidx&"|"&iodetailidx&"|"&iorderserial&"|"&isongjangno&"|"&isongjangdiv&"|"&chgsongjangno&"|"&chgsongjangdiv

			sqlStr = " exec [db_order].[dbo].[usp_Ten_Delivery_Trace_ChgOrderSongjang] "&iodetailidx&",'"&iorderserial&"','"&isongjangno&"',"&isongjangdiv&",'"&chgsongjangno&"',"&chgsongjangdiv&",'"&session("ssBctId")&"'"
			dbget.Execute sqlStr

		 next
	else
		if (request("chkix")="") then
			response.write "잘못된 접근입니다.[003]"
			dbget.close() : response.end
		end if

		chkidx = requestCheckvar(request("chkix"),10)
		chgsongjangno = requestCheckvar(request("chgsongjangno"),32)
		chgsongjangdiv = requestCheckvar(request("chgsongjangdiv"),10)
		iodetailidx  = requestCheckvar(request("odetailidx"),20)
		iorderserial = requestCheckvar(request("orderserial"),20)
		isongjangno = requestCheckvar(request("songjangno"),32)
		isongjangdiv = requestCheckvar(request("songjangdiv"),32)

		''rw chkidx&"|"&iorderserial&"|"&isongjangno&"|"&isongjangdiv&"|"&chgsongjangno&"|"&chgsongjangdiv
		sqlStr = " exec [db_order].[dbo].[usp_Ten_Delivery_Trace_ChgOrderSongjang] "&iodetailidx&",'"&iorderserial&"','"&isongjangno&"',"&isongjangdiv&",'"&chgsongjangno&"',"&chgsongjangdiv&",'"&session("ssBctId")&"'"
		dbget.Execute sqlStr
	end if

	response.write	"<script language='javascript'>" &_
					"	alert('수정 요청 되었습니다.'); " &_
					"	location.replace('" + CStr(refer) + "'); " &_
					"</script>"

elseif (mode="receivedata") then

	' 'sqlStr = " exec [db_cs].[dbo].[usp_Ten_DeliveryTrackingList_Get] "
	' db3_dbget.Execute sqlStr

	' response.write	"<script language='javascript'>" &_
	' 				"	alert('저장되었습니다.'); " &_
	' 				"	location.replace('" + CStr(refer) + "'); " &_
	' 				"</script>"
else

	response.write "잘못된 접근입니다.[mode=" & mode & "]"

end if

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
