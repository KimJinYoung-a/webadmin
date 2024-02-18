<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%

dim refer, i
refer = request.ServerVariables("HTTP_REFERER")

dim sqlStr
dim mode
dim startdate, enddate
dim sellsite, yyyymm, yyyymmdd, addval

mode = requestCheckVar(request("mode"), 32)
sellsite = requestCheckVar(request("sellsite"), 32)
yyyymm = requestCheckVar(request("yyyymm"), 32)
yyyymmdd = requestCheckVar(request("yyyymmdd"), 10)
addval	= requestCheckVar(request("addval"), 10)

Dim extOrderserial : extOrderserial = requestCheckVar(request("extOrderserial"), 32)
Dim extOrderserSeq : extOrderserSeq = requestCheckVar(request("extOrderserSeq"), 32)
Dim newSliceNo     : newSliceNo = requestCheckVar(request("newSliceNo"), 10)

Dim Orderserial : Orderserial = requestCheckVar(request("Orderserial"), 11)
Dim itemid : itemid = requestCheckVar(request("itemid"), 10)
Dim itemoption : itemoption = requestCheckVar(request("itemoption"), 4)
Dim addcomment : addcomment = requestCheckVar(request("addcomment"), 100)

Dim outmallorderseq : outmallorderseq = requestCheckVar(request("outmallorderseq"), 10)
Dim chgval : chgval = requestCheckVar(request("chgval"), 10)

dim ChkIxCnt : ChkIxCnt = request("chkix").count
dim extorgorderserial : extorgorderserial = requestCheckVar(request("extorgorderserial"), 32)
dim odetailidx : odetailidx = requestCheckVar(request("odetailidx"), 10)
dim chgprice   : chgprice = requestCheckVar(request("chgprice"), 9)
dim chkidx, sellsiteArr, extOrderserialArr, extOrderserSeqArr, OrgOrderserialArr, itemidArr, itemoptionArr
dim objCmd, returnValue, retErrText, retText

select case mode
	case "addcommission"
		sqlStr = " exec [db_jungsan].[dbo].[usp_Ten_OUTAMLL_Jungsan_AddVal] '"&mode&"','"&sellsite & "','" & yyyymm & "',"&addval&""
		dbget.Execute sqlStr

		response.write	"<script language='javascript'>" &_
		"	alert('처리되었습니다.'); " &_
		"	opener.location.reload(); window.close(); " &_
		"</script>"
	case "jungsanfixdateUpd"
		sqlStr = " exec [db_jungsan].[dbo].[usp_Ten_OUTAMLL_JungsanFixdate_Upd] '"&sellsite & "' "
		dbget.Execute sqlStr
		response.write	"<script language='javascript'>" &_
		"	alert('처리되었습니다.'); " &_
		"	location.replace('" + CStr(refer) + "'); " &_
		"</script>"
	case "delmeachulbyday"
		sqlStr = " exec [db_jungsan].[dbo].[usp_Ten_OUTAMLL_Jungsan_DelByDate] '" & sellsite & "', '" & yyyymmdd & "' "
		dbget.Execute sqlStr

		response.write	"<script language='javascript'>" &_
		"	alert('처리되었습니다.'); " &_
		"	opener.location.reload(); window.close(); " &_
		"</script>"

	case "extjungsandiffmake"
		sqlStr = " exec [db_dataSummary].[dbo].[usp_Ten_OUT_Jungsan_DIFF_Make_YYYYMM] '" & yyyymm & "', '" & sellsite & "' "
		db3_dbget.Execute sqlStr

		response.write	"<script language='javascript'>" &_
		"	alert('작성되었습니다.'); " &_
		"	location.replace('" + CStr(refer) + "'); " &_
		"</script>"
	case "extjungsanerrmake"
		sqlStr = " exec [db_dataSummary].[dbo].[usp_Ten_OUT_Jungsan_ERR_Make_YYYYMM] '" & yyyymm & "', '" & sellsite & "' "
		db3_dbget.Execute sqlStr

		response.write	"<script language='javascript'>" &_
		"	alert('작성되었습니다.'); " &_
		"	location.replace('" + CStr(refer) + "'); " &_
		"</script>"
	case "extjungsanaccDetailmake"
		sqlStr = " exec [db_dataSummary].[dbo].[usp_Ten_OUT_Jungsan_DIFF_Make_DetailList] '" & sellsite & "' "
		db3_dbget.Execute sqlStr

		response.write	"<script language='javascript'>" &_
		"	alert('작성되었습니다.'); " &_
		"	location.replace('" + CStr(refer) + "'); " &_
		"</script>"
		
	case "chgmaporder"
		if (ChkIxCnt>0) then
			for i=1 to ChkIxCnt
				chkidx			= request("chkix")(i)
				sellsiteArr		= Trim(request("sellsiteArr")(chkidx+1))
				extOrderserialArr	= Trim(request("extOrderserialArr")(chkidx+1))
				extOrderserSeqArr	= Trim(request("extOrderserSeqArr")(chkidx+1))

				OrgOrderserialArr	= Trim(request("OrgOrderserialArr")(chkidx+1))
				itemidArr			= Trim(request("itemidArr")(chkidx+1))
				itemoptionArr		= Trim(request("itemoptionArr")(chkidx+1))
				
				sqlStr = " exec [db_jungsan].[dbo].[usp_Ten_OUTAMLL_Jungsan_MapByHand] '"&sellsiteArr&"','"&extOrderserialArr&"','"&extOrderserSeqArr&"','"&OrgOrderserialArr&"',"&itemidArr&",'"&itemoptionArr&"'"
				dbget.Execute sqlStr

			next
		else
			if (request("chkix")="") then
				response.write "잘못된 접근입니다."
				dbget.close() : response.end
			end if

			chkidx = requestCheckvar(request("chkix"),10)
			sellsiteArr = requestCheckvar(request("sellsiteArr"),32)
			extOrderserialArr = requestCheckvar(request("extOrderserialArr"),32)
			extOrderserSeqArr  = requestCheckvar(request("extOrderserSeqArr"),32)

			OrgOrderserialArr = requestCheckvar(request("OrgOrderserialArr"),11)
			itemidArr = requestCheckvar(request("itemidArr"),10)
			itemoptionArr = requestCheckvar(request("itemoptionArr"),4)

			sqlStr = " exec [db_jungsan].[dbo].[usp_Ten_OUTAMLL_Jungsan_MapByHand] '"&sellsiteArr&"','"&extOrderserialArr&"','"&extOrderserSeqArr&"','"&OrgOrderserialArr&"',"&itemidArr&",'"&itemoptionArr&"'"
			dbget.Execute sqlStr
		end if

		''"	//alert('처리되었습니다.'); " &_
		response.write	"<script language='javascript'>" &_
		"	location.replace('" + CStr(refer) + "'); " &_
		"</script>"
	case "extorgorderserialedit"
		sqlStr = " exec [db_jungsan].[dbo].[usp_Ten_OUTAMLL_Jungsan_ExtOrgOrderserialEdit] '" & sellsite & "', '" & extOrderserial & "', '" & extOrderserSeq & "', '" & extorgorderserial & "' "
		dbget.Execute sqlStr

		''"	//alert('처리되었습니다.'); " &_
		response.write	"<script language='javascript'>" &_
		"	location.replace('" + CStr(refer) + "'); " &_
		"</script>"
	case "slicejitemno"
		sqlStr = " exec [db_jungsan].[dbo].[usp_Ten_OUTAMLL_Jungsan_SLICE_Itemno] '" & sellsite & "', '" & extOrderserial & "', '" & extOrderserSeq & "', '" & newSliceNo & "' "
		dbget.Execute sqlStr

		response.write	"<script language='javascript'>" &_
		"	alert('처리되었습니다.'); " &_
		"	location.replace('" + CStr(refer) + "'); " &_
		"</script>"

    case "addcmt"
		sqlStr = " exec [db_dataSummary].[dbo].[usp_Ten_OUTAMLL_Jungsan_Comment_add] '" & Orderserial & "', " & itemid & ", '" & itemoption & "','" & addcomment & "','"&session("ssBctId")&"'"
		db3_dbget.Execute sqlStr

		response.write	"<script language='javascript'>" &_
		"	alert('처리되었습니다.'); " &_
		"	location.replace('" + CStr(refer) + "'); " &_
		"</script>"
	case "delcmt"
		sqlStr = " EXEC  [db_dataSummary].[dbo].[usp_Ten_OUTAMLL_Jungsan_Comment_Del] "&requestCheckVar(request("rowidx"), 10)&""
		db3_dbget.Execute sqlStr

		response.write	"<script language='javascript'>" &_
		"	alert('처리되었습니다.'); " &_
		"	location.replace('" + CStr(refer) + "'); " &_
		"</script>"
	case "chgrealsellprice", "chgmatchitemoption"
		sqlStr = " EXEC  [db_jungsan].[dbo].[usp_Ten_OUTAMLL_XSiteOrderTmp_ChangVal] '"&mode&"','"&outmallorderseq&"','"&chgval&"'"
		dbget.Execute sqlStr

		response.write	"<script language='javascript'>" &_
		"	alert('처리되었습니다.'); " &_
		"	location.replace('" + CStr(refer) + "'); " &_
		"</script>"
	case "chgRealOrderRealsellprice"
		Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdStoredProc
			.CommandText = "db_jungsan.[dbo].[sp_Ten_OUTAMLL_Jungsan_realOrder_modify_CancelDtlCpn]"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Parameters.Append .CreateParameter("@orderserial", adVarchar, adParamInput, 11, Orderserial)
			.Parameters.Append .CreateParameter("@itemid", adInteger, adParamInput, , itemid)
			.Parameters.Append .CreateParameter("@itemoption", adVarchar, adParamInput, 4, itemoption)
			.Parameters.Append .CreateParameter("@chgval", adCurrency, adParamInput, , chgval)
			.Parameters.Append .CreateParameter("@retErrText", adVarchar, adParamOutput, 100, retErrText)

			.Execute, , adExecuteNoRecords
			End With
			returnValue = objCmd.Parameters("RETURN_VALUE").Value
			retErrText  = objCmd.Parameters("@retErrText").Value
		Set objCmd = nothing
	
		IF (returnValue<1) then
			retText = retText + retErrText &"\r\n"
		end if
		
		if (retText<>"")  then
			response.write "<script language='javascript'>" &_
						"	alert('"&retText&"'); " &_
						"</script>"
		else
			response.write	"<script language='javascript'>" &_
						"	alert('수정 되었습니다.'); " &_
						"</script>"
		end if
	case "chgCancelOrderJFixdtNULL"
		Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdStoredProc
			.CommandText = "db_jungsan.[dbo].[sp_Ten_OUTAMLL_Jungsan_cancelOrder_DlvDt2NULL]"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Parameters.Append .CreateParameter("@orderserial", adVarchar, adParamInput, 11, Orderserial)
			.Parameters.Append .CreateParameter("@retErrText", adVarchar, adParamOutput, 100, retErrText)

			.Execute, , adExecuteNoRecords
			End With
			returnValue = objCmd.Parameters("RETURN_VALUE").Value
			retErrText  = objCmd.Parameters("@retErrText").Value
		Set objCmd = nothing
	
		IF (returnValue<1) then
			retText = retText + retErrText &"\r\n"
		end if
		
		if (retText<>"")  then
			response.write "<script language='javascript'>" &_
						"	alert('"&retText&"'); " &_
						"</script>"
		else
			response.write	"<script language='javascript'>" &_
						"	alert('수정 되었습니다.'); " &_
						"</script>"
		end if
	case "mapcpnbyorderserial"
		Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdStoredProc
			.CommandText = "db_jungsan.[dbo].[sp_Ten_OUTAMLL_Jungsan_realOrder_modify_by_xSite_JungsanData_ONE]"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Parameters.Append .CreateParameter("@sellsite", adVarchar, adParamInput, 32, sellsite)
			.Parameters.Append .CreateParameter("@extOrderserial", adVarchar, adParamInput, 32, extOrderserial)
			.Parameters.Append .CreateParameter("@tenorderserial", adVarchar, adParamInput, 11, Orderserial)
			.Parameters.Append .CreateParameter("@retErrText", adVarchar, adParamOutput, 100, retErrText)

			.Execute, , adExecuteNoRecords
			End With
			returnValue = objCmd.Parameters("RETURN_VALUE").Value
			retErrText  = objCmd.Parameters("@retErrText").Value
		Set objCmd = nothing
	
		IF (returnValue<1) then
			retText = retText + retErrText &"\r\n"
		end if
		
		if (retText<>"")  then
			response.write "<script language='javascript'>" &_
						"	alert('"&retText&"'); " &_
						"</script>"
		else
			response.write	"<script language='javascript'>" &_
						"	alert('수정 되었습니다.'); " &_
						"</script>"
		end if
	case "chgkakaodtl"

		Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdStoredProc
			.CommandText = "db_jungsan.[dbo].[usp_Ten_OUTAMLL_KakaoOrderEdit_WithDlvPrice]"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Parameters.Append .CreateParameter("@orderserial", adVarchar, adParamInput, 11, Orderserial)
			.Parameters.Append .CreateParameter("@odetailidx", adBigInt, adParamInput, , odetailidx)
			.Parameters.Append .CreateParameter("@chgprice", adCurrency, adParamInput, , chgprice)
			.Parameters.Append .CreateParameter("@retErrText", adVarchar, adParamOutput, 100, retErrText)

			.Execute, , adExecuteNoRecords
			End With
			returnValue = objCmd.Parameters("RETURN_VALUE").Value
			retErrText  = objCmd.Parameters("@retErrText").Value
		Set objCmd = nothing
	
		IF (returnValue<1) then
			retText = retText + retErrText &"\r\n"
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
		
		
	case "remapCpnValbyJungsan"

	case else
		'//
		response.write "invalid-"&mode
end select

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
