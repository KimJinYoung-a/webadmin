<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
Dim objCmd, returnValue, sMode
Dim ipartType, sPartTypeName, sPartName, arrPartsn, iPartsn
Dim iOrderNo, sOutBank, sOutAccNo, sOutName, sarap_cd, sbizsection_cd, scust_cd
Dim iOpExpPartidx,blnUsing, sAdminId
Dim sCardCo, sCardNo
Dim intLoop,menupos

dim strSQL
dim arrDepartmentID, iDepartmentID

sMode		= requestCheckvar(Request("hidM"),1)
menupos		= requestCheckvar(Request("menupos"),10)
ipartType	= requestCheckvar(Request("selPT"),10)
sPartTypeName=requestCheckvar(Request("sPTN"),60)
sPartName 	= requestCheckvar(Request("sPN"),60)
sOutBank 	= requestCheckvar(Request("selOB"),50)
sOutAccNo 	= requestCheckvar(Request("sOAN"),50)
sOutName 	= requestCheckvar(Request("sON"),30)
sarap_cd	= requestCheckvar(Request("dAC"),13)
sbizsection_cd= requestCheckvar(Request("selUP"),10)
iOrderNo 	= requestCheckvar(Request("iON"),10)
iOpExpPartidx= requestCheckvar(Request("hidOEP"),10)
sAdminId	= requestCheckvar(Request("hidAI"),32)
blnUsing 	= requestCheckvar(Request("rdoU"),1)

arrPartsn	= ReplaceRequestSpecialChar(Request("hidPsn"))
arrDepartmentID	= ReplaceRequestSpecialChar(Request("hidDPid"))

iPartsn 		= split(arrPartsn,",")
iDepartmentID 	= split(arrDepartmentID,",")

scust_cd	= requestCheckvar(Request("hidcustcd"),13)
sCardCo		= requestCheckvar(Request("selCCo"),50)
sCardNo		= requestCheckvar(Request("sCNo"),20)

 IF sbizsection_cd = "" THEN sbizsection_cd = 0
 IF iOrderNo = "" THEN iOrderNo = 0
 IF sarap_cd = "" THEN sarap_cd = 0
SELECT CASE sMode
Case "I"
   	dbget.beginTrans
   	'1. 타입등록
   if ipartType = 0 then
 	Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_OpExpPartType_Insert]('"&sPartTypeName&"')}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
	Set objCmd = nothing

	IF returnValue > 0 THEN
		ipartType = returnValue
	ELSE
	  dbget.RollBackTrans
	  Call Alert_return ("데이터 처리에 문제가 발생하였습니다.1")
	END IF
   end if

   '2.운영비관리 부서(팀)등록
   Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_OpExpPart_Insert]("&ipartType&",'"&sPartName&"','"&sOutBank&"','"&sOutAccNo&"','"&sOutName&"','"&sbizsection_cd&"', "&sarap_cd&" ,'"&scust_cd&"',"&iOrderNo&",'"&sCardCo&"','"&sCardNo&"','"&sAdminId&"')}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
	Set objCmd = nothing
	IF returnValue > 0 THEN
		iOpExpPartidx = returnValue
	ELSE
	  dbget.RollBackTrans
	  Call Alert_return ("데이터 처리에 문제가 발생하였습니다.2")
	END IF

	'3.부서 등록
	if arrPartsn <> "" then
	For intLoop = 0 To UBound(iPartsn)
	Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_OpExpPartInfo_Insert]("&iOpExpPartidx&","&trim(iPartsn(intLoop))&")}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
	Set objCmd = nothing

	IF returnValue = 0 THEN
	  dbget.RollBackTrans
	  Call Alert_return ("데이터 처리에 문제가 발생하였습니다.2")
	END IF

	Next
	end if

	'// 부서NEW
	if arrDepartmentID <> "" then
		strSQL = ""
		For intLoop = 0 To UBound(iDepartmentID)
			if (strSQL = "") then
				strSQL = " select " + CStr(iOpExpPartidx) + ", " + CStr(trim(iDepartmentID(intLoop))) + ", 'Y', getdate(), getdate() " & vbCrLf
			else
				strSQL = strSQL + " union all " & vbCrLf
				strSQL = strSQL + " select " + CStr(iOpExpPartidx) + ", " + CStr(trim(iDepartmentID(intLoop))) + ", 'Y', getdate(), getdate() " & vbCrLf
			end if
		Next

		strSQL = "insert into db_partner.dbo.tbl_OpExpDepartmentInfo(OpExpPartIdx, department_id, useYN, regdate, lastupdate)" & vbCrLf & strSQL
		rsget.Open strSQL, dbget, 1
	end if


		dbget.CommitTrans
	call Alert_closenreload("등록되었습니다.")
	response.end
Case "U"
	 dbget.beginTrans
   	'1. 타입등록
   if ipartType = 0 then
 	Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_OpExpPartType_Insert]('"&sPartTypeName&"')}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
	Set objCmd = nothing

	IF returnValue = 1 THEN
		ipartType = returnValue
	ELSE
	  dbget.RollBackTrans
	  Call Alert_return ("데이터 처리에 문제가 발생하였습니다.1")
	END IF
   end if

   '2.운영비관리 부서(팀)등록
   Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_OpExpPart_Update]("&iOpExpPartidx&","&ipartType&",'"&sPartName&"','"&sOutBank&"','"&sOutAccNo&"','"&sOutName&"','"&sbizsection_cd&"', "&sarap_cd&" ,'"&scust_cd&"',"&iOrderNo&",'"&sCardCo&"','"&sCardNo&"','"&sAdminId&"',"&blnUsing&")}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
	Set objCmd = nothing
	IF returnValue = 0 THEN
	  dbget.RollBackTrans
	  Call Alert_return ("데이터 처리에 문제가 발생하였습니다.2")
	END IF

	'3.부서 등록
	Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_OpExpPartInfo_Delete]("&iOpExpPartidx&")}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
	Set objCmd = nothing

	IF returnValue = 0 THEN
	  dbget.RollBackTrans
	  Call Alert_return ("데이터 처리에 문제가 발생하였습니다.3")
	END IF

	if arrPartsn <> "" then
	For intLoop = 0 To UBound(iPartsn)
	IF trim(iPartsn(intLoop)) <> "" THEN
	Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_OpExpPartInfo_Insert]("&iOpExpPartidx&","&trim(iPartsn(intLoop))&")}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
	Set objCmd = nothing

	IF returnValue = 0 THEN
	  dbget.RollBackTrans
	  Call Alert_return ("데이터 처리에 문제가 발생하였습니다.3")
	END IF
	END IF
	Next
	end if

	'// 부서NEW
	if arrDepartmentID <> "" then
		strSQL = " update db_partner.dbo.tbl_OpExpDepartmentInfo "
		strSQL = strSQL + " set useYN = 'N' "
		strSQL = strSQL + " where OpExpPartIdx = " + CStr(iOpExpPartidx) + " and useYN = 'Y' and department_id not in (" + CStr(arrDepartmentID) + ") "
		rsget.Open strSQL, dbget, 1

		strSQL = " select T.* " & vbCrLf
		strSQL = strSQL + " from " & vbCrLf
		strSQL = strSQL + " 	( " & vbCrLf

		For intLoop = 0 To UBound(iDepartmentID)
			if (intLoop = 0) then
				strSQL = strSQL + " select " + CStr(iOpExpPartidx) + " as OpExpPartIdx, " + CStr(trim(iDepartmentID(intLoop))) + " as department_id, 'Y' as useYN, getdate() as regdate, getdate() as lastupdate " & vbCrLf
			else
				strSQL = strSQL + " union all " & vbCrLf
				strSQL = strSQL + " select " + CStr(iOpExpPartidx) + ", " + CStr(trim(iDepartmentID(intLoop))) + ", 'Y', getdate(), getdate() " & vbCrLf
			end if
		Next

		strSQL = strSQL + " ) T " & vbCrLf
		strSQL = strSQL + " left join db_partner.dbo.tbl_OpExpDepartmentInfo e " & vbCrLf
		strSQL = strSQL + " on " & vbCrLf
		strSQL = strSQL + "	1 = 1 " & vbCrLf
		strSQL = strSQL + " and T.OpExpPartIdx = e.OpExpPartIdx " & vbCrLf
		strSQL = strSQL + " and T.department_id = e.department_id " & vbCrLf
		strSQL = strSQL + " and T.useYN = e.useYN " & vbCrLf
		strSQL = strSQL + " where " & vbCrLf
		strSQL = strSQL + " 1 = 1 " & vbCrLf
		strSQL = strSQL + " and e.idx is NULL " & vbCrLf

		strSQL = "insert into db_partner.dbo.tbl_OpExpDepartmentInfo(OpExpPartIdx, department_id, useYN, regdate, lastupdate)" & vbCrLf & strSQL

		rsget.Open strSQL, dbget, 1
	else
		strSQL = " update db_partner.dbo.tbl_OpExpDepartmentInfo "
		strSQL = strSQL + " set useYN = 'N' "
		strSQL = strSQL + " where OpExpPartIdx = " + CStr(iOpExpPartidx) + " and useYN = 'Y' "
		rsget.Open strSQL, dbget, 1
	end if

 dbget.CommitTrans
	call Alert_closenreload("수정되었습니다.")

CASE "T"	'파트타입 수정
 	Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_OpExpPartType_Update]("&ipartType&",'"&sPartTypeName&"',"&blnUsing&")}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
	Set objCmd = nothing
	IF returnValue = 0 THEN
		 Call Alert_return ("데이터 처리에 문제가 발생하였습니다.3")
	END IF
	call Alert_closenreload("수정되었습니다.")
CASE ELSE
	Call Alert_return ("데이터 처리에 문제가 발생하였습니다.0")
END SELECT
%>
