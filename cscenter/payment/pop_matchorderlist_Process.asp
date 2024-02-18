<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/new_ordercls.asp"-->
<%

dim refer
refer = request.ServerVariables("HTTP_REFERER")

dim mode
dim ipkumidx, orderserial, finishstr, reguserid, outmallorderserial
dim bankdate

dim ipkumCause

mode = requestCheckVar(request("mode"), 32)
ipkumidx = requestCheckVar(request("ipkumidx"), 32)
orderserial = requestCheckVar(request("orderserial"), 32)
finishstr = requestCheckVar(request("finishstr"), 32)
reguserid = session("ssBctId")
bankdate = requestCheckVar(request("bankdate"), 32)

ipkumCause = requestCheckVar(request("ipkumCause"), 32)
if (ipkumCause = "직접입력") then
	ipkumCause = requestCheckVar(request("ipkumCauseText"), 32)
end if

dim paramInfo, sqlStr, retParamInfo, RetErr, RetErrStr


if (mode = "matchWithOrder") then
	'==============================================================================
	paramInfo = Array(Array("@RETURN_VALUE"	, adInteger , adParamReturnValue,,0) 				_
				,Array("@idx"       		, adInteger	, adParamInput,, ipkumidx)			_
				,Array("@BackUserID"		, adVarchar	, adParamInput, 32, reguserid)		_
				,Array("@MatchOrderSerial"	, adVarchar	, adParamInput, 11, orderserial)	_
				,Array("@RetVal"			, adInteger , adParamOutput,, 0) 				_
			)

	sqlStr = "db_order.dbo.sp_Ten_IpkumConfirm_ByHand_Proc"
	retParamInfo = fnExecSPOutput(sqlStr,paramInfo)

	RetErr   = GetValue(retParamInfo, "@RetVal")         '

	if (RetErr <> 1) then
		response.write	"<script language='javascript'>" &_
						"	alert('매칭실패\n\n매칭에 실패했습니다[" + CStr(RetErr) + "].'); history.back(); " &_
						"</script>"
	else
		response.write	"<script language='javascript'>" &_
						"	alert('매칭되었습니다.'); history.back(); " &_
						"</script>"
	end if
elseif (mode = "matchByHand") then

    orderserial = finishstr
    if (Len(orderserial) >= 12) or (Len(orderserial) <= 10) then
	    '// 제휴몰 주문번호 -> 주문번호
	    outmallorderserial = orderserial
	    Call GetOrderserialWithOutmallOrderserial(outmallorderserial, orderserial)
	    if (orderserial = "") then
		    orderserial = outmallorderserial
	    end if
        finishstr = orderserial
    end if

	sqlStr = " update [db_order].[dbo].tbl_ipkum_list " & vbCrlf
	sqlStr = sqlStr + " set ipkumstate='7'" & vbCrlf
	sqlStr = sqlStr + " ,finishstr='" + html2db(finishstr) + "'" & vbCrlf
	sqlStr = sqlStr + " ,finishuser='" + CStr(reguserid) + "'" & vbCrlf
	sqlStr = sqlStr + " ,ipkumCause='" + html2db(CStr(ipkumCause)) + "'" & vbCrlf
	sqlStr = sqlStr + " ,lastupdate=getdate()" & vbCrlf
	sqlStr = sqlStr + " where idx=" + CStr(ipkumidx)
	rsget.Open sqlStr, dbget, 1

	response.write	"<script language='javascript'>" &_
					"	alert('매칭되었습니다.'); opener.location.reload(); opener.focus(); window.close(); " &_
					"</script>"

end if

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
