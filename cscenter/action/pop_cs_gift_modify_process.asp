<% option Explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/email/maillib.asp" -->
<!-- #include virtual="/cscenter/lib/csAsfunction.asp" -->
<%

dim mode, orderserial, gift_code, modi_gift_code, modi_giftkind_code, itemlist

mode = requestCheckVar(request("mode"), 32)
orderserial = requestCheckVar(request("orderserial"), 32)

gift_code = requestCheckVar(request("gift_code"), 32)
modi_gift_code = requestCheckVar(request("modi_gift_code"), 32)
modi_giftkind_code = requestCheckVar(request("modi_giftkind_code"), 32)
itemlist = requestCheckVar(request("itemlist"), 320)

dim sqlStr
dim tmpStr

function GetGiftStr(orderserial, gift_code)
	dim resultStr : resultStr = ""

	sqlStr = " select top 1 chg_giftSTR + '(' + convert(varchar,gift_code) + ', ' + convert(varchar,giftkind_code) + ')' as giftStr "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " [db_order].[dbo].[tbl_order_gift] "
	sqlStr = sqlStr + " where orderserial = '" & orderserial & "' and gift_code = " & gift_code
	''rsget.Open sqlStr, dbget, 1
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
	if not rsget.EOF then
		resultStr = rsget("giftStr")
	end if
	rsget.close

	GetGiftStr = resultStr
end function

select case mode
	case "del"
		tmpStr = GetGiftStr(orderserial, gift_code)
		if (tmpStr = "") then
			response.write "시스템오류"
			dbget.close()	:	response.End
		end if

		call AddCsMemo(orderserial,"1","",session("ssBctId"),"사은품삭제" + VbCrlf + "사은품을 삭제했습니다." + VbCrlf + tmpStr)

		sqlStr = " delete "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " [db_order].[dbo].[tbl_order_gift] "
		sqlStr = sqlStr + " where orderserial = '" & orderserial & "' and gift_code = " & gift_code
		dbget.Execute sqlStr

		response.write "<script>alert('사은품을 삭제 했습니다.'); location.href = 'pop_cs_gift_modify.asp?orderserial=" & orderserial & "&mode=chk&itemlist=" & itemlist & "'</script>"
	case "modi"
		if (modi_gift_code = gift_code) then
			response.write "시스템오류 : 동일한 기프트코드입니다."
			dbget.close()	:	response.End
		end if

		tmpStr = GetGiftStr(orderserial, gift_code)
		if (tmpStr = "") then
			response.write "시스템오류"
			dbget.close()	:	response.End
		end if

		sqlStr = " delete "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " [db_order].[dbo].[tbl_order_gift] "
		sqlStr = sqlStr + " where orderserial = '" & orderserial & "' and gift_code = " & gift_code
		''response.write sqlStr & "<br />"
		dbget.Execute sqlStr

		sqlStr = " exec [db_order].[dbo].[sp_Ten_order_OpenGiftMODI] '" & orderserial & "', " & modi_gift_code & ", " & modi_giftkind_code
		''response.write sqlStr & "<br />"
		dbget.Execute sqlStr

		tmpStr = tmpStr & " => " & GetGiftStr(orderserial, modi_gift_code)

		call AddCsMemo(orderserial,"1","",session("ssBctId"),"사은품변경" + VbCrlf + "사은품을 변경했습니다." + VbCrlf + tmpStr)

		response.write "<script>alert('사은품을 변경 했습니다.'); location.href = 'pop_cs_gift_modify.asp?orderserial=" & orderserial & "&mode=chk&itemlist=" & itemlist & "'</script>"
	case else
		response.write "시스템오류 : " & mode
end select



%>


<!-- #include virtual="/lib/db/dbclose.asp" -->
