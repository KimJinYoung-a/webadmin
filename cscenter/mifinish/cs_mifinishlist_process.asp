<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/email/smsLib.asp"-->
<!-- #include virtual="/lib/email/maillib.asp"-->
<!-- #include virtual="/lib/email/mailLib2.asp"-->
<!-- #include virtual="/lib/email/mailFunc_Designer.asp"-->
<!-- #include virtual="/cscenter/lib/cs_action_mail_Function.asp"-->
<!-- #include virtual="/lib/classes/cscenter/oldmisendcls.asp"-->
<!-- #include virtual="/cscenter/lib/csAsfunction.asp" -->
<%

dim refer : refer = request.servervariables("http_referer")

dim sqlStr, i
dim reguserid

dim mode
mode = request("mode")

dim makerid, itemid
dim return_comment, return_reguserid
dim return_changemindyn

dim arrcsdetailidx, regReturnReason, nextactday, ckSendSMS, ckSendEmail, sendsmsmsg, sendmailmsg
dim lastorderserial, currorderserial, curritemid, curritemname, currreqhp, currbuyemail
dim csdetailidx
dim tmp_sendmailmsg, tmp_sendsmsmsg

dim mailTitle, smsTitle

makerid = request("makerid")
itemid = request("itemid")
return_comment = requestCheckVar(html2db(request("return_comment")),8000)
return_changemindyn = request("return_changemindyn")

arrcsdetailidx = request("arrcsdetailidx")
regReturnReason = request("regReturnReason")
nextactday = request("nextactday")
ckSendSMS = request("ckSendSMS")
ckSendEmail = request("ckSendEmail")
sendsmsmsg = request("sendsmsmsg")
sendmailmsg = request("sendmailmsg")

reguserid = session("ssBctId")



if (mode = "modifybrandmemo") then
	'// ========================================================================
	'// 업체 반품 메모
	return_reguserid = reguserid

	sqlStr = " IF EXISTS (SELECT brandid FROM [db_cs].[dbo].tbl_cs_brand_memo WHERE brandid = '" + CStr(makerid) + "') "
	sqlStr = sqlStr & "	update [db_cs].[dbo].tbl_cs_brand_memo set return_modifyday = getdate() "
	sqlStr = sqlStr & " ,return_comment = '" & return_comment & "' "
	sqlStr = sqlStr & " ,return_reguserid = '" & return_reguserid & "' "
	sqlStr = sqlStr & " where brandid = '" & makerid & "' "
	sqlStr = sqlStr & " ELSE "
	sqlStr = sqlStr & " insert into [db_cs].[dbo].tbl_cs_brand_memo(brandid, return_comment, return_modifyday, return_reguserid) "
	sqlStr = sqlStr & "  values('" & makerid & "', '" & return_comment & "', getdate(), '" & return_reguserid & "') "
	''response.write sqlStr
	rsget.Open sqlStr,dbget,1

	response.write "<script>alert('저장되었습니다.'); location.replace('" + CStr(refer) + "');</script>"

elseif (mode = "modifyitemmemo") then
	'// ========================================================================
	'// 업체 상품 반품 메모
	return_reguserid = reguserid

	sqlStr = " IF EXISTS (SELECT itemid FROM [db_cs].[dbo].tbl_cs_item_memo WHERE itemid = " + CStr(itemid) + ") "
	sqlStr = sqlStr & "	update [db_cs].[dbo].tbl_cs_item_memo set return_modifyday = getdate() "
	sqlStr = sqlStr & " ,return_changemindyn = '" & return_changemindyn & "' "
	sqlStr = sqlStr & " ,return_comment = '" & return_comment & "' "
	sqlStr = sqlStr & " ,return_reguserid = '" & return_reguserid & "' "
	sqlStr = sqlStr & " where itemid = '" & itemid & "' "
	sqlStr = sqlStr & " ELSE "
	sqlStr = sqlStr & " insert into [db_cs].[dbo].tbl_cs_item_memo(itemid, return_changemindyn, return_modifyday, return_reguserid, return_comment) "
	sqlStr = sqlStr & "  values(" & itemid & ", '" & return_changemindyn & "', getdate(), '" & return_reguserid & "', '" & return_comment & "') "
	''response.write sqlStr
	rsget.Open sqlStr,dbget,1

	response.write "<script>alert('저장되었습니다.'); location.replace('" + CStr(refer) + "');</script>"

elseif (mode = "regallreturnreason") then
	'// ========================================================================
	'// 반품안내 일괄전송
	return_reguserid = reguserid

	ckSendSMS   = CHKIIF(request("ckSendSMS")="on","Y","N")
    ckSendEmail = CHKIIF(request("ckSendEmail")="on","Y","N")

	if (regReturnReason = "25") then
		'송장입력 안내
		smsTitle = "텐바이텐 업체반품안내"
		mailTitle = "[텐바이텐 반품안내] 업체배송상품 반품안내메일입니다"
	else
		'반품불가 안내
		smsTitle = "텐바이텐 반품철회안내"
		mailTitle = "[텐바이텐 반품철회안내] 제작상품 변심반품철회 안내메일입니다"
	end if

	arrcsdetailidx = Split(arrcsdetailidx, ",")

	currorderserial = ""
	lastorderserial = ""
	for i = 0 to UBound(arrcsdetailidx)

		if (Trim(arrcsdetailidx(i)) <> "") then

			csdetailidx = Trim(arrcsdetailidx(i))
			currorderserial = ""
			curritemid = ""
			curritemname = ""
			currreqhp = ""
			currbuyemail = ""

			sqlStr = " select top 1 d.orderserial, d.itemid, d.itemname, o.reqhp, o.buyemail "
			sqlStr = sqlStr & " from "
			sqlStr = sqlStr & " 	db_cs.dbo.tbl_new_as_detail d "
			sqlStr = sqlStr & " 	join db_cs.dbo.tbl_new_as_list m "
			sqlStr = sqlStr & " 	on "
			sqlStr = sqlStr & " 		d.masterid = m.id "
			sqlStr = sqlStr & " 	join db_order.dbo.tbl_order_master o "
			sqlStr = sqlStr & " 	on "
			sqlStr = sqlStr & " 		m.orderserial = o.orderserial "
			sqlStr = sqlStr & " where d.id  = " + CStr(csdetailidx) + " "
			rsget.CursorLocation = adUseClient
			rsget.Open sqlStr, dbget, adOpenForwardOnly
			if Not rsget.Eof then
				currorderserial = rsget("orderserial")
				curritemid = rsget("itemid")
				curritemname = db2html(rsget("itemname"))
				currreqhp = rsget("reqhp")
				currbuyemail = rsget("buyemail")
			end if
			rsget.close

			'// ================================================================
			sqlStr = " IF Exists(select idx from [db_temp].dbo.tbl_csmifinish_list where csdetailidx=" & csdetailidx & ")"
			sqlStr = sqlStr + " BEGIN "
			sqlStr = sqlStr + "	    update [db_temp].dbo.tbl_csmifinish_list"
			sqlStr = sqlStr + "	    set code='" & regReturnReason & "'"
			if (nextactday <> "") then
				sqlStr = sqlStr + "	,ipgodate='" & nextactday & "'"
			else
				sqlStr = sqlStr + "	,ipgodate=NULL"
			end if
			sqlStr = sqlStr + "	, reguserid = '" & session("ssBctId") & "' "
			sqlStr = sqlStr + "	, lastupdate = getdate() "
			sqlStr = sqlStr + "	    where csdetailidx=" & csdetailidx
			sqlStr = sqlStr + " END "
			sqlStr = sqlStr + " ELSE "
			sqlStr = sqlStr + " BEGIN "
			sqlStr = sqlStr + "	    insert into [db_temp].dbo.tbl_csmifinish_list"
			sqlStr = sqlStr + "	    (csdetailidx, orderserial, itemid, itemoption,"
			sqlStr = sqlStr + "	    itemno, itemlackno, code, ipgodate, reqstr, "
			sqlStr = sqlStr + "	    itemname, itemoptionname, reguserid, lastupdate)"
			sqlStr = sqlStr + "	    select d.id, m.orderserial, d.itemid, d.itemoption,"
			sqlStr = sqlStr + "	    d.regitemno, d.regitemno, '" & regReturnReason & "',"
			if (nextactday<>"") then
				sqlStr = sqlStr + "	'" & nextactday & "','',"
			else
				sqlStr = sqlStr + "	NULL,'',"
			end if
			sqlStr = sqlStr + "	    d.itemname, d.itemoptionname, '" & session("ssBctId") & "', getdate() "
			sqlStr = sqlStr + " from db_cs.dbo.tbl_new_as_list m "
			sqlStr = sqlStr + " 	join db_cs.dbo.tbl_new_as_detail d "
			sqlStr = sqlStr + " 	on "
			sqlStr = sqlStr + " 		m.id = d.masterid "
			sqlStr = sqlStr + "	    where d.id=" & csdetailidx
			sqlStr = sqlStr + " END "
			''rw   sqlStr
			dbget.Execute sqlStr


			'' '// ================================================================
			tmp_sendsmsmsg = sendsmsmsg
			tmp_sendmailmsg = sendmailmsg

			tmp_sendsmsmsg = Replace(tmp_sendsmsmsg, "[상품명]", curritemname)
			tmp_sendsmsmsg = Replace(tmp_sendsmsmsg, "[상품코드]", curritemid)

			tmp_sendmailmsg = Replace(tmp_sendmailmsg, "[상품명]", curritemname)
			tmp_sendmailmsg = Replace(tmp_sendmailmsg, "[상품코드]", curritemid)


			if (ckSendSMS="Y") then
				if (application("Svr_Info")<>"Dev") then
					''SMS 발송

					if (lastorderserial <> currorderserial) then
						lastorderserial = currorderserial

						if LenB(tmp_sendsmsmsg) > 80 then
							Call SendNormalLMS(currreqhp, smsTitle, "", tmp_sendsmsmsg)
						else
							Call SendNormalSMS(currreqhp, "", tmp_sendsmsmsg)
						end if

						'// 메모저장
						Call AddCsMemo(currorderserial,"1","",session("ssBctId"),"[SMS "+ currreqhp + "]" + html2db(tmp_sendsmsmsg))
					else
						'// 중복발송 않함.
						'// 상품 한개만 발송
					end if

				end if
			end if

			if (ckSendEmail="Y") then
				if (application("Svr_Info")<>"Dev") then
					''EMail발송
					Call sendmailCS(currbuyemail,mailTitle,nl2br(tmp_sendmailmsg))

					Call AddCsMemo(currorderserial,"1","",session("ssBctId"),"[Mail]" + mailTitle + VbCrlf + html2db(tmp_sendmailmsg))
				end if
			end if

		end if
	next

	if (ckSendSMS="Y") and (ckSendEmail="Y") then
		response.write "<script language='javascript'>alert('SMS와 메일이 발송 되었습니다.');</script>"
	elseif (ckSendSMS="Y") then
		response.write "<script language='javascript'>alert('SMS가 발송 되었습니다.');</script>"
	elseif (ckSendSMS="Y") then
		response.write "<script language='javascript'>alert('메일이 발송 되었습니다.');</script>"
	else
		response.write "<script language='javascript'>alert('처리 되었습니다.');</script>"
	end if

	response.write "<script>alert('저장되었습니다.'); location.replace('" + CStr(refer) + "');</script>"

end if

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
