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
	'// ��ü ��ǰ �޸�
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

	response.write "<script>alert('����Ǿ����ϴ�.'); location.replace('" + CStr(refer) + "');</script>"

elseif (mode = "modifyitemmemo") then
	'// ========================================================================
	'// ��ü ��ǰ ��ǰ �޸�
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

	response.write "<script>alert('����Ǿ����ϴ�.'); location.replace('" + CStr(refer) + "');</script>"

elseif (mode = "regallreturnreason") then
	'// ========================================================================
	'// ��ǰ�ȳ� �ϰ�����
	return_reguserid = reguserid

	ckSendSMS   = CHKIIF(request("ckSendSMS")="on","Y","N")
    ckSendEmail = CHKIIF(request("ckSendEmail")="on","Y","N")

	if (regReturnReason = "25") then
		'�����Է� �ȳ�
		smsTitle = "�ٹ����� ��ü��ǰ�ȳ�"
		mailTitle = "[�ٹ����� ��ǰ�ȳ�] ��ü��ۻ�ǰ ��ǰ�ȳ������Դϴ�"
	else
		'��ǰ�Ұ� �ȳ�
		smsTitle = "�ٹ����� ��ǰöȸ�ȳ�"
		mailTitle = "[�ٹ����� ��ǰöȸ�ȳ�] ���ۻ�ǰ ���ɹ�ǰöȸ �ȳ������Դϴ�"
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

			tmp_sendsmsmsg = Replace(tmp_sendsmsmsg, "[��ǰ��]", curritemname)
			tmp_sendsmsmsg = Replace(tmp_sendsmsmsg, "[��ǰ�ڵ�]", curritemid)

			tmp_sendmailmsg = Replace(tmp_sendmailmsg, "[��ǰ��]", curritemname)
			tmp_sendmailmsg = Replace(tmp_sendmailmsg, "[��ǰ�ڵ�]", curritemid)


			if (ckSendSMS="Y") then
				if (application("Svr_Info")<>"Dev") then
					''SMS �߼�

					if (lastorderserial <> currorderserial) then
						lastorderserial = currorderserial

						if LenB(tmp_sendsmsmsg) > 80 then
							Call SendNormalLMS(currreqhp, smsTitle, "", tmp_sendsmsmsg)
						else
							Call SendNormalSMS(currreqhp, "", tmp_sendsmsmsg)
						end if

						'// �޸�����
						Call AddCsMemo(currorderserial,"1","",session("ssBctId"),"[SMS "+ currreqhp + "]" + html2db(tmp_sendsmsmsg))
					else
						'// �ߺ��߼� ����.
						'// ��ǰ �Ѱ��� �߼�
					end if

				end if
			end if

			if (ckSendEmail="Y") then
				if (application("Svr_Info")<>"Dev") then
					''EMail�߼�
					Call sendmailCS(currbuyemail,mailTitle,nl2br(tmp_sendmailmsg))

					Call AddCsMemo(currorderserial,"1","",session("ssBctId"),"[Mail]" + mailTitle + VbCrlf + html2db(tmp_sendmailmsg))
				end if
			end if

		end if
	next

	if (ckSendSMS="Y") and (ckSendEmail="Y") then
		response.write "<script language='javascript'>alert('SMS�� ������ �߼� �Ǿ����ϴ�.');</script>"
	elseif (ckSendSMS="Y") then
		response.write "<script language='javascript'>alert('SMS�� �߼� �Ǿ����ϴ�.');</script>"
	elseif (ckSendSMS="Y") then
		response.write "<script language='javascript'>alert('������ �߼� �Ǿ����ϴ�.');</script>"
	else
		response.write "<script language='javascript'>alert('ó�� �Ǿ����ϴ�.');</script>"
	end if

	response.write "<script>alert('����Ǿ����ϴ�.'); location.replace('" + CStr(refer) + "');</script>"

end if

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
