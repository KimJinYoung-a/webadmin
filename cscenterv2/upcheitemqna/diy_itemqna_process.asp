<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/cscenterv2/lib/classes/upcheitemqna/diy_item_qnacls.asp"-->

<!-- #include virtual="/lectureadmin/lib/email/mailLib2.asp" -->
<%

dim mailcontent

dim itemqna
dim boarditem
dim id, mode, replycontents, replyuser
dim emailok, extsitename

id = RequestCheckvar(request("id"),10)
mode = RequestCheckvar(request("mode"),16)
replycontents = html2db(request("replycontents"))
replyuser = session("ssBctId")
if replycontents <> "" then
	if checkNotValidHTML(replycontents) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end if
emailok = RequestCheckvar(request("emailok"),2)
extsitename = request("extsitename")

dim sql


set itemqna = new CItemQna
itemqna.FRectID = id
itemqna.getOneItemQna

if (mode = "firstreply") then

	if Not IsNULL(itemqna.FOneItem.Freplydate) then
		response.write "<script>alert('이미 답변이 된 내용입니다.');</script>"
		response.write "<script>location.replace('diy_itemqna_view.asp?id=" + id + "')</script>"
		dbACADEMYget.close()	:	response.End
	end if

end if


if (mode = "reply") or (mode = "firstreply") then
		sql = "update db_academy.dbo.tbl_diy_item_qna " + VbCRlf
        sql = sql + " set replycontents = '" + replycontents + "'" + VbCRlf
        sql = sql + " , replyuser = '" + replyuser + "'" + VbCRlf
        sql = sql + " , replydate = getdate()" + VbCRlf
        sql = sql + " where idx = '" + Cstr(id) + "'"

        rsACADEMYget.Open sql, dbACADEMYget, 1

	IF (emailok = "Y") Then

		dim MailTo_Nm,MailTo
		MailTo_Nm=	itemqna.FOneItem.Fusername
		MailTo = itemqna.FOneItem.Fusermail
		dim oMail
		dim MailHTML
		dim MailTypeNo

		set oMail = New MailCls

		oMail.MailType 		= 16 '메일 종류별 고정값 (mailLib2.asp 참고)
		oMail.MailTitles 	= "[Academy]" & MailTo_Nm & "님께서 문의하신 내용에 대한 답변입니다."  '"즐거움이 가득한 쇼핑몰, 텐바이텐 [10X10=tenbyten]"
		oMail.SenderMail 	= "customer@thefingers.co.kr"
		oMail.SenderNm 		= "아카데미"

		oMail.AddrType 		= "string"
		oMail.ReceiverNm 	= MailTo_Nm
		oMail.ReceiverMail 	= MailTo

		MailHTML = oMail.getMailTemplate()

		IF MailHTML="" Then
			response.write "<script>alert('메일발송이 실패 하였습니다.')</script>"
	    	response.write "<script>location.replace('diy_itemqna_view.asp?id=" + id + "')</script>"
			dbACADEMYget.close()	:	response.End
		End IF

		MailHTML =replace(MailHTML,"[$USER_NAME$]",MailTo_Nm)
		MailHTML =replace(MailHTML,"[$ITEMMAKER_NAME$]",itemqna.FOneItem.FBrandName)
		MailHTML =replace(MailHTML,"[$ITEM_NAME$]",itemqna.FOneItem.FItemName)
		MailHTML =replace(MailHTML,"[$ITEM_CODE$]",itemqna.FOneItem.FItemID)
		MailHTML =replace(MailHTML,"[$ITEM_PRICE$]",itemqna.FOneItem.FSellcash)
		MailHTML =replace(MailHTML,"[$ITEMIMG_URL$]",itemqna.FOneItem.Flistimage)
		MailHTML =replace(MailHTML,"[$QUESTION_TIME$]",itemqna.FOneItem.Fregdate)
		MailHTML =replace(MailHTML,"[$QUESTION_CONTENTS$]","<b>[질문내용]</b><br><br>"& nl2br(server.HTMLEncode(db2html(itemqna.FOneItem.Fcontents))))
		MailHTML =replace(MailHTML,"[$ANSWER_TIME$]",now())
		MailHTML =replace(MailHTML,"[$ANSWER_CONTENTS$]","<b>[답변내용]</b><br><br>"& nl2br(server.HTMLEncode(db2html(replycontents))))
		MailHTML =replace(MailHTML,"[$ANSWER_NOTICE$]","")
		MailHTML =replace(MailHTML,"[$KEYVAL$]","")

		oMail.MailConts = MailHTML

		On Error Resume Next
		'oMail.Send()
		oMail.Send_CDO()
		'oMail.Send_CDONT()
		On Error Goto 0

		set oMail = nothing
		response.write "<script>alert('답변메일이 발송되었습니다.')</script>"
	    response.write "<script>location.replace('diy_itemqna_view.asp?id=" + id + "')</script>"
		'dbACADEMYget.close()	:	response.End

	 End IF

elseif  (mode = "del") then
        sql = "update db_academy.dbo.tbl_diy_item_qna " + VbCRlf
        sql = sql + " set isusing = 'N'" + VbCRlf
        sql = sql + " , replyuser = '" + replyuser + "'" + VbCRlf
        sql = sql + " , replydate = getdate()" + VbCRlf
        sql = sql + " where idx = '" + Cstr(id) + "'"

        rsACADEMYget.Open sql, dbACADEMYget, 1
        response.write "<script>location.replace('diy_itemqna_list.asp')</script>"
end if

Set itemqna = Nothing
%>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
