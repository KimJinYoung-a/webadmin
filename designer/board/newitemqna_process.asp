<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/email/maillib.asp" -->
<!-- #include virtual="/lib/classes/board/item_qnacls.asp" -->

<!-- #include virtual="/lib/util/scm_myalarmlib.asp" -->

<!-- #include virtual="/lib/email/mailLib2.asp" -->
<%

dim mailcontent

dim itemqna
dim boarditem
dim id, mode, replycontents, replyuser
dim emailok, extsitename

id = requestCheckVar(request("id"),10)
mode = requestCheckVar(request("mode"),32)
replycontents = html2db(request("replycontents"))
replyuser = session("ssBctId")

emailok = requestCheckVar(request("emailok"),32)
extsitename = requestCheckVar(request("extsitename"),32)

dim sql



set itemqna = new CItemQna
itemqna.FRectID = id
itemqna.FRectMakerid = replyuser ''2017/04/10 추가
itemqna.getOneItemQna

''유효성 체크// 해당 브랜드 상품이 맞는지. 2017/04/10 eastone---------------
if (itemqna.FREsultCount<1) then
    response.write "<script>alert('유효한 상품이 아닙니다.');</script>"
	response.write "<script>location.replace('newitemqna_view.asp?menupos="&menupos&"&id=" + id + "')</script>"
	dbget.close()	:	response.End
end if
'' ---------------------------------------------------------------------------

if (mode = "firstreply") then

	if Not IsNULL(itemqna.FOneItem.Freplydate) then
		response.write "<script>alert('이미 답변이 된 내용입니다.');</script>"
		response.write "<script>location.replace('newitemqna_view.asp?id=" + id + "')</script>"
		dbget.close()	:	response.End
	end if


end if


if itemqna.FOneItem.Fisusing = "N" then
	response.write "<script>alert('이미 삭제된 내용입니다.');</script>"
	response.write "<script>location.replace('newitemqna_view.asp?id=" + id + "')</script>"
	dbget.close()	:	response.End
end if


if (mode = "reply") or (mode = "firstreply") then
		sql = "update [db_cs].[dbo].tbl_my_item_qna " + VbCRlf
        sql = sql + " set replycontents = '" + replycontents + "'" + VbCRlf
        sql = sql + " , replyuser = '" + replyuser + "'" + VbCRlf
        sql = sql + " , replydate = getdate()" + VbCRlf
        sql = sql + " where id = '" + Cstr(id) + "'"

        rsget.Open sql, dbget, 1


		'### 글 등록된지 30일이 지나면 알림, 메일 안보냄. 20170410
		If DateDiff("d",itemqna.FOneItem.Fregdate,now()) < 30 Then
			'// MY알림
			dim myalarmtitle, myalarmsubtitle, myalarmcontents, myalarmwwwTargetURL

			if (itemqna.FOneItem.Fuserid <> "") then
				myalarmtitle = "<상품 Q&A>"
				myalarmsubtitle = itemqna.FOneItem.FContents
				if (Len(myalarmsubtitle) > 20) then
					myalarmsubtitle = Left(myalarmsubtitle, 20) & " ..."
				end if

				myalarmcontents = "고객님 문의에 대해 답변드립니다."
				myalarmwwwTargetURL = "/my10x10/myitemqna.asp"

				Call MyAlarm_InsertMyAlarm_SCM(itemqna.FOneItem.Fuserid, "006", myalarmtitle, myalarmsubtitle, myalarmcontents, myalarmwwwTargetURL)
			end if

			''// 메일 발송
			''IF (emailok = "Y") Then
		    IF (IsNULL(itemqna.FOneItem.Freplydate)) and (itemqna.FOneItem.FEmailOK="Y") Then '' 답변된것은 추가로 안보냄,EmailOK="Y" 추가  2017/04/10
				dim MailTo_Nm,MailTo
				MailTo_Nm=	itemqna.FOneItem.Fusername
				MailTo = itemqna.FOneItem.Fusermail
				dim oMail
				dim MailHTML
				dim MailTypeNo

				set oMail = New MailCls

				oMail.MailType 		= 16 '메일 종류별 고정값 (mailLib2.asp 참고)
				oMail.MailTitles 	= "[텐바이텐]" & MailTo_Nm & "님께서 문의하신 내용에 대한 답변입니다."  '"즐거움이 가득한 쇼핑몰, 텐바이텐 [10X10=tenbyten]"
				oMail.SenderMail 	= "customer@10x10.co.kr"
				oMail.SenderNm 		= "텐바이텐"

				oMail.AddrType 		= "string"
				oMail.ReceiverNm 	= MailTo_Nm
				oMail.ReceiverMail 	= MailTo

				MailHTML = oMail.getMailTemplate()

				IF MailHTML="" Then
					response.write "<script>alert('메일발송이 실패 하였습니다.')</script>"
			    	response.write "<script>location.replace('newitemqna_view.asp?id=" + id + "')</script>"
					dbget.close()	:	response.End
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
			    
				'dbget.close()	:	response.End

			 End IF
		End If
	 response.write "<script>location.replace('newitemqna_list.asp')</script>" ''위치변경 2017/04/10

elseif  (mode = "del") then
        sql = "update [db_cs].[dbo].tbl_my_item_qna " + VbCRlf
        sql = sql + " set isusing = 'N'" + VbCRlf
        sql = sql + " , replyuser = '" + replyuser + "'" + VbCRlf
        sql = sql + " , replydate = getdate()" + VbCRlf
        sql = sql + " where id = '" + Cstr(id) + "'"

        rsget.Open sql, dbget, 1
        response.write "<script>location.replace('newitemqna_list.asp')</script>"
end if

Set itemqna = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
