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
itemqna.FRectMakerid = replyuser ''2017/04/10 �߰�
itemqna.getOneItemQna

''��ȿ�� üũ// �ش� �귣�� ��ǰ�� �´���. 2017/04/10 eastone---------------
if (itemqna.FREsultCount<1) then
    response.write "<script>alert('��ȿ�� ��ǰ�� �ƴմϴ�.');</script>"
	response.write "<script>location.replace('newitemqna_view.asp?menupos="&menupos&"&id=" + id + "')</script>"
	dbget.close()	:	response.End
end if
'' ---------------------------------------------------------------------------

if (mode = "firstreply") then

	if Not IsNULL(itemqna.FOneItem.Freplydate) then
		response.write "<script>alert('�̹� �亯�� �� �����Դϴ�.');</script>"
		response.write "<script>location.replace('newitemqna_view.asp?id=" + id + "')</script>"
		dbget.close()	:	response.End
	end if


end if


if itemqna.FOneItem.Fisusing = "N" then
	response.write "<script>alert('�̹� ������ �����Դϴ�.');</script>"
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


		'### �� ��ϵ��� 30���� ������ �˸�, ���� �Ⱥ���. 20170410
		If DateDiff("d",itemqna.FOneItem.Fregdate,now()) < 30 Then
			'// MY�˸�
			dim myalarmtitle, myalarmsubtitle, myalarmcontents, myalarmwwwTargetURL

			if (itemqna.FOneItem.Fuserid <> "") then
				myalarmtitle = "<��ǰ Q&A>"
				myalarmsubtitle = itemqna.FOneItem.FContents
				if (Len(myalarmsubtitle) > 20) then
					myalarmsubtitle = Left(myalarmsubtitle, 20) & " ..."
				end if

				myalarmcontents = "���� ���ǿ� ���� �亯�帳�ϴ�."
				myalarmwwwTargetURL = "/my10x10/myitemqna.asp"

				Call MyAlarm_InsertMyAlarm_SCM(itemqna.FOneItem.Fuserid, "006", myalarmtitle, myalarmsubtitle, myalarmcontents, myalarmwwwTargetURL)
			end if

			''// ���� �߼�
			''IF (emailok = "Y") Then
		    IF (IsNULL(itemqna.FOneItem.Freplydate)) and (itemqna.FOneItem.FEmailOK="Y") Then '' �亯�Ȱ��� �߰��� �Ⱥ���,EmailOK="Y" �߰�  2017/04/10
				dim MailTo_Nm,MailTo
				MailTo_Nm=	itemqna.FOneItem.Fusername
				MailTo = itemqna.FOneItem.Fusermail
				dim oMail
				dim MailHTML
				dim MailTypeNo

				set oMail = New MailCls

				oMail.MailType 		= 16 '���� ������ ������ (mailLib2.asp ����)
				oMail.MailTitles 	= "[�ٹ�����]" & MailTo_Nm & "�Բ��� �����Ͻ� ���뿡 ���� �亯�Դϴ�."  '"��ſ��� ������ ���θ�, �ٹ����� [10X10=tenbyten]"
				oMail.SenderMail 	= "customer@10x10.co.kr"
				oMail.SenderNm 		= "�ٹ�����"

				oMail.AddrType 		= "string"
				oMail.ReceiverNm 	= MailTo_Nm
				oMail.ReceiverMail 	= MailTo

				MailHTML = oMail.getMailTemplate()

				IF MailHTML="" Then
					response.write "<script>alert('���Ϲ߼��� ���� �Ͽ����ϴ�.')</script>"
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
				MailHTML =replace(MailHTML,"[$QUESTION_CONTENTS$]","<b>[��������]</b><br><br>"& nl2br(server.HTMLEncode(db2html(itemqna.FOneItem.Fcontents))))
				MailHTML =replace(MailHTML,"[$ANSWER_TIME$]",now())
				MailHTML =replace(MailHTML,"[$ANSWER_CONTENTS$]","<b>[�亯����]</b><br><br>"& nl2br(server.HTMLEncode(db2html(replycontents))))
				MailHTML =replace(MailHTML,"[$ANSWER_NOTICE$]","")
				MailHTML =replace(MailHTML,"[$KEYVAL$]","")

				oMail.MailConts = MailHTML

				On Error Resume Next
				'oMail.Send()
				oMail.Send_CDO()
				'oMail.Send_CDONT()
				On Error Goto 0

				set oMail = nothing
				response.write "<script>alert('�亯������ �߼۵Ǿ����ϴ�.')</script>"
			    
				'dbget.close()	:	response.End

			 End IF
		End If
	 response.write "<script>location.replace('newitemqna_list.asp')</script>" ''��ġ���� 2017/04/10

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
