<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/cscenter/lib/csAsfunction.asp"-->

<!-- #include virtual="/lib/util/md5.asp" -->
<!-- #include virtual="/lib/email/maillib.asp" -->
<!-- #include virtual="/lib/classes/smscls.asp" -->

<!-- #include virtual="/lib/classes/giftcard/giftcard_MyCardInfoCls.asp" -->

<!-- #include virtual="/lib/classes/cscenter/sp_tenGiftCardCls.asp" -->

<!-- #include virtual="/cscenter/lib/cs_action_mail_Function.asp"-->
<!-- #include virtual="/lib/email/MailLib2.asp"-->
<%


dim giftorderserial, mode
dim buyname, buyphone, buyhp, buyemail, accountname
dim reqname, reqphone, reqhp, reqemail, reqzipcode, reqzipaddr, reqaddress, comment, requserid
dim cardribbon, message, fromname, yyyy, mm, dd, tt,  reqdate, reqtime
dim sqlStr
dim iAsID, divcd, reguserid, title, gubun01, gubun02, contents_jupsu, finishuser, contents_finish
dim ipkumdiv, jumundiv, userid, cancelyn, emailok, smsok
dim requiredetail, detailidx

dim nextjumundiv, nextipkumdiv, prevjumundiv, previpkumdiv

dim bookingYN, bookingDate, bookingDateHH, sendhp, MMSTitle, MMSContent

dim sendDiv, sendemail, emailTitle, emailContent, senddate
dim iscreatenewcode

dim masterCardCd
dim result

Dim paydateid, orgpaydateid

''' html2db : �Է½� ���. : 2���� Case RegCSMaster������ html2db ������� ����.


giftorderserial = request("giftorderserial")
mode        = request("mode")

buyname     = request("buyname")
buyphone    = request("buyphone")
buyhp       = request("buyhp")
buyemail    = request("buyemail")
accountname = request("accountname")
reguserid   = session("ssbctid")
requserid 	= request("requserid")

reqname     = request("reqname")
reqphone    = request("reqphone")
reqhp       = request("reqhp")
reqemail    = request("reqemail")
reqzipcode  = request("reqzipcode")
reqzipaddr  = request("reqzipaddr")
reqaddress  = request("reqaddress")
comment     = request("comment")


cardribbon  = request("cardribbon")
message     = request("message")
fromname    = request("fromname")
yyyy        = request("yyyy")
mm          = request("mm")
dd          = request("dd")
tt          = request("tt")

reqdate     = yyyy + "-" + dd + "-" + dd
reqtime     = tt

ipkumdiv    = request("ipkumdiv")
jumundiv    = request("jumundiv")
userid      = request("userid")

emailok     = request("emailok")
smsok       = request("smsok")

requiredetail = request("requiredetail")
detailidx     = request("detailidx")

bookingYN     	= request("bookingYN")
bookingDate     = request("bookingDate")
bookingDateHH   = request("bookingDateHH")
sendhp     		= request("sendhp")
MMSTitle     	= request("MMSTitle")
MMSContent     	= request("MMSContent")

paydateid     	= request("paydateid")


sendDiv = request("sendDiv")
sendemail = request("sendemail")
reqEmail = request("reqEmail")
emailTitle = request("emailTitle")
emailContent = request("emailContent")
iscreatenewcode = request("iscreatenewcode")


dim errcode
dim osms
const CNEXT = " => "

dim myorderdetail,i

if (mode = "modifybuyerinfo") then
    On Error Resume Next
    dbget.beginTrans

    If Err.Number = 0 Then
        errcode = "001"

        divcd   = "A900"
        title   = "�ֹ��� ���� ����"
        gubun01 = "C004"
        gubun02 = "CD99"

        contents_jupsu = ""
        finishuser      = reguserid
        contents_finish = ""


        sqlStr = " select top 1 IsNULL(buyname,'') as buyname"
        sqlStr = sqlStr + " ,IsNULL(buyphone,'') as buyphone"
        sqlStr = sqlStr + " ,IsNULL(buyhp,'') as buyhp"
        sqlStr = sqlStr + " ,IsNULL(buyemail,'') as buyemail"
        sqlStr = sqlStr + " ,IsNULL(accountname,'') as accountname"
        sqlStr = sqlStr + " from db_order.dbo.tbl_giftcard_order "
        sqlStr = sqlStr + " where giftorderserial='" + CStr(giftorderserial) + "' " + VbCrlf
        rsget.Open sqlStr,dbget,1
        if Not rsget.Eof then
            contents_jupsu = contents_jupsu & "���� ����" & VbCrlf

            if (db2html(rsget("buyname"))<>buyname) then
                contents_jupsu = contents_jupsu & "�ֹ��ڸ�: " & rsget("buyname") & CNEXT & buyname & VbCrlf
            end if

            if (rsget("buyphone")<>buyphone) then
                contents_jupsu = contents_jupsu & "�ֹ�����ȭ: " & rsget("buyphone") & CNEXT & buyphone & VbCrlf
            end if

            if (rsget("buyhp")<>buyhp) then
                contents_jupsu = contents_jupsu & "�ֹ����ڵ���: " & rsget("buyhp") & CNEXT & buyhp & VbCrlf
            end if

            if (db2html(rsget("buyemail"))<>buyemail) then
                contents_jupsu = contents_jupsu & "�ֹ����̸���: " & rsget("buyemail") & CNEXT & buyemail & VbCrlf
            end if

            if (db2html(rsget("accountname"))<>accountname) then
                contents_jupsu = contents_jupsu & "�Ա��ڸ�: " & rsget("accountname") & CNEXT & accountname & VbCrlf
            end if
        end if
        rsget.Close

        contents_finish = contents_jupsu

    end if

    If Err.Number = 0 Then
        errcode = "002"


        sqlStr = " update db_order.dbo.tbl_giftcard_order "     + VbCrlf
        sqlStr = sqlStr + " set buyname='" + html2db(buyname) + "' "   + VbCrlf
        sqlStr = sqlStr + " ,buyphone = '" + CStr(buyphone) + "' "  + VbCrlf
        sqlStr = sqlStr + " ,buyhp = '" + CStr(buyhp) + "' "        + VbCrlf
        sqlStr = sqlStr + " ,buyemail = '" + html2db(buyemail) + "' "  + VbCrlf
        sqlStr = sqlStr + " ,accountname = '" + html2db(accountname) + "' "    + VbCrlf
        sqlStr = sqlStr + " where giftorderserial='" + CStr(giftorderserial) + "' " + VbCrlf

        dbget.Execute sqlStr

    end if


    If Err.Number = 0 Then
        errcode = "003"
        ''html2db ������� ����.
        iAsID = RegCSMaster(divcd , giftorderserial, reguserid, title, contents_jupsu, gubun01, gubun02)
    end if

    If Err.Number = 0 Then
        errcode = "004"
        Call FinishCSMaster(iAsid, finishuser, html2db(contents_finish))

        Call AddCustomerOpenContents(iAsid, html2db(contents_finish))
    end if

    If Err.Number = 0 Then
        dbget.CommitTrans
    Else
        dbget.RollBackTrans
        response.write "<script>alert('����Ÿ�� �����ϴ� ���߿� ������ �߻��Ͽ����ϴ�.\r\n(�����ڵ� : " + CStr(errcode) + ")')</script>"
        response.write "<script>history.back()</script>"
        dbget.close()	:	response.End
    End If

    Call SendCsActionMail(iAsID)

    On Error Goto 0

    response.write "<script>alert('���� �Ǿ����ϴ�.');</script>"
    response.write "<script>opener.location.reload();</script>"
    response.write "<script>opener.focus(); window.close();</script>"
    dbget.close()	:	response.End
elseif (mode="modifyreceiverinfo") then
    On Error Resume Next
    dbget.beginTrans

    If Err.Number = 0 Then
        errcode = "001"

        divcd   = "A900"
        title   = "������ ���� ����"
        gubun01 = "C004"
        gubun02 = "CD99"

        contents_jupsu = ""
        finishuser      = reguserid
        contents_finish = ""


        sqlStr = " select top 1 IsNULL(reqhp,'') as reqhp"
        sqlStr = sqlStr + " ,IsNULL(reqemail,'') as reqemail"
        sqlStr = sqlStr + " from db_order.dbo.tbl_giftcard_order "
        sqlStr = sqlStr + " where giftorderserial='" + CStr(giftorderserial) + "' " + VbCrlf
        rsget.Open sqlStr,dbget,1
        if Not rsget.Eof then
            contents_jupsu = contents_jupsu & "������ ����" & VbCrlf
            if (rsget("reqhp")<>reqhp) then
                contents_jupsu = contents_jupsu & "�������ڵ���: " & rsget("reqhp") & CNEXT & reqhp & VbCrlf
            end if

            if (rsget("reqemail")<>reqemail) then
                contents_jupsu = contents_jupsu & "�������̸���: " & rsget("reqemail") & CNEXT & reqemail & VbCrlf
            end if
        end if
        rsget.Close

        contents_finish = contents_jupsu
    end if

     If Err.Number = 0 Then
        errcode = "002"


        sqlStr = " update db_order.dbo.tbl_giftcard_order "     + VbCrlf
        sqlStr = sqlStr + " set reqhp = '" + CStr(reqhp) + "' "        + VbCrlf
        sqlStr = sqlStr + " ,reqemail = '" + CStr(reqemail) + "' "  + VbCrlf
        sqlStr = sqlStr + " where giftorderserial='" + CStr(giftorderserial) + "' " + VbCrlf
        dbget.Execute sqlStr

    end if


    If Err.Number = 0 Then
        errcode = "003"
        iAsID = RegCSMaster(divcd , giftorderserial, reguserid, title, contents_jupsu, gubun01, gubun02)
    end if

    If Err.Number = 0 Then
        errcode = "004"
        Call FinishCSMaster(iAsid, finishuser, html2db(contents_finish))

        Call AddCustomerOpenContents(iAsid, html2db(contents_finish))
    end if

    If Err.Number = 0 Then
        dbget.CommitTrans
    Else
        dbget.RollBackTrans
        response.write "<script>alert('����Ÿ�� �����ϴ� ���߿� ������ �߻��Ͽ����ϴ�.\r\n(�����ڵ� : " + CStr(errcode) + ")')</script>"
        response.write "<script>history.back()</script>"
        dbget.close()	:	response.End
    End If

    Call SendCsActionMail(iAsID)

    On Error Goto 0

    response.write "<script>alert('���� �Ǿ����ϴ�.');</script>"
    response.write "<script>opener.location.reload();</script>"
    response.write "<script>opener.focus(); window.close();</script>"
    dbget.close()	:	response.End

elseif (mode="jumundivnextstep") then

	if (jumundiv = "1") and (ipkumdiv = "2") then
		'�����Ϸ�����
	    nextjumundiv = "3"
	    nextipkumdiv = "4"
	    title = "�����Ϸ� ��������"

	    if (bookingYN = "N") then
		    nextjumundiv = "5"
		    nextipkumdiv = "8"
		    title = "�����Ϸ� �� ���ۿϷ����� ��������"
	    end if
	elseif (jumundiv = "3") then
		'���ۿϷ�����
	    nextjumundiv = "5"
	    nextipkumdiv = "8"
	    title = "���ۿϷ� ��������"
	elseif (jumundiv = "5") then
		'��ϿϷ�����
	    nextjumundiv = "7"
	    nextipkumdiv = "8"
	    title = "��ϿϷ� ��������"
	else
        response.write "<script>alert('���� ���·� ���� �� �� �����ϴ�.');</script>"
        response.write "<script>history.back()</script>"
        dbget.close()	:	response.End
	end if

    divcd   = "A900"
    gubun01 = "C004"
    gubun02 = "CD99"

    contents_jupsu = ""
    finishuser      = reguserid
    contents_finish = ""

    sqlStr = "select top 1 userid, buyname, buyhp, buyemail, cancelyn, bookingYN, sendDiv, reqhp, reqemail, senddate, masterCardCode as masterCardCd "
    sqlStr = sqlStr + " from db_order.dbo.tbl_giftcard_order"
    sqlStr = sqlStr + " where giftorderserial='" + giftorderserial + "'"

    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        userid      = rsget("userid")
    	buyname     = db2html(rsget("buyname"))
    	buyhp       = db2html(rsget("buyhp"))
    	buyemail    = db2html(rsget("buyemail"))
    	cancelyn    = rsget("cancelyn")
    	sendDiv     = rsget("sendDiv")
    	reqhp     	= rsget("reqhp")
    	reqemail    = rsget("reqemail")
    	senddate    = rsget("senddate")
    	masterCardCd    = rsget("masterCardCd")


    end if
    rsget.close

    if (cancelyn="") or (cancelyn="Y") or (cancelyn="D") then
        response.write "<script>alert('���� �ֹ����� �ƴϰų� ��ҵ� �����Դϴ�.')</script>"
        response.write "<script>history.back()</script>"
        dbget.close()	:	response.End
    end if

    ''On Error Resume Next
    ''dbget.beginTrans

    If Err.Number = 0 Then
        errcode = "001"

		if (nextjumundiv = "5") then
			'�������

			'MMS�߼�
			set osms = new CSMSClass
			Call osms.sendGiftCardLMSMsg(giftOrderSerial)
			set osms=Nothing
			contents_jupsu = contents_jupsu & "SMS ���ۿϷ�(�޴º�HP:" + db2html(reqhp) + ")" & VbCrlf

			if (sendDiv = "E") then
				Call sendGiftCardEmail_SMTP(giftOrderSerial)
				contents_jupsu = contents_jupsu & "�̸��� ���ۿϷ�(�޴º�Email:" + db2html(reqemail) + ")" & VbCrlf
			end if

			sqlStr = " update [db_order].[dbo].tbl_giftcard_order" + vbCrlf
			sqlStr = sqlStr + " Set senddate=IsNull(senddate,getdate())" + vbCrlf
			sqlStr = sqlStr + " Where giftOrderSerial='" & CStr(giftorderserial) & "'"
			dbget.Execute sqlStr

			contents_finish = contents_jupsu
		end if

		if (nextjumundiv = "7") then

			dim GetLoginUserID
			GetLoginUserID = requserid					'ī������ ���̵�
			result = procGiftCardReg(masterCardCd)
			if (result = "W") or (result = "E") then

				dbget.RollBackTrans
				if (result = "W") then
					response.write "<script>alert('��Ͻ��� : ����Ʈī���ȣ�� ���ų� �߸��� �ڵ��Դϴ�')</script>"
				else
					response.write "<script>alert('��Ͻ��� : ��ȿ�Ⱓ�� ����� ī���Դϴ�')</script>"
				end if
				'response.write "<script>history.back()</script>"
		        dbget.close()	:	response.End

			end if

			contents_jupsu = contents_jupsu & "Giftī�尡 ��ϵǾ����ϴ�." & VbCrlf
			contents_finish = contents_jupsu
		end if


        sqlStr =	"update db_order.dbo.tbl_giftcard_order " & vbCrlf
		sqlStr = sqlStr & " set jumundiv = '" + CStr(nextjumundiv) + "' " & vbCrlf
		sqlStr = sqlStr & " ,ipkumdiv = '" + CStr(nextipkumdiv) + "' " & vbCrlf
		sqlStr = sqlStr & " ,ipkumdate=IsNull(ipkumdate,getdate()) " & vbCrlf
		sqlStr = sqlStr & " where giftorderserial='" & giftorderserial & "'"
		dbget.Execute sqlStr
    end IF

	If (Err.Number = 0) and (smsok<>"") Then
        errcode = "003"

	    '' SMS �߼�
        set osms = new CSMSClass

        if (jumundiv = "1") then
            osms.SendAcctIpkumOkMsg buyhp,giftorderserial
        end if
        set osms = Nothing

    end IF

	If (Err.Number = 0) and (emailok<>"") Then
        errcode = "004"

	    '' Email �߼�

	    if (jumundiv = "1") then
	        Call SendMailBankOk_GIFTCard(buyemail,buyname,giftorderserial)
	    end if
	end IF

    If Err.Number = 0 Then
        errcode = "005"
        iAsID = RegCSMaster(divcd , giftorderserial, reguserid, title, contents_jupsu, gubun01, gubun02)
    end if

    If Err.Number = 0 Then
        errcode = "006"
        Call FinishCSMaster(iAsid, finishuser, html2db(contents_finish))

        Call AddCustomerOpenContents(iAsid, html2db(contents_finish))
    end if

	If Err.Number = 0 Then
       '' dbget.CommitTrans
    Else
       '' dbget.RollBackTrans
        response.write "<script>alert(" + Chr(34) + "����Ÿ�� �����ϴ� ���߿� ������ �߻��Ͽ����ϴ�.\r\n(�����ڵ� : " + CStr(errcode) + "|" + Replace(Err.description,Vbcrlf," ") + ")" + Chr(34) + ")</script>"
        response.write "<script>history.back()</script>"
        dbget.close()	:	response.End
    End If


    response.write "<script>alert('���� �Ǿ����ϴ�.');</script>"
    response.write "<script>opener.location.reload();</script>"
    response.write "<script>opener.focus(); window.close();</script>"
    'dbget.close()	:	response.End

elseif (mode="resendcardcode") then

	if (iscreatenewcode = "Y") then
		title = "�ű������ڵ� ����"
	else
		title = "���������ڵ� ������"
	end if


	if (jumundiv = "5") then
		'
	else
        response.write "<script>alert('���ۿϷ� �����϶��� �������� �����մϴ�.');</script>"
        response.write "<script>history.back()</script>"
        dbget.close()	:	response.End
	end if

    divcd   = "A900"
    gubun01 = "C004"
    gubun02 = "CD99"

    contents_jupsu = ""
    finishuser      = reguserid
    contents_finish = ""

    sqlStr = "select top 1 userid, buyname, buyhp, buyemail, cancelyn, bookingYN, sendDiv, reqhp, reqemail, senddate, resendCnt, idx "
    sqlStr = sqlStr + " from db_order.dbo.tbl_giftcard_order"
    sqlStr = sqlStr + " where giftorderserial='" + giftorderserial + "'"

    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        userid      = rsget("userid")
    	buyname     = db2html(rsget("buyname"))
    	buyhp       = db2html(rsget("buyhp"))
    	buyemail    = db2html(rsget("buyemail"))
    	cancelyn    = rsget("cancelyn")
    	bookingYN   = rsget("bookingYN")
    	sendDiv     = rsget("sendDiv")
    	reqhp     	= rsget("reqhp")
    	reqemail    = rsget("reqemail")
    	senddate    = rsget("senddate")
    end if
    rsget.close

    if (cancelyn="") or (cancelyn="Y") or (cancelyn="D") then
        response.write "<script>alert('���� �ֹ����� �ƴϰų� ��ҵ� �����Դϴ�.')</script>"
        response.write "<script>history.back()</script>"
        dbget.close()	:	response.End
    end if

    dim resendCnt, idx
'rw giftorderserial
'rw resendCnt
'rw idx
'response.end

    ''On Error Resume Next
    ''dbget.beginTrans

    If Err.Number = 0 Then
        errcode = "001"

		masterCardCd = getGiftCardMasterCD(giftorderserial,resendCnt,idx)

		if (masterCardCd = "W") or (masterCardCd = "A") or (masterCardCd = "R") or (masterCardCd = "C") or (masterCardCd = "E") then

			dbget.RollBackTrans
			if (result = "W") then
				response.write "<script>alert('���۽��� : ��ҵǾ��ų� ���� ī���Դϴ�')</script>"
			elseif (result = "A") then
				response.write "<script>alert('���۽��� : �Ա� ������� ī���Դϴ�')</script>"
			elseif (result = "R") then
				response.write "<script>alert('���۽��� : �ֹ���ҵ� ī���Դϴ�')</script>"
			elseif (result = "C") then
				response.write "<script>alert('���۽��� : ����� �Ϸ�� ī��� ������ȣ �������� �� �� �����ϴ�')</script>"
			elseif (result = "E") then
				response.write "<script>alert('���۽��� : ��ȿ�Ⱓ�� ����� ī���Դϴ�')</script>"
			end if
			response.write "<script>history.back()</script>"
	        dbget.close()	:	response.End

		end if

		if (iscreatenewcode = "Y") then
			masterCardCd = getMasterCode(idx,16,resendCnt+1)

			'# ����Ʈī�� ������ȣ �߱� �α� ����
			Call putGiftCardMasterCDLog(giftOrderSerial,masterCardCd,resendCnt)

			contents_jupsu = contents_jupsu & "ī�������ڵ� �ű� �����Ϸ�." & VbCrlf
		end if

		Call chgOrderInfoResendMasterCD(giftOrderSerial,masterCardCd)

		'MMS�߼� /// Ʈ����� �ȵ�..
		set osms = new CSMSClass
		Call osms.sendGiftCardLMSMsg(giftOrderSerial)
		set osms = Nothing
		contents_jupsu = contents_jupsu & "SMS ���ۿϷ�(�޴º�HP:" + db2html(reqhp) + ")" & VbCrlf

		if (sendDiv = "E") then
			Call sendGiftCardEmail_SMTP(giftOrderSerial)
			contents_jupsu = contents_jupsu & "�̸��� ���ۿϷ�(�޴º�Email:" + db2html(reqemail) + ")" & VbCrlf
		end if

		contents_finish = contents_jupsu
    end IF

    If Err.Number = 0 Then
        errcode = "005"
        iAsID = RegCSMaster(divcd , giftorderserial, reguserid, title, contents_jupsu, gubun01, gubun02)
    end if

    If Err.Number = 0 Then
        errcode = "006"
        Call FinishCSMaster(iAsid, finishuser, html2db(contents_finish))

        Call AddCustomerOpenContents(iAsid, html2db(contents_finish))
    end if

	If Err.Number = 0 Then
        ''dbget.CommitTrans
    Else
        ''dbget.RollBackTrans
        response.write "<script>alert(" + Chr(34) + "����Ÿ�� �����ϴ� ���߿� ������ �߻��Ͽ����ϴ�.\r\n(�����ڵ� : " + CStr(errcode) + "|" + Replace(Err.description,Vbcrlf," ") + ")" + Chr(34) + ")</script>"
        response.write "<script>history.back()</script>"
        dbget.close()	:	response.End
    End If


    response.write "<script>alert('���� �Ǿ����ϴ�.');</script>"
    response.write "<script>opener.location.reload();</script>"
    response.write "<script>opener.focus(); window.close();</script>"
    'dbget.close()	:	response.End

elseif (mode="jumundivprevstep") then

	if (jumundiv = "7") then
		'���ۿϷ�����
	    prevjumundiv = "5"
	    title = "���ۿϷ� ������ȯ"
	else
        response.write "<script>alert('���� ���·� ��ȯ �� �� �����ϴ�.');</script>"
        response.write "<script>history.back()</script>"
        dbget.close()	:	response.End
	end if

    divcd   = "A900"
    gubun01 = "C004"
    gubun02 = "CD99"

    contents_jupsu = ""
    finishuser      = reguserid
    contents_finish = ""

    sqlStr = "select top 1 userid, buyname, buyhp, buyemail, cancelyn "
    sqlStr = sqlStr + " from db_order.dbo.tbl_giftcard_order"
    sqlStr = sqlStr + " where giftorderserial='" + giftorderserial + "'"

    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        userid      = rsget("userid")
    	buyname     = db2html(rsget("buyname"))
    	buyhp       = db2html(rsget("buyhp"))
    	buyemail    = db2html(rsget("buyemail"))
    	cancelyn    = rsget("cancelyn")
    end if
    rsget.close

    if (cancelyn="") or (cancelyn="Y") or (cancelyn="D") then
        response.write "<script>alert('���� �ֹ����� �ƴϰų� ��ҵ� �����Դϴ�.')</script>"
        response.write "<script>history.back()</script>"
        dbget.close()	:	response.End
    end if

response.write "������"
response.end

    On Error Resume Next
    dbget.beginTrans

    If Err.Number = 0 Then
        errcode = "001"

		'������ �������� Ȯ��(������ �� �ܾ��� 0�� �̻�����)


''		db_user.dbo.tbl_giftcard_current



        sqlStr =	"update db_order.dbo.tbl_giftcard_order " & vbCrlf
		sqlStr = sqlStr & " set jumundiv = '" + CStr(prevjumundiv) + "' " & vbCrlf
		sqlStr = sqlStr & " where giftorderserial='" & giftorderserial & "'"
		dbget.Execute sqlStr




'		delete from db_user.dbo.tbl_giftcard_regList
'		where
'
'1. tbl_giftcard_regList ����
'3. tbl_giftcard_log �α�����
'2. tbl_giftcard_current ��������Ʈ ����
'3.5 tbl_giftcard_order �ֹ����� ����(7��5)
'4. �� ������ȣ �߼�







    end IF

	If Err.Number = 0 Then
        errcode = "005"
        call AddCsMemo(giftorderserial,"1",userid,reguserid,title)
    end if

    If Err.Number = 0 Then
        errcode = "006"
        iAsID = RegCSMaster(divcd , giftorderserial, reguserid, title, contents_jupsu, gubun01, gubun02)
    end if

    If Err.Number = 0 Then
        errcode = "007"
        Call FinishCSMaster(iAsid, finishuser, html2db(contents_finish))

        Call AddCustomerOpenContents(iAsid, html2db(contents_finish))
    end if

	If Err.Number = 0 Then
        dbget.CommitTrans
    Else
        dbget.RollBackTrans
        response.write "<script>alert(" + Chr(34) + "����Ÿ�� �����ϴ� ���߿� ������ �߻��Ͽ����ϴ�.\r\n(�����ڵ� : " + CStr(errcode) + "|" + Replace(Err.description,Vbcrlf," ") + ")" + Chr(34) + ")</script>"
        response.write "<script>history.back()</script>"
        dbget.close()	:	response.End
    End If


    response.write "<script>alert('���� �Ǿ����ϴ�.');</script>"
    response.write "<script>opener.location.reload();</script>"
    response.write "<script>opener.focus(); window.close();</script>"
    'dbget.close()	:	response.End



elseif (mode="modipgkey") then

    On Error Resume Next
    dbget.beginTrans

    paydateid = Replace(paydateid, vbTab, " ")
    paydateid = Trim(paydateid)

    If Err.Number = 0 Then
        errcode = "001"

		''orgpaydateid

		sqlStr = "select top 1 paydateid "
		sqlStr = sqlStr + " from db_order.dbo.tbl_giftcard_order"
		sqlStr = sqlStr + " where giftorderserial='" + giftorderserial + "'"

		rsget.Open sqlStr,dbget,1
		if Not rsget.Eof then
			orgpaydateid      = rsget("paydateid")
		end if
		rsget.close

        sqlStr =	"update db_order.dbo.tbl_giftcard_order " & vbCrlf
		sqlStr = sqlStr & " set paydateid = '" + CStr(paydateid) + "' " & vbCrlf
		sqlStr = sqlStr & " where giftorderserial='" & giftorderserial & "'"
		dbget.Execute sqlStr

    end IF

	If Err.Number = 0 Then
        errcode = "005"
        call AddCsMemo(giftorderserial,"1",userid,reguserid,"PGKey ���� : " & orgpaydateid & " ==&gt; " & paydateid)
    end if

	If Err.Number = 0 Then
        dbget.CommitTrans
    Else
        dbget.RollBackTrans
        response.write "<script>alert(" + Chr(34) + "����Ÿ�� �����ϴ� ���߿� ������ �߻��Ͽ����ϴ�.\r\n(�����ڵ� : " + CStr(errcode) + "|" + Replace(Err.description,Vbcrlf," ") + ")" + Chr(34) + ")</script>"
        response.write "<script>history.back()</script>"
        dbget.close()	:	response.End
    End If

    response.write "<script>alert('���� �Ǿ����ϴ�.');</script>"
    response.write "<script>opener.location.reload();</script>"
    response.write "<script>opener.focus(); window.close();</script>"
    'dbget.close()	:	response.End

elseif (mode="modifymmsinfo") then

	title   = "MMS ���� ����"
    divcd   = "A900"
    gubun01 = "C004"
    gubun02 = "CD99"

    contents_jupsu = ""
    finishuser      = reguserid
    contents_finish = ""

    sqlStr = " select top 1 bookingYn, IsNULL(bookingDate,'') as bookingDate"
    sqlStr = sqlStr + " ,IsNULL(sendhp,'') as sendhp"
    sqlStr = sqlStr + " ,IsNULL(MMSTitle,'') as MMSTitle"
    sqlStr = sqlStr + " ,IsNULL(MMSContent,'') as MMSContent"
    sqlStr = sqlStr + " from db_order.dbo.tbl_giftcard_order "
    sqlStr = sqlStr + " where giftorderserial='" + CStr(giftorderserial) + "' " + VbCrlf
    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        contents_jupsu = contents_jupsu & "������ ����" & VbCrlf

        if (rsget("bookingYn") = "Y") then
	        if (Left(rsget("bookingDate"), 10) <> bookingDate) or (bookingDateHH*1 <> Hour(rsget("bookingDate"))) then
	            contents_jupsu = contents_jupsu & "�����Ͻ�: " & Left(rsget("bookingDate"), 10) & " " & Right(("0" & Hour(rsget("bookingDate"))), 2) & CNEXT & bookingDate & " " & bookingDateHH & VbCrlf
	        end if
        end if

        if (rsget("sendhp")<>sendhp) then
            contents_jupsu = contents_jupsu & "�����º�HP: " & rsget("sendhp") & CNEXT & sendhp & VbCrlf
        end if

        if (db2html(rsget("MMSTitle"))<>MMSTitle) then
            contents_jupsu = contents_jupsu & "MMS ����: " & db2html(rsget("MMSTitle")) & CNEXT & MMSTitle & VbCrlf
        end if

        if (db2html(rsget("MMSContent"))<>MMSContent) then
            contents_jupsu = contents_jupsu & "MMS ����: " & db2html(rsget("MMSContent")) & vbCrLf & CNEXT & vbCrLf & MMSContent & VbCrlf
        end if

    end if
    rsget.Close

    contents_finish = contents_jupsu

    On Error Resume Next
    dbget.beginTrans

    If Err.Number = 0 Then
        errcode = "001"

        sqlStr =	"update db_order.dbo.tbl_giftcard_order " & vbCrlf
		sqlStr = sqlStr & " set sendhp = '" + CStr(sendhp) + "' " & vbCrlf
		sqlStr = sqlStr & " , MMSTitle = '" + CStr(html2db(MMSTitle)) + "' " & vbCrlf
		sqlStr = sqlStr & " , MMSContent = '" + CStr(html2db(MMSContent)) + "' " & vbCrlf
		if (bookingDate <> "") then
			sqlStr = sqlStr & " , bookingDate = '" + CStr(bookingDate) + " " + bookingDateHH + ":00:00" + "' " & vbCrlf
		end if
		sqlStr = sqlStr & " where giftorderserial='" & giftorderserial & "'"
		dbget.Execute sqlStr
    end IF

	If Err.Number = 0 Then
        errcode = "005"
        call AddCsMemo(giftorderserial,"1",userid,reguserid,title)
    end if

    If Err.Number = 0 Then
        errcode = "006"
        iAsID = RegCSMaster(divcd , giftorderserial, reguserid, title, contents_jupsu, gubun01, gubun02)
    end if

    If Err.Number = 0 Then
        errcode = "007"
        Call FinishCSMaster(iAsid, finishuser, html2db(contents_finish))

        Call AddCustomerOpenContents(iAsid, html2db(contents_finish))
    end if

	If Err.Number = 0 Then
        dbget.CommitTrans
    Else
        dbget.RollBackTrans
        response.write "<script>alert(" + Chr(34) + "����Ÿ�� �����ϴ� ���߿� ������ �߻��Ͽ����ϴ�.\r\n(�����ڵ� : " + CStr(errcode) + "|" + Replace(Err.description,Vbcrlf," ") + ")" + Chr(34) + ")</script>"
        response.write "<script>history.back()</script>"
        dbget.close()	:	response.End
    End If


    response.write "<script>alert('���� �Ǿ����ϴ�.');</script>"
    response.write "<script>opener.location.reload();</script>"
    response.write "<script>opener.focus(); window.close();</script>"
    'dbget.close()	:	response.End


elseif (mode="modifyemailinfo") then

	title   = "����Email ���� ����"
    divcd   = "A900"
    gubun01 = "C004"
    gubun02 = "CD99"

    contents_jupsu 	= ""
    finishuser      = reguserid
    contents_finish = ""

    sqlStr = " select top 1 IsNULL(sendDiv,'S') as sendDiv"
    sqlStr = sqlStr + " ,IsNULL(sendemail,'') as sendemail"
    sqlStr = sqlStr + " ,IsNULL(reqEmail,'') as reqEmail"
    sqlStr = sqlStr + " ,IsNULL(emailTitle,'') as emailTitle"
    sqlStr = sqlStr + " ,IsNULL(emailContent,'') as emailContent"
    sqlStr = sqlStr + " from db_order.dbo.tbl_giftcard_order "
    sqlStr = sqlStr + " where giftorderserial='" + CStr(giftorderserial) + "' " + VbCrlf
    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        contents_jupsu = contents_jupsu & "������ ����" & VbCrlf

        if (rsget("sendDiv")<>sendDiv) then
        	if (sendDiv = "E") then
        		contents_jupsu = contents_jupsu & "���ۿ���: " & "�߼۾���" & CNEXT & "��������" & VbCrlf
        	else
        		contents_jupsu = contents_jupsu & "���ۿ���: " & "�߼۾���" & CNEXT & "��������" & VbCrlf
        	end if
        end if

        if (rsget("sendemail")<>sendemail) then
            contents_jupsu = contents_jupsu & "�����º�Email: " & rsget("sendemail") & CNEXT & sendemail & VbCrlf
        end if

        if (rsget("reqEmail")<>reqEmail) then
            contents_jupsu = contents_jupsu & "�޴º�Email: " & rsget("reqEmail") & CNEXT & reqEmail & VbCrlf
        end if

        if (db2html(rsget("emailTitle"))<>emailTitle) then
            contents_jupsu = contents_jupsu & "Email����: " & db2html(rsget("emailTitle")) & CNEXT & emailTitle & VbCrlf
        end if

        if (db2html(rsget("emailContent"))<>emailContent) then
            contents_jupsu = contents_jupsu & "Email����: " & db2html(rsget("emailContent")) & vbCrLf & CNEXT & vbCrLf & emailContent & VbCrlf
        end if

    end if
    rsget.Close

    contents_finish = contents_jupsu

    On Error Resume Next
    dbget.beginTrans

    If Err.Number = 0 Then
        errcode = "001"

        sqlStr =	"update db_order.dbo.tbl_giftcard_order " & vbCrlf
		sqlStr = sqlStr & " set sendDiv = '" + CStr(sendDiv) + "' " & vbCrlf
		sqlStr = sqlStr & " , sendemail = '" + CStr(html2db(sendemail)) + "' " & vbCrlf
		sqlStr = sqlStr & " , reqEmail = '" + CStr(html2db(reqEmail)) + "' " & vbCrlf
		sqlStr = sqlStr & " , emailTitle = '" + CStr(html2db(emailTitle)) + "' " & vbCrlf
		sqlStr = sqlStr & " , emailContent = '" + CStr(html2db(emailContent)) + "' " & vbCrlf
		sqlStr = sqlStr & " where giftorderserial='" & giftorderserial & "'"
		dbget.Execute sqlStr
    end IF

	If Err.Number = 0 Then
        errcode = "005"
        call AddCsMemo(giftorderserial,"1",userid,reguserid,title)
    end if

    If Err.Number = 0 Then
        errcode = "006"
        iAsID = RegCSMaster(divcd , giftorderserial, reguserid, title, contents_jupsu, gubun01, gubun02)
    end if

    If Err.Number = 0 Then
        errcode = "007"
        Call FinishCSMaster(iAsid, finishuser, html2db(contents_finish))

        Call AddCustomerOpenContents(iAsid, html2db(contents_finish))
    end if

	If Err.Number = 0 Then
        dbget.CommitTrans
    Else
        dbget.RollBackTrans
        response.write "<script>alert(" + Chr(34) + "����Ÿ�� �����ϴ� ���߿� ������ �߻��Ͽ����ϴ�.\r\n(�����ڵ� : " + CStr(errcode) + "|" + Replace(Err.description,Vbcrlf," ") + ")" + Chr(34) + ")</script>"
        response.write "<script>history.back()</script>"
        dbget.close()	:	response.End
    End If


    response.write "<script>alert('���� �Ǿ����ϴ�.');</script>"
    response.write "<script>opener.location.reload();</script>"
    response.write "<script>opener.focus(); window.close();</script>"
    'dbget.close()	:	response.End


elseif (mode="edithandmadereq") then

    ' �����ϴµ�
    response.end
    set myorderdetail = new COrderMaster
    myorderdetail.FRectgiftorderserial = giftorderserial
    myorderdetail.FRectDetailIdx = detailidx
    myorderdetail.GetOneOrderDetail


    ''������ ������ �϶�
    if (myorderdetail.FOneItem.FItemNo>1) then
        requiredetail = ""
        for i=0 to myorderdetail.FOneItem.FItemNo-1
            if (request.form("requiredetail"&i)<>"") then
                requiredetail = requiredetail & request.form("requiredetail"&i) & CAddDetailSpliter
            end if
        next

        if Right(requiredetail,2)=CAddDetailSpliter then
            requiredetail = Left(requiredetail,Len(requiredetail)-2)
        end if
    end if
    set myorderdetail = Nothing

    On Error Resume Next
    dbget.beginTrans

    If Err.Number = 0 Then
        errcode = "001"

        divcd       = "A900"
        title       = "�ֹ����� ��ǰ ���� ����"
        gubun01     = "C004"
        gubun02     = "CD99"

        contents_jupsu  = ""
        contents_finish = ""
        finishuser      = CFINISH_SYSTEM

        sqlStr = " select top 1 isnull(dd.requiredetailUTF8,d.requiredetail) as requiredetail"
        sqlStr = sqlStr + " ,IsNULL(d.itemname,'') as itemname"
        sqlStr = sqlStr + " ,IsNULL(d.itemoptionname,'') as itemoptionname"
        sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_detail d"
		sqlStr = sqlStr & " LEFT JOIN db_order.dbo.tbl_order_require dd" & vbcrlf
		sqlStr = sqlStr & "     ON d.idx = dd.detailidx" & vbcrlf
        sqlStr = sqlStr + " where d.giftorderserial='" + CStr(giftorderserial) + "' " + VbCrlf
        sqlStr = sqlStr + " and d.idx=" + CStr(detailidx)


        rsget.Open sqlStr,dbget,1
        if Not rsget.Eof then

            contents_jupsu = contents_jupsu & "���� ����" & VbCrlf

            if (db2html(rsget("requiredetail"))<>requiredetail) then
                contents_jupsu = contents_jupsu & "��ǰ��(�ɼ�): " & db2html(rsget("itemname"))
                if (rsget("itemoptionname")<>"") then
                    contents_jupsu = contents_jupsu & "(" & db2html(rsget("itemoptionname")) & ")" & VbCrlf
                end if
                contents_jupsu = contents_jupsu & "����: " & rsget("requiredetail") & VbCrlf & CNEXT & VbCrlf & requiredetail & VbCrlf
            end if

        end if
        rsget.Close

    end if

    if (contents_jupsu="") then
        response.write "<script language='javascript'>alert('�����Ͻ� ������ ���� ������ ��ġ�մϴ�. �������� �ʾҽ��ϴ�.');</script>"
        response.write "<script language='javascript'>history.back();</script>"
        dbget.close()	:	response.End
    else
        contents_jupsu = "���� ����" & VbCrlf & contents_jupsu
        contents_finish = contents_jupsu
    end if

    If Err.Number = 0 Then
        errcode = "002"

        sqlStr = "update [db_order].[dbo].tbl_order_detail" + VbCrlf
        sqlStr = sqlStr + " set requiredetail='" + html2db(requiredetail) + "'" + VbCrlf
        sqlStr = sqlStr + " where idx=" + CStr(detailidx)

        'response.write sqlStr & "<br>"
        dbget.Execute sqlStr

        sqlStr = "if exists(" & VbCrlf
        sqlStr = sqlStr & " select top 1 requiredetailUTF8 from [db_order].[dbo].tbl_order_require where detailidx="& detailidx &"" & VbCrlf
        sqlStr = sqlStr & " )" & VbCrlf
        sqlStr = sqlStr & "     begin" & VbCrlf
        sqlStr = sqlStr & "     update [db_order].[dbo].tbl_order_require set requiredetailUTF8=N'" & trim(html2db(requiredetail)) & "' , lastupdate=getdate() where detailidx="& detailidx &"" & VbCrlf
        sqlStr = sqlStr & "     end" & VbCrlf
        sqlStr = sqlStr & " else" & VbCrlf
        sqlStr = sqlStr & "     begin" & VbCrlf
        sqlStr = sqlStr & "     insert into [db_order].[dbo].tbl_order_require (detailidx, requiredetailUTF8, regdate, lastupdate) values (" & VbCrlf
        sqlStr = sqlStr & "     "& trim(detailidx) &", N'" & trim(html2db(requiredetail)) & "', getdate(), getdate())" & VbCrlf
        sqlStr = sqlStr & "     end" & VbCrlf

        'response.write sqlStr & "<br>"
        dbget.Execute sqlStr
    end if

    If Err.Number = 0 Then
        errcode = "003"
        '' html2db ������� ����.
        iAsID = RegCSMaster(divcd , giftorderserial, reguserid, title, contents_jupsu, gubun01, gubun02)
    end if

    If Err.Number = 0 Then
        errcode = "004"
        Call FinishCSMaster(iAsid, finishuser, html2db(contents_finish))

    end if

    If Err.Number = 0 Then
        dbget.CommitTrans

        response.write "<script language='javascript'>alert('���� �Ǿ����ϴ�.');</script>"
        response.write "<script language='javascript'>opener.location.reload();</script>"
        response.write "<script language='javascript'>window.close();</script>"
        'dbget.close()	:	response.End
    Else
        dbget.RollBackTrans
        response.write "<script>alert('����Ÿ�� �����ϴ� ���߿� ������ �߻��Ͽ����ϴ�.\r\n(�����ڵ� : " + CStr(errcode) + ")')</script>"
        response.write "<script>history.back()</script>"
        dbget.close()	:	response.End
    End If
    On Error Goto 0
end if

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
