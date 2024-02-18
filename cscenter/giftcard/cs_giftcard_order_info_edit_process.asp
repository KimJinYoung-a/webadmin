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

''' html2db : 입력시 사용. : 2가지 Case RegCSMaster에서는 html2db 사용하지 말것.


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
        title   = "주문자 정보 수정"
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
            contents_jupsu = contents_jupsu & "변경 내역" & VbCrlf

            if (db2html(rsget("buyname"))<>buyname) then
                contents_jupsu = contents_jupsu & "주문자명: " & rsget("buyname") & CNEXT & buyname & VbCrlf
            end if

            if (rsget("buyphone")<>buyphone) then
                contents_jupsu = contents_jupsu & "주문자전화: " & rsget("buyphone") & CNEXT & buyphone & VbCrlf
            end if

            if (rsget("buyhp")<>buyhp) then
                contents_jupsu = contents_jupsu & "주문자핸드폰: " & rsget("buyhp") & CNEXT & buyhp & VbCrlf
            end if

            if (db2html(rsget("buyemail"))<>buyemail) then
                contents_jupsu = contents_jupsu & "주문자이메일: " & rsget("buyemail") & CNEXT & buyemail & VbCrlf
            end if

            if (db2html(rsget("accountname"))<>accountname) then
                contents_jupsu = contents_jupsu & "입금자명: " & rsget("accountname") & CNEXT & accountname & VbCrlf
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
        ''html2db 사용하지 말것.
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
        response.write "<script>alert('데이타를 저장하는 도중에 에러가 발생하였습니다.\r\n(에러코드 : " + CStr(errcode) + ")')</script>"
        response.write "<script>history.back()</script>"
        dbget.close()	:	response.End
    End If

    Call SendCsActionMail(iAsID)

    On Error Goto 0

    response.write "<script>alert('수정 되었습니다.');</script>"
    response.write "<script>opener.location.reload();</script>"
    response.write "<script>opener.focus(); window.close();</script>"
    dbget.close()	:	response.End
elseif (mode="modifyreceiverinfo") then
    On Error Resume Next
    dbget.beginTrans

    If Err.Number = 0 Then
        errcode = "001"

        divcd   = "A900"
        title   = "수령인 정보 수정"
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
            contents_jupsu = contents_jupsu & "변경전 내역" & VbCrlf
            if (rsget("reqhp")<>reqhp) then
                contents_jupsu = contents_jupsu & "수령인핸드폰: " & rsget("reqhp") & CNEXT & reqhp & VbCrlf
            end if

            if (rsget("reqemail")<>reqemail) then
                contents_jupsu = contents_jupsu & "수령인이메일: " & rsget("reqemail") & CNEXT & reqemail & VbCrlf
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
        response.write "<script>alert('데이타를 저장하는 도중에 에러가 발생하였습니다.\r\n(에러코드 : " + CStr(errcode) + ")')</script>"
        response.write "<script>history.back()</script>"
        dbget.close()	:	response.End
    End If

    Call SendCsActionMail(iAsID)

    On Error Goto 0

    response.write "<script>alert('수정 되었습니다.');</script>"
    response.write "<script>opener.location.reload();</script>"
    response.write "<script>opener.focus(); window.close();</script>"
    dbget.close()	:	response.End

elseif (mode="jumundivnextstep") then

	if (jumundiv = "1") and (ipkumdiv = "2") then
		'결제완료진행
	    nextjumundiv = "3"
	    nextipkumdiv = "4"
	    title = "결제완료 수기진행"

	    if (bookingYN = "N") then
		    nextjumundiv = "5"
		    nextipkumdiv = "8"
		    title = "결제완료 및 전송완료진행 수기진행"
	    end if
	elseif (jumundiv = "3") then
		'전송완료진행
	    nextjumundiv = "5"
	    nextipkumdiv = "8"
	    title = "전송완료 수기진행"
	elseif (jumundiv = "5") then
		'등록완료진행
	    nextjumundiv = "7"
	    nextipkumdiv = "8"
	    title = "등록완료 수기진행"
	else
        response.write "<script>alert('다음 상태로 진행 할 수 없습니다.');</script>"
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
        response.write "<script>alert('정상 주문건이 아니거나 취소된 내역입니다.')</script>"
        response.write "<script>history.back()</script>"
        dbget.close()	:	response.End
    end if

    ''On Error Resume Next
    ''dbget.beginTrans

    If Err.Number = 0 Then
        errcode = "001"

		if (nextjumundiv = "5") then
			'즉시전송

			'MMS발송
			set osms = new CSMSClass
			Call osms.sendGiftCardLMSMsg(giftOrderSerial)
			set osms=Nothing
			contents_jupsu = contents_jupsu & "SMS 전송완료(받는분HP:" + db2html(reqhp) + ")" & VbCrlf

			if (sendDiv = "E") then
				Call sendGiftCardEmail_SMTP(giftOrderSerial)
				contents_jupsu = contents_jupsu & "이메일 전송완료(받는분Email:" + db2html(reqemail) + ")" & VbCrlf
			end if

			sqlStr = " update [db_order].[dbo].tbl_giftcard_order" + vbCrlf
			sqlStr = sqlStr + " Set senddate=IsNull(senddate,getdate())" + vbCrlf
			sqlStr = sqlStr + " Where giftOrderSerial='" & CStr(giftorderserial) & "'"
			dbget.Execute sqlStr

			contents_finish = contents_jupsu
		end if

		if (nextjumundiv = "7") then

			dim GetLoginUserID
			GetLoginUserID = requserid					'카드사용자 아이디
			result = procGiftCardReg(masterCardCd)
			if (result = "W") or (result = "E") then

				dbget.RollBackTrans
				if (result = "W") then
					response.write "<script>alert('등록실패 : 기프트카드번호가 없거나 잘못된 코드입니다')</script>"
				else
					response.write "<script>alert('등록실패 : 유효기간이 만료된 카드입니다')</script>"
				end if
				'response.write "<script>history.back()</script>"
		        dbget.close()	:	response.End

			end if

			contents_jupsu = contents_jupsu & "Gift카드가 등록되었습니다." & VbCrlf
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

	    '' SMS 발송
        set osms = new CSMSClass

        if (jumundiv = "1") then
            osms.SendAcctIpkumOkMsg buyhp,giftorderserial
        end if
        set osms = Nothing

    end IF

	If (Err.Number = 0) and (emailok<>"") Then
        errcode = "004"

	    '' Email 발송

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
        response.write "<script>alert(" + Chr(34) + "데이타를 저장하는 도중에 에러가 발생하였습니다.\r\n(에러코드 : " + CStr(errcode) + "|" + Replace(Err.description,Vbcrlf," ") + ")" + Chr(34) + ")</script>"
        response.write "<script>history.back()</script>"
        dbget.close()	:	response.End
    End If


    response.write "<script>alert('수정 되었습니다.');</script>"
    response.write "<script>opener.location.reload();</script>"
    response.write "<script>opener.focus(); window.close();</script>"
    'dbget.close()	:	response.End

elseif (mode="resendcardcode") then

	if (iscreatenewcode = "Y") then
		title = "신규인증코드 전송"
	else
		title = "기존인증코드 재전송"
	end if


	if (jumundiv = "5") then
		'
	else
        response.write "<script>alert('전송완료 상태일때만 재전송이 가능합니다.');</script>"
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
        response.write "<script>alert('정상 주문건이 아니거나 취소된 내역입니다.')</script>"
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
				response.write "<script>alert('전송실패 : 취소되었거나 없는 카드입니다')</script>"
			elseif (result = "A") then
				response.write "<script>alert('전송실패 : 입금 대기중인 카드입니다')</script>"
			elseif (result = "R") then
				response.write "<script>alert('전송실패 : 주문취소된 카드입니다')</script>"
			elseif (result = "C") then
				response.write "<script>alert('전송실패 : 등록이 완료된 카드는 인증번호 재전송을 할 수 없습니다')</script>"
			elseif (result = "E") then
				response.write "<script>alert('전송실패 : 유효기간이 만료된 카드입니다')</script>"
			end if
			response.write "<script>history.back()</script>"
	        dbget.close()	:	response.End

		end if

		if (iscreatenewcode = "Y") then
			masterCardCd = getMasterCode(idx,16,resendCnt+1)

			'# 기프트카드 인증번호 발급 로그 저장
			Call putGiftCardMasterCDLog(giftOrderSerial,masterCardCd,resendCnt)

			contents_jupsu = contents_jupsu & "카드인증코드 신규 생성완료." & VbCrlf
		end if

		Call chgOrderInfoResendMasterCD(giftOrderSerial,masterCardCd)

		'MMS발송 /// 트랜잭션 안됨..
		set osms = new CSMSClass
		Call osms.sendGiftCardLMSMsg(giftOrderSerial)
		set osms = Nothing
		contents_jupsu = contents_jupsu & "SMS 전송완료(받는분HP:" + db2html(reqhp) + ")" & VbCrlf

		if (sendDiv = "E") then
			Call sendGiftCardEmail_SMTP(giftOrderSerial)
			contents_jupsu = contents_jupsu & "이메일 전송완료(받는분Email:" + db2html(reqemail) + ")" & VbCrlf
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
        response.write "<script>alert(" + Chr(34) + "데이타를 저장하는 도중에 에러가 발생하였습니다.\r\n(에러코드 : " + CStr(errcode) + "|" + Replace(Err.description,Vbcrlf," ") + ")" + Chr(34) + ")</script>"
        response.write "<script>history.back()</script>"
        dbget.close()	:	response.End
    End If


    response.write "<script>alert('수정 되었습니다.');</script>"
    response.write "<script>opener.location.reload();</script>"
    response.write "<script>opener.focus(); window.close();</script>"
    'dbget.close()	:	response.End

elseif (mode="jumundivprevstep") then

	if (jumundiv = "7") then
		'전송완료진행
	    prevjumundiv = "5"
	    title = "전송완료 수기전환"
	else
        response.write "<script>alert('이전 상태로 전환 할 수 없습니다.');</script>"
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
        response.write "<script>alert('정상 주문건이 아니거나 취소된 내역입니다.')</script>"
        response.write "<script>history.back()</script>"
        dbget.close()	:	response.End
    end if

response.write "수정중"
response.end

    On Error Resume Next
    dbget.beginTrans

    If Err.Number = 0 Then
        errcode = "001"

		'등록취소 가능한지 확인(등록취소 후 잔액이 0원 이상인지)


''		db_user.dbo.tbl_giftcard_current



        sqlStr =	"update db_order.dbo.tbl_giftcard_order " & vbCrlf
		sqlStr = sqlStr & " set jumundiv = '" + CStr(prevjumundiv) + "' " & vbCrlf
		sqlStr = sqlStr & " where giftorderserial='" & giftorderserial & "'"
		dbget.Execute sqlStr




'		delete from db_user.dbo.tbl_giftcard_regList
'		where
'
'1. tbl_giftcard_regList 삭제
'3. tbl_giftcard_log 로그저장
'2. tbl_giftcard_current 적립포인트 수정
'3.5 tbl_giftcard_order 주문상태 변경(7→5)
'4. 새 인증번호 발송







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
        response.write "<script>alert(" + Chr(34) + "데이타를 저장하는 도중에 에러가 발생하였습니다.\r\n(에러코드 : " + CStr(errcode) + "|" + Replace(Err.description,Vbcrlf," ") + ")" + Chr(34) + ")</script>"
        response.write "<script>history.back()</script>"
        dbget.close()	:	response.End
    End If


    response.write "<script>alert('수정 되었습니다.');</script>"
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
        call AddCsMemo(giftorderserial,"1",userid,reguserid,"PGKey 변경 : " & orgpaydateid & " ==&gt; " & paydateid)
    end if

	If Err.Number = 0 Then
        dbget.CommitTrans
    Else
        dbget.RollBackTrans
        response.write "<script>alert(" + Chr(34) + "데이타를 저장하는 도중에 에러가 발생하였습니다.\r\n(에러코드 : " + CStr(errcode) + "|" + Replace(Err.description,Vbcrlf," ") + ")" + Chr(34) + ")</script>"
        response.write "<script>history.back()</script>"
        dbget.close()	:	response.End
    End If

    response.write "<script>alert('수정 되었습니다.');</script>"
    response.write "<script>opener.location.reload();</script>"
    response.write "<script>opener.focus(); window.close();</script>"
    'dbget.close()	:	response.End

elseif (mode="modifymmsinfo") then

	title   = "MMS 정보 수정"
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
        contents_jupsu = contents_jupsu & "변경전 내역" & VbCrlf

        if (rsget("bookingYn") = "Y") then
	        if (Left(rsget("bookingDate"), 10) <> bookingDate) or (bookingDateHH*1 <> Hour(rsget("bookingDate"))) then
	            contents_jupsu = contents_jupsu & "예약일시: " & Left(rsget("bookingDate"), 10) & " " & Right(("0" & Hour(rsget("bookingDate"))), 2) & CNEXT & bookingDate & " " & bookingDateHH & VbCrlf
	        end if
        end if

        if (rsget("sendhp")<>sendhp) then
            contents_jupsu = contents_jupsu & "보내는분HP: " & rsget("sendhp") & CNEXT & sendhp & VbCrlf
        end if

        if (db2html(rsget("MMSTitle"))<>MMSTitle) then
            contents_jupsu = contents_jupsu & "MMS 제목: " & db2html(rsget("MMSTitle")) & CNEXT & MMSTitle & VbCrlf
        end if

        if (db2html(rsget("MMSContent"))<>MMSContent) then
            contents_jupsu = contents_jupsu & "MMS 내용: " & db2html(rsget("MMSContent")) & vbCrLf & CNEXT & vbCrLf & MMSContent & VbCrlf
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
        response.write "<script>alert(" + Chr(34) + "데이타를 저장하는 도중에 에러가 발생하였습니다.\r\n(에러코드 : " + CStr(errcode) + "|" + Replace(Err.description,Vbcrlf," ") + ")" + Chr(34) + ")</script>"
        response.write "<script>history.back()</script>"
        dbget.close()	:	response.End
    End If


    response.write "<script>alert('수정 되었습니다.');</script>"
    response.write "<script>opener.location.reload();</script>"
    response.write "<script>opener.focus(); window.close();</script>"
    'dbget.close()	:	response.End


elseif (mode="modifyemailinfo") then

	title   = "전송Email 정보 수정"
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
        contents_jupsu = contents_jupsu & "변경전 내역" & VbCrlf

        if (rsget("sendDiv")<>sendDiv) then
        	if (sendDiv = "E") then
        		contents_jupsu = contents_jupsu & "전송여부: " & "발송안함" & CNEXT & "동시전송" & VbCrlf
        	else
        		contents_jupsu = contents_jupsu & "전송여부: " & "발송안함" & CNEXT & "동시전송" & VbCrlf
        	end if
        end if

        if (rsget("sendemail")<>sendemail) then
            contents_jupsu = contents_jupsu & "보내는분Email: " & rsget("sendemail") & CNEXT & sendemail & VbCrlf
        end if

        if (rsget("reqEmail")<>reqEmail) then
            contents_jupsu = contents_jupsu & "받는분Email: " & rsget("reqEmail") & CNEXT & reqEmail & VbCrlf
        end if

        if (db2html(rsget("emailTitle"))<>emailTitle) then
            contents_jupsu = contents_jupsu & "Email제목: " & db2html(rsget("emailTitle")) & CNEXT & emailTitle & VbCrlf
        end if

        if (db2html(rsget("emailContent"))<>emailContent) then
            contents_jupsu = contents_jupsu & "Email내용: " & db2html(rsget("emailContent")) & vbCrLf & CNEXT & vbCrLf & emailContent & VbCrlf
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
        response.write "<script>alert(" + Chr(34) + "데이타를 저장하는 도중에 에러가 발생하였습니다.\r\n(에러코드 : " + CStr(errcode) + "|" + Replace(Err.description,Vbcrlf," ") + ")" + Chr(34) + ")</script>"
        response.write "<script>history.back()</script>"
        dbget.close()	:	response.End
    End If


    response.write "<script>alert('수정 되었습니다.');</script>"
    response.write "<script>opener.location.reload();</script>"
    response.write "<script>opener.focus(); window.close();</script>"
    'dbget.close()	:	response.End


elseif (mode="edithandmadereq") then

    ' 사용안하는듯
    response.end
    set myorderdetail = new COrderMaster
    myorderdetail.FRectgiftorderserial = giftorderserial
    myorderdetail.FRectDetailIdx = detailidx
    myorderdetail.GetOneOrderDetail


    ''갯수가 여러개 일때
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
        title       = "주문제작 상품 문구 수정"
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

            contents_jupsu = contents_jupsu & "변경 내역" & VbCrlf

            if (db2html(rsget("requiredetail"))<>requiredetail) then
                contents_jupsu = contents_jupsu & "상품명(옵션): " & db2html(rsget("itemname"))
                if (rsget("itemoptionname")<>"") then
                    contents_jupsu = contents_jupsu & "(" & db2html(rsget("itemoptionname")) & ")" & VbCrlf
                end if
                contents_jupsu = contents_jupsu & "문구: " & rsget("requiredetail") & VbCrlf & CNEXT & VbCrlf & requiredetail & VbCrlf
            end if

        end if
        rsget.Close

    end if

    if (contents_jupsu="") then
        response.write "<script language='javascript'>alert('수정하실 내역이 기존 내역과 일치합니다. 수정되지 않았습니다.');</script>"
        response.write "<script language='javascript'>history.back();</script>"
        dbget.close()	:	response.End
    else
        contents_jupsu = "변경 내역" & VbCrlf & contents_jupsu
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
        '' html2db 사용하지 말것.
        iAsID = RegCSMaster(divcd , giftorderserial, reguserid, title, contents_jupsu, gubun01, gubun02)
    end if

    If Err.Number = 0 Then
        errcode = "004"
        Call FinishCSMaster(iAsid, finishuser, html2db(contents_finish))

    end if

    If Err.Number = 0 Then
        dbget.CommitTrans

        response.write "<script language='javascript'>alert('변경 되었습니다.');</script>"
        response.write "<script language='javascript'>opener.location.reload();</script>"
        response.write "<script language='javascript'>window.close();</script>"
        'dbget.close()	:	response.End
    Else
        dbget.RollBackTrans
        response.write "<script>alert('데이타를 저장하는 도중에 에러가 발생하였습니다.\r\n(에러코드 : " + CStr(errcode) + ")')</script>"
        response.write "<script>history.back()</script>"
        dbget.close()	:	response.End
    End If
    On Error Goto 0
end if

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
