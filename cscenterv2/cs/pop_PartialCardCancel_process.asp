<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/cscenterv2/lib/incSessionAdminCS.asp" -->
<!-- #include virtual="/cscenterv2/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/cscenterv2/lib/function.asp"-->
<!-- #include virtual="/cscenterv2/lib/classes/cs/cs_aslistcls.asp"-->
<!-- #include virtual="/cscenterv2/lib/csAsfunction.asp" -->
<!-- #include virtual="/cscenterv2/lib/cs_action_mail_Function.asp"-->
<!-- #include virtual="/lib/email/MailLib2.asp"-->
<!-- #include virtual="/cscenterv2/lib/classes/order/ordercls.asp"-->
<!-- #include virtual="/cscenterv2/lib/KCP/site_conf_inc.asp" -->
<!-- #include virtual="/cscenterv2/lib/KCP/pp_cli_hub_lib.asp" -->
<%
''
''

dim id, finishuserid, msg, force
id           = RequestCheckVar(request("id"),10)
msg          = RequestCheckVar(request("msg"),50)
finishuserid = session("ssBctID")
force = RequestCheckVar(request("force"),10)

if msg <> "" then
	if checkNotValidHTML(msg) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end if
if (msg="") and (IsAutoScript) then msg="배송전취소"

dim orderserial, returnmethod
dim orgOrderserial, jumundiv, pggubun
dim sqlStr

dim ocsaslist
set ocsaslist = New CCSASList
ocsaslist.FRectCsAsID = id

if (id<>"") then
    ocsaslist.GetOneCSASMaster
end if


dim orefund
set orefund = New CCSASList
orefund.FRectCsAsID = id

if (id<>"") then
    orefund.GetOneRefundInfo
end if

if (ocsaslist.FResultCount<1) or (orefund.FResultCount<1) then
    if (IsAutoScript) then
        response.write "S_ERR|환불내역이 없거나 유효하지 않은 내역입니다."
    else
        response.write "<script>alert('환불내역이 없거나 유효하지 않은 내역입니다.');</script>"
        response.write "<script>window.close();</script>"
    end if
    dbget.close() : dbget_CS.close()	:	response.End
end if

if (ocsaslist.FOneItem.FCurrstate<>"B001") then
    if (IsAutoScript) then
        response.write "S_ERR|접수 상태가 아닙니다."
    else
        response.write "<script>alert('접수 상태가 아닙니다.');</script>"
        response.write "<script>window.close();</script>"
    end if
    dbget.close() : dbget_CS.close()	:	response.End
end if


orderserial = ocsaslist.FOneItem.FOrderserial
returnmethod = orefund.FOneItem.Freturnmethod

if (returnmethod<>"R120") and (returnmethod<>"R420") and (returnmethod<>"R022") Then
    if (IsAutoScript) then
        response.write "S_ERR|신용카드, 실시간이체 부분취소만 가능합니다."
    else
        response.write "<script>alert('신용카드, 실시간이체 부분취소만 가능합니다.');</script>"
        response.write "<script>window.close();</script>"
    end if
    dbget.close()	:	dbget_CS.close() : response.End
end if

''주문 마스타
dim oordermaster
set oordermaster = new COrderMaster

if (ocsaslist.FResultCount>0) then
    oordermaster.FRectOrderSerial = ocsaslist.FOneItem.FOrderSerial
    oordermaster.QuickSearchOrderMaster
end if

if (oordermaster.FResultCount<1) then
    response.write "<script>alert('올바른 주문건이 아닙니다..');</script>"
    response.write "<script>window.close();</script>"
    dbget.close()	:	response.End
end if


'' IniPay 만 취소만 가능 => KCP도 가능
if (Left(orefund.FOneItem.FpaygateTid,10)<>"IniTechPG_") AND (orefund.FOneItem.Freturnmethod<>"R400") AND (oordermaster.FoneItem.FPgGubun<>"KP") then
    if (IsAutoScript) then
        response.write "S_ERR|이니시스, KCP 거래만 취소 가능합니다."
    else
        response.write "<script>alert('이니시스, KCP 거래만 취소 가능합니다.');</script>"
        response.write "<script>window.close();</script>"
    end if
    dbget.close()	: dbget_CS.close() :	response.End
end if


'==============================================================================
'최초 주결제수단금액
dim accountdiv : accountdiv ="100"
dim omainpayment, mainpaymentorg
dim cardcancelok, cardcancelerrormsg, cardcancelcount, cardcancelsum, cardcode

if (returnmethod = "R400") or (returnmethod = "R420") then
	accountdiv = "400"
elseif (returnmethod = "R022") then  ''2016/06/21 추가.
    accountdiv = "20"
end if


set omainpayment = new COrderMaster

mainpaymentorg = 0			'// 최초 결제금액
cardcancelok = "N"
cardcancelerrormsg = ""
cardcancelcount = 0
cardcancelsum   = 0			'// 취소금액합계
cardcode = ""

if (orderserial<>"") then
	omainpayment.FRectOrderSerial = orderserial

	if ((accountdiv = "100") or (accountdiv = "20")) then ''네이버페이 실시간 이체 부분취소 가능
		Call omainpayment.getMainPaymentInfo(accountdiv, mainpaymentorg, cardcancelok, cardcancelerrormsg, cardcancelcount,cardcancelsum, cardcode)
	else
		'' Call omainpayment.getMainPaymentInfoPhone(accountdiv, mainpaymentorg, cardcancelok, cardcancelerrormsg, cardcancelcount,cardcancelsum, cardcode)
	end if

	orgOrderserial = orderserial

	'// 교환주문( jumundiv = 6 )이면 원주문에서 결제정보 가져온다.
	sqlStr = " select top 1 m.jumundiv, IsNull(m.pggubun,'') as pggubun "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_academy.dbo.tbl_academy_order_master m "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and m.orderserial = '" & orderserial & "' "
	rsget.Open sqlStr,dbget,1
	if Not rsget.Eof then
		jumundiv = rsget("jumundiv")
		pggubun = rsget("pggubun")
	end if
	rsget.close

	if (jumundiv = "6") then
		sqlStr = " select top 1 c.orgorderserial "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	db_academy.dbo.tbl_academy_change_order c "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + " 	and c.chgorderserial = '" & orderserial & "' "
		rsget.Open sqlStr,dbget,1
		if Not rsget.Eof then
			orgOrderserial = rsget("orgorderserial")
		end if
		rsget.close
	end if
end if


set omainpayment = nothing



IF (cardcancelok<>"Y") then
    if (IsAutoScript) then
        response.write "S_ERR|"&cardcancelerrormsg
    else
        response.write cardcancelerrormsg
        response.write "<script>alert('"&TRIM(cardcancelerrormsg)&"');</script>"
        response.write "<script>window.close();</script>"
    end if
    dbget.close()	:	response.End
ENd IF


''''원거래ID,취소금액    , 재승인금액,  구매자email
dim ioldtid, iCancelPrice, iconfirm_price

ioldtid        = orefund.FOneItem.FpaygateTid
iCancelPrice   = orefund.FOneItem.Frefundrequire
iconfirm_price = (mainpaymentorg-cardcancelsum) - orefund.FOneItem.Frefundrequire

'RW iCancelPrice
'RW iconfirm_price
'RW mainpaymentorg
'RW cardcancelsum


'''결과 값.
dim Tid
dim OldTid, CancelPrice, RepayPrice, CntRepay


IF (ioldtid="") or (iCancelPrice<1) or (iconfirm_price<0) THEN
    if (IsAutoScript) then
        response.write "S_ERR|부분 취소금액 오류 또는 TID오류"
    else
        response.write "<script>alert('부분 취소금액 오류 또는 TID오류"&iCancelPrice&"."&iconfirm_price&"');</script>"
        ''response.write "<script>window.close();</script>"
    end if
    dbget.close()	:	response.End
END IF


''response.end

''통신전 저장======================================================================
dim ResultCode, ResultMsg
dim CancelDate, CancelTime
Dim clogIdx

sqlStr = "select * from [db_academy].[dbo].tbl_academy_card_cancel_log where 1=0"
rsget.Open sqlStr,dbget,1,3
    rsget.AddNew
    rsget("orderserial") = orgOrderserial
    rsget("orgtid")      = ioldtid
    rsget("cancelprice") = iCancelPrice
    rsget("repayprice")  = iconfirm_price
    rsget("usermail")    = "" ''buyemail

    rsget.update
	clogIdx = rsget("clogIdx")
rsget.close


''KCPPAY
Dim IsKCPPya   : IsKCPPya = (pggubun = "KP")
''카카오페이
Dim IsKakaoPay : IsKakaoPay = (pggubun = "KA")
''네이버페이
Dim iprimaryPayRestAmount, inpointRestAmount, itotalRestAmount
Dim IsNaverPay : IsNaverPay = (pggubun = "NP")

if (IsKakaoPay) then
    response.write "<script>alert('지원되지 않는거래 pggubun:"&pggubun&"."&iCancelPrice&"."&iconfirm_price&"');</script>"
    dbget.close()	:	response.End

'	CALL PartialCancelKakaoPay(orefund.FOneItem.FpaygateTid, (mainpaymentorg-cardcancelsum),orefund.FOneItem.Frefundrequire,"",retval,ResultCode,ResultMsg,CancelDate,CancelTime, Tid)
'
'	'// 리턴값중 CcPartCl 일단 무시(0:부분취소불가, 1:부분취소가능)
'	CntRepay = "2"
'
'	RepayPrice = iconfirm_price
'	CancelPrice = orefund.FOneItem.Frefundrequire
'
'	OldTid = ioldtid
elseif (IsNaverPay) then
    response.write "<script>alert('지원되지 않는거래 pggubun:"&pggubun&"."&iCancelPrice&"."&iconfirm_price&"');</script>"
    dbget.close()	:	response.End
'    dim nPayCancelRequester : nPayCancelRequester = fnGetNPayCancelRequester(id)
'    CALL PartialCancelNaverPay(orefund.FOneItem.FpaygateTid, orefund.FOneItem.Frefundrequire,"",retval,ResultCode,ResultMsg,CancelDate,CancelTime, Tid, iprimaryPayRestAmount, inpointRestAmount, itotalRestAmount, nPayCancelRequester)
'
'	'// 리턴값중 CcPartCl 일단 무시(0:부분취소불가, 1:부분취소가능)
'	CntRepay = "2"
'
'	RepayPrice = iconfirm_price
'	CancelPrice = orefund.FOneItem.Frefundrequire
'
'	OldTid = ioldtid
elseif (IsKCPPya) then
    dim ret_amount, ret_panc_mod_mny, ret_panc_rem_mny
    Call fnKCPCancelProc(False ,ioldtid,msg, iCancelPrice,(mainpaymentorg-cardcancelsum),ResultCode,ResultMsg,CancelDate,CancelTime,ret_amount, ret_panc_mod_mny, ret_panc_rem_mny)

    CancelPrice = ret_panc_mod_mny
    RepayPrice = ret_panc_rem_mny
    OldTid = ioldtid
else
    response.write "<script>alert('지원되지 않는거래 pggubun:"&pggubun&"."&iCancelPrice&"."&iconfirm_price&"');</script>"
    dbget.close()	:	response.End
end if


if (CntRepay="") then CntRepay=1 ''2013/06/27 추가 :: ''키워드 'where' 근처의 구문이 잘못되었습니다.

sqlStr = "update [db_academy].[dbo].tbl_academy_card_cancel_log"&VbCRLF
sqlStr = sqlStr & " set newtid='"&Tid&"'"&VbCRLF
sqlStr = sqlStr & " ,resultcode='"&ResultCode&"'"&VbCRLF
sqlStr = sqlStr & " ,resultmsg='"&Replace(ResultMsg,"'","")&"'"&VbCRLF
sqlStr = sqlStr & " ,cancelrequestcount="&CntRepay&VbCRLF
sqlStr = sqlStr & " where clogIdx="&clogIdx&""

dbget.Execute sqlStr


''부분취소 정상처리시.tbl_order_paymentETc 업데이트.
if (ResultCode="00") or (ResultCode="0000") or (IsKakaoPay and (resultCode="2001")) then
    IF (IsNaverPay) then
        sqlStr = " update db_academy.dbo.tbl_academy_order_paymentETc"&VbCRLF
        sqlStr = sqlStr & " set realPayedSum="&iprimaryPayRestAmount&VbCRLF
        sqlStr = sqlStr & " where orderserial='"&orgOrderserial&"'"&VbCRLF
        sqlStr = sqlStr & " and acctdiv='"&accountdiv&"'"&VbCRLF

        dbget.Execute sqlStr

        sqlStr = " update db_academy.dbo.tbl_academy_order_paymentETc"&VbCRLF
        sqlStr = sqlStr & " set realPayedSum="&inpointRestAmount&VbCRLF
        sqlStr = sqlStr & " where orderserial='"&orgOrderserial&"'"&VbCRLF
        sqlStr = sqlStr & " and acctdiv='120'"&VbCRLF

        dbget.Execute sqlStr

    ELSE
        sqlStr = " update db_academy.dbo.tbl_academy_order_paymentETc"&VbCRLF
        sqlStr = sqlStr & " set realPayedSum=realPayedSum-"&CancelPrice&VbCRLF
        sqlStr = sqlStr & " where orderserial='"&orgOrderserial&"'"&VbCRLF
        sqlStr = sqlStr & " and acctdiv='"&accountdiv&"'"&VbCRLF

        dbget.Execute sqlStr
    END IF
end if

rw "Tid="&Tid  ''부분취소 TID
rw "ResultCode="&ResultCode
rw "ResultMsg="&ResultMsg
rw "OldTid="&OldTid
rw "CancelPrice="&CancelPrice
rw "RepayPrice="&RepayPrice
rw "CntRepay="&CntRepay

set oordermaster = Nothing
set ocsaslist = Nothing
set orefund = Nothing



dim refundrequire,refundresult,userid
dim iorderserial, ibuyhp
dim contents_finish

contents_finish = "결과 " & "[" & ResultCode & "]" & ResultMsg & VbCrlf
if (ResultCode="00") or (ResultCode="0000") or (IsKakaoPay and (resultCode="2001")) then
    IF (IsNaverPay) then
        contents_finish = contents_finish & "취소금액 : " &FormatNumber(CancelPrice,0) & VbCrlf
        contents_finish = contents_finish & "남은금액 : " &FormatNumber(itotalRestAmount,0) & VbCrlf
        ''contents_finish = contents_finish & "취소일시 : " & CancelDate & " " & CancelTime & VbCrlf
        contents_finish = contents_finish & "취소자 ID " & finishuserid
    else
        contents_finish = contents_finish & "취소금액 : " &FormatNumber(CancelPrice,0) & VbCrlf
        contents_finish = contents_finish & "남은금액 : " &FormatNumber(RepayPrice,0) & VbCrlf
        ''contents_finish = contents_finish & "취소일시 : " & CancelDate & " " & CancelTime & VbCrlf
        contents_finish = contents_finish & "취소자 ID " & finishuserid
    end if
end if

if (ResultCode="00") or (ResultCode="0000") or (IsKakaoPay and (resultCode="2001")) then
    sqlStr = "select r.*, a.userid, m.orderserial, m.buyhp from "
    sqlStr = sqlStr + " [db_academy].[dbo].tbl_academy_as_refund_info r,"
    sqlStr = sqlStr + " [db_academy].dbo.tbl_academy_as_list a"
    sqlStr = sqlStr + "     left join db_academy.dbo.tbl_academy_order_master m "
	sqlStr = sqlStr + "     on a.orderserial=m.orderserial"
    sqlStr = sqlStr + " where r.asid=" + CStr(id)
    sqlStr = sqlStr + " and r.asid=a.id"

    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        returnmethod    = rsget("returnmethod")
        refundrequire   = rsget("refundrequire")
        refundresult    = rsget("refundresult")
        userid          = rsget("userid")
        iorderserial    = rsget("orderserial")
        ibuyhp          = rsget("buyhp")
    end if
    rsget.Close


    sqlStr = " update [db_academy].[dbo].tbl_academy_as_refund_info"
    sqlStr = sqlStr + " set refundresult=" + CStr(refundrequire)
    sqlStr = sqlStr + " where asid=" + CStr(id)
    dbget.Execute sqlStr

    Call AddCustomerOpenContents(id, "부분 취소 완료: " & FormatNumber(CancelPrice,0) & VbCRLF & "남은 승인 금액: "& FormatNumber(RepayPrice,0)) '''CStr(refundrequire))


    Call FinishCSMaster(id, finishuserid, contents_finish)

    ''승인 취소 요청 SMS 발송
    if (iorderserial<>"") and (ibuyhp<>"") then
        'sqlStr = "Insert into [db_sms].[ismsuser].em_tran(tran_phone, tran_callback, tran_status, tran_date, tran_msg ) "
    	'sqlStr = sqlStr + " values('" + ibuyhp + "',"
    	'sqlStr = sqlStr + " '1644-6030',"
    	'sqlStr = sqlStr + " '1',"
    	'sqlStr = sqlStr + " getdate(),"
    	'sqlStr = sqlStr + " '[텐바이텐]승인 부분 취소 되었습니다. 승인취소액 : " + FormatNumber(CancelPrice,0) + " 주문번호 : " + iorderserial + "')"

    	''sqlStr = " exec [db_sms].[dbo].[usp_SendSMS] '"+ibuyhp+"','1644-1557','[더 핑거스]승인 부분 취소 되었습니다. 승인취소액 : " + FormatNumber(CancelPrice,0) + " 주문번호 : " + iorderserial + "'"
    	
    	sqlStr = "insert into [SMSDB].[db_infoSMS].dbo.em_smt_tran (date_client_req, content, callback, service_type, broadcast_yn, msg_status,recipient_num) "
	    sqlStr = sqlStr + " values(getdate(),'[더 핑거스]승인 부분 취소 되었습니다. 주문번호 : " + iorderserial + "','1644-1557','0','N','1','"+ibuyhp+"')"
        dbget_CS.Execute sqlStr
    end if

    ''메일
    Call SendCsActionMail(id)

    if (IsAutoScript) then
        response.write "S_OK"
    else
        response.write "<script>alert('" & TRIM(ResultMsg) & "');</script>"
        'response.write "<script>opener.location.reload();</script>"
        'response.write "<script>window.close();</script>"
    end if
    dbget.close()	:	response.End

else
    if (IsAutoScript) then
        response.write "S_ERR|"&ResultMsg
    else
        response.write ResultCode & "<br>"
        response.write ResultMsg & "<br>"

		response.write "<br><br>* <font color=red>반복적으로 부분취소실패 오류</font>가 발생하는 경우 시스템팀 문의요망<br>(중복 취소일 수 있습니다.)"

    end if
end if

%>

<!-- #include virtual="/cscenterv2/lib/db/dbclose.asp" -->
