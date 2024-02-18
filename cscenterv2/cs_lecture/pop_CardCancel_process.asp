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

''주문 마스타
dim oordermaster
set oordermaster = new COrderMaster

if (ocsaslist.FResultCount>0) then
    oordermaster.FRectOrderSerial = ocsaslist.FOneItem.FOrderSerial
    oordermaster.QuickSearchOrderMaster
end if

if (oordermaster.FResultCount>0) then
    if (oordermaster.FOneItem.FCancelYn="N") and (oordermaster.FOneItem.Fjumundiv<>"9")  then
        response.write "<script>alert('반품주문건 이거나, 취소된 거래만 취소 가능합니다.');</script>"
        response.write "<script>window.close();</script>"
        dbget.close()	:	response.End
    end if
else
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

''=============전체취소만 가능함.. 부분취소등 취소안됨..=============
dim sqlStr, isSameMoney
dim t_refundrequire, t_MaybeOrgPayPrice
isSameMoney = false
sqlStr = " select r.refundrequire, "
''sqlStr = sqlStr & " (select sum(d.itemcost*d.itemno) from "&TABLE_ORDERDETAIL&" d where d.orderserial=l.orderserial and d.cancelyn<>'A')-(select tencardspend+miletotalprice+allatdiscountprice from "&TABLE_ORDERMASTER&" m where m.orderserial=l.orderserial) as MaybeOrgPayPrice"
sqlStr = sqlStr & " (select sum(d.itemcost*d.itemno) from "&TABLE_ORDERDETAIL&" d where d.orderserial=l.orderserial and d.cancelyn<>'A')-(select (CASE WHEN m.bcpnIDX is Not NULL THEN isNULL(tencardspend,0) ELSE 0 END)+miletotalprice+allatdiscountprice from "&TABLE_ORDERMASTER&" m where m.orderserial=l.orderserial) as MaybeOrgPayPrice"
sqlStr = sqlStr & " from "&TABLE_CSMASTER&" l"
sqlStr = sqlStr & " 	Join "&TABLE_CS_REFUND&" r"
sqlStr = sqlStr & " 	on l.id=r.asid"
sqlStr = sqlStr & " 	and r.returnmethod  in ('R100','R020','R400')"
sqlStr = sqlStr & " where l.id="&id
sqlStr = sqlStr & " and l.divcd='A007'"

rsget.Open sqlStr,dbget,1
if Not rsget.Eof then
    t_refundrequire=rsget("refundrequire")
    t_MaybeOrgPayPrice=rsget("MaybeOrgPayPrice")
    isSameMoney    = (t_refundrequire=Abs(t_MaybeOrgPayPrice))
end if
rsget.Close

IF  (Not isSameMoney) THEN
    IF (force="on") then
        response.write "취소금액과 원금액 상이<br><br>"
    ELSE
        if (IsAutoScript) then
            response.write "S_ERR|취소금액과 원금액 상이"
        else
            response.write "<script>alert('취소금액과 원금액 상이 - 관리자 문의 요망."&t_refundrequire&":"&t_MaybeOrgPayPrice&"');</script>"
            response.write "<script>window.close();</script>"
        end if
        dbget.close() : dbget_CS.close()	:	response.End
    End IF
END IF
'''=================================================================


'' Pg_Mid
dim MctID
MctID = Mid(orefund.FOneItem.FpaygateTid,11,10)
'' response.write MctID

dim INIpay, PInst
dim ResultCode, ResultMsg, CancelDate, CancelTime, Rcash_cancel_noappl


'############################################################## 핸드폰 결제 취소 ##############################################################
If orefund.FOneItem.Freturnmethod = "R400" Then

	Dim McashCancelObj, Mrchid, Svcid, Tradeid, Prdtprice, Mobilid, retval

	Set McashCancelObj = Server.CreateObject("Mcash_Cancel.Cancel.1")

	Mrchid      = "10030289"
	If Request("rdsite") = "mobile" Then
		Svcid       = "100302890002"
	Else
		Svcid       = "100302890001"
	End If
	Tradeid     = Split(orefund.FOneItem.FpaygateTid,"|")(0)
	Prdtprice   = orefund.FOneItem.Frefundrequire
	Mobilid     = Split(orefund.FOneItem.FpaygateTid,"|")(1)

	McashCancelObj.Mrchid			= Mrchid
	McashCancelObj.Svcid			= Svcid
	McashCancelObj.Tradeid			= Tradeid
	McashCancelObj.Prdtprice		= Prdtprice
	McashCancelObj.Mobilid	        = Mobilid

	retval = McashCancelObj.CancelData

	set McashCancelObj = nothing

	If retval = "0" Then
		ResultCode 	= "00"
		ResultMsg	= "정상처리"
	Else
		ResultCode = retval
		Select Case ResultCode
			Case "14"
				ResultMsg = "해지"
			Case "20"
				ResultMsg = "휴대폰 등록정보 오류(PG사) (LGT의 경우 사용자정보변경에 의한 인증실패)"
			Case "41"
				ResultMsg = "거래내역 미존재"
			Case "42"
				ResultMsg = "취소기간경과"
			Case "43"
				ResultMsg = "승인내역오류 ( 인증정보와의 불일치, 승인번호 유효시간 초과( 3분 ) )"
			Case "44"
				ResultMsg = "중복 취소 요청"
			Case "45"
				ResultMsg = "취소 요청 시 취소 정보 불일치"
			Case "97"
				ResultMsg = "요청자료 오류"
			Case "98"
				ResultMsg = "통신사 통신오류"
			Case "99"
				ResultMsg = "기타"
			Case Else
				ResultMsg = ""
		End Select
	End If

	CancelDate	= year(now) & "년 " & month(now) & "월 " & day(now) & "일"
	CancelTime	= hour(now) & "시 " & minute(now) & "분 " & second(now) & "초"
ELSEIF (oordermaster.FoneItem.FPgGubun="KP") then
    dim ret_amount, ret_panc_mod_mny, ret_panc_rem_mny ''부분취소 관련 파람.
    Call fnKCPCancelProc(True ,orefund.FOneItem.FpaygateTid,msg, "","",ResultCode,ResultMsg,CancelDate,CancelTime,ret_amount, ret_panc_mod_mny, ret_panc_rem_mny)

Else
'############################################################## 카드, 실시간 결제 취소 ##############################################################
		'###############################################################################
		'# 1. 객체 생성 #
		'################
		Set INIpay = Server.CreateObject("INItx41.INItx41.1")


		'###############################################################################
		'# 2. 인스턴스 초기화 #
		'######################
		PInst = INIpay.Initialize("")

		'###############################################################################
		'# 3. 거래 유형 설정 #
		'#####################
		INIpay.SetActionType CLng(PInst), "CANCEL"

		'###############################################################################
		'# 4. 정보 설정 #
		'################
		INIpay.SetField CLng(PInst), "pgid", "IniTechPG_" 'PG ID (고정)
		INIpay.SetField CLng(PInst), "spgip", "203.238.3.10" '예비 PG IP (고정)
		INIpay.SetField CLng(PInst), "mid", MctID '상점아이디
		INIpay.SetField CLng(PInst), "admin", "1111" '키패스워드(상점아이디에 따라 변경)
		INIpay.SetField CLng(PInst), "tid", Request("tid") '취소할 거래번호(TID)
		INIpay.SetField CLng(PInst), "msg", msg '취소 사유
		INIpay.SetField CLng(PInst), "uip", Request.ServerVariables("REMOTE_ADDR") 'IP
		INIpay.SetField CLng(PInst), "debug", "false" '로그모드("true"로 설정하면 상세한 로그를 남김)
		INIpay.SetField CLng(PInst), "merchantreserved", "예비" '예비

		'###############################################################################
		'# 5. 취소 요청 #
		'################
		INIpay.StartAction(CLng(PInst))

		'###############################################################################
		'# 6. 취소 결과 #
		'################
		ResultCode = INIpay.GetResult(CLng(PInst), "resultcode") '결과코드 ("00"이면 취소성공)
		ResultMsg  = INIpay.GetResult(CLng(PInst), "resultmsg") '결과내용
		CancelDate = INIpay.GetResult(CLng(PInst), "pgcanceldate") '이니시스 취소날짜
		CancelTime = INIpay.GetResult(CLng(PInst), "pgcanceltime") '이니시스 취소시각
		Rcash_cancel_noappl = INIpay.GetResult(CLng(PInst), "rcash_cancel_noappl") '현금영수증 취소 승인번호

		'###############################################################################
		'# 7. 인스턴스 해제 #
		'####################
		INIpay.Destroy CLng(PInst)
End If




dim returnmethod,refundrequire,refundresult,userid
dim iorderserial, ibuyhp
dim contents_finish

contents_finish = "결과 " & "[" & ResultCode & "]" & ResultMsg & VbCrlf
contents_finish = contents_finish & "취소일시 : " & CancelDate & " " & CancelTime & VbCrlf
contents_finish = contents_finish & "취소자 ID " & finishuserid

if (ResultCode="00") then

    sqlStr = "select r.*, a.userid, m.orderserial, m.buyhp from "
    sqlStr = sqlStr + " "&TABLE_CS_REFUND&" r,"
    sqlStr = sqlStr + " "&TABLE_CSMASTER&" a"
    sqlStr = sqlStr + "     left join "&TABLE_ORDERMASTER&" m "
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


    sqlStr = " update "&TABLE_CS_REFUND&""
    sqlStr = sqlStr + " set refundresult=" + CStr(refundrequire)
    sqlStr = sqlStr + " where asid=" + CStr(id)
    dbget.Execute sqlStr

    Call AddCustomerOpenContents(id, "환불(취소) 완료: " & CStr(refundrequire))


    Call FinishCSMaster(id, finishuserid, contents_finish)

    ''승인 취소 요청 SMS 발송
    if (iorderserial<>"") and (ibuyhp<>"") then
'        sqlStr = "Insert into [db_sms].[ismsuser].em_tran(tran_phone, tran_callback, tran_status, tran_date, tran_msg ) "
'    	sqlStr = sqlStr + " values('" + ibuyhp + "',"
'    	sqlStr = sqlStr + " '1644-6030',"
'    	sqlStr = sqlStr + " '1',"
'    	sqlStr = sqlStr + " getdate(),"
'    	sqlStr = sqlStr + " '[핑거스]승인 취소 되었습니다. 주문번호 : " + iorderserial + "')"

        ''2015/10/16
    	sqlStr = "insert into [SMSDB].[db_infoSMS].dbo.em_smt_tran (date_client_req, content, callback, service_type, broadcast_yn, msg_status,recipient_num) "
	    sqlStr = sqlStr + " values(getdate(),'[핑거스]승인 취소 되었습니다. 주문번호 : " + iorderserial + "','027419070','0','N','1','"+ibuyhp+"')"

        dbget_CS.Execute sqlStr
    end if

    ''메일 :: 일단재낌.
    Call SendCsActionMail(id)

    if (IsAutoScript) then
        response.write "S_OK"
    else
        response.write "<script>alert('" & ResultMsg & "');</script>"
        response.write "<script>opener.location.reload();</script>"
        response.write "<script>window.close();</script>"
    end if
    dbget.close() : dbget_CS.close()	:	response.End

else
    if (IsAutoScript) then
        response.write "S_ERR|"&ResultMsg
    else
        response.write ResultCode & "<br>"
        response.write ResultMsg & "<br>"
        response.write CancelDate & "<br>"
        response.write CancelTime & "<br>"
        response.write Rcash_cancel_noappl & "<br>"
    end if
end if
%>



<%
set oordermaster = Nothing
set ocsaslist = Nothing
set orefund = Nothing
%>
<!-- #include virtual="/cscenterv2/lib/db/dbclose.asp" -->
