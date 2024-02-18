<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_aslistcls.asp" -->
<!-- #include virtual="/cscenter/lib/csAsfunction.asp" -->
<!-- #include virtual="/cscenter/lib/cs_action_mail_Function.asp"-->
<!-- #include virtual="/lib/email/MailLib2.asp"-->

<%
'''신용카드 부분취소 R120 => 다른 페이지에서 따로 처리.

dim id, finishuserid, msg, force
dim orgOrderSerial
id           = RequestCheckVar(request("id"),10)
msg          = RequestCheckVar(request("msg"),50)
finishuserid = session("ssBctID")
force = RequestCheckVar(request("force"),10)

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
    dbget.close()	:	response.End
end if

if (ocsaslist.FOneItem.FCurrstate<>"B001") then
    if (IsAutoScript) then
        response.write "S_ERR|접수 상태가 아닙니다."
    else
        response.write "<script>alert('접수 상태가 아닙니다.');</script>"
        response.write "<script>window.close();</script>"
    end if
    dbget.close()	:	response.End
end if

'' 신용카드 취소만 가능
'if (orefund.FOneItem.Freturnmethod<>"R100") then
'    response.write "<script>alert('현재 신용카드 거래만 취소 가능합니다.');</script>"
'    response.write "<script>window.close();</script>"
'    dbget.close()	:	response.End
'end if

Dim returnmethod, IsCardPartialCancel
returnmethod = orefund.FOneItem.Freturnmethod

if Not ((returnmethod="R100") or (returnmethod="R020") or (returnmethod="R400")) Then
    if (IsAutoScript) then
        response.write "S_ERR|신용카드 전체취소, 실시간이체 취소, 휴대폰 전체 취소 만 가능."
    else
        response.write "<script>alert('신용카드 전체취소, 실시간이체 취소, 휴대폰 전체 취소 만 가능합니다.');</script>"
        response.write "<script>window.close();</script>"
    end if
    dbget.close()	:	response.End
end if


'' IniPay 만 취소만 가능
if (Left(orefund.FOneItem.FpaygateTid,10)<>"IniTechPG_") AND Left(orefund.FOneItem.FpaygateTid,6)<>"Stdpay" AND Left(orefund.FOneItem.FpaygateTid,6)<>"INIMX_" AND Left(orefund.FOneItem.FpaygateTid,10)<>"INIAPICARD" AND orefund.FOneItem.Freturnmethod<>"R400" then
    if (IsAutoScript) then
        response.write "S_ERR|이니시스 거래만 취소 가능합니다."
    else
        response.write "<script>alert('이니시스 거래만 취소 가능합니다.');</script>"
        response.write "<script>window.close();</script>"
    end if
    dbget.close()	:	response.End
end if

''=============전체취소만 가능함.. 부분취소등 취소안됨..=============
dim sqlStr, isSameMoney
dim t_refundrequire, t_MaybeOrgPayPrice
isSameMoney = false

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




dim refundrequire,refundresult,userid
dim iorderserial, ibuyhp
dim contents_finish

contents_finish = "결과 " & "[" & ResultCode & "]" & ResultMsg & VbCrlf
contents_finish = contents_finish & "취소일시 : " & CancelDate & " " & CancelTime & VbCrlf
contents_finish = contents_finish & "취소자 ID " & finishuserid

if (ResultCode="00") then

    sqlStr = "select r.*, a.userid, m.giftorderserial, m.buyhp from "
    sqlStr = sqlStr + " [db_cs].[dbo].tbl_as_refund_info r,"
    sqlStr = sqlStr + " [db_cs].dbo.tbl_new_as_list a"
    sqlStr = sqlStr + "     left join [db_order].[dbo].tbl_giftcard_order m "
	sqlStr = sqlStr + "     on a.orderserial=m.giftorderserial"
    sqlStr = sqlStr + " where r.asid=" + CStr(id)
    sqlStr = sqlStr + " and r.asid=a.id"

    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        returnmethod    = rsget("returnmethod")
        refundrequire   = rsget("refundrequire")
        refundresult    = rsget("refundresult")
        userid          = rsget("userid")
        iorderserial    = rsget("giftorderserial")
        ibuyhp          = rsget("buyhp")
    end if
    rsget.Close


    sqlStr = " update [db_cs].[dbo].tbl_as_refund_info"
    sqlStr = sqlStr + " set refundresult=" + CStr(refundrequire)
    sqlStr = sqlStr + " where asid=" + CStr(id)
    dbget.Execute sqlStr

    Call AddCustomerOpenContents(id, "환불(취소) 완료: " & CStr(refundrequire))


    Call FinishCSMaster(id, finishuserid, contents_finish)

    ''승인 취소 요청 SMS 발송
    if (iorderserial<>"") and (ibuyhp<>"") then
        sqlStr = "Insert into [db_sms].[ismsuser].em_tran(tran_phone, tran_callback, tran_status, tran_date, tran_msg ) "
    	sqlStr = sqlStr + " values('" + ibuyhp + "',"
    	sqlStr = sqlStr + " '1644-6030',"
    	sqlStr = sqlStr + " '1',"
    	sqlStr = sqlStr + " getdate(),"
    	sqlStr = sqlStr + " '[텐바이텐]승인 취소 되었습니다. 주문번호 : " + iorderserial + "')"
        dbget.Execute sqlStr
    end if

    ''메일
    Call SendCsActionMail_GiftCard(id)

    if (IsAutoScript) then
        response.write "S_OK"
    else
        response.write "<script>alert('" & ResultMsg & "');</script>"
        response.write "<script>opener.location.reload();</script>"
        response.write "<script>window.close();</script>"
    end if
    dbget.close()	:	response.End

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
set ocsaslist = Nothing
set orefund = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
