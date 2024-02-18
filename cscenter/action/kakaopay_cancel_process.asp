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
<!-- #include virtual="/lib/email/smsLib.asp"-->
<!-- #include virtual="/cscenter/action/incKakaopayCommonNew.asp"-->
<%
''KaKao 신용카드 취소
function CanCelNewKakaoPay(ipaygatetid, irefundrequire, irdSite, byref iretval, byref iResultCode, byref iResultMsg, byref iCancelDate, byref iCancelTime)
    Dim KPay_Result, cancelYmdt, Status

    Set KPay_Result = fnCallKakaoPayCancel(ipaygatetid, irefundrequire, Status)
    'response.write KPay_Result.code
    'response.end
    if Status = "200" then
        iResultCode = "00"                                  ''00 사용할것.
        cancelYmdt = KPay_Result.canceled_at                 ''결제 취소 시각 
        iCancelDate = LEFT(cancelYmdt,10)
        iCancelTime = RIGHT(cancelYmdt,8)
        iResultMsg = KPay_Result.status                        ''결제상태값
    else
        iResultCode = KPay_Result.code                        ''실패코드
        iResultMsg = KPay_Result.message                     ''실패 메세지
    end if

    Set KPay_Result = Nothing

end function

dim paygateTid, refundrequire
dim retVal, ResultCode, ResultMsg, CancelDate, CancelTime
paygateTid           = RequestCheckVar(request("paygateTid"),20)
refundrequire          = RequestCheckVar(request("refundrequire"),10)

CALL CanCelNewKakaoPay(paygateTid, refundrequire, "", retval, ResultCode, ResultMsg, CancelDate, CancelTime)

response.write "S_OK" & "<br>"
response.write ResultCode & "<br>"
response.write ResultMsg & "<br>"
response.write CancelDate & "<br>"
response.write CancelTime & "<br>"
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->