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
<!-- #include virtual="/cscenter/lib/giftiConCls.asp"-->
<%
response.write "T"
response.end

dim pinNo : pinNo="999693381018"
dim refundrequire : refundrequire = 10000

dim oGicon
	dim ret, bufStr, ResultCode, ResultMsg, CancelDate, CancelTime

	set oGicon = new CGiftiCon
	ret = oGicon.reqCouponCancel(CStr(pinNo), "100100", refundrequire) ''������ȣ, ������ȣ, ��ǰ ��ȯ��

	if (ret) then
		ResultCode = oGicon.FConResult.getResultCode
		ResultMsg = oGicon.FConResult.FMESSAGE

		CancelDate	= year(now) & "�� " & month(now) & "�� " & day(now) & "��"
		CancelTime	= hour(now) & "�� " & minute(now) & "�� " & second(now) & "��"

	    ''bufStr =          "SERVICE_CODE:" & oGicon.FConResult.FSERVICE_CODE & VbCRLF
	    ''bufStr = bufStr & "COUPON_NUMBER:" & oGicon.FConResult.FCOUPON_NUMBER & VbCRLF
	    ''bufStr = bufStr & "ERROR_CODE:" & oGicon.FConResult.getResultCode & VbCRLF
	    ''bufStr = bufStr & "MESSAGE:" & oGicon.FConResult.FMESSAGE & VbCRLF
	    ''bufStr = bufStr & "EXCHANGE_COUNT:" & oGicon.FConResult.FEXCHANGE_COUNT & VbCRLF

	    ''bufStr = bufStr & "ApprovNO:" & oGicon.FConResult.FApprovNO & VbCRLF
	    ''bufStr = bufStr & "ExchangePrice:" & oGicon.FConResult.FExchangePrice & VbCRLF

	end if
	set oGicon = Nothing
	
	response.write CancelTime
	response.write CancelDate
%>
	