<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/cscenter/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/tenEncUtil.asp" -->
<!-- #include virtual="/lib/util/base64unicode.asp" -->
<!-- #include virtual="/lib/classes/cscenter/cs_aslistcls.asp" -->
<!-- #include virtual="/lib/classes/order/new_ordercls.asp"-->
<!-- #include virtual="/cscenter/lib/csAsfunction.asp"-->
<!-- #include virtual="/cscenter/lib/cs_action_mail_Function.asp"-->
<!-- #include virtual="/lib/email/MailLib2.asp"-->
<!-- #include virtual="/lib/util/DcCyberAcctUtil.asp"-->
<!-- #include virtual="/lib/classes/cscenter/sp_tenGiftCardCls.asp" -->

<%

dim divcd, id, reguserid, ipkumdiv
dim title, gubun01, gubun02, contents_jupsu

''취소 관련
dim refundmileagesum, refundcouponsum, allatsubtractsum
dim refunditemcostsum, canceltotal, nextsubtotal
dim refundbeasongpay, remainbeasongpay, refunddeliverypay, refundadjustpay
dim remainitemcostsum
dim refundgiftcardsum, refunddepositsum

''환불 관련 maybe (refundrequire==canceltotal)
dim refundrequire, returnmethod
dim rebankname, rebankaccount, rebankownername, paygateTid, encmethod

''원주문 금액
dim orgitemcostsum, orgbeasongpay, orgmileagesum, orgcouponsum, orgallatdiscountsum, orgsubtotalprice, orggiftcardsum, orgdepositsum

dim ScanErr, errcode


dim userid, orderserial, currentCash
dim sqlStr

userid      	= request("userid")
orderserial 	= request("orderserial")
refundrequire 	= request("refundrequire")



'==============================================================================
dim oTenGiftCard

set oTenGiftCard = new CTenGiftCard

oTenGiftCard.FRectUserID = userid

currentCash = 0
if (userid<>"") then
    oTenGiftCard.getUserCurrentTenGiftCard

    currentCash = oTenGiftCard.FcurrentCash
end if



'==============================================================================
if (userid = "") then
	Response.Write "<script>alert('아이디가 없습니다.');</script>"
	dbget.close()
	Response.End
end if

if (orderserial = "") and (userid<>"danbi2612") and (userid<>"setjddms") and (userid<>"dadareda") then
	Response.Write "<script>alert('주문번호가 없습니다.');</script>"
	dbget.close()
	Response.End
end if

if (CLng(FormatNumber((100*oTenGiftCard.FspendCash/oTenGiftCard.FgainCash),0)) < 60) and (userid<>"danbi2612") and (userid<>"setjddms") and (userid<>"dadareda") and (userid<>"eiddr0705") then
	Response.Write "<script>alert('Gift카드사용비율( = 상품구매총액/등록총액) 이 60% 이상인 경우만 잔액의 예치금전환이 가능합니다.');</script>"
	dbget.close()
	Response.End
end if

if (currentCash*1 < refundrequire*1) then
	Response.Write "<script>alert('Gift카드 잔액보다 예치금전환액이 더 큽니다.');</script>"
	dbget.close()
	Response.End
end if



'==============================================================================
reguserid   = session("ssbctid")

divcd 		= "A003"
title 		= "Gift카드 무통장 환불"

gubun01 	= "C004"
gubun02 	= "CD99"

returnmethod = "R007"

contents_jupsu = ""

orgsubtotalprice 	= 0
orgitemcostsum 		= 0
orgbeasongpay 		= 0
orgmileagesum 		= 0
orgcouponsum 		= 0
orgallatdiscountsum = 0
canceltotal 		= 0
refunditemcostsum 	= 0
refundmileagesum 	= 0
refundcouponsum 	= 0
allatsubtractsum 	= 0
refundbeasongpay 	= 0
refunddeliverypay 	= 0
refundadjustpay 	= 0
rebankname 			= html2db(request("rebankname"))
rebankaccount 		= html2db(request("rebankaccount"))
rebankownername 	= html2db(request("rebankownername"))
paygateTid 			= "0"

orggiftcardsum		= 0
orgdepositsum		= 0
refundgiftcardsum	= 0
refunddepositsum	= 0


On Error Resume Next
    dbget.beginTrans

    If (Err.Number = 0) and (ScanErr="") Then
        errcode = "001"

        '' CS Master 접수
        id = RegCSMaster(divcd, orderserial, reguserid, title, contents_jupsu, gubun01, gubun02)
    end if

    If (Err.Number = 0) and (ScanErr="") Then
        errcode = "002"

        'CS Master 환불 관련정보 저장
        Call RegCSMasterRefundInfo(id, returnmethod, refundrequire , orgsubtotalprice, orgitemcostsum, orgbeasongpay , orgmileagesum, orgcouponsum, orgallatdiscountsum  , canceltotal, refunditemcostsum, refundmileagesum, refundcouponsum, allatsubtractsum , refundbeasongpay, refunddeliverypay, refundadjustpay , rebankname, rebankaccount, rebankownername, paygateTid)
        Call AddCSMasterRefundInfo(id, orggiftcardsum, orgdepositsum, refundgiftcardsum, refunddepositsum)

        if (rebankaccount <> "") then
        	Call EditCSMasterRefundEncInfo(id, "AE2", rebankaccount)
        end if
    End if

    sqlStr = "insert into [db_user].[dbo].tbl_giftcard_log"
    sqlStr = sqlStr + " (userid, useCash, jukyocd, jukyo, orderserial, deleteyn, reguserid)"
    sqlStr = sqlStr + " values('" + userid + "'," + CStr(refundrequire*-1) + ",'400','" & title & "','" + orderserial + "','N', '" & reguserid & "')"
    dbget.Execute sqlStr

	Call updateUserGiftCard(userid)

    If (Err.Number = 0) and (ScanErr="") Then
        dbget.CommitTrans
    Else
        dbget.RollBackTrans
        response.write "<script>alert(" + Chr(34) + "데이타를 저장하는 도중에 에러가 발생하였습니다. 관리자 문의 요망.(에러코드 : " + CStr(errcode) + ":" + Err.Description + "|" + ScanErr + ")" + Chr(34) + ")</script>"
        'response.write "<script>history.back()</script>"
        dbget.close()	:	response.End
    End If

on error Goto 0

%>
<script language="javascript">
	alert("접수되었습니다.");
	opener.location.reload();
	window.close();
</script>

<!-- #include virtual="/cscenter/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
