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

''��� ����
dim refundmileagesum, refundcouponsum, allatsubtractsum
dim refunditemcostsum, canceltotal, nextsubtotal
dim refundbeasongpay, remainbeasongpay, refunddeliverypay, refundadjustpay
dim remainitemcostsum
dim refundgiftcardsum, refunddepositsum

''ȯ�� ���� maybe (refundrequire==canceltotal)
dim refundrequire, returnmethod
dim rebankname, rebankaccount, rebankownername, paygateTid, encmethod

''���ֹ� �ݾ�
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
	Response.Write "<script>alert('���̵� �����ϴ�.');</script>"
	dbget.close()
	Response.End
end if

if (orderserial = "") and (userid<>"danbi2612") and (userid<>"setjddms") and (userid<>"dadareda") then
	Response.Write "<script>alert('�ֹ���ȣ�� �����ϴ�.');</script>"
	dbget.close()
	Response.End
end if

if (CLng(FormatNumber((100*oTenGiftCard.FspendCash/oTenGiftCard.FgainCash),0)) < 60) and (userid<>"danbi2612") and (userid<>"setjddms") and (userid<>"dadareda") and (userid<>"eiddr0705") then
	Response.Write "<script>alert('Giftī�������( = ��ǰ�����Ѿ�/����Ѿ�) �� 60% �̻��� ��츸 �ܾ��� ��ġ����ȯ�� �����մϴ�.');</script>"
	dbget.close()
	Response.End
end if

if (currentCash*1 < refundrequire*1) then
	Response.Write "<script>alert('Giftī�� �ܾ׺��� ��ġ����ȯ���� �� Ů�ϴ�.');</script>"
	dbget.close()
	Response.End
end if



'==============================================================================
reguserid   = session("ssbctid")

divcd 		= "A003"
title 		= "Giftī�� ������ ȯ��"

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

        '' CS Master ����
        id = RegCSMaster(divcd, orderserial, reguserid, title, contents_jupsu, gubun01, gubun02)
    end if

    If (Err.Number = 0) and (ScanErr="") Then
        errcode = "002"

        'CS Master ȯ�� �������� ����
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
        response.write "<script>alert(" + Chr(34) + "����Ÿ�� �����ϴ� ���߿� ������ �߻��Ͽ����ϴ�. ������ ���� ���.(�����ڵ� : " + CStr(errcode) + ":" + Err.Description + "|" + ScanErr + ")" + Chr(34) + ")</script>"
        'response.write "<script>history.back()</script>"
        dbget.close()	:	response.End
    End If

on error Goto 0

%>
<script language="javascript">
	alert("�����Ǿ����ϴ�.");
	opener.location.reload();
	window.close();
</script>

<!-- #include virtual="/cscenter/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
