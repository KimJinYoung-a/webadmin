<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : cs����
' History : 2009.04.17 �̻� ����
'			2016.06.30 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db_TPLOpen.asp" -->
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

<%

'[�⺻����]
'
'if (mode="regcsas") then
'	'==========================================================================
'	'CS ����
'
'elseif (mode="deletecsas") then
'	'==========================================================================
'	'CS ���� ����
'
'elseif (mode="editcsas") then
'	'==========================================================================
'	'CS ���� ���� ����
'
'elseif (mode="finishededitcsas") then
'	'==========================================================================
'	'�Ϸ�� ���� ����
'
'elseif (mode="delfinishedcsas") then
'	'==========================================================================
'	'�Ϸ�� ���� ����
'
'elseif (mode="finishcsas") then
'	'==========================================================================
'	'CS ���� ���� �Ϸ�ó��
'
'elseif (mode="state2jupsu") then
'	'==========================================================================
'	'��ü ��Ÿ���� �������·� ����
'
'elseif (mode="addupchejungsanEdit") then
'	'==========================================================================
'	'��ü�߰����� ����
'
'elseif (mode="upcheconfirm2jupsu") then
'	'==========================================================================
'	'��ü ó���Ϸ� => �������·κ���
'
'elseif (mode="changeorderreg") then
'	'==========================================================================
'	'��ȯ�ֹ� �������
'
'elseif (mode="changedivcdtoa004") then
'	'==========================================================================
'	'�� ������ǰ ��ȯ(A010 -> A004)
'
'elseif (mode="changedivcdtoa010") then
'	'==========================================================================
'	'ȸ����û ��ȯ(A004 -> A010)
'
'else
'	'==========================================================================
'    '����
'
'end if



'[�ڵ�����]
'------------------------------------------------------------------------------
'A008			�ֹ����
'
'A004			��ǰ����(��ü���)
'A010			ȸ����û(�ٹ����ٹ��)
'
'A001			������߼�
'A002			���񽺹߼�

'A200			��Ÿȸ��
'
'A000			�±�ȯ���
'
'A009			��Ÿ����
'A006			�������ǻ���
'A700			��ü��Ÿ����
'
'A003			ȯ��
'A005			�ܺθ�ȯ�ҿ�û
'
'A011			�±�ȯȸ��(�ٹ����ٹ��)



dim mode, modeflag2, divcd, id, reguserid, ipkumdiv
dim title, orderserial, gubun01, gubun02, contents_jupsu
dim finishuser, contents_finish

dim requireupche, requiremakerid, ForceReturnByTen
dim detailitemlist
dim csdetailitemlist

''��� ����
dim refundmileagesum, refundcouponsum, allatsubtractsum
dim refunditemcostsum, canceltotal, nextsubtotal
dim refundbeasongpay, remainbeasongpay, refunddeliverypay, refundadjustpay
dim remainitemcostsum
dim refundgiftcardsum, refunddepositsum

''ȯ�� ���� maybe (refundrequire==canceltotal)
dim refundrequire, returnmethod
dim rebankname, rebankaccount, rebankownername, paygateTid, encmethod

''��ü �߰� �����
dim add_upchejungsandeliverypay, add_upchejungsancause, add_upchejungsancauseText

''���ֹ� �ݾ�
dim orgitemcostsum, orgbeasongpay, orgmileagesum, orgcouponsum, orgallatdiscountsum, orgsubtotalprice, orggiftcardsum, orgdepositsum

''�� Open msg
dim opentitle, opencontents

''�߰�����ID
dim buf_requiremakerid

''�߰��� ��ϵ� CSID
dim newasid

''CS���Ϲ߼�����
dim isCsMailSend

'�� �߰���ۺ�(��ǰ���� �±�ȯ)
dim add_customeraddmethod, add_customeraddbeasongpay, customerrealbeasongpay

'�� �߰���ۺ�(��ǰ)
dim addbeasongpay, addmethod

'��ȯ�ֹ�
dim changeorderserial, changeorderstate

'ǰ����� ��ǰ���� ����
dim modifyitemstockoutyn

dim isCSServiceRefund

'���� �����
dim copycouponinfo

dim ocsaslist
dim oRefCSASList, refminusorderserial

dim regDetailRows

dim needChkYN

newasid = -1

mode        = request.Form("mode")
modeflag2   = request.Form("modeflag2")
divcd       = request.Form("divcd")
id          = request.Form("id")
ipkumdiv    = request.Form("ipkumdiv")
reguserid   = session("ssbctid")
finishuser  = reguserid
title       = html2DB(request.Form("title"))
orderserial = request.Form("orderserial")
gubun01     = request.Form("gubun01")
gubun02     = request.Form("gubun02")
contents_jupsu  	= html2DB(request.Form("contents_jupsu"))
detailitemlist  	= html2db(request.Form("detailitemlist"))
csdetailitemlist  	= html2db(request.Form("csdetailitemlist"))
contents_finish 	= html2DB(request.Form("contents_finish"))

''��ü ó�� ��û
requireupche = request.Form("requireupche")
requiremakerid = request.Form("requiremakerid")
ForceReturnByTen = request.Form("ForceReturnByTen")

orgitemcostsum      = request.Form("orgitemcostsum")
orgbeasongpay       = request.Form("orgbeasongpay")
orgmileagesum       = request.Form("miletotalprice")
orgcouponsum        = request.Form("tencardspend")
orgallatdiscountsum = request.Form("allatdiscountprice")
orgsubtotalprice    = request.Form("subtotalprice")

orggiftcardsum    	= request.Form("giftcardsum")			'�÷������� �����Ѵ�.
refundgiftcardsum   = request.Form("refundgiftcardsum")
orgdepositsum    	= request.Form("depositsum")			'�÷������� �����Ѵ�.
refunddepositsum    = request.Form("refunddepositsum")

refunditemcostsum   = request.Form("refunditemcostsum")
nextsubtotal        = request.Form("nextsubtotal")
canceltotal         = request.Form("canceltotal")

refundbeasongpay    = request.Form("refundbeasongpay")
remainbeasongpay    = request.Form("remainbeasongpay")
refunddeliverypay   = request.Form("refunddeliverypay")

refundmileagesum    = request.Form("refundmileagesum")
refundcouponsum     = request.Form("refundcouponsum")
allatsubtractsum    = request.Form("allatsubtractsum")
refundadjustpay     = request.Form("refundadjustpay")
remainitemcostsum   = request.Form("remainitemcostsum")



''ȯ�ҿ�û��
refundrequire       = request.Form("refundrequire")
returnmethod        = request.Form("returnmethod")

rebankname          = request.Form("rebankname")
rebankaccount       = request.Form("rebankaccount")
rebankownername     = request.Form("rebankownername")

encmethod 			= "AE2" ''"PH1"

paygateTid          = request.Form("paygateTid")

add_upchejungsandeliverypay = request.Form("add_upchejungsandeliverypay")
add_upchejungsancause       = request.Form("add_upchejungsancause")
add_upchejungsancauseText   = request.Form("add_upchejungsancauseText")

buf_requiremakerid  = request.Form("buf_requiremakerid")


isCsMailSend = (request.Form("csmailsend")="on")

add_customeraddmethod   	= request.Form("add_customeraddmethod")
add_customeraddbeasongpay   = request.Form("add_customeraddbeasongpay")
customerrealbeasongpay   	= request.Form("customerrealbeasongpay")

modifyitemstockoutyn   		= request.Form("modifyitemstockoutyn")

copycouponinfo        		= request.Form("copycouponinfo")

addbeasongpay   			= request.Form("addbeasongpay")
addmethod   				= request.Form("addmethod")

needChkYN   				= request.Form("needChkYN")

if (add_upchejungsancause="�����Է�") then add_upchejungsancause = add_upchejungsancauseText


if (Not IsNumeric(orgitemcostsum)) or (orgitemcostsum="") then orgitemcostsum     = 0
if (Not IsNumeric(orgbeasongpay)) or (orgbeasongpay="") then orgbeasongpay      = 0
if (Not IsNumeric(orgmileagesum)) or (orgmileagesum="") then orgmileagesum   = 0
if (Not IsNumeric(orgcouponsum)) or (orgcouponsum="") then orgcouponsum    = 0
if (Not IsNumeric(orgallatdiscountsum)) or (orgallatdiscountsum="") then orgallatdiscountsum= 0
if (Not IsNumeric(orgsubtotalprice)) or (orgsubtotalprice="") then orgsubtotalprice   = 0

if (Not IsNumeric(refunditemcostsum)) or (refunditemcostsum="") then refunditemcostsum  = 0
if (Not IsNumeric(refundmileagesum)) or (refundmileagesum="") then refundmileagesum = 0
if (Not IsNumeric(refundcouponsum)) or (refundcouponsum="") then refundcouponsum = 0
if (Not IsNumeric(allatsubtractsum)) or (allatsubtractsum="") then allatsubtractsum = 0

if (Not IsNumeric(refundbeasongpay)) or (refundbeasongpay="") then refundbeasongpay = 0
if (Not IsNumeric(remainbeasongpay)) or (remainbeasongpay="") then remainbeasongpay = 0
if (Not IsNumeric(refunddeliverypay)) or (refunddeliverypay="") then refunddeliverypay = 0

if (Not IsNumeric(refundadjustpay)) or (refundadjustpay="") then refundadjustpay = 0
if (Not IsNumeric(remainitemcostsum)) or (remainitemcostsum="") then remainitemcostsum = 0
if (Not IsNumeric(canceltotal)) or (canceltotal="") then canceltotal = 0
if (Not IsNumeric(refundrequire)) or (refundrequire="") then refundrequire = 0

if (Not IsNumeric(orggiftcardsum)) or (orggiftcardsum="") then orggiftcardsum = 0
if (Not IsNumeric(refundgiftcardsum)) or (refundgiftcardsum="") then refundgiftcardsum = 0
if (Not IsNumeric(orgdepositsum)) or (orgdepositsum="") then orgdepositsum = 0
if (Not IsNumeric(refunddepositsum)) or (refunddepositsum="") then refunddepositsum = 0

if (Not IsNumeric(add_customeraddbeasongpay)) or (add_customeraddbeasongpay="") then add_customeraddbeasongpay = 0
if (Not IsNumeric(customerrealbeasongpay)) or (customerrealbeasongpay="") then customerrealbeasongpay = 0

if (Not IsNumeric(addbeasongpay)) or (addbeasongpay="") then addbeasongpay = 0

if (returnmethod="") then returnmethod="R000"
if (copycouponinfo="") then copycouponinfo="N"


''�ÿ�ī������.. -��ǰ���� ����.

dim sqlStr, errcode, i
dim ScanErr
dim ResultMsg, ReturnUrl, EtcStr
dim ProceedFinish
dim ResultCount

ScanErr = ""
ProceedFinish = False

dim IsAllCancel
dim CancelValidResultMessage

''���� �ֹ� �������� Check
GC_IsOLDOrder = CheckIsOldOrder(orderserial)



'==============================================================================
''�ֹ� ����Ÿ
dim oordermaster

set oordermaster = new COrderMaster

oordermaster.FRectOrderSerial = orderserial

if Left(orderserial,1)="A" then
    set oordermaster.FOneItem = new COrderMasterItem
else
    oordermaster.QuickSearchOrderMaster_3PL
end if

'' ���� 6���� ���� ���� �˻�
if (oordermaster.FResultCount<1) and (Len(orderserial)=11) and (IsNumeric(orderserial)) then
    oordermaster.FRectOldOrder = "on"
    oordermaster.QuickSearchOrderMaster_3PL
end if

if (oordermaster.FResultCount<1) then
	rw "ERROR!!"
	response.end
end if

if (mode="regcsas") then
    '==========================================================================

	rw "�۾��ȵǾ� ����!!"
	dbget.close : dbget_TPL.close : response.end

	'CS ����
    if (divcd="A008") then
		'----------------------------------------------------------------------
        'CS ���� - �ֹ����
        'On Error Resume Next
        dbget.beginTrans

		'// 6���� ���� �ֹ� ��Ұ���(2014-03-31)
        ''if (GC_IsOLDOrder) then ScanErr = "6���� ���� ���� ���� ��� �Ұ� - ������ ���� ���"

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "001"

            'CS Master ����
            id = RegCSMaster(divcd, orderserial, reguserid, title, contents_jupsu, gubun01, gubun02)
        end if

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "002"

            'CS Master ȯ�� �������� ����
	        Call RegCSMasterRefundInfo(id, returnmethod, refundrequire , orgsubtotalprice, orgitemcostsum, orgbeasongpay , orgmileagesum, orgcouponsum, orgallatdiscountsum  , canceltotal, refunditemcostsum, refundmileagesum, refundcouponsum, allatsubtractsum , refundbeasongpay, refunddeliverypay, refundadjustpay , rebankname, rebankaccount, rebankownername, paygateTid)
	        Call AddCSMasterRefundInfo(id, orggiftcardsum, orgdepositsum, refundgiftcardsum, refunddepositsum)

            '''���� ��ȣȭ �߰�.
	        Call EditCSMasterRefundEncInfo(id, encmethod, rebankaccount)

			''���� ���������
			Call EditCSCopyCouponInfo(id, copycouponinfo)
	    End if

	    If (Err.Number = 0) and (ScanErr="") Then
            errcode = "003"

            'CS Detail ����(���� ��ǰ����)
	        Call AddCSDetailByArrStr(detailitemlist, id, orderserial)

        end if

	    If (Err.Number = 0) and (ScanErr="") Then
            errcode = "005"

            if (remainitemcostsum = 0) then
            	'��ü��� : �ܿ���ǰ�Ѿ��� ���� ���
            	IsAllCancel = true
            	CancelValidResultMessage = GetAllCancelRegValidResult(id, orderserial)
            else
	            IsAllCancel = false
	            CancelValidResultMessage = GetPartialCancelRegValidResult(id, orderserial)
	        end if

			if (CancelValidResultMessage <> "") then
				ScanErr = CancelValidResultMessage
			end if
        End If


        '���Ϸ� �Ǵ� ��ҵ� ��ǰ�� ���� ���, ��������(�ֹ���� �Ұ�)
        '���Ϸ�� ��ǰ�� ��ǰ�� �����ϴ�.
        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "006"

            ''��� �Ϸ� �Ǵ� ��ҵ� ������ �ִ��� Ȯ��
            if Not (IsCancelValidState(id, orderserial)) then
                ScanErr = "��� ���� ����. - ���� ������ �ְų� ��ҵ� ������ �ֽ��ϴ�."
            end if
        end if

        '' �Ϸ�ó�� �ٷ� �������� ����
        '' ��ü Ȯ���� ���°� �ִ°�� - > �����θ� ����
        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "007"

        	''�ٷ� �Ϸ�ó���� ���� ���� ���� - AsDetail �Է��� �˻�
            ProceedFinish   = IsDirectProceedFinish(divcd, id, orderserial, EtcStr)
            contents_finish = ""
        End If

        ResultMsg = ResultMsg + "->. [�ֹ� ��� CS] ����\n\n"

        '' �Ϸ�ó�� ���μ���
        'TODO : �������� ������ �����ִ� ��� :
        If (ProceedFinish) then
            If (Err.Number = 0) and (ScanErr="") Then
                errcode = "008"

                Call CancelProcess(id, orderserial)

                ResultMsg = ResultMsg + "->. �ֹ��� ��� �Ϸ�\n\n"
            End IF

            ''����?. ����?
            If (Err.Number = 0) and (ScanErr="") Then
                errcode = "009"

                'ȯ�� ������ �ִ��� üũ �� ������ȯ��/���ϸ���ȯ��/�ſ�ī����� CS ���� ���
                newasid = CheckNRegRefund(id, orderserial,reguserid)

                If (newasid>0) then
                    ResultMsg = ResultMsg + "->. ȯ�� ���� �Ϸ�\n\n"
                end if
            End If

            If (Err.Number = 0) and (ScanErr="") Then
                errcode = "010"

                Call FinishCSMaster(id, reguserid, contents_finish)

                ResultMsg = ResultMsg + "->. [�ֹ� ��� CS] �Ϸ� ó��\n\n"

				if CheckAndCopyBonusCoupon(id, reguserid) = True then
					ResultMsg = ResultMsg + "->. [���� ��߱�] �Ϸ� ó��\n\n"
				end if
            End If
        ELSE
            ResultMsg = ResultMsg + "->. ��ǰ �غ��� ������ ��ǰ�� �����ϹǷ�\n\n �ֹ� ��� ������ ���� �Ǿ����ϴ�.\n\n Ȯ���� �Ϸ� ó���ϼž� �մϴ�."
        End If

	    If (Err.Number = 0) and (ScanErr="") Then
            dbget.CommitTrans
        Else
            dbget.RollBackTrans
            response.write "<script>alert(" + Chr(34) + "����Ÿ�� �����ϴ� ���߿� ������ �߻��Ͽ����ϴ�. ������ ���� ���.(�����ڵ� : " + CStr(errcode) + ":" + Err.Description + "|" + ScanErr + ")" + Chr(34) + ")</script>"
            response.write "����Ÿ�� �����ϴ� ���߿� ������ �߻��Ͽ����ϴ�. ������ ���� ���.(�����ڵ� : " + CStr(errcode) + ":" + Err.Description + "|" + ScanErr + ")"
            ''response.write "<script>history.back()</script>"
            dbget.close()	:	response.End
        End If

		'' �����ǰ ��ǰ�غ����̸� ��üȮ�� �� ����ϴ� �δܰ��ε� �ϴ��� �����Ѵ�.
		''If (ProceedFinish) then

			''������� �ݾ�/������ ����
			Call CheckNChangeCyberAcct(orderserial)

			if IsAllCancel = true then
				''���ں��� ���
				Call CheckNUsafeCancel(orderserial)
			end if

			'' ��� ���� ��ǰ�� ǰ����ǰ�� ��� ��ǰ������ ǰ������
			if (modifyitemstockoutyn = "Y") then
				ResultCount   = SetStockOutByCsAs(id)
				if (ResultCount > 0) then
					ResultMsg = ResultMsg + "->. [��ǰ���� ǰ�� ����] �Ϸ� ó��\n\n"
				end if
			end if

			''��� �� �������� ����(2007-09-01 ������ �߰�)
			''Call LimitItemRecover(orderserial) : ����
			if (IsAllCancel) then
				''��ü ����ΰ��
				sqlStr = " exec [db_summary].[dbo].sp_Ten_RealtimeStock_cancelOrderAll '" & orderserial & "'"
				dbget.Execute sqlStr
			else
				''�κ� ����ΰ��
				sqlStr = " select itemid,itemoption,regitemno "
				sqlStr = sqlStr & " from [db_cs].[dbo].tbl_new_as_detail "
				sqlStr = sqlStr & " where masterid=" & id
				sqlStr = sqlStr & " and orderserial='" & orderserial & "'"

				rsget.Open sqlStr,dbget,1
				if Not rsget.Eof then
					regDetailRows = rsget.getRows()
				end if
				rsget.Close

				if IsArray(regDetailRows) then
					for i=0 to UBound(regDetailRows,2)
						sqlStr = " exec [db_summary].[dbo].sp_Ten_RealtimeStock_cancelOrderPartial '" & orderserial & "'," & regDetailRows(0,i) & ",'" & regDetailRows(1,i) & "'," & regDetailRows(2,i)
						dbget.Execute sqlStr
					Next
				end if
			end if

		''end if

        ''�̸��� �߼�. �ٷ� �Ϸ��ΰ�츸.
        If (isCsMailSend) then
            If (ProceedFinish) then
                ''�ֹ���� �Ϸ� ����
                Call SendCsActionMail(id)

                ''ȯ�� ���� ����
                if (newasid>0) then
                    Call SendCsActionMail(newasid)
                end if
            End If
        End IF
        'on error Goto 0

        ReturnUrl = "/admin/etc/3pl/cscenter/action/pop_cs_action_new_3PL.asp?id="  + CStr(id) + "&divcd=" + divcd + "&mode=editreginfo"

    elseif (divcd="A004") or (divcd="A010") then
    	'----------------------------------------------------------------------
        'CS ���� - ��ǰ ���� �Ǵ� ȸ����û.
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

		        '''���� ��ȣȭ �߰�.
	            Call EditCSMasterRefundEncInfo(id, encmethod, rebankaccount)

				''���� ���������
				Call EditCSCopyCouponInfo(id, copycouponinfo)
		    End if


		    If (Err.Number = 0) and (ScanErr="") Then
	            errcode = "003"

	            'CS Detail ����(���� ��ǰ����)
		        Call AddCSDetailByArrStr(detailitemlist, id, orderserial)
	        end if

	        '' Check - ��ü��۰� ��ü����� ���� �������� ����.
	        ''       - ��ü����� ������ ��� �Ѱ��� �귣�常 ���� �ؾ���.
	        If (Err.Number = 0) and (ScanErr="") Then
	            errcode = "004"

	            if (IsReturnRegValid(id, orderserial, ScanErr, requiremakerid)) then
	                '��ü����� ��� ���� ��ü �귣�� ���̵� �Է�(requiremakerid)
	                if ((requiremakerid<>"") and (ForceReturnByTen="")) then
	                    call RegCSMasterAddUpche(id, requiremakerid)
	                end if

	                ResultMsg = ResultMsg + "->. [��ǰ / ȸ�� CS] ����\n\n"
	            end if
	        End if

	        ''��ü �߰� ����� 2008.11.10
	        If (Err.Number = 0) and (ScanErr="") Then
	            errcode = "005"

				'// ��ǰ����� ��� ����ݾ� 0��, skyer9, 2015-09-02
	            if (add_upchejungsandeliverypay<>"0") and (add_upchejungsandeliverypay<>"")  then
	                call RegCSUpcheAddJungsanPay(id, add_upchejungsandeliverypay, add_upchejungsancause, buf_requiremakerid)
	            end if
	        end If

			If (Err.Number = 0) and (ScanErr="") Then
				errcode = "006"

				'' �� �߰���ۺ�
				if (divcd = "A004") or (divcd = "A010")  then
					Call SetCustomerAddBeasongPay(id, addmethod, addbeasongpay, "Y", 0)
				end if

			end if

	        ReturnUrl = "/admin/etc/3pl/cscenter/action/pop_cs_action_new_3PL.asp?id="  + CStr(id) + "&divcd=" + divcd

	        If (Err.Number = 0) and (ScanErr="") Then
	            dbget.CommitTrans
	        Else
	            dbget.RollBackTrans
	            response.write "<script>alert(" + Chr(34) + "����Ÿ�� �����ϴ� ���߿� ������ �߻��Ͽ����ϴ�. ������ ���� ���.(�����ڵ� : " + CStr(errcode) + ":" + Err.Description + "|" + ScanErr + ")" + Chr(34) + ")</script>"
	            'response.write "<script>history.back()</script>"
	            dbget.close()	:	response.End
	        End If

	        ''�̸��� �߼�. ��ǰ ȸ�� ����
	        If (isCsMailSend) then
	            Call SendCsActionMail(id)
	        end if
        on error Goto 0

    elseif (divcd="A001") or (divcd="A002") or (divcd="A200") then
    	'----------------------------------------------------------------------
        'CS ���� - ������߼�, ���񽺹߼�, ��Ÿȸ��
        On Error Resume Next
	        dbget.beginTrans

	        If (Err.Number = 0) and (ScanErr="") Then
	            errcode = "001"

	            '' CS Master ����
	            id = RegCSMaster(divcd, orderserial, reguserid, title, contents_jupsu, gubun01, gubun02)
	        end if

	        If (Err.Number = 0) and (ScanErr="") Then
	            errcode = "002"

	            'CS Detail ����(���� ��ǰ����)
		        Call AddCSDetailByArrStr(detailitemlist, id, orderserial)
	        end if

			'��ü����� ��� ���� ��ü �귣�� ���̵� �Է�(requiremakerid)
	        if (requiremakerid<>"") then
	            call RegCSMasterAddUpche(id, requiremakerid)
	        end if

	        ''��ü �߰� ����� 2012-06-25(skyer9)
	        If (Err.Number = 0) and (ScanErr="") Then
	            errcode = "003"

	            if (add_upchejungsandeliverypay<>"0") and (add_upchejungsandeliverypay<>"")  then
	                call RegCSUpcheAddJungsanPay(id, add_upchejungsandeliverypay, add_upchejungsancause, buf_requiremakerid)
	            end if
	        end if

	        ResultMsg = "�����Ϸ�"
	        ReturnUrl = "/admin/etc/3pl/cscenter/action/pop_cs_action_new_3PL.asp?id="  + CStr(id) + "&divcd=" + divcd

	        If (Err.Number = 0) and (ScanErr="") Then
	            dbget.CommitTrans
	        Else
	            dbget.RollBackTrans
	            response.write "<script>alert(" + Chr(34) + "����Ÿ�� �����ϴ� ���߿� ������ �߻��Ͽ����ϴ�. ������ ���� ���.(�����ڵ� : " + CStr(errcode) + ":" + Err.Description + "|" + ScanErr + ")" + Chr(34) + ")</script>"
	            'response.write "<script>history.back()</script>"
	            dbget.close()	:	response.End
	        End If

	        ''�̸��� �߼� ���� ���� ����
	        If (isCsMailSend) then
				if Not (divcd="A200") then
					'// ��Ÿȸ�� ���Ϲ߼� ����
	            	Call SendCsActionMail(id)
				end if
	        End If
        on error Goto 0

    elseif (divcd="A000") then
		'----------------------------------------------------------------------
        'CS ���� - �±�ȯ���
        On Error Resume Next
	        dbget.beginTrans

	        If (Err.Number = 0) and (ScanErr="") Then
	            errcode = "001"

	            '' CS Master ����
	            id = RegCSMaster(divcd, orderserial, reguserid, title, contents_jupsu, gubun01, gubun02)
	        end if

	        If (Err.Number = 0) and (ScanErr="") Then
	            errcode = "002"

	            'CS Detail ����(���� ��ǰ����)
		        Call AddCSDetailByArrStr(detailitemlist, id, orderserial)
	        end if

	        If (Err.Number = 0) and (ScanErr="") Then
	            errcode = "003"

				'// CS �±�ȯ���(���ϻ�ǰ, ��ǰ���� - A000, A100) ������ ���Ǵ� ��ǰ ��������
				Call ApplyLimitItemByCS(id)
	        end if

	        If (Err.Number = 0) and (ScanErr="") Then
	            errcode = "004"

	            if (requiremakerid<>"") then

	                '��ü����� ��� ���� ��ü �귣�� ���̵� �Է�(requiremakerid)
	                Call RegCSMasterAddUpche(id, requiremakerid)

	                '��ü����� ��� �±�ȯ ȸ�� ����
                    newasid = RegCSMaster("A012", orderserial, reguserid, "��ȯȸ��(��ü���) ����", contents_jupsu, gubun01, gubun02)

					'��ü����� ��� ���� ��ü �귣�� ���̵� �Է�(requiremakerid)
                    Call RegCSMasterAddUpche(newasid, requiremakerid)

                    Call AddCSDetailByArrStr(detailitemlist, newasid, orderserial)

					'// asid ����
					Call SetRefAsid(newasid, id)

                    ResultMsg = "��ȯ ��� ���� �� ȸ������ �Ϸ� - ��ü ���"

	            else

	                '�ٹ����� ����� ��� �±�ȯ ȸ�� ����
	                newasid = RegCSMaster("A011", orderserial, reguserid, "��ȯȸ�� ����", contents_jupsu, gubun01, gubun02)

	                Call AddCSDetailByArrStr(detailitemlist, newasid, orderserial)

					'// asid ����
					Call SetRefAsid(newasid, id)

	                ResultMsg = "��ȯ ��� ���� �� ȸ������ �Ϸ� - �ٹ����� ���"

	            end if
	        end if

	        ''��ü �߰� ����� 2008.11.10
	        If (Err.Number = 0) and (ScanErr="") Then
	            errcode = "005"

	            if (add_upchejungsandeliverypay<>"0") and (add_upchejungsandeliverypay<>"")  then
	                call RegCSUpcheAddJungsanPay(id, add_upchejungsandeliverypay, add_upchejungsancause, buf_requiremakerid)
	            end if
	        end if

	        ReturnUrl = "/admin/etc/3pl/cscenter/action/pop_cs_action_new_3PL.asp?id="  + CStr(id) + "&divcd=" + divcd

	        If (Err.Number = 0) and (ScanErr="") Then
	            dbget.CommitTrans
	        Else
	            dbget.RollBackTrans
	            response.write "<script>alert(" + Chr(34) + "����Ÿ�� �����ϴ� ���߿� ������ �߻��Ͽ����ϴ�. ������ ���� ���.(�����ڵ� : " + CStr(errcode) + ":" + Err.Description + "|" + ScanErr + ")" + Chr(34) + ")</script>"
	            'response.write "<script>history.back()</script>"
	            dbget.close()	:	response.End
	        End If

	        ''�̸��� �߼� �±�ȯ ����
	        if (isCsMailSend) then
	            Call SendCsActionMail(id)

	            ''�±�ȯ ȸ���� �������
	            if (newasid>0) and (requiremakerid = "") then
	            	'// ��ü����� ������ �ʴ´�.
   	                Call SendCsActionMail(newasid)
	            end if
	        End If
        on error Goto 0

    elseif (divcd="A009") or (divcd="A006") or (divcd="A700") then
    	'----------------------------------------------------------------------
        'CS ���� - ��Ÿ���� / �������ǻ��� / ��ü �߰� �����
        On Error Resume Next
	        dbget.beginTrans

	        If (Err.Number = 0) and (ScanErr="") Then
	            errcode = "001"

	            '' CS Master ����
	            id = RegCSMaster(divcd, orderserial, reguserid, title, contents_jupsu, gubun01, gubun02)
	        end if

	        If (Err.Number = 0) and (ScanErr="") Then
	            errcode = "002"

	            'CS Detail ����(���� ��ǰ����)
		        Call AddCSDetailByArrStr(detailitemlist, id, orderserial)
	        end if

	        ''��ü �߰� ����� : 2008.11.10
	        If (Err.Number = 0) and (ScanErr="") Then
	            errcode = "003"

				'// ��ǰ����� ��� ����ݾ� 0��, skyer9, 2015-09-02
				if ((add_upchejungsandeliverypay<>"0") or ((divcd = "A700") And (add_upchejungsancause = "��ǰ���"))) and (add_upchejungsandeliverypay<>"")  then
	                call RegCSUpcheAddJungsanPay(id, add_upchejungsandeliverypay, add_upchejungsancause, buf_requiremakerid)
	            end if
	        end if

	        ''��ü����.
	        If (Err.Number = 0) and (ScanErr="") Then
	            errcode = "004"

	            '��ü����� ��� ���� ��ü �귣�� ���̵� �Է�(requiremakerid)
	            if (requiremakerid<>"") then
	                call RegCSMasterAddUpche(id, requiremakerid)
	            end if
	         end if

	        ResultMsg = "��ϵǾ����ϴ�."
	        ReturnUrl = "/admin/etc/3pl/cscenter/action/pop_cs_action_new_3PL.asp?id="  + CStr(id) + "&divcd=" + divcd

	        If (Err.Number = 0) and (ScanErr="") Then
	            dbget.CommitTrans
	        Else
	            dbget.RollBackTrans
	            response.write "<script>alert(" + Chr(34) + "����Ÿ�� �����ϴ� ���߿� ������ �߻��Ͽ����ϴ�. ������ ���� ���.(�����ڵ� : " + CStr(errcode) + ":" + Err.Description + "|" + ScanErr + ")" + Chr(34) + ")</script>"
	            'response.write "<script>history.back()</script>"
	            dbget.close()	:	response.End
	        End If
        on error Goto 0

    elseif (divcd="A003") or (divcd="A005") then
    	'----------------------------------------------------------------------
        'CS ���� - ȯ������ / �ܺθ� ȯ������
		isCSServiceRefund = False
        On Error Resume Next
	        dbget.beginTrans

	        if (divcd="A005") and (Not IsExtSiteOrder(orderserial)) then
	            ScanErr = "�ܺθ� ȯ�������� �ܺθ� �ֹ��Ǹ� �����մϴ�."
	        end if

	        If (Err.Number = 0) and (ScanErr="") Then
	            errcode = "001"

	            '' CS Master ����
	            if (divcd="A003") then
					if (returnmethod = "R900") then
						title = title & "(���ϸ���)"
					elseif (returnmethod = "R910") then
						title = title & "(��ġ��)"
					elseif (returnmethod = "R100") then
						title = title & "(�ſ�ī�����)"
					elseif (returnmethod = "R550") then
						title = title & "(���������)"
					elseif (returnmethod = "R560") then
						title = title & "(����Ƽ�����)"
					elseif (returnmethod = "R120") then
						title = title & "(�ſ�ī��κ����)"
					elseif (returnmethod = "R022") then
						title = title & "(�ǽð���ü�κ����)"
					elseif (returnmethod = "R007") then
						title = title & "(������)"
					end if

					'// ����ȯ������
					isCSServiceRefund = GetIsCSServiceRefund(id, divcd, title)

					title = GetCSRefundTitle(id, divcd, orderserial, returnmethod, title)
				end if

	            id = RegCSMaster(divcd, orderserial, reguserid, title, contents_jupsu, gubun01, gubun02)
	        end if

	        If (Err.Number = 0) and (ScanErr="") Then
	            errcode = "002"

	            'CS Master ȯ�� �������� ����
		        Call RegCSMasterRefundInfo(id, returnmethod, refundrequire , orgsubtotalprice, orgitemcostsum, orgbeasongpay , orgmileagesum, orgcouponsum, orgallatdiscountsum  , canceltotal, refunditemcostsum, refundmileagesum, refundcouponsum, allatsubtractsum , refundbeasongpay, refunddeliverypay, refundadjustpay , rebankname, rebankaccount, rebankownername, paygateTid)
		        Call AddCSMasterRefundInfo(id, orggiftcardsum, orgdepositsum, refundgiftcardsum, refunddepositsum)

				if (isCSServiceRefund) then
					Call SetCSServiceRefund(id)
				end if

		        '''���� ��ȣȭ �߰�.
	            Call EditCSMasterRefundEncInfo(id, encmethod, rebankaccount)
		    End if


	        ResultMsg = "��ϵǾ����ϴ�."
	        ReturnUrl = "/admin/etc/3pl/cscenter/action/pop_cs_action_new_3PL.asp?id="  + CStr(id) + "&divcd=" + divcd

	        If (Err.Number = 0) and (ScanErr="") Then
	            dbget.CommitTrans
	        Else
	            dbget.RollBackTrans
	            response.write "<script>alert(" + Chr(34) + "����Ÿ�� �����ϴ� ���߿� ������ �߻��Ͽ����ϴ�. ������ ���� ���.(�����ڵ� : " + CStr(errcode) + ":" + Err.Description + "|" + ScanErr + ")" + Chr(34) + ")</script>"
	            'response.write "<script>history.back()</script>"
	            dbget.close()	:	response.End
	        End If

	        ''�̸��� �߼� ȯ������
	        If (isCsMailSend) then
	            if (divcd="A003") then
	                Call SendCsActionMail(id)
	            end if
	        End If
        on error Goto 0

    else
        ResultMsg = "���ǵ��� �ʾҽ��ϴ�[1]. : mode=" + mode + " , divcd=" + divcd
        response.write "<script>alert('" + ResultMsg + "');</script>"
        response.write "<script>history.back();</script>"
        dbget.close()	:	response.End
    end If

	If (id > 0) And needChkYN <> "" Then
		Call EditCSMasterAddInfo(id, Array( Array("needChkYN", needChkYN) ))
	End If


elseif (mode="deletecsas") then
	'==========================================================================
	'CS ����

	rw "�۾��ȵǾ� ����!!"
	dbget.close : dbget_TPL.close : response.end

	set ocsaslist = New CCSASList
	ocsaslist.FRectCsAsID = id
	ocsaslist.GetOneCSASMaster

	if (ocsaslist.FOneItem.Fdeleteyn = "Y") then
	    response.write "<script>alert(" + Chr(34) + "�̹� ������ �����Դϴ�." + Chr(34) + ")</script>"
	    response.write "�̹� ������ �����Դϴ�."
	    dbget.close()	:	response.End
	elseif (ocsaslist.FOneItem.Fcurrstate = "B007") then
		response.write "<script>alert(" + Chr(34) + "�̹� �Ϸ�� �����Դϴ�." + Chr(34) + ")</script>"
		response.write "�̹� �Ϸ�� �����Դϴ�."
		dbget.close()	:	response.End
	end if

    On Error Resume Next
        dbget.beginTrans

        ''Check Valid Delete - ����� B006 ��üó���Ϸ� , B007 �Ϸ� ������ ���(����) �Ұ�
        if (NOT ValidDeleteCS(id)) then
            response.write "<script>alert(" + Chr(34) + "���� ��� ���� ���°� �ƴմϴ�. ������ ���� ���." + Chr(34) + ")</script>"
            response.write "<script>history.back()</script>"
            dbget.close()	:	response.End
        end if

		if (divcd = "A111") or (divcd = "A112") then
			'// ��ǰ���� ��ȯȸ��
			Call GetChangeOrderInfo(id, changeorderserial, changeorderstate, ResultMsg)

			if (changeorderserial <> "") then
				if Not IsCancelChangeOrderValidState(changeorderserial) then
					ResultMsg = "���� ���� ���°� �ƴմϴ�. ���� ��ȯ�ֹ�[" + CStr(changeorderserial) + "] ���� CS �� �����ϼ���."
		            response.write "<script>alert(" + Chr(34) + ResultMsg + Chr(34) + ")</script>"
		            response.write ResultMsg
		            ''response.write "<script>history.back()</script>"
		            dbget.close()	:	response.End
				end if

				Call setCancelMaster(id, changeorderserial)
			end if
			'
		end if

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "001"

            If Not DeleteCSProcess(id, reguserid) then
                ScanErr = "������ ������ ����"
            else
                ResultMsg = ResultMsg + "->. [CS ó���� ����] ó��\n\n"
            End if
        end if

        ResultMsg = "OK"
        ReturnUrl = "/admin/etc/3pl/cscenter/action/pop_cs_action_new_3PL.asp?id="  + CStr(id) + "&divcd=" + divcd + "&mode=editreginfo"

        If (Err.Number = 0) and (ScanErr="") Then
            dbget.CommitTrans
        Else
            dbget.RollBackTrans
            response.write "<script>alert(" + Chr(34) + "����Ÿ�� �����ϴ� ���߿� ������ �߻��Ͽ����ϴ�. ������ ���� ���.(�����ڵ� : " + CStr(errcode) + ":" + Err.Description + "|" + ScanErr + ")" + Chr(34) + ")</script>"
            'response.write "<script>history.back()</script>"
            dbget.close()	:	response.End
        End If
    On error Goto 0



elseif (mode="editcsas") then
	'==========================================================================
    ''���� ���� ����

	rw "�۾��ȵǾ� ����!!"
	dbget.close : dbget_TPL.close : response.end

    On Error Resume Next
        dbget.beginTrans

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "001"

            '' CS Master ����
            if (divcd = "A003") or (divcd = "A007")  then
            	title = GetCSRefundTitle(id, divcd, orderserial, returnmethod, title)
            end if

            Call EditCSMaster(id, reguserid, title, contents_jupsu, gubun01, gubun02)

            ''ȯ�ҹ���� �ٲ� ���.. 2011-07-20 �������߰�
            if (divcd="A007") and Not ((returnmethod="R020") or (returnmethod="R022") or (returnmethod="R080") or (returnmethod="R100") or (returnmethod="R550") or (returnmethod="R560") or (returnmethod="R120") or (returnmethod="R400") or (returnmethod="R420")) then
                sqlStr = " update [db_cs].[dbo].tbl_new_as_list"
                sqlStr = sqlStr + " set divcd='A003'"
                sqlStr = sqlStr + " where id=" + CStr(id)

                dbget.Execute sqlStr
            end if

            if (divcd="A003") and ((returnmethod="R020") or (returnmethod="R022") or (returnmethod="R080") or (returnmethod="R100") or (returnmethod="R550") or (returnmethod="R560") or (returnmethod="R120") or (returnmethod="R400") or (returnmethod="R420")) then
                sqlStr = " update [db_cs].[dbo].tbl_new_as_list"
                sqlStr = sqlStr + " set divcd='A007'"
                sqlStr = sqlStr + " where id=" + CStr(id)

                dbget.Execute sqlStr
            end if
        end if

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "002"

			'���� ���� ������ ��ü ���Է�
			''Call DeleteAllCSDetail(id, orderserial)

			if (divcd="A100") then
				'// �߰��Ǵ� ��ǰ��� ���
				Call ModiCSDetailWithoutOrderDetailByArrStr(csdetailitemlist, id, orderserial)
			else
	            'CS Detail ����(���� ��ǰ����)
		        Call ModiCSDetailByArrStr(detailitemlist, id, orderserial)
			end if

        End if

        ResultMsg = ResultMsg + "->. [CS ó���� ����] ó��\n\n"

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "003"

            '' ȯ�� ���� ����
            if (CheckNEditRefundInfo(id,returnmethod,rebankaccount,rebankownername,rebankname,paygateTid,refundrequire, orgsubtotalprice, orgitemcostsum, orgbeasongpay , orgmileagesum, orgcouponsum, orgallatdiscountsum , canceltotal, refunditemcostsum, refundmileagesum, refundcouponsum, allatsubtractsum , refundbeasongpay, refunddeliverypay, refundadjustpay)) then
            	Call AddCSMasterRefundInfo(id, orggiftcardsum, orgdepositsum, refundgiftcardsum, refunddepositsum)

            	'''���� ��ȣȭ �߰�.
	            Call EditCSMasterRefundEncInfo(id, encmethod, rebankaccount)
                ResultMsg = ResultMsg + "->. [ȯ������ ����] ó��\n\n"

				''���� ���������
				Call EditCSCopyCouponInfo(id, copycouponinfo)
            end if
        end If

        ''��ü �߰� ����� 2008.11.10
        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "004"

            if (add_upchejungsandeliverypay<>"") then
                call EditCSUpcheAddJungsanPay(id, add_upchejungsandeliverypay, add_upchejungsancause, buf_requiremakerid)
            end if
        end if

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "005"

            '' �� �߰���ۺ�
            if (divcd = "A100") or (divcd = "A111")  then
            	Call SetCustomerAddBeasongPay(id, add_customeraddmethod, add_customeraddbeasongpay, "Y", customerrealbeasongpay)
			elseif (divcd = "A004") or (divcd = "A010") then
				Call SetCustomerAddBeasongPay(id, addmethod, addbeasongpay, "Y", 0)
            end if

        end if

        ReturnUrl = "/admin/etc/3pl/cscenter/action/pop_cs_action_new_3PL.asp?id="  + CStr(id) + "&divcd=" + divcd + "&mode=editreginfo"

        If (Err.Number = 0) and (ScanErr="") Then
            dbget.CommitTrans
        Else
            dbget.RollBackTrans
            response.write "<script>alert(" + Chr(34) + "����Ÿ�� �����ϴ� ���߿� ������ �߻��Ͽ����ϴ�. ������ ���� ���.(�����ڵ� : " + CStr(errcode) + ":" + Err.Description + "|" + ScanErr + ")" + Chr(34) + ")</script>"
            'response.write "<script>history.back()</script>"
            dbget.close()	:	response.End
        End If
    On error Goto 0

	If (id > 0) And needChkYN <> "" Then
		Call EditCSMasterAddInfo(id, Array( Array("needChkYN", needChkYN) ))
	End If


elseif (mode="finishededitcsas") then
	'==========================================================================
    ''�Ϸ�� ���� ����

	rw "�۾��ȵǾ� ����!!"
	dbget.close : dbget_TPL.close : response.end

    On Error Resume Next
        dbget.beginTrans

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "001"

            '' ����Ÿ ����
            Call EditCSMasterFinished(id, title, contents_jupsu, gubun01, gubun02, reguserid, contents_finish)
        end if

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "002"

            '' ������ ����
            Call EditCSDetailByArrStr(detailitemlist, id, orderserial)
        End if

        ResultMsg = "���� �Ǿ����ϴ�."
        ReturnUrl = "/admin/etc/3pl/cscenter/action/pop_cs_action_new_3PL.asp?id="  + CStr(id) + "&divcd=" + divcd + "&mode=editreginfo"

        If (Err.Number = 0) and (ScanErr="") Then
            dbget.CommitTrans
        Else
            dbget.RollBackTrans
            response.write "<script>alert(" + Chr(34) + "����Ÿ�� �����ϴ� ���߿� ������ �߻��Ͽ����ϴ�. ������ ���� ���.(�����ڵ� : " + CStr(errcode) + ":" + Err.Description + "|" + ScanErr + ")" + Chr(34) + ")</script>"
            'response.write "<script>history.back()</script>"
            dbget.close()	:	response.End
        End If
    On error Goto 0


elseif (mode="delfinishedcsas") then
	'==========================================================================
    ''�Ϸ�� ���� ����

	ScanErr = ""

	set ocsaslist = New CCSASList
	ocsaslist.FRectCsAsID = id
	ocsaslist.GetOneCSASMaster_3PL

	rw "�۾��ȵǾ� ����!!"
	dbget.close : dbget_TPL.close : response.end

	if (ocsaslist.FOneItem.Fdeleteyn = "Y") then
	    response.write "<script>alert(" + Chr(34) + "�̹� ������ �����Դϴ�." + Chr(34) + ")</script>"
	    response.write "�̹� ������ �����Դϴ�."
	    dbget.close()	:	response.End
	end if

	if (divcd="A008") then
		'// ��ҿϷ�CS ����

		On Error Resume Next
			dbget.beginTrans

			If (Err.Number = 0) and (ScanErr="") Then
				errcode = "001"

				set oRefCSASList = new CCSASList
				oRefCSASList.FRectCsRefAsID = id
				oRefCSASList.GetOneCSASMaster

				if (oRefCSASList.FResultCount > 0) then
					if (oRefCSASList.FOneItem.Fdeleteyn = "N") then
						ScanErr = "���� ���� ȯ��CS �� �����ϼ���."
					end if
				end if
			end if

			If (Err.Number = 0) and (ScanErr="") Then
				errcode = "002"

				If Not DeleteFinishedCSProcess(id) then
					ScanErr = "������ ������ ����"
				else
					Call AddCsMemo(ocsaslist.FOneItem.Forderserial, "1", ocsaslist.FOneItem.Fuserid, session("ssBctId"), "��ҿϷ�CS ����")
					ResultMsg = ResultMsg + "->. [CSó���Ϸ�� ����] ó��\n\n"
				End if
			end if

			If (Err.Number = 0) and (ScanErr="") Then
				errcode = "005"

				If Not RestoreCancelProcess(id, orderserial) then
					ScanErr = "����ֹ� ������ ����"
				else
					Call AddCsMemo(ocsaslist.FOneItem.Forderserial, "1", ocsaslist.FOneItem.Fuserid, session("ssBctId"), "����ֹ� ����")
					ResultMsg = ResultMsg + "->. [����ֹ� ����] ó��\n\n"
				End if
			end if

			ResultMsg = "���� �Ǿ����ϴ�."
			ReturnUrl = "/admin/etc/3pl/cscenter/action/pop_cs_action_new_3PL.asp?id="  + CStr(id) + "&divcd=" + divcd + "&mode=editreginfo"

			If (Err.Number = 0) and (ScanErr="") Then
				dbget.CommitTrans
			Else
				dbget.RollBackTrans
				response.write "<script>alert(" + Chr(34) + "����Ÿ�� �����ϴ� ���߿� ������ �߻��Ͽ����ϴ�. ������ ���� ���.(�����ڵ� : " + CStr(errcode) + ":" + Err.Description + "|" + ScanErr + ")" + Chr(34) + ")</script>"
				'response.write "<script>history.back()</script>"
				dbget.close()	:	response.End
			End If
		On error Goto 0
	elseif (divcd="A004") or (divcd="A010") then
		'// ��ǰ�Ϸ�CS ����

		On Error Resume Next
			dbget.beginTrans

			refminusorderserial = ""

			If (Err.Number = 0) and (ScanErr="") Then
				errcode = "001"

				set oRefCSASList = new CCSASList
				oRefCSASList.FRectCsRefAsID = id
				oRefCSASList.GetOneCSASMaster

				if (oRefCSASList.FResultCount > 0) then
					refminusorderserial = oRefCSASList.FOneItem.Frefminusorderserial

					if (oRefCSASList.FOneItem.Fdeleteyn = "N") then
						ScanErr = "���� ���� ȯ��CS �� �����ϼ���."
					end if
				else
					refminusorderserial = ocsaslist.FOneItem.Frefminusorderserial
				end if
			end if

			If (Err.Number = 0) and (ScanErr="") Then
				errcode = "002"

				if (refminusorderserial = "") then ScanErr = "���̳ʽ� �ֹ���ȣ ���� - ������ ���� ���"
			end if

			If (Err.Number = 0) and (ScanErr="") Then
				errcode = "003"

				'// 0. ����� ������ �����Ұ�
				if (CheckJungsanExists(refminusorderserial) = True) then
					ScanErr = "�����Ұ� - ���곻���� �ֽ��ϴ�."
				end if
			end if

			If (Err.Number = 0) and (ScanErr="") Then
				errcode = "004"

				If Not DeleteFinishedCSProcess(id) then
					ScanErr = "������ ������ ����"
				else
					Call AddCsMemo(ocsaslist.FOneItem.Forderserial, "1", ocsaslist.FOneItem.Fuserid, session("ssBctId"), "��ǰ�Ϸ�CS ����")
					ResultMsg = ResultMsg + "->. [CSó���Ϸ�� ����] ó��\n\n"
				End if
			end if

			If (Err.Number = 0) and (ScanErr="") Then
				errcode = "005"

				if (Left(ocsaslist.FOneItem.Ffinishdate,7) < Left(Now(),7)) and ChkMaeipItemExist(id) then
					'// ���Ի�ǰ �ְ�, �Ϸ����ڰ� �������̸�
					ScanErr = "���Ի�ǰ�̰� ��ǰ���ڰ� �������̸� ��ǰ��ҺҰ�"
				Else
					If Not CancelMinusOrderProcess(refminusorderserial) then
						ScanErr = "������ ������ ����"
					else
						Call AddCsMemo(ocsaslist.FOneItem.Forderserial, "1", ocsaslist.FOneItem.Fuserid, session("ssBctId"), "���̳ʽ� �ֹ� ����")
						ResultMsg = ResultMsg + "->. [���̳ʽ� �ֹ� ����] ó��\n\n"
					End if
				End If
			end if

			ResultMsg = "CS�������� �� ���̳ʽ��ֹ����� �Ǿ����ϴ�."
			ReturnUrl = "/admin/etc/3pl/cscenter/action/pop_cs_action_new_3PL.asp?id="  + CStr(id) + "&divcd=" + divcd + "&mode=editreginfo"

			If (Err.Number = 0) and (ScanErr="") Then
				dbget.CommitTrans
			Else
				dbget.RollBackTrans
				response.write "<script>alert(" + Chr(34) + "����Ÿ�� �����ϴ� ���߿� ������ �߻��Ͽ����ϴ�. ������ ���� ���.(�����ڵ� : " + CStr(errcode) + ":" + Err.Description + "|" + ScanErr + ")" + Chr(34) + ")</script>"
				'response.write "<script>history.back()</script>"
				dbget.close()	:	response.End
			End If
		On error Goto 0
	elseif C_ADMIN_AUTH or (((divcd = "A700") or (divcd = "A005")) and (Left(ocsaslist.FOneItem.Fregdate,7) = Left(Now(),7))) then

		'// �Ǵ� ��ü��Ÿ���� ����Ϸ��
		'// �Ǵ� ������ȯ�� ����Ϸ��
		On Error Resume Next
			dbget.beginTrans

			If (Err.Number = 0) and (ScanErr="") Then
				errcode = "001"

				'' ����Ÿ ����
				Call DeleteFinishedCSForce(id)

				Call AddCsMemo(ocsaslist.FOneItem.Forderserial, "1", ocsaslist.FOneItem.Fuserid, session("ssBctId"), "�Ϸ�CS ����")
			end if

			ResultMsg = "���� �Ǿ����ϴ�."
			ReturnUrl = "/admin/etc/3pl/cscenter/action/pop_cs_action_new_3PL.asp?id="  + CStr(id) + "&divcd=" + divcd + "&mode=editreginfo"

			If (Err.Number = 0) and (ScanErr="") Then
				dbget.CommitTrans
			Else
				dbget.RollBackTrans
				response.write "<script>alert(" + Chr(34) + "����Ÿ�� �����ϴ� ���߿� ������ �߻��Ͽ����ϴ�. ������ ���� ���.(�����ڵ� : " + CStr(errcode) + ":" + Err.Description + "|" + ScanErr + ")" + Chr(34) + ")</script>"
				'response.write "<script>history.back()</script>"
				dbget.close()	:	response.End
			End If
		On error Goto 0

	else
        ResultMsg = "���ǵ��� �ʾҽ��ϴ�[2]. : mode=" + mode + " , divcd=" + divcd
        response.write "<script>alert('" + ResultMsg + "');</script>"
        ''response.write "<script>history.back();</script>"
        dbget.close()	:	response.End
	end if


elseif (mode="finishcsas") then
	'==========================================================================
    'CS ���� ���� �Ϸ�ó��

	set ocsaslist = New CCSASList
	ocsaslist.FRectCsAsID = id
	ocsaslist.GetOneCSASMaster_3PL

	if (ocsaslist.FOneItem.Fdeleteyn = "Y") then
	    response.write "<script>alert(" + Chr(34) + "�̹� ������ �����Դϴ�." + Chr(34) + ")</script>"
	    response.write "�̹� ������ �����Դϴ�."
	    dbget.close()	:	response.End
	elseif (ocsaslist.FOneItem.Fcurrstate = "B007") then
		response.write "<script>alert(" + Chr(34) + "�̹� �Ϸ�� �����Դϴ�." + Chr(34) + ")</script>"
		response.write "�̹� �Ϸ�� �����Դϴ�."
		dbget.close()	:	response.End
	end if

    if (divcd="A008") then
		'----------------------------------------------------------------------
		'CS ���� ���� �Ϸ�ó�� - �ֹ����

		rw "�۾��ȵǾ� ����!!"
		dbget.close : dbget_TPL.close : response.end

        On Error Resume Next
	    	dbget.beginTrans
			'// ������(2014-03-31)
	        ''if (GC_IsOLDOrder) then ScanErr = "6���� ���� ���� ���� ��� �Ұ� - ������ ���� ���"

	        If (Err.Number = 0) and (ScanErr="") Then
	            errcode = "001"

	            Call CancelProcess(id, orderserial)

				IsAllCancel = False
				if (remainitemcostsum = 0) then
					'��ü��� : �ܿ���ǰ�Ѿ��� ���� ���
            		IsAllCancel = True
				end if

				'// 2018-01-12, skyer9
				if (IsAllCancel) then
					''��ü ����ΰ��
					sqlStr = " exec [db_summary].[dbo].sp_Ten_RealtimeStock_cancelOrderAll '" & orderserial & "'"
					dbget.Execute sqlStr
				end if

				'// �����ǰ
				if (oordermaster.FOneItem.Fjumundiv = "3") then
					Call AddCsMemoRequest(ocsaslist.FOneItem.Forderserial, ocsaslist.FOneItem.Fuserid, "59", session("ssBctId"), "�����ǰ �߱��� ���")
				end if
	        End IF

	        If (Err.Number = 0) and (ScanErr="") Then
	            errcode = "008"

	            'ȯ�� ������ �ִ��� üũ �� ������ȯ��/���ϸ���ȯ��/�ſ�ī����� CS ���� ���
	            newasid = CheckNRegRefund(id, orderserial, reguserid)
	            if (newasid>0) then
	                ResultMsg = ResultMsg + "->. [ȯ�� ���] ó��\n\n"
	            end if
	        End If

	        If (Err.Number = 0) and (ScanErr="") Then
	            errcode = "009"

	            Call FinishCSMaster(id, reguserid, contents_finish)

				Call CheckAndCopyBonusCoupon(id, reguserid)
	        End If

	        ReturnUrl = "/admin/etc/3pl/cscenter/action/pop_cs_action_new_3PL.asp?id="  + CStr(id) + "&divcd=" + divcd + "&mode=editfinishedinfo"

	        If (Err.Number = 0) and (ScanErr="") Then
	            dbget.CommitTrans
	        Else
	            dbget.RollBackTrans
	            response.write "<script>alert(" + Chr(34) + "����Ÿ�� �����ϴ� ���߿� ������ �߻��Ͽ����ϴ�. ������ ���� ���.(�����ڵ� : " + CStr(errcode) + ":" + Err.Description + "|" + ScanErr + ")" + Chr(34) + ")</script>"
	            'response.write "<script>history.back()</script>"
	            dbget.close()	:	response.End
	        End If

	        ''�ֹ���� �Ϸ� ����
	        If (isCsMailSend) then
	            Call SendCsActionMail(id)

	            ''ȯ�� ���� ����
	            if (newasid>0) then
	                Call SendCsActionMail(newasid)
	            end if
	        End IF
        On error Goto 0
    elseif (divcd="A003") or (divcd="A007") then
    	'----------------------------------------------------------------------
        'CS ���� ���� �Ϸ�ó�� - ȯ�ҿ�û / ī��,��ü,�޴�����ҿ�û
        dim RefreturnMethod, Refrealrefund

		rw "�۾��ȵǾ� ����!!"
		dbget.close : dbget_TPL.close : response.end

        On Error Resume Next
	        dbget.beginTrans

	        If (Err.Number = 0) and (ScanErr="") Then
	            errcode = "001"

	            Call FinishCSMaster(id, reguserid, contents_finish)
	        End If

	        If (Err.Number = 0) and (ScanErr="") Then
	            errcode = "002"

				'���ϸ��� ȯ�� �� ��ġ����ȯ�� ������ ȯ���� ó��������, �� �ۿ� �ſ�ī��/������ ���� ȯ���� ���� ȯ�� ���μ������� ó���ȴ�.
				'���� �Ϸ�ó���Ѵٰ� �ؼ� ������ ȯ���� �Ͼ�� �ʴ´�.
	            Call CheckRefundFinish(id, orderserial, RefreturnMethod, Refrealrefund)
	        End If

	        ResultMsg = "ó�� �Ϸ�"
	        if (RefreturnMethod="R007") and (Refrealrefund>0) then
	            ReturnUrl = "/admin/etc/3pl/cscenter/action/pop_cs_action_new_3PL.asp?id="  + CStr(id) + "&divcd=" + divcd + "&finishtype=1"
	        else
	            ReturnUrl = "/admin/etc/3pl/cscenter/action/pop_cs_action_new_3PL.asp?id="  + CStr(id) + "&divcd=" + divcd
	        end if

	        If (Err.Number = 0) and (ScanErr="") Then
	            dbget.CommitTrans
	        Else
	            dbget.RollBackTrans
	            response.write "<script>alert(" + Chr(34) + "����Ÿ�� �����ϴ� ���߿� ������ �߻��Ͽ����ϴ�. ������ ���� ���.(�����ڵ� : " + CStr(errcode) + ":" + Err.Description + "|" + ScanErr + ")" + Chr(34) + ")</script>"
	            response.write "<script>history.back()</script>"
	            dbget.close()	:	response.End
	        End If

	        ''ȯ�� �Ϸ� ����
	        If (isCsMailSend) then
	            Call SendCsActionMail(id)
	        End IF
        On error Goto 0

    elseif (divcd="A004") or (divcd="A010") then
		'----------------------------------------------------------------------
        'CS ���� ���� �Ϸ�ó�� - ��ǰ����(��ü���)  // ȸ����û(�ٹ����ٹ��)
        dim MinusOrderserial

        On Error Resume Next
	        dbget_TPL.beginTrans

	        If (modeflag2<>"norefund") and (Err.Number = 0) and (ScanErr="") Then
	            errcode = "001"

	            ''���̳ʽ� �ֹ� ���
	            if (CheckNAddMinusOrder_3PL(id, orderserial, reguserid, MinusOrderserial, ScanErr)) then
					'// ���̳ʽ� �ֹ���ũ �ִ� �Լ��� üũ�Լ� ������ �̵�, skyer9, 2015-06-24
					'' Call AddminusOrderLink(id, MinusOrderserial)
	                ResultMsg = ResultMsg + "->. [��ǰ �ֹ�] ���\n\n"
	            end if
	        End If

            '//�ٹ� ��ǰȸ����û ����ó��
            '/���̳ʽ�/ȯ�ҿ�û �ִ°�
            if modeflag2<>"norefund" and divcd="A010" then
		        If (Err.Number = 0) and (ScanErr="") Then
		            errcode = "011"

					'/���� �ø���, �Ǹ� ���� �����Ѵ�.	'/2016.07.15 �ѿ�� ����
					''call setItemLimitcs(id, orderserial, "UP")

					''ResultMsg = ResultMsg + "->. [���� ó��] �ٹ� ��ǰȸ�� ����ó�� �Ϸ�\n\n"
		        End If
			End If

	        If (Err.Number = 0) and (ScanErr="") Then
	            errcode = "009"

	            Call FinishCSMaster_3PL(id, reguserid, contents_finish)

	            if (divcd="A004") then
	                ResultMsg = ResultMsg + "->. ��ǰ ó�� �Ϸ�\n\n"
	            elseif (divcd="A010") then
	                ResultMsg = ResultMsg + "->. ȸ�� ó�� �Ϸ�\n\n"
	            end if
	        End If

	        ResultMsg = ResultMsg
	        ReturnUrl = "/admin/etc/3pl/cscenter/action/pop_cs_action_new_3PL.asp?id="  + CStr(id) + "&divcd=" + divcd
	        If (Err.Number = 0) and (ScanErr="") Then
	            dbget_TPL.CommitTrans

				sqlStr = " exec [db_summary].[dbo].[usp_TPL_RealtimeStock_minusOrder] '" & MinusOrderserial & "'"
				dbget.Execute sqlStr
	        Else
	            dbget_TPL.RollBackTrans
	            response.write "<script>alert(" + Chr(34) + "����Ÿ�� �����ϴ� ���߿� ������ �߻��Ͽ����ϴ�. ������ ���� ���.(�����ڵ� : " + CStr(errcode) + ":" + Err.Description + "|" + ScanErr + ")" + Chr(34) + ")</script>"
	            response.write "<script>history.back()</script>"
	            dbget.close() : dbget_TPL.close()	:	response.End
	        End If

        On error Goto 0
    elseif  (divcd="A011") or (divcd="A012") then
    	'----------------------------------------------------------------------
    	'CS ���� ���� �Ϸ�ó�� - �±�ȯȸ��(�ٹ����ٹ��), �±�ȯȸ��(��ü���)
        On Error Resume Next
	        dbget_TPL.beginTrans

			'//�ٹ� ��ȯȸ�� ����ó��
			if divcd="A011" then
		        If (Err.Number = 0) and (ScanErr="") Then
		            errcode = "011"

					'/���� �ø���, �Ǹ� ���� �����Ѵ�.	'/2016.07.15 �ѿ�� ����
					''call setItemLimitcs(id, orderserial, "UP")

					''ResultMsg = ResultMsg + "->. [���� ó��] �ٹ� ��ȯȸ�� ����ó�� �Ϸ�\n\n"
		        End If
			End If

	        If (Err.Number = 0) and (ScanErr="") Then
	            errcode = "009"

	            Call FinishCSMaster_3PL(id, reguserid, contents_finish)
	        End If

	        ResultMsg = ResultMsg + "ó�� �Ϸ�"
	        ReturnUrl = "/admin/etc/3pl/cscenter/action/pop_cs_action_new_3PL.asp?id="  + CStr(id) + "&divcd=" + divcd

	        If (Err.Number = 0) and (ScanErr="") Then
				dbget_TPL.CommitTrans
	        Else
	            dbget_TPL.RollBackTrans
	            response.write "<script>alert(" + Chr(34) + "����Ÿ�� �����ϴ� ���߿� ������ �߻��Ͽ����ϴ�. ������ ���� ���.(�����ڵ� : " + CStr(errcode) + ":" + Err.Description + "|" + ScanErr + ")" + Chr(34) + ")</script>"
	           ' response.write "<script>history.back()</script>"
	            dbget_TPL.close()	:	response.End
	        End If

        On error Goto 0
    elseif (divcd="A111") or (divcd="A112") then
    	'----------------------------------------------------------------------
    	'CS ���� ���� �Ϸ�ó�� - ��ǰ���� �±�ȯȸ��(�ٹ����ٹ��), ��ǰ���� �±�ȯȸ��(��ü���)

		rw "�۾��ȵǾ� ����!!"
		dbget.close : dbget_TPL.close : response.end

        On Error Resume Next
	        dbget.beginTrans

	        If (Err.Number = 0) and (ScanErr="") Then
	            errcode = "009"

				'// �Լ� �ȿ��� CS����� �Ѵ�.
	            Call FinishCSMaster(id, reguserid, contents_finish)
	        End If

	        If (Err.Number = 0) and (ScanErr="") Then
	            errcode = "010"

				Call GetChangeOrderInfo(id, changeorderserial, changeorderstate,  ScanErr)

				if (ScanErr = "") then
					if (changeorderserial = "") then
						'// ��ȯ�ֹ� ����Ѵ�.
						'// �±�ȯ��� ���¿� �����ϰ� �±�ȯȸ���� �Ǹ� ��ȯ�ֹ� ����Ѵ�.
						'// �ٹ��� ��� ���� �±�ȯ�� ��� ȸ�����Ŀ� �±�ȯ����Ѵ�(http://logics.10x10.co.kr/v2/online/m_re_chulgo.asp ����)
						changeorderserial = CheckAndAddChangeOrder(id, orderserial, ScanErr)

			            if (changeorderserial <> "") then
			            	Call AddChangeOrderLink(id, changeorderserial)
			                ResultMsg = ResultMsg + "->. [��ǰ���� �±�ȯ ��ȯ�ֹ�] ���Ϸ� ���\n\n"
			            end if
					else
						if (changeorderstate <> "8") then
							Call FinishChangeOrder(changeorderserial)
			            	Call AddChangeOrderChulgoLink(id, changeorderserial)
			                ResultMsg = ResultMsg + "->. [��ǰ���� �±�ȯ ��ȯ�ֹ�] ���Ϸ� ��ȯ\n\n"
						end if
					end if
				end if

	        End If

            '//�ٹ� ��ȯȸ�� ��ǰ���� ����ó��
            if divcd="A111" then
		        If (Err.Number = 0) and (ScanErr="") Then
		            errcode = "011"

					'/���� �ø���, �Ǹ� ���� �����Ѵ�.	'/2016.07.15 �ѿ�� ����
					call setItemLimitcs(id, orderserial, "UP")

					ResultMsg = ResultMsg + "->. [���� ó��] �ٹ� ��ȯȸ�� ��ǰ���� ����ó�� �Ϸ�\n\n"
		        End If
			End If

			ResultMsg = ResultMsg + "ó�� �Ϸ�"
	        ReturnUrl = "/admin/etc/3pl/cscenter/action/pop_cs_action_new_3PL.asp?id="  + CStr(id) + "&divcd=" + divcd

	        If (Err.Number = 0) and (ScanErr="") Then
	            dbget.CommitTrans
	        Else
	            dbget.RollBackTrans
	            response.write "<script>alert(" + Chr(34) + "����Ÿ�� �����ϴ� ���߿� ������ �߻��Ͽ����ϴ�. ������ ���� ���.(�����ڵ� : " + CStr(errcode) + ":" + Err.Description + "|" + ScanErr + ")" + Chr(34) + ")</script>"
	           ' response.write "<script>history.back()</script>"
	            dbget.close()	:	response.End
	        End If

	        ''�±�ȯ �Ϸ� ����
	        If (isCsMailSend) and (divcd <> "A112") then
	        	'// ��ü����� ������ �ʴ´�.
   	            Call SendCsActionMail(id)
	        End If
        On error Goto 0
    elseif  (divcd="A000") or (divcd="A100") or (divcd="A001") or (divcd="A002") or (divcd="A200") or (divcd="A009") or (divcd="A006") or (divcd="A005") or (divcd="A700") then
    	'----------------------------------------------------------------------
        'CS ���� ���� �Ϸ�ó�� - �±�ȯ ��� / ��ǰ���� �±�ȯ��� / ���� / ���� �߼� / ��Ÿ /  ���� ���ǻ���

		''rw "aaaa"
		''dbget.close : dbget_TPL.close : response.end

        On Error Resume Next
	        dbget_TPL.BeginTrans

	        If (divcd="A100") and (modeflag2<>"norefund") and (Err.Number = 0) and (ScanErr="") Then
	            errcode = "09"

	            ''newasid = CheckNRegRefund(id, orderserial, reguserid)

	            if (newasid>0) then
	                ResultMsg = ResultMsg + "->. [���(ȯ��)����] ó��\n\n"
	            end if
	        End If

	        If (Err.Number = 0) and (ScanErr="") Then
	            errcode = "010"

	            Call FinishCSMaster_3PL(id, reguserid, contents_finish)
	        End If


	        ResultMsg = "ó�� �Ϸ�"
	        ReturnUrl = "/admin/etc/3pl/cscenter/action/pop_cs_action_new_3PL.asp?id="  + CStr(id) + "&divcd=" + divcd

	        If (Err.Number = 0) and (ScanErr="") Then
	            dbget_TPL.CommitTrans
	        Else
	            dbget_TPL.RollBackTrans
	            response.write "<script>alert(" + Chr(34) + "����Ÿ�� �����ϴ� ���߿� ������ �߻��Ͽ����ϴ�. ������ ���� ���.(�����ڵ� : " + CStr(errcode) + ":" + Err.Description + "|" + ScanErr + ")" + Chr(34) + ")</script>"
	           ' response.write "<script>history.back()</script>"
	            dbget_TPL.close()	:	response.End
	        End If

        On error Goto 0
    else
        ResultMsg = "���ǵ��� �ʾҽ��ϴ�[3]. : mode=" + mode + " , divcd=" + divcd
        response.write "<script>alert('" + ResultMsg + "');</script>"
        response.write "<script>history.back();</script>"
        dbget.close()	:	response.End
    end if



elseif (mode="state2jupsu") then

	rw "�۾��ȵǾ� ����!!"
	dbget.close : dbget_TPL.close : response.end

	'==========================================================================
    if (divcd="A700") then
    	'----------------------------------------------------------------------
        '' ��ü ��Ÿ���� �������·� ����
        sqlStr = " select top 1 * from db_jungsan.dbo.tbl_designer_jungsan_detail"
        sqlStr = sqlStr + " where gubuncd='witakchulgo'"
        sqlStr = sqlStr + " and detailidx<>0"
        sqlStr = sqlStr + " and itemid=0"
        sqlStr = sqlStr + " and detailidx=" & id

        rsget.Open sqlStr,dbget,1
	        if not rsget.Eof then
			    ResultMsg = "���� ������ �����մϴ�. ���� ���� �Ұ�"
			else
			    ResultMsg = ""
			end if
		rsget.Close

        if (ResultMsg="") then
            sqlStr = " update db_cs.dbo.tbl_new_as_list"
            sqlStr = sqlStr + " set currstate='B001'"
            sqlStr = sqlStr + " ,finishdate=NULL"
            sqlStr = sqlStr + " where id=" & CStr(id)
            sqlStr = sqlStr + " and currstate='B007'"
            'response.write sqlStr
            dbget.Execute sqlStr

            ResultMsg = "ó�� �Ϸ�"
            ReturnUrl = "/admin/etc/3pl/cscenter/action/pop_cs_action_new_3PL.asp?id="  + CStr(id) + "&divcd=" + divcd
        else
            response.write "<script>alert('" + ResultMsg + "');</script>"
            response.write "<script>history.back();</script>"
            dbget.close()	:	response.End
        end if
    else
        ResultMsg = "���ǵ��� �ʾҽ��ϴ�[4]. : mode=" + mode + " , divcd=" + divcd
        response.write "<script>alert('" + ResultMsg + "');</script>"
        response.write "<script>history.back();</script>"
        dbget.close()	:	response.End
    end if



elseif (mode="addupchejungsanEdit") then

	rw "�۾��ȵǾ� ����!!"
	dbget.close : dbget_TPL.close : response.end

	'==========================================================================
    '' ��ü ��Ÿ���� �������·� ����
    sqlStr = " select top 1 * from db_jungsan.dbo.tbl_designer_jungsan_detail"
    sqlStr = sqlStr + " where gubuncd='witakchulgo'"
    sqlStr = sqlStr + " and detailidx<>0"
    sqlStr = sqlStr + " and itemid=0"
    sqlStr = sqlStr + " and detailidx=" & id

    rsget.Open sqlStr,dbget,1
	    if not rsget.Eof then
		    ResultMsg = "���� ������ �����մϴ�. ���� �Ұ�"
		else
		    ResultMsg = ""
		end if
	rsget.Close

    if (ResultMsg="") then
        if (add_upchejungsandeliverypay<>"") then
            call EditCSUpcheAddJungsanPay(id, add_upchejungsandeliverypay, add_upchejungsancause, buf_requiremakerid)
        end if

        ResultMsg = "ó�� �Ϸ�"
        ReturnUrl = "/cscenter/action/pop_AddUpchejungsanEdit.asp?id="  + CStr(id)
    else
        response.write "<script>alert('" + ResultMsg + "');</script>"
        response.write "<script>history.back();</script>"
        dbget.close()	:	response.End
    end if



elseif (mode="upcheconfirm2jupsu") then

	rw "�۾��ȵǾ� ����!!"
	dbget.close : dbget_TPL.close : response.end

	'==========================================================================
    '' ��ü ó���Ϸ� => �������·κ���
    sqlStr = " select top 1 currstate from [db_cs].[dbo].tbl_new_as_list"
    sqlStr = sqlStr + " where id=" + CStr(id)

    rsget.Open sqlStr,dbget,1
	    if not rsget.Eof then
	        ResultMsg = ""
	        if (rsget("currstate")<>"B006") then
	            ResultMsg = "��ü ó�� �Ϸ� ���°� �ƴմϴ�. ���� �Ұ�"
	        end if
		else
		    ResultMsg = "�ڵ����. ���� �Ұ�"
		end if
	rsget.Close

    if (ResultMsg="") then
        sqlStr = " update [db_cs].[dbo].tbl_new_as_list" + VbCrlf
        sqlStr = sqlStr + "set currstate='B001', confirmdate=NULL" + VbCrlf
        sqlStr = sqlStr + ",contents_jupsu='" + (contents_jupsu) + "'" + VbCrlf
        sqlStr = sqlStr + " where id=" + CStr(id)
        dbget.Execute sqlStr

        sqlStr = " update [db_cs].[dbo].tbl_new_as_detail" + VbCrlf
        sqlStr = sqlStr + "set currstate='B001'" + VbCrlf
        sqlStr = sqlStr + " where masterid=" + CStr(id)
        dbget.Execute sqlStr

        ResultMsg = "ó�� �Ϸ�"
        ReturnUrl = "/admin/etc/3pl/cscenter/action/pop_cs_action_new_3PL.asp?id="  + CStr(id) + "&divcd=" + divcd
    else
        response.write "<script>alert('" + ResultMsg + "');</script>"
        response.write "<script>history.back();</script>"
        dbget.close()	:	response.End
    end if



elseif (mode="changeorderreg") then
	'==========================================================================
    '' ��ȯ�ֹ� �������

	rw "�۾��ȵǾ� ����!!"
	dbget.close : dbget_TPL.close : response.end

	Call GetChangeOrderInfo(id, changeorderserial, changeorderstate,  ResultMsg)

	if (ResultMsg="") and (changeorderserial <> "") then
		ResultMsg = "��ȯ�ֹ��� �̹� ��ϵǾ� �ֽ��ϴ�."
	end if

    if (ResultMsg="") then
		'// ��ȯ�ֹ� ����Ѵ�.
		'// ��ȯ��� �� ȸ�����¿� �����ϰ� �������Ѵ�.(�ֹ���������)
		'// �ٹ��� ��� ���� �±�ȯ�� ��� ȸ�����Ŀ� �±�ȯ����Ѵ�(http://logics.10x10.co.kr/v2/online/m_re_chulgo.asp ����)
		changeorderserial = CheckAndAddChangeOrderJupsu(id, orderserial, ScanErr)

        if (changeorderserial <> "") then
        	Call AddChangeOrderJupsuLink(id, changeorderserial)
            ResultMsg = ResultMsg + "->. [��ǰ���� �±�ȯ ��ȯ�ֹ�] �ֹ����� ���\n\n"
        end if

        ResultMsg = ResultMsg + "ó�� �Ϸ�"
        ReturnUrl = "/admin/etc/3pl/cscenter/action/pop_cs_action_new_3PL.asp?id="  + CStr(id) + "&divcd=" + divcd
    else
        response.write "<script>alert('" + ResultMsg + "');</script>"
        response.write "<script>history.back();</script>"
        dbget.close()	:	response.End
    end if

elseif (mode="changedivcdtoa004") then

	rw "�۾��ȵǾ� ����!!"
	dbget.close : dbget_TPL.close : response.end

	'==========================================================================
    '' �� ������ǰ ��ȯ(A010 -> A004)

    sqlStr = " select top 1 currstate, deleteyn from [db_cs].[dbo].tbl_new_as_list"
    sqlStr = sqlStr + " where id=" + CStr(id)

    rsget.Open sqlStr,dbget,1
	    if not rsget.Eof then
	        ResultMsg = ""
	        if (rsget("deleteyn")="Y") then
	            ResultMsg = "������ �����Դϴ�. ���� �Ұ�"
	        else
		        if (rsget("currstate")<>"B001") then
		            ResultMsg = "�̹� �ù�翡 ���۵� �����Դϴ�. ���� �Ұ�"
		        end if
	        end if
		else
		    ResultMsg = "�ڵ����. ���� �Ұ�"
		end if
	rsget.Close

    if (ResultMsg="") then
    	divcd = "A004"

        sqlStr = " update [db_cs].[dbo].tbl_new_as_list" + VbCrlf
        sqlStr = sqlStr + "set divcd='" + CStr(divcd) + "'" + VbCrlf
        sqlStr = sqlStr + ", requireupche='Y' " + VbCrlf
        sqlStr = sqlStr + ", makerid='10x10logistics' " + VbCrlf
        sqlStr = sqlStr + ", title='�� ������ǰ ��ȯ' " + VbCrlf
        sqlStr = sqlStr + ", opentitle='��ǰ����(��ü���)' " + VbCrlf
        sqlStr = sqlStr + " where id=" + CStr(id)
        dbget.Execute sqlStr

        ResultMsg = "ó�� �Ϸ�"
        ReturnUrl = "/admin/etc/3pl/cscenter/action/pop_cs_action_new_3PL.asp?id="  + CStr(id) + "&divcd=" + divcd
    else
        response.write "<script>alert('" + ResultMsg + "');</script>"
        response.write "<script>history.back();</script>"
        dbget.close()	:	response.End
    end if



elseif (mode="changedivcdtoa010") then
	'==========================================================================
    '' ȸ����û ��ȯ(A004 -> A010)

	rw "�۾��ȵǾ� ����!!"
	dbget.close : dbget_TPL.close : response.end

    sqlStr = " select top 1 currstate, deleteyn from [db_cs].[dbo].tbl_new_as_list"
    sqlStr = sqlStr + " where id=" + CStr(id)

    rsget.Open sqlStr,dbget,1
	    if not rsget.Eof then
	        ResultMsg = ""
	        if (rsget("deleteyn")="Y") then
	            ResultMsg = "������ �����Դϴ�. ���� �Ұ�"
	        else
		        if (rsget("currstate")<>"B001") then
		            ResultMsg = "�̹� �ù�翡 ���۵� �����Դϴ�. ���� �Ұ�"
		        end if
	        end if
		else
		    ResultMsg = "�ڵ����. ���� �Ұ�"
		end if
	rsget.Close

    if (ResultMsg="") then
    	divcd = "A010"

        sqlStr = " update [db_cs].[dbo].tbl_new_as_list" + VbCrlf
        sqlStr = sqlStr + "set divcd='" + CStr(divcd) + "'" + VbCrlf
        sqlStr = sqlStr + ", requireupche='N' " + VbCrlf
        sqlStr = sqlStr + ", makerid=NULL " + VbCrlf
        sqlStr = sqlStr + ", title='ȸ����û ��ȯ' " + VbCrlf
        sqlStr = sqlStr + ", opentitle='ȸ����û(�ٹ����ٹ��)' " + VbCrlf
        sqlStr = sqlStr + " where id=" + CStr(id)
        dbget.Execute sqlStr

        ResultMsg = "ó�� �Ϸ�"
        ReturnUrl = "/admin/etc/3pl/cscenter/action/pop_cs_action_new_3PL.asp?id="  + CStr(id) + "&divcd=" + divcd
    else
        response.write "<script>alert('" + ResultMsg + "');</script>"
        response.write "<script>history.back();</script>"
        dbget.close()	:	response.End
    end if



else
	'==========================================================================
    ResultMsg = "���ǵ��� �ʾҽ��ϴ�[5]. : mode=" + mode + " , divcd=" + divcd
    response.write "<script>alert('" + ResultMsg + "');</script>"
    response.write "<script>history.back();</script>"
    dbget.close()	:	response.End
end if

'=============================================================
'			���� SMS �߼� ����
'=============================================================
'' ���� �߰��� ����

''dim isMailProc '// ���� �߼ۿ���
''dim isSmsProc	'// SMS �߼� ����
''
''isMailProc = False
''isSmsProc = False
''
''IF mode="regcsas" THEN
''	IF divcd ="A000" or divcd ="A001" or divcd ="A002" or divcd ="A003" or divcd ="A004" or divcd ="A007" or divcd ="A010" THEN
''		isMailProc=True
''	END IF
''ELSEIF mode="finishcsas" THEN
''	IF divcd ="A000" or divcd ="A001" or divcd ="A002" or divcd ="A003" or divcd ="A004" or divcd ="A007" or divcd ="A008" or divcd ="A010" or divcd="A900" THEN
''		isMailProc=True
''	END IF
''END IF
''
''IF mode="regcsas" THEN
''	IF divcd ="A000" or divcd ="A001" or divcd ="A002" or divcd ="A003" or divcd ="A004" or divcd ="A007" or divcd ="A010" THEN
''		isSmsProc= True
''	END IF
''
''ELSEIF mode="finishcsas" THEN
''	IF divcd ="A000" or divcd ="A001" or divcd ="A002" or divcd ="A003" or divcd ="A004" or divcd ="A007" or divcd ="A008" or divcd ="A010" or divcd="A900" THEN
''		isSmsProc= True
''	END IF
''END IF
''
'''//=======  SMS �߼� ���� =========/
''IF isSmsProc THEN
''	'oCsAction.sendSMS "",""
''	'oCsAction.sendSMS "010-8831-6240",""
''END IF


if (mode <> "regcsas") and (id <> "") then
	'// ���� ó���� ���̵� ����
	Call SaveCSListHistory(id)
end if

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db_TPLClose.asp" -->
<%
response.write "<script>alert('" + ResultMsg + "');</script>"
response.write "<script>location.replace('" + ReturnUrl + "');</script>"
%>


<%

''''    	If (Err.Number = 0) and (ScanErr="") Then
''''            errcode = "003"
''''
''''            if (divcd="A020") then
''''        	    ''��ü ����ΰ��
''''        	    ''1- ��ü ��� ���� ��ȿ�� üũ
''''        	    if Not (IsAllCancelRegValid(id, orderserial)) then
''''        	        ScanErr = "��ü ��� ���� ����. - ��ü ��� �ƴ�."
''''        	    end if
''''
''''
''''        	elseif (divcd="A021") then
''''        	    ''�κ� ����ΰ��
''''        	    ''1- �κ� ��� ���� ��ȿ�� üũ
''''        	    if Not (IsPartialCancelRegValid(id, orderserial)) then
''''        	        ScanErr = "��ü ��� ���� ����. - �κ� ��� �ƴϰų� ��������."
''''        	    end if
''''        	end if
''''        end if
%>
