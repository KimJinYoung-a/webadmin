<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
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

''��� ����
dim refundmileagesum, refundcouponsum, allatsubtractsum
dim refunditemcostsum, canceltotal, nextsubtotal
dim refundbeasongpay, recalcubeasongpay, refunddeliverypay, refundadjustpay

''ȯ�� ���� maybe (refundrequire==canceltotal)
dim refundrequire, returnmethod
dim rebankname, rebankaccount, rebankownername, paygateTid

''��ü �߰� �����
dim add_upchejungsandeliverypay, add_upchejungsancause, add_upchejungsancauseText

''���ֹ� �ݾ�
dim orgitemcostsum, orgbeasongpay, orgmileagesum, orgcouponsum, orgallatdiscountsum, orgsubtotalprice

''�� Open msg
dim opentitle, opencontents

''�߰�����ID
dim buf_requiremakerid

''�߰��� ��ϵ� CSID
dim newasid

''CS���Ϲ߼�����
dim isCsMailSend

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
contents_jupsu  = html2DB(request.Form("contents_jupsu"))
detailitemlist  = html2db(request.Form("detailitemlist"))
contents_finish = html2DB(request.Form("contents_finish"))

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


refunditemcostsum   = request.Form("refunditemcostsum")
nextsubtotal        = request.Form("nextsubtotal")
canceltotal         = request.Form("canceltotal")

refundbeasongpay    = request.Form("refundbeasongpay")
recalcubeasongpay   = request.Form("recalcubeasongpay")
refunddeliverypay   = request.Form("refunddeliverypay")

refundmileagesum    = request.Form("refundmileagesum")
refundcouponsum     = request.Form("refundcouponsum")
allatsubtractsum    = request.Form("allatsubtractsum")
refundadjustpay     = request.Form("refundadjustpay")


''ȯ�ҿ�û��
refundrequire       = request.Form("refundrequire")
returnmethod        = request.Form("returnmethod")

rebankname          = request.Form("rebankname")
rebankaccount       = request.Form("rebankaccount")
rebankownername     = request.Form("rebankownername")

paygateTid          = request.Form("paygateTid")

add_upchejungsandeliverypay = request.Form("add_upchejungsandeliverypay")
add_upchejungsancause       = request.Form("add_upchejungsancause")
add_upchejungsancauseText   = request.Form("add_upchejungsancauseText")

buf_requiremakerid  = request.Form("buf_requiremakerid")


isCsMailSend = (request.Form("csmailsend")="on")

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
if (Not IsNumeric(recalcubeasongpay)) or (recalcubeasongpay="") then recalcubeasongpay = 0
if (Not IsNumeric(refunddeliverypay)) or (refunddeliverypay="") then refunddeliverypay = 0

if (Not IsNumeric(refundadjustpay)) or (refundadjustpay="") then refundadjustpay = 0
if (Not IsNumeric(canceltotal)) or (canceltotal="") then canceltotal = 0
if (Not IsNumeric(refundrequire)) or (refundrequire="") then refundrequire = 0

if (returnmethod="") then returnmethod="R000"

''�ÿ�ī������.. -��ǰ���� ����.

dim sqlStr, errcode, i
dim ScanErr
dim ResultMsg, ReturnUrl, EtcStr
dim ProceedFinish

ScanErr = ""
ProceedFinish = False

dim IsAllCancel

''���� �ֹ� �������� Check
GC_IsOLDOrder = CheckIsOldOrder(orderserial)



if (mode="regcsas") then
    '==========================================================================
	'CS ����
    if (divcd="A008") then
		'----------------------------------------------------------------------
        'CS ���� - �ֹ����
        'On Error Resume Next
        dbget.beginTrans

        if (GC_IsOLDOrder) then ScanErr = "6���� ���� ���� ���� ��� �Ұ� - ������ ���� ���"

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "001"

            'CS Master ����
            id = RegCSMaster(divcd, orderserial, reguserid, title, contents_jupsu, gubun01, gubun02)
        end if

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "002"

            'CS Master ȯ�� �������� ����
	        Call RegCSMasterRefundInfo(id, returnmethod, refundrequire , orgsubtotalprice, orgitemcostsum, orgbeasongpay , orgmileagesum, orgcouponsum, orgallatdiscountsum  , canceltotal, refunditemcostsum, refundmileagesum, refundcouponsum, allatsubtractsum , refundbeasongpay, refunddeliverypay, refundadjustpay , rebankname, rebankaccount, rebankownername, paygateTid)
	    End if

	    If (Err.Number = 0) and (ScanErr="") Then
            errcode = "003"

            'CS Detail ����(���� ��ǰ����)
	        Call AddCSDetailByArrStr(detailitemlist, id, orderserial)

        end if

		'''����..
	    If (Err.Number = 0) and (ScanErr="") Then
            errcode = "005"
            ''��ü ������� ���� Ȯ�� - AsDetail �Է��� �˻� �ؾ� ��.
        	IsAllCancel     = IsAllCancelRegValid(id, orderserial)

        	if (IsAllCancel) And (orgsubtotalprice<>canceltotal) then
        	    ScanErr = "��� �ݾ� ����ġ - ��ü��ҽ� ��ұݾװ� �����ݾ��� ��ġ�ؾ���"
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
            End If
        ELSE
            ResultMsg = ResultMsg + "->. ��ǰ �غ��� ������ ��ǰ�� �����ϹǷ�\n\n �ֹ� ��� ������ ���� �Ǿ����ϴ�.\n\n Ȯ���� �Ϸ� ó���ϼž� �մϴ�."
        End If

	    If (Err.Number = 0) and (ScanErr="") Then
            dbget.CommitTrans
        Else
            dbget.RollBackTrans
            response.write "<script>alert(" + Chr(34) + "����Ÿ�� �����ϴ� ���߿� ������ �߻��Ͽ����ϴ�. ������ ���� ���.(�����ڵ� : " + CStr(errcode) + ":" + Err.Description + "|" + ScanErr + ")" + Chr(34) + ")</script>"
            ''response.write "<script>history.back()</script>"
            dbget.close()	:	response.End
        End If

        ''������� �ݾ�/������ ����
        Call CheckNChangeCyberAcct(orderserial)

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

        ReturnUrl = "/cscenter/action/pop_cs_action_reg.asp?id="  + CStr(id) + "&divcd=" + divcd + "&mode=editreginfo"

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

	            if (add_upchejungsandeliverypay<>"0") and (add_upchejungsandeliverypay<>"")  then
	                call RegCSUpcheAddJungsanPay(id, add_upchejungsandeliverypay, add_upchejungsancause, buf_requiremakerid)
	            end if
	        end if

	        ReturnUrl = "/cscenter/action/pop_cs_action_reg.asp?id="  + CStr(id) + "&divcd=" + divcd

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

    elseif (divcd="A001") or (divcd="A002") then
    	'----------------------------------------------------------------------
        'CS ���� - ������߼�, ���񽺹߼�
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

	        ResultMsg = "�����Ϸ�"
	        ReturnUrl = "/cscenter/action/pop_cs_action_reg.asp?id="  + CStr(id) + "&divcd=" + divcd

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
	            Call SendCsActionMail(id)
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

	            if (requiremakerid<>"") then
	                '��ü����� ��� ���� ��ü �귣�� ���̵� �Է�(requiremakerid)

	                call RegCSMasterAddUpche(id, requiremakerid)

	                ResultMsg = "�±�ȯ �����Ϸ� - ��ü���"
	            else
	                '�ٹ����� ����� ��� �±�ȯ ȸ�� ����
	                newasid = RegCSMaster("A011", orderserial, reguserid, "�±�ȯ ȸ������", contents_jupsu, gubun01, gubun02)

	                Call AddCSDetailByArrStr(detailitemlist, newasid, orderserial)

	                 ResultMsg = "�±�ȯ ��� ���� �� ȸ������ �Ϸ� - �ٹ����� ���"
	            end if
	        end if

	        ''��ü �߰� ����� 2008.11.10
	        If (Err.Number = 0) and (ScanErr="") Then
	            errcode = "004"

	            if (add_upchejungsandeliverypay<>"0") and (add_upchejungsandeliverypay<>"")  then
	                call RegCSUpcheAddJungsanPay(id, add_upchejungsandeliverypay, add_upchejungsancause, buf_requiremakerid)
	            end if
	        end if

	        ReturnUrl = "/cscenter/action/pop_cs_action_reg.asp?id="  + CStr(id) + "&divcd=" + divcd

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
	            if (newasid>0) then
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

	            if (add_upchejungsandeliverypay<>"0") and (add_upchejungsandeliverypay<>"")  then
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
	        ReturnUrl = "/cscenter/action/pop_cs_action_reg.asp?id="  + CStr(id) + "&divcd=" + divcd

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
					elseif (returnmethod = "R100") then
						title = title & "(�ſ�ī�����)"
					elseif (returnmethod = "R007") then
						title = title & "(������)"
					end if
				end if

	            id = RegCSMaster(divcd, orderserial, reguserid, title, contents_jupsu, gubun01, gubun02)
	        end if

	        If (Err.Number = 0) and (ScanErr="") Then
	            errcode = "002"

	            'CS Master ȯ�� �������� ����
		        Call RegCSMasterRefundInfo(id, returnmethod, refundrequire , orgsubtotalprice, orgitemcostsum, orgbeasongpay , orgmileagesum, orgcouponsum, orgallatdiscountsum  , canceltotal, refunditemcostsum, refundmileagesum, refundcouponsum, allatsubtractsum , refundbeasongpay, refunddeliverypay, refundadjustpay , rebankname, rebankaccount, rebankownername, paygateTid)
		    End if


	        ResultMsg = "��ϵǾ����ϴ�."
	        ReturnUrl = "/cscenter/action/pop_cs_action_reg.asp?id="  + CStr(id) + "&divcd=" + divcd

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
        ResultMsg = "���ǵ��� �ʾҽ��ϴ�. : mode=" + mode + " , divcd=" + divcd
        response.write "<script>alert('" + ResultMsg + "');</script>"
        response.write "<script>history.back();</script>"
        dbget.close()	:	response.End
    end if



elseif (mode="deletecsas") then
	'==========================================================================
	'CS ����
    On Error Resume Next
        dbget.beginTrans

        ''Check Valid Delete - ����� B006 ��üó���Ϸ� , B007 �Ϸ� ������ ���(����) �Ұ�
        if (NOT ValidDeleteCS(id)) then
            response.write "<script>alert(" + Chr(34) + "���� ��� ���� ���°� �ƴմϴ�. ������ ���� ���." + Chr(34) + ")</script>"
            response.write "<script>history.back()</script>"
            dbget.close()	:	response.End
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
        ReturnUrl = "/cscenter/action/pop_cs_action_reg.asp?id="  + CStr(id) + "&divcd=" + divcd + "&mode=editreginfo"

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
    On Error Resume Next
        dbget.beginTrans

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "001"

            '' CS Master ����
            Call EditCSMaster(divcd, orderserial, reguserid, title, contents_jupsu, gubun01, gubun02)

        end if

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "002"

            '' CS Detail ����
            Call EditCSDetailByArrStr(detailitemlist, id, orderserial)
        End if

        ResultMsg = ResultMsg + "->. [CS ó���� ����] ó��\n\n"

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "003"

            '' ȯ�� ���� ����
            if (CheckNEditRefundInfo(id,returnmethod,rebankaccount,rebankownername,rebankname,paygateTid,refundrequire, orgsubtotalprice, orgitemcostsum, orgbeasongpay , orgmileagesum, orgcouponsum, orgallatdiscountsum , canceltotal, refunditemcostsum, refundmileagesum, refundcouponsum, allatsubtractsum , refundbeasongpay, refunddeliverypay, refundadjustpay)) then
                ResultMsg = ResultMsg + "->. [ȯ������ ����] ó��\n\n"
            end if
        end If

        ''��ü �߰� ����� 2008.11.10
        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "004"

            if (add_upchejungsandeliverypay<>"") then
                call EditCSUpcheAddJungsanPay(id, add_upchejungsandeliverypay, add_upchejungsancause, buf_requiremakerid)
            end if
        end if

        ReturnUrl = "/cscenter/action/pop_cs_action_reg.asp?id="  + CStr(id) + "&divcd=" + divcd + "&mode=editreginfo"

        If (Err.Number = 0) and (ScanErr="") Then
            dbget.CommitTrans
        Else
            dbget.RollBackTrans
            response.write "<script>alert(" + Chr(34) + "����Ÿ�� �����ϴ� ���߿� ������ �߻��Ͽ����ϴ�. ������ ���� ���.(�����ڵ� : " + CStr(errcode) + ":" + Err.Description + "|" + ScanErr + ")" + Chr(34) + ")</script>"
            'response.write "<script>history.back()</script>"
            dbget.close()	:	response.End
        End If
    On error Goto 0



elseif (mode="finishededitcsas") then
	'==========================================================================
    ''�Ϸ�� ���� ����
    On Error Resume Next
        dbget.beginTrans

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "001"

            '' ����Ÿ ����
            Call EditCSMasterFinished(divcd, orderserial, reguserid, title, contents_jupsu, gubun01, gubun02, reguserid, contents_finish)
        end if

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "002"

            '' ������ ����
            Call EditCSDetailByArrStr(detailitemlist, id, orderserial)
        End if

        ResultMsg = "���� �Ǿ����ϴ�."
        ReturnUrl = "/cscenter/action/pop_cs_action_reg.asp?id="  + CStr(id) + "&divcd=" + divcd + "&mode=editreginfo"

        If (Err.Number = 0) and (ScanErr="") Then
            dbget.CommitTrans
        Else
            dbget.RollBackTrans
            response.write "<script>alert(" + Chr(34) + "����Ÿ�� �����ϴ� ���߿� ������ �߻��Ͽ����ϴ�. ������ ���� ���.(�����ڵ� : " + CStr(errcode) + ":" + Err.Description + "|" + ScanErr + ")" + Chr(34) + ")</script>"
            'response.write "<script>history.back()</script>"
            dbget.close()	:	response.End
        End If
    On error Goto 0



elseif (mode="finishcsas") then
	'==========================================================================
    'CS ���� ���� �Ϸ�ó��

    if (divcd="A008") then
		'----------------------------------------------------------------------
		'CS ���� ���� �Ϸ�ó�� - �ֹ����
        On Error Resume Next
	        dbget.beginTrans
	        if (GC_IsOLDOrder) then ScanErr = "6���� ���� ���� ���� ��� �Ұ� - ������ ���� ���"

	        If (Err.Number = 0) and (ScanErr="") Then
	            errcode = "001"

	            Call CancelProcess(id, orderserial)
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
	        End If

	        ReturnUrl = "/cscenter/action/pop_cs_action_reg.asp?id="  + CStr(id) + "&divcd=" + divcd + "&mode=editfinishedinfo"

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

        On Error Resume Next
	        dbget.beginTrans

	        If (Err.Number = 0) and (ScanErr="") Then
	            errcode = "001"

	            Call FinishCSMaster(id, reguserid, contents_finish)
	        End If

	        If (Err.Number = 0) and (ScanErr="") Then
	            errcode = "002"

				'���ϸ��� ȯ���� ������ ȯ���� ó��������, �� �ۿ� �ſ�ī��/������ ���� ȯ���� ���� ȯ�� ���μ������� ó���ȴ�.
				'���� �Ϸ�ó���Ѵٰ� �ؼ� ������ ȯ���� �Ͼ�� �ʴ´�.
	            Call CheckRefundFinish(id, orderserial, RefreturnMethod, Refrealrefund)
	        End If

	        ResultMsg = "ó�� �Ϸ�"
	        if (RefreturnMethod="R007") and (Refrealrefund>0) then
	            ReturnUrl = "/cscenter/action/pop_cs_action_reg.asp?id="  + CStr(id) + "&divcd=" + divcd + "&finishtype=1"
	        else
	            ReturnUrl = "/cscenter/action/pop_cs_action_reg.asp?id="  + CStr(id) + "&divcd=" + divcd
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
	        dbget.beginTrans

	        If (modeflag2<>"norefund") and (Err.Number = 0) and (ScanErr="") Then
	            errcode = "001"

	            ''���̳ʽ� �ֹ� ���
	            if (CheckNAddMinusOrder(id, orderserial, reguserid, MinusOrderserial, ScanErr)) then
	                ResultMsg = ResultMsg + "->. [��ǰ �ֹ�] ���\n\n"
	            end if
	        End If

	        If (modeflag2<>"norefund") and (Err.Number = 0) and (ScanErr="") Then
	            errcode = "002"

	            'ȯ�� ������ �ִ��� üũ �� ������ȯ��/���ϸ���ȯ��/�ſ�ī����� CS ���� ���
	            newasid = CheckNRegRefund(id, MinusOrderserial, reguserid)
	            if (newasid>0) then
	                ResultMsg = ResultMsg + "->. [���(ȯ��)����] ó��\n\n"
	            end if
	        End If


	        If (Err.Number = 0) and (ScanErr="") Then
	            errcode = "009"

	            Call FinishCSMaster(id, reguserid, contents_finish)

	            if (divcd="A004") then
	                ResultMsg = ResultMsg + "->. ��ǰ ó�� �Ϸ�\n\n"
	            elseif (divcd="A010") then
	                ResultMsg = ResultMsg + "->. ȸ�� ó�� �Ϸ�\n\n"
	            end if
	        End If

	        ResultMsg = ResultMsg
	        ReturnUrl = "/cscenter/action/pop_cs_action_reg.asp?id="  + CStr(id) + "&divcd=" + divcd
	        If (Err.Number = 0) and (ScanErr="") Then
	            dbget.CommitTrans
	        Else
	            dbget.RollBackTrans
	            response.write "<script>alert(" + Chr(34) + "����Ÿ�� �����ϴ� ���߿� ������ �߻��Ͽ����ϴ�. ������ ���� ���.(�����ڵ� : " + CStr(errcode) + ":" + Err.Description + "|" + ScanErr + ")" + Chr(34) + ")</script>"
	            response.write "<script>history.back()</script>"
	            dbget.close()	:	response.End
	        End If

	        ''ȸ�� �Ϸ� ����
	        If (isCsMailSend) then
	            Call SendCsActionMail(id)

	            ''ȯ�� ���� ����
	            if (newasid>0) then
	                Call SendCsActionMail(newasid)
	            end if
	        End If
        On error Goto 0
    elseif  (divcd="A011") then
    	'----------------------------------------------------------------------
    	'CS ���� ���� �Ϸ�ó�� - �±�ȯȸ��(�ٹ����ٹ��)
        On Error Resume Next
	        dbget.beginTrans

	        If (Err.Number = 0) and (ScanErr="") Then
	            errcode = "009"

	            Call FinishCSMaster(id, reguserid, contents_finish)
	        End If

	        ResultMsg = "ó�� �Ϸ�"
	        ReturnUrl = "/cscenter/action/pop_cs_action_reg.asp?id="  + CStr(id) + "&divcd=" + divcd

	        If (Err.Number = 0) and (ScanErr="") Then
	            dbget.CommitTrans
	        Else
	            dbget.RollBackTrans
	            response.write "<script>alert(" + Chr(34) + "����Ÿ�� �����ϴ� ���߿� ������ �߻��Ͽ����ϴ�. ������ ���� ���.(�����ڵ� : " + CStr(errcode) + ":" + Err.Description + "|" + ScanErr + ")" + Chr(34) + ")</script>"
	           ' response.write "<script>history.back()</script>"
	            dbget.close()	:	response.End
	        End If

	        ''�±�ȯ �Ϸ� ����
	        If (isCsMailSend) then
	            Call SendCsActionMail(id)
	        End If
        On error Goto 0
    elseif  (divcd="A000") or (divcd="A001") or (divcd="A002") or (divcd="A009") or (divcd="A006") or (divcd="A005") or (divcd="A700") then
    	'----------------------------------------------------------------------
        'CS ���� ���� �Ϸ�ó�� - �±�ȯ ��� / ���� / ���� �߼� / ��Ÿ /  ���� ���ǻ���
        On Error Resume Next
	        dbget.beginTrans

	        If (Err.Number = 0) and (ScanErr="") Then
	            errcode = "009"

	            Call FinishCSMaster(id, reguserid, contents_finish)
	        End If


	        ResultMsg = "ó�� �Ϸ�"
	        ReturnUrl = "/cscenter/action/pop_cs_action_reg.asp?id="  + CStr(id) + "&divcd=" + divcd

	        If (Err.Number = 0) and (ScanErr="") Then
	            dbget.CommitTrans
	        Else
	            dbget.RollBackTrans
	            response.write "<script>alert(" + Chr(34) + "����Ÿ�� �����ϴ� ���߿� ������ �߻��Ͽ����ϴ�. ������ ���� ���.(�����ڵ� : " + CStr(errcode) + ":" + Err.Description + "|" + ScanErr + ")" + Chr(34) + ")</script>"
	           ' response.write "<script>history.back()</script>"
	            dbget.close()	:	response.End
	        End If

	        If (isCsMailSend) then
	            if ((divcd="A000") or (divcd="A001") or (divcd="A002")) then
	                ''�±�ȯ/����/���� �Ϸ� ����
	                Call SendCsActionMail(id)
	            end if
	        End If
        On error Goto 0
    else
        ResultMsg = "���ǵ��� �ʾҽ��ϴ�. : mode=" + mode + " , divcd=" + divcd
        response.write "<script>alert('" + ResultMsg + "');</script>"
        response.write "<script>history.back();</script>"
        dbget.close()	:	response.End
    end if



elseif (mode="state2jupsu") then
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
            ReturnUrl = "/cscenter/action/pop_cs_action_reg.asp?id="  + CStr(id) + "&divcd=" + divcd
        else
            response.write "<script>alert('" + ResultMsg + "');</script>"
            response.write "<script>history.back();</script>"
            dbget.close()	:	response.End
        end if
    else
        ResultMsg = "���ǵ��� �ʾҽ��ϴ�. : mode=" + mode + " , divcd=" + divcd
        response.write "<script>alert('" + ResultMsg + "');</script>"
        response.write "<script>history.back();</script>"
        dbget.close()	:	response.End
    end if



elseif (mode="addupchejungsanEdit") then
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
        sqlStr = sqlStr + "set currstate='B001'" + VbCrlf
        sqlStr = sqlStr + ",contents_jupsu='" + (contents_jupsu) + "'" + VbCrlf
        sqlStr = sqlStr + " where id=" + CStr(id)
        dbget.Execute sqlStr

        ResultMsg = "ó�� �Ϸ�"
        ReturnUrl = "/cscenter/action/pop_cs_action_reg.asp?id="  + CStr(id) + "&divcd=" + divcd
    else
        response.write "<script>alert('" + ResultMsg + "');</script>"
        response.write "<script>history.back();</script>"
        dbget.close()	:	response.End
    end if



else
	'==========================================================================
    ResultMsg = "���ǵ��� �ʾҽ��ϴ�. : mode=" + mode + " , divcd=" + divcd
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


%>

<!-- #include virtual="/lib/db/dbclose.asp" -->

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