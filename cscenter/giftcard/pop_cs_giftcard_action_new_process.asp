<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/tenEncUtil.asp" -->
<!-- #include virtual="/lib/util/base64unicode.asp" -->
<!-- #include virtual="/lib/classes/cscenter/cs_aslistcls.asp" -->
<!-- #include virtual="/lib/classes/cscenter/cs_giftcard_ordercls.asp" -->
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
'elseif (mode="finishcsas") then
'	'==========================================================================
'	'CS ���� ���� �Ϸ�ó��
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


dim mode, modeflag2, divcd, id, reguserid, ipkumdiv
dim title, giftorderserial, gubun01, gubun02, contents_jupsu
dim finishuser, contents_finish

dim requireupche, requiremakerid, ForceReturnByTen
dim detailitemlist

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

newasid = -1

mode        = request.Form("mode")
modeflag2   = request.Form("modeflag2")
divcd       = request.Form("divcd")
id          = request.Form("id")
ipkumdiv    = request.Form("ipkumdiv")
reguserid   = session("ssbctid")
finishuser  = reguserid
title       = html2DB(request.Form("title"))
giftorderserial = request.Form("giftorderserial")
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
orgbeasongpay       = 0
orgmileagesum       = 0
orgcouponsum        = 0
orgallatdiscountsum = 0
orgsubtotalprice    = request.Form("subtotalprice")

orggiftcardsum    	= 0
refundgiftcardsum   = 0
orgdepositsum    	= 0
refunddepositsum    = 0

refunditemcostsum   = request.Form("refunditemcostsum")
nextsubtotal        = request.Form("nextsubtotal")
canceltotal         = request.Form("canceltotal")

refundbeasongpay    = 0
remainbeasongpay    = 0
refunddeliverypay   = 0

refundmileagesum    = 0
refundcouponsum     = 0
allatsubtractsum    = 0
refundadjustpay     = 0
remainitemcostsum   = 0



''ȯ�ҿ�û��
refundrequire       = request.Form("refundrequire")
returnmethod        = request.Form("returnmethod")

rebankname          = request.Form("rebankname")
rebankaccount       = request.Form("rebankaccount")
rebankownername     = request.Form("rebankownername")

encmethod 			= "AE2"

paygateTid          = request.Form("paygateTid")

add_upchejungsandeliverypay = 0
add_upchejungsancause       = ""
add_upchejungsancauseText   = ""

buf_requiremakerid  = ""


isCsMailSend = (request.Form("csmailsend")="on")

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

if (returnmethod="") then returnmethod="R000"

''�ÿ�ī������.. -��ǰ���� ����.

dim sqlStr, errcode, i
dim ScanErr
dim ResultMsg, ReturnUrl, EtcStr
dim ProceedFinish

ScanErr = ""
ProceedFinish = False

dim IsAllCancel
dim CancelValidResultMessage



'==============================================================================
''�ֹ� ����Ÿ
dim ogiftcardordermaster

set ogiftcardordermaster = new cGiftCardOrder

ogiftcardordermaster.FRectgiftorderserial = giftorderserial

ogiftcardordermaster.getCSGiftcardOrderDetail



if (mode="regcsas") then
    '==========================================================================
	'CS ����
    if (divcd="A008") then

		'----------------------------------------------------------------------
        'CS ���� - �ֹ����
        'On Error Resume Next
        dbget.beginTrans

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "001"

            'CS Master ����
            id = RegCSMaster(divcd, giftorderserial, reguserid, title, contents_jupsu, gubun01, gubun02)
        end if

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "002"

            'CS Master ȯ�� �������� ����
	        Call RegCSMasterRefundInfo(id, returnmethod, refundrequire , orgsubtotalprice, orgitemcostsum, orgbeasongpay , orgmileagesum, orgcouponsum, orgallatdiscountsum  , canceltotal, refunditemcostsum, refundmileagesum, refundcouponsum, allatsubtractsum , refundbeasongpay, refunddeliverypay, refundadjustpay , rebankname, rebankaccount, rebankownername, paygateTid)
	        Call AddCSMasterRefundInfo(id, orggiftcardsum, orgdepositsum, refundgiftcardsum, refunddepositsum)

            '''���� ��ȣȭ �߰�.
	        Call EditCSMasterRefundEncInfo(id, encmethod, rebankaccount)
	    End if

	    If (Err.Number = 0) and (ScanErr="") Then
            errcode = "005"

        	IsAllCancel = true
        	CancelValidResultMessage = ""
        	if (ogiftcardordermaster.FOneItem.FCancelyn <> "N") then
        		CancelValidResultMessage = "��ҵ� �ֹ��Դϴ�."
        	end if

			if (ogiftcardordermaster.FOneItem.Fjumundiv = "7") then
				CancelValidResultMessage = "��ϵ� Giftī���ֹ��� ����� �� �����ϴ�. ������� ���·� ��ȯ�ϼ���."
			end if

			if (CancelValidResultMessage <> "") then
				ScanErr = CancelValidResultMessage
			end if
        End If

        ResultMsg = ResultMsg + "->. [�ֹ� ��� CS] ����\n\n"

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "008"

		    sqlStr = "update [db_order].[dbo].tbl_giftcard_order " + VbCrlf
		    sqlStr = sqlStr + " set cancelyn='Y'" + VbCrlf
		    sqlStr = sqlStr + " ,canceldate=IsNULL(canceldate,getdate())" + VbCrlf
		    sqlStr = sqlStr + " where giftorderserial='" + giftorderserial + "'" + VbCrlf
		    dbget.Execute sqlStr

		    ''���ں����� �߱޵� ��� ���
		    if (ogiftcardordermaster.FOneItem.FInsureCd="0") then
		        Call UsafeCancel(giftorderserial)
		    end if

            ResultMsg = ResultMsg + "->. �ֹ��� ��� �Ϸ�\n\n"
        End IF

        ''����?. ����?
        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "009"

            'ȯ�� ������ �ִ��� üũ �� ������ȯ��/���ϸ���ȯ��/�ſ�ī����� CS ���� ���
            newasid = CheckNRegRefund(id, giftorderserial,reguserid)

            If (newasid>0) then
                ResultMsg = ResultMsg + "->. ȯ�� ���� �Ϸ�\n\n"
            end if
        End If

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "010"

            Call FinishCSMaster(id, reguserid, contents_finish)

            ResultMsg = ResultMsg + "->. [�ֹ� ��� CS] �Ϸ� ó��\n\n"
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

        ''������� �ݾ�/������ ����
        ''''Call CheckNChangeCyberAcct(giftorderserial)
        response.write "<script>alert('TODO : ������� �ݾ�/������ ����')</script>"


        ''�̸��� �߼�. �ٷ� �Ϸ��ΰ�츸.
        If (isCsMailSend) then
            If (ProceedFinish) then
                ''�ֹ���� �Ϸ� ����
                '''Call SendCsActionMail(id)
                response.write "<script>alert('TODO : SendCsActionMail')</script>"

                ''ȯ�� ���� ����
                if (newasid>0) then
                    '''''Call SendCsActionMail(newasid)
                    response.write "<script>alert('TODO : SendCsActionMail')</script>"
                end if
            End If
        End IF
        'on error Goto 0

        ReturnUrl = "/cscenter/giftcard/pop_cs_giftcard_action_new.asp?id="  + CStr(id) + "&divcd=" + divcd + "&mode=editreginfo"

    else
        ResultMsg = "���ǵ��� �ʾҽ��ϴ�. : mode=" + mode + " , divcd=" + divcd
        response.write "<script>alert('" + ResultMsg + "');</script>"
        response.write "<script>history.back();</script>"
        dbget.close()	:	response.End
    end if



elseif (mode="deletecsas") then
	'==========================================================================
	'CS ���� ����

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
        ReturnUrl = "/cscenter/giftcard/pop_cs_giftcard_action_new.asp?id="  + CStr(id) + "&divcd=" + divcd + "&mode=editreginfo"

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
	'CS ���� ���� ����

    On Error Resume Next
        dbget.beginTrans

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "001"

            '' CS Master ����
            if (divcd = "A003") or (divcd = "A007")  then
            	title = GetCSRefundTitle(id, divcd, giftorderserial, returnmethod, title)
            end if

            Call EditCSMaster(id, reguserid, title, contents_jupsu, gubun01, gubun02)

            ''ȯ�ҹ���� �ٲ� ���.. 2011-07-20 �������߰�
            if (divcd="A007") and Not ((returnmethod="R020") or (returnmethod="R080") or (returnmethod="R100") or (returnmethod="R120") or (returnmethod="R400")) then
                sqlStr = " update [db_cs].[dbo].tbl_new_as_list"
                sqlStr = sqlStr + " set divcd='A003'"
                sqlStr = sqlStr + " where id=" + CStr(id)

                dbget.Execute sqlStr
            end if

            if (divcd="A003") and ((returnmethod="R020") or (returnmethod="R080") or (returnmethod="R100") or (returnmethod="R120") or (returnmethod="R400")) then
                sqlStr = " update [db_cs].[dbo].tbl_new_as_list"
                sqlStr = sqlStr + " set divcd='A007'"
                sqlStr = sqlStr + " where id=" + CStr(id)

                dbget.Execute sqlStr
            end if
        end if

        ResultMsg = ResultMsg + "->. [CS ó���� ����] ó��\n\n"

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "003"

            '' ȯ�� ���� ����
            if (CheckNEditRefundInfo(id,returnmethod,rebankaccount,rebankownername,rebankname,paygateTid,refundrequire, orgsubtotalprice, orgitemcostsum, orgbeasongpay , orgmileagesum, orgcouponsum, orgallatdiscountsum , canceltotal, refunditemcostsum, refundmileagesum, refundcouponsum, allatsubtractsum , refundbeasongpay, refunddeliverypay, refundadjustpay)) then
            	Call AddCSMasterRefundInfo(id, orggiftcardsum, orgdepositsum, refundgiftcardsum, refunddepositsum)

            	'''���� ��ȣȭ �߰�.
	            Call EditCSMasterRefundEncInfo(id, encmethod, rebankaccount)
                ResultMsg = ResultMsg + "->. [ȯ������ ����] ó��\n\n"
            end if
        end If

        ReturnUrl = "/cscenter/giftcard/pop_cs_giftcard_action_new.asp?id="  + CStr(id) + "&divcd=" + divcd + "&mode=editreginfo"

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

	if (divcd="A003") or (divcd="A007") then
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

				'���ϸ��� ȯ�� �� ��ġ����ȯ�� ������ ȯ���� ó��������, �� �ۿ� �ſ�ī��/������ ���� ȯ���� ���� ȯ�� ���μ������� ó���ȴ�.
				'���� �Ϸ�ó���Ѵٰ� �ؼ� ������ ȯ���� �Ͼ�� �ʴ´�.
	            Call CheckRefundFinish(id, giftorderserial, RefreturnMethod, Refrealrefund)
	        End If

	        ResultMsg = "ó�� �Ϸ�"
	        if (RefreturnMethod="R007") and (Refrealrefund>0) then
	            ReturnUrl = "/cscenter/giftcard/pop_cs_giftcard_action_new.asp?id="  + CStr(id) + "&divcd=" + divcd + "&finishtype=1"
	        else
	            ReturnUrl = "/cscenter/giftcard/pop_cs_giftcard_action_new.asp?id="  + CStr(id) + "&divcd=" + divcd
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
	            'newasid = CheckNRegRefund(id, MinusOrderserial, reguserid)

	            '���ֹ��� ���� CS����Ѵ�. ���̳ʽ� �ֹ����� CS�� ����� �� ����.
	            newasid = CheckNRegRefund(id, orderserial, reguserid)
	            call AddminusOrderLink(newasid,MinusOrderserial)

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
	        ReturnUrl = "/cscenter/action/pop_cs_action_new.asp?id="  + CStr(id) + "&divcd=" + divcd
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
    elseif  (divcd="A011") or (divcd="A012") then
    	'----------------------------------------------------------------------
    	'CS ���� ���� �Ϸ�ó�� - �±�ȯȸ��(�ٹ����ٹ��)
        On Error Resume Next
	        dbget.beginTrans

	        If (Err.Number = 0) and (ScanErr="") Then
	            errcode = "009"

	            Call FinishCSMaster(id, reguserid, contents_finish)
	        End If

	        ResultMsg = "ó�� �Ϸ�"
	        ReturnUrl = "/cscenter/action/pop_cs_action_new.asp?id="  + CStr(id) + "&divcd=" + divcd

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
	            if (divcd="A011") then
    	            Call SendCsActionMail(id)
    	        end if
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
	        ReturnUrl = "/cscenter/action/pop_cs_action_new.asp?id="  + CStr(id) + "&divcd=" + divcd

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

else
	'==========================================================================
    ResultMsg = "���ǵ��� �ʾҽ��ϴ�. : mode=" + mode + " , divcd=" + divcd
    response.write "<script>alert('" + ResultMsg + "');</script>"
    response.write "<script>history.back();</script>"
    dbget.close()	:	response.End
end if

%>

<!-- #include virtual="/lib/db/dbclose.asp" -->

<%
response.write "<script>alert('" + ResultMsg + "');</script>"
response.write "<script>location.replace('" + ReturnUrl + "');</script>"
%>
