<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/cscenterv2/lib/incSessionAdminCS.asp" -->
<!-- #include virtual="/cscenterv2/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/cscenterv2/lib/classes/lecture/lecturecls.asp"-->
<!-- #include virtual="/cscenterv2/lib/csAsLecturefunction.asp"-->
<!-- #include virtual="/cscenterv2/lib/classes/cs_lecture/lec_cs_aslistcls.asp"-->
<!-- #include virtual="/cscenterv2/lib/cs_action_mail_Function.asp"-->
<!-- #include virtual="/lib/email/MailLib2.asp"-->
<%

dim mode, modeflag2, divcd, id, reguserid, ipkumdiv
dim title, orderserial, gubun01, gubun02, contents_jupsu
dim finishuser, contents_finish

dim requireupche, requiremakerid, ForceReturnByTen
dim detailitemlist

''��� ����
dim refundmileagesum, refundcouponsum, allatsubtractsum
dim refunditemcostsum, canceltotal, nextsubtotal
dim refundbeasongpay, recalcubeasongpay, refunddeliverypay, refundadjustpay
dim refundmatdiv

''ȯ�� ���� maybe (refundrequire==canceltotal)
dim refundrequire, returnmethod
dim rebankname, rebankaccount, rebankownername, paygateTid, encmethod

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

mode        = RequestCheckvar(request.Form("mode"),16)
modeflag2   = RequestCheckvar(request.Form("modeflag2"),16)
divcd       = RequestCheckvar(request.Form("divcd"),4)
id          = RequestCheckvar(request.Form("id"),10)
ipkumdiv    = RequestCheckvar(request.Form("ipkumdiv"),2)
reguserid   = session("ssbctid")
finishuser  = reguserid
title       = html2DB(request.Form("title"))
if title <> "" then
	if checkNotValidHTML(title) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
	response.write "</script>"
	response.End
	end if
end if
orderserial = RequestCheckvar(request.Form("orderserial"),16)
gubun01     = RequestCheckvar(request.Form("gubun01"),4)
gubun02     = RequestCheckvar(request.Form("gubun02"),4)
contents_jupsu  = html2DB(request.Form("contents_jupsu"))
detailitemlist  = html2db(request.Form("detailitemlist"))
contents_finish = html2DB(request.Form("contents_finish"))
if contents_jupsu <> "" then
	if checkNotValidHTML(contents_jupsu) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
	response.write "</script>"
	response.End
	end if
end If
if detailitemlist <> "" then
	if checkNotValidHTML(detailitemlist) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
	response.write "</script>"
	response.End
	end if
end If
if contents_finish <> "" then
	if checkNotValidHTML(contents_finish) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
	response.write "</script>"
	response.End
	end if
end if
''��ü ó�� ��û
requireupche = RequestCheckvar(request.Form("requireupche"),1)
requiremakerid = RequestCheckvar(request.Form("requiremakerid"),32)
ForceReturnByTen = RequestCheckvar(request.Form("ForceReturnByTen"),32)

orgitemcostsum      = RequestCheckvar(request.Form("orgitemcostsum"),10)
orgbeasongpay       = RequestCheckvar(request.Form("orgbeasongpay"),10)
orgmileagesum       = RequestCheckvar(request.Form("miletotalprice"),10)
orgcouponsum        = RequestCheckvar(request.Form("tencardspend"),10)
orgallatdiscountsum = RequestCheckvar(request.Form("allatdiscountprice"),10)
orgsubtotalprice    = RequestCheckvar(request.Form("subtotalprice"),10)


refunditemcostsum   = RequestCheckvar(request.Form("refunditemcostsum"),10)
nextsubtotal        = RequestCheckvar(request.Form("nextsubtotal"),10)
canceltotal         = RequestCheckvar(request.Form("canceltotal"),10)

refundbeasongpay    = RequestCheckvar(request.Form("refundbeasongpay"),10)
recalcubeasongpay   = RequestCheckvar(request.Form("recalcubeasongpay"),10)
refunddeliverypay   = RequestCheckvar(request.Form("refunddeliverypay"),10)

refundmileagesum    = RequestCheckvar(request.Form("refundmileagesum"),10)
refundcouponsum     = RequestCheckvar(request.Form("refundcouponsum"),10)
allatsubtractsum    = RequestCheckvar(request.Form("allatsubtractsum"),10)
refundadjustpay     = RequestCheckvar(request.Form("refundadjustpay"),10)

'����ȯ�ҹ��
refundmatdiv     	= RequestCheckvar(request.Form("ckmaterialpay20"),1)
if (refundmatdiv = "") then
	refundmatdiv = RequestCheckvar(request.Form("ckmaterialpay10"),1)
end if
if (refundmatdiv = "") then
	refundmatdiv = RequestCheckvar(request.Form("cklecturepay0"),1)
end if



''ȯ�ҿ�û��
refundrequire       = RequestCheckvar(request.Form("refundrequire"),10)
returnmethod        = RequestCheckvar(request.Form("returnmethod"),4)

rebankname          = RequestCheckvar(request.Form("rebankname"),32)
rebankaccount       = RequestCheckvar(request.Form("rebankaccount"),32)
rebankownername     = RequestCheckvar(request.Form("rebankownername"),64)
if rebankownername <> "" then
	if checkNotValidHTML(rebankownername) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
	response.write "</script>"
	response.End
	end if
end if
encmethod 			= "PH1"

paygateTid          = RequestCheckvar(request.Form("paygateTid"),64)
if paygateTid <> "" then
	if checkNotValidHTML(paygateTid) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
	response.write "</script>"
	response.End
	end if
end if
add_upchejungsandeliverypay = RequestCheckvar(request.Form("add_upchejungsandeliverypay"),10)
add_upchejungsancause       = RequestCheckvar(request.Form("add_upchejungsancause"),32)
add_upchejungsancauseText   = RequestCheckvar(request.Form("add_upchejungsancauseText"),32)

buf_requiremakerid  = RequestCheckvar(request.Form("buf_requiremakerid"),32)


isCsMailSend = (RequestCheckvar(request.Form("csmailsend"),10)="on")

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

''response.write mode & "_" & divcd
''response.end

if (mode="regcsas") then
    '' ���
    if (divcd="A008") then
        ''On Error Resume Next
        dbget.beginTrans

        if (GC_IsOLDOrder) then ScanErr="6���� ���� ���� ���� ��� �Ұ� - ������ ���� ���"

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "001"

            '' CS Master ����
            id = RegCSMaster(divcd, orderserial, reguserid, title, contents_jupsu, gubun01, gubun02)
        end if

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "002"

            '' CS Master ���/ ȯ�� ��������
	        Call RegCSMasterRefundInfoLecture(id, returnmethod, refundrequire , orgsubtotalprice, orgitemcostsum, orgbeasongpay , orgmileagesum, orgcouponsum, orgallatdiscountsum  , canceltotal, refunditemcostsum, refundmileagesum, refundcouponsum, allatsubtractsum , refundbeasongpay, refunddeliverypay, refundadjustpay , rebankname, rebankaccount, rebankownername, paygateTid, refundmatdiv)

			'''���� ��ȣȭ �߰�.
			Call EditCSMasterRefundEncInfo(id, encmethod, rebankaccount)
	    End if

	    If (Err.Number = 0) and (ScanErr="") Then
            errcode = "003"

            '' CS Detail ����
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


        '���Ϸ��� ��ǰ�� ���� ���, ��������(�ֹ���� �Ұ�)

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
                ''ȯ�� ��ϰ��� �ִ��� üũ �� ȯ�ҿ�û/�ſ�ī�� ��ҿ�û ���
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
        'TODO : ������� �ϴ� ��ŵ
        'Call CheckNChangeCyberAcct(orderserial)

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
        on error Goto 0


        ReturnUrl = "pop_lec_cs_register.asp?id="  + CStr(id) + "&divcd=" + divcd + "&mode=editreginfo"
    elseif (divcd="A004") or (divcd="A010") then
        ''����Ȯ���� ��� ���� �Ǵ� ȸ����û.
        On Error Resume Next
        dbget.beginTrans

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "001"

            '' CS Master ����
            id = RegCSMaster(divcd, orderserial, reguserid, title, contents_jupsu, gubun01, gubun02)
        end if

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "002"

            '' CS Master ���/ ȯ�� ��������
	        Call RegCSMasterRefundInfoLecture(id, returnmethod, refundrequire , orgsubtotalprice, orgitemcostsum, orgbeasongpay , orgmileagesum, orgcouponsum, orgallatdiscountsum  , canceltotal, refunditemcostsum, refundmileagesum, refundcouponsum, allatsubtractsum , refundbeasongpay, refunddeliverypay, refundadjustpay , rebankname, rebankaccount, rebankownername, paygateTid, refundmatdiv)

			'''���� ��ȣȭ �߰�.
			Call EditCSMasterRefundEncInfo(id, encmethod, rebankaccount)
	    End if


	    If (Err.Number = 0) and (ScanErr="") Then
            errcode = "003"

            '' CS Detail ����
	        Call AddCSDetailByArrStr(detailitemlist, id, orderserial)

        end if

        '' To Do : ��ǰ, ȸ�� ������ : ��ü��۰� ��ü ����� ���� ������� -> ��ü����� �귣�庰�� �Է� ��ü����� ���� �Է�
        '' Check - ��ü��۰� ��ü����� ���� �������� ����.
        ''       - ��ü����� ������ ��� MakerID�� 1���� ���� �ؾ���.

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "004"

            if (IsReturnRegValid(id, orderserial, ScanErr, requiremakerid)) then
                ''��ü ��û�� ��� ��ü require, ���
                if ((requiremakerid<>"") and (ForceReturnByTen="")) then
                    call RegCSMasterAddUpche(id, requiremakerid)
                end if

                ResultMsg = ResultMsg + "->. ����Ȯ�� �� �Ϻ���� ����\n\n"
            end if
        End if

        ''��ü �߰� ����� 2008.11.10
        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "005"

            if (add_upchejungsandeliverypay<>"0") and (add_upchejungsandeliverypay<>"")  then
                call RegCSUpcheAddJungsanPay(id, add_upchejungsandeliverypay, add_upchejungsancause, buf_requiremakerid)
            end if
        end if

        ReturnUrl = "pop_lec_cs_register.asp?id="  + CStr(id) + "&divcd=" + divcd


        If (Err.Number = 0) and (ScanErr="") Then
            dbget.CommitTrans
        Else
            dbget.RollBackTrans
            response.write "<script>alert(" + Chr(34) + "����Ÿ�� �����ϴ� ���߿� ������ �߻��Ͽ����ϴ�. ������ ���� ���.(�����ڵ� : " + CStr(errcode) + ":" + Err.Description + "|" + ScanErr + ")" + Chr(34) + ")</script>"
            'response.write "<script>history.back()</script>"
            dbget.close()	:	response.End
        End If

        ''�̸��� �߼�. ����Ȯ������� ȸ�� ����
        If (isCsMailSend) then
            '�̸��Ϲ߼� ����(ȯ�������� �̸��� �߼�)
            'Call SendCsActionMail(id)
        end if
        on error Goto 0
    elseif (divcd="A001") or (divcd="A002") then
        ''������߼�, ���񽺹߼�
        On Error Resume Next
        dbget.beginTrans

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "001"
            '' CS Master ����
            id = RegCSMaster(divcd, orderserial, reguserid, title, contents_jupsu, gubun01, gubun02)
        end if

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "002"
            '' CS Detail ����
	        Call AddCSDetailByArrStr(detailitemlist, id, orderserial)

        end if

        if (requiremakerid<>"") then
            call RegCSMasterAddUpche(id, requiremakerid)
        end if

        ResultMsg = "�����Ϸ�"
        ReturnUrl = "pop_lec_cs_register.asp?id="  + CStr(id) + "&divcd=" + divcd


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
        ''�±�ȯ���
        On Error Resume Next
        dbget.beginTrans

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "001"
            '' CS Master ����
            id = RegCSMaster(divcd, orderserial, reguserid, title, contents_jupsu, gubun01, gubun02)
        end if

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "002"
            '' CS Detail ����
	        Call AddCSDetailByArrStr(detailitemlist, id, orderserial)

        end if


        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "003"
            if (requiremakerid<>"") then
                call RegCSMasterAddUpche(id, requiremakerid)

                ResultMsg = "�±�ȯ �����Ϸ� - ��ü���"
            else
                ''�±�ȯ ȸ�� ����
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


        ReturnUrl = "pop_lec_cs_register.asp?id="  + CStr(id) + "&divcd=" + divcd


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
        ''��Ÿ���� / �������ǻ��� / ��ü �߰� �����
        On Error Resume Next
        dbget.beginTrans

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "001"
            '' CS Master ����
            id = RegCSMaster(divcd, orderserial, reguserid, title, contents_jupsu, gubun01, gubun02)
        end if

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "002"
            '' CS Detail ����
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
            if (requiremakerid<>"") then
                call RegCSMasterAddUpche(id, requiremakerid)
            end if
         end if

        ResultMsg = "��ϵǾ����ϴ�."
        ReturnUrl = "pop_lec_cs_register.asp?id="  + CStr(id) + "&divcd=" + divcd

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
        ''ȯ������ / �ܺθ� ȯ������
        On Error Resume Next
        dbget.beginTrans

        if (divcd="A005") and (Not IsExtSiteOrder(orderserial)) then
            ScanErr = "�ܺθ� ȯ�������� �ܺθ� �ֹ��Ǹ� �����մϴ�."
        end if

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "001"
            '' CS Master ����
            id = RegCSMaster(divcd, orderserial, reguserid, title, contents_jupsu, gubun01, gubun02)
        end if

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "002"

            '' CS Master ���/ ȯ�� ��������
	        Call RegCSMasterRefundInfoLecture(id, returnmethod, refundrequire , orgsubtotalprice, orgitemcostsum, orgbeasongpay , orgmileagesum, orgcouponsum, orgallatdiscountsum  , canceltotal, refunditemcostsum, refundmileagesum, refundcouponsum, allatsubtractsum , refundbeasongpay, refunddeliverypay, refundadjustpay , rebankname, rebankaccount, rebankownername, paygateTid, refundmatdiv)

			'''���� ��ȣȭ �߰�.
			Call EditCSMasterRefundEncInfo(id, encmethod, rebankaccount)
	    End if


        ResultMsg = "��ϵǾ����ϴ�."
        ReturnUrl = "pop_lec_cs_register.asp?id="  + CStr(id) + "&divcd=" + divcd

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
        ReturnUrl = "pop_lec_cs_register.asp?id="  + CStr(id) + "&divcd=" + divcd + "&mode=editreginfo"

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
    ''���� ���� ����
    On Error Resume Next
        dbget.beginTrans

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "001"
            '' ����Ÿ ����
            Call EditCSMaster(divcd, orderserial, reguserid, title, contents_jupsu, gubun01, gubun02)

        end if

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "002"
            '' ������ ����
            Call EditCSDetailByArrStr(detailitemlist, id, orderserial)
        End if

        ResultMsg = ResultMsg + "->. [CS ó���� ����] ó��\n\n"

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "003"
            '' ȯ�� ���� ����

            if (CheckNEditRefundInfo(id,returnmethod,rebankaccount,rebankownername,rebankname,paygateTid,refundrequire, orgsubtotalprice, orgitemcostsum, orgbeasongpay , orgmileagesum, orgcouponsum, orgallatdiscountsum , canceltotal, refunditemcostsum, refundmileagesum, refundcouponsum, allatsubtractsum , refundbeasongpay, refunddeliverypay, refundadjustpay, refundmatdiv)) Then
				'''���� ��ȣȭ �߰�.
				Call EditCSMasterRefundEncInfo(id, encmethod, rebankaccount)

                ResultMsg = ResultMsg + "->. [ȯ������ ����] ó��\n\n"
            end if
        end If

        ''��ü �߰� ����� 2008.11.10
        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "004"

            if (add_upchejungsandeliverypay<>"") then
               '���¿� ��ü�߰� ��ۺ� ������ ����.
                'call EditCSUpcheAddJungsanPay(id, add_upchejungsandeliverypay, add_upchejungsancause, buf_requiremakerid)
            end if
        end if

        ReturnUrl = "pop_lec_cs_register.asp?id="  + CStr(id) + "&divcd=" + divcd + "&mode=editreginfo"


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
        ReturnUrl = "pop_lec_cs_register.asp?id="  + CStr(id) + "&divcd=" + divcd + "&mode=editreginfo"


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
    ''�Ϸ� ó�� ����

    ''����ΰ��
    if (divcd="A008") then
        On Error Resume Next
        dbget.beginTrans
        if (GC_IsOLDOrder) then ScanErr="6���� ���� ���� ���� ��� �Ұ� - ������ ���� ���"

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "001"
            Call CancelProcess(id, orderserial)
        End IF


        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "008"
            ''ȯ�� ��ϰ��� �ִ��� üũ �� ȯ�ҿ�û/�ſ�ī�� ��ҿ�û ���
            newasid = CheckNRegRefund(id, orderserial, reguserid)
            if (newasid>0) then
                ResultMsg = ResultMsg + "->. [ȯ�� ���] ó��\n\n"
            end if
        End If


        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "009"
            Call FinishCSMaster(id, reguserid, contents_finish)
        End If


        ReturnUrl = "pop_lec_cs_register.asp?id="  + CStr(id) + "&divcd=" + divcd + "&mode=editfinishedinfo"


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
        ''ȯ�ҿ�û
        dim RefreturnMethod, Refrealrefund
        On Error Resume Next
        dbget.beginTrans

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "001"
            Call FinishCSMaster(id, reguserid, contents_finish)
        End If

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "002"
            Call CheckRefundFinish(id, orderserial, RefreturnMethod, Refrealrefund)
        End If

        ResultMsg = "ó�� �Ϸ�"
        if (RefreturnMethod="R007") and (Refrealrefund>0) then
            ReturnUrl = "pop_lec_cs_register.asp?id="  + CStr(id) + "&divcd=" + divcd + "&finishtype=1"
        else
            ReturnUrl = "pop_lec_cs_register.asp?id="  + CStr(id) + "&divcd=" + divcd
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

        ''����Ȯ������� ����(��ü���)  // ȸ����û(�ٹ����ٹ��) // �±�ȯȸ��(�ٹ����ٹ��)

        dim MinusOrderserial

        On Error Resume Next
        dbget.beginTrans


        If (modeflag2<>"norefund") and (Err.Number = 0) and (ScanErr="") Then
            errcode = "001"
            ''���̳ʽ� �ֹ� ���
            if (CheckNAddMinusOrder(id, orderserial, reguserid, MinusOrderserial, ScanErr)) then
                ResultMsg = ResultMsg + "->. [����Ȯ�� �� �Ϻ����] ���\n\n"
            end if
        End If


        If (modeflag2<>"norefund") and (Err.Number = 0) and (ScanErr="") Then
            errcode = "002"
            ''ȯ�� ��ϰ��� �ִ��� üũ �� ȯ�ҿ�û/�ſ�ī�� ��ҿ�û ���
            newasid = CheckNRegRefund(id, MinusOrderserial, reguserid)
            if (newasid>0) then
                ResultMsg = ResultMsg + "->. [���(ȯ��)����] ó��\n\n"
            end if
        End If


        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "009"
            Call FinishCSMaster(id, reguserid, contents_finish)

            if (divcd="A004") then
                ResultMsg = ResultMsg + "->. ����Ȯ�� �� �Ϻ���� ó�� �Ϸ�\n\n"
            elseif (divcd="A010") then
                ResultMsg = ResultMsg + "->. ȸ�� ó�� �Ϸ�\n\n"
            end if
        End If


        ResultMsg = ResultMsg
        ReturnUrl = "pop_lec_cs_register.asp?id="  + CStr(id) + "&divcd=" + divcd
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
        On Error Resume Next
        dbget.beginTrans

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "009"
            Call FinishCSMaster(id, reguserid, contents_finish)
        End If


        ResultMsg = "ó�� �Ϸ�"
        ReturnUrl = "pop_lec_cs_register.asp?id="  + CStr(id) + "&divcd=" + divcd
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
        ''�±�ȯ ��� / ���� / ���� �߼� / ��Ÿ /  ���� ���ǻ���
        On Error Resume Next
        dbget.beginTrans

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "009"
            Call FinishCSMaster(id, reguserid, contents_finish)
        End If


        ResultMsg = "ó�� �Ϸ�"
        ReturnUrl = "pop_lec_cs_register.asp?id="  + CStr(id) + "&divcd=" + divcd
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
    if (divcd="A700") then
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

            response.write sqlStr
            dbget.Execute sqlStr

            ResultMsg = "ó�� �Ϸ�"
            ReturnUrl = "pop_lec_cs_register.asp?id="  + CStr(id) + "&divcd=" + divcd
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
    '' ��ü ó���Ϸ�=>�������·κ���
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
        ReturnUrl = "pop_lec_cs_register.asp?id="  + CStr(id) + "&divcd=" + divcd
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
