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
	response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
	response.write "</script>"
	response.End
	end if
end if
if (msg="") and (IsAutoScript) then msg="��������"

dim orderserial, returnmethod
dim orgOrderserial, jumundiv, pggubun
dim sqlStr

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
        response.write "S_ERR|ȯ�ҳ����� ���ų� ��ȿ���� ���� �����Դϴ�."
    else
        response.write "<script>alert('ȯ�ҳ����� ���ų� ��ȿ���� ���� �����Դϴ�.');</script>"
        response.write "<script>window.close();</script>"
    end if
    dbget.close() : dbget_CS.close()	:	response.End
end if

if (ocsaslist.FOneItem.FCurrstate<>"B001") then
    if (IsAutoScript) then
        response.write "S_ERR|���� ���°� �ƴմϴ�."
    else
        response.write "<script>alert('���� ���°� �ƴմϴ�.');</script>"
        response.write "<script>window.close();</script>"
    end if
    dbget.close() : dbget_CS.close()	:	response.End
end if


orderserial = ocsaslist.FOneItem.FOrderserial
returnmethod = orefund.FOneItem.Freturnmethod

if (returnmethod<>"R120") and (returnmethod<>"R420") and (returnmethod<>"R022") Then
    if (IsAutoScript) then
        response.write "S_ERR|�ſ�ī��, �ǽð���ü �κ���Ҹ� �����մϴ�."
    else
        response.write "<script>alert('�ſ�ī��, �ǽð���ü �κ���Ҹ� �����մϴ�.');</script>"
        response.write "<script>window.close();</script>"
    end if
    dbget.close()	:	dbget_CS.close() : response.End
end if

''�ֹ� ����Ÿ
dim oordermaster
set oordermaster = new COrderMaster

if (ocsaslist.FResultCount>0) then
    oordermaster.FRectOrderSerial = ocsaslist.FOneItem.FOrderSerial
    oordermaster.QuickSearchOrderMaster
end if

if (oordermaster.FResultCount<1) then
    response.write "<script>alert('�ùٸ� �ֹ����� �ƴմϴ�..');</script>"
    response.write "<script>window.close();</script>"
    dbget.close()	:	response.End
end if


'' IniPay �� ��Ҹ� ���� => KCP�� ����
if (Left(orefund.FOneItem.FpaygateTid,10)<>"IniTechPG_") AND (orefund.FOneItem.Freturnmethod<>"R400") AND (oordermaster.FoneItem.FPgGubun<>"KP") then
    if (IsAutoScript) then
        response.write "S_ERR|�̴Ͻý�, KCP �ŷ��� ��� �����մϴ�."
    else
        response.write "<script>alert('�̴Ͻý�, KCP �ŷ��� ��� �����մϴ�.');</script>"
        response.write "<script>window.close();</script>"
    end if
    dbget.close()	: dbget_CS.close() :	response.End
end if


'==============================================================================
'���� �ְ������ܱݾ�
dim accountdiv : accountdiv ="100"
dim omainpayment, mainpaymentorg
dim cardcancelok, cardcancelerrormsg, cardcancelcount, cardcancelsum, cardcode

if (returnmethod = "R400") or (returnmethod = "R420") then
	accountdiv = "400"
elseif (returnmethod = "R022") then  ''2016/06/21 �߰�.
    accountdiv = "20"
end if


set omainpayment = new COrderMaster

mainpaymentorg = 0			'// ���� �����ݾ�
cardcancelok = "N"
cardcancelerrormsg = ""
cardcancelcount = 0
cardcancelsum   = 0			'// ��ұݾ��հ�
cardcode = ""

if (orderserial<>"") then
	omainpayment.FRectOrderSerial = orderserial

	if ((accountdiv = "100") or (accountdiv = "20")) then ''���̹����� �ǽð� ��ü �κ���� ����
		Call omainpayment.getMainPaymentInfo(accountdiv, mainpaymentorg, cardcancelok, cardcancelerrormsg, cardcancelcount,cardcancelsum, cardcode)
	else
		'' Call omainpayment.getMainPaymentInfoPhone(accountdiv, mainpaymentorg, cardcancelok, cardcancelerrormsg, cardcancelcount,cardcancelsum, cardcode)
	end if

	orgOrderserial = orderserial

	'// ��ȯ�ֹ�( jumundiv = 6 )�̸� ���ֹ����� �������� �����´�.
	sqlStr = " select top 1 m.jumundiv, IsNull(m.pggubun,'') as pggubun "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_academy.dbo.tbl_academy_order_master m "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and m.orderserial = '" & orderserial & "' "
	rsget.Open sqlStr,dbget,1
	if Not rsget.Eof then
		jumundiv = rsget("jumundiv")
		pggubun = rsget("pggubun")
	end if
	rsget.close

	if (jumundiv = "6") then
		sqlStr = " select top 1 c.orgorderserial "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	db_academy.dbo.tbl_academy_change_order c "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + " 	and c.chgorderserial = '" & orderserial & "' "
		rsget.Open sqlStr,dbget,1
		if Not rsget.Eof then
			orgOrderserial = rsget("orgorderserial")
		end if
		rsget.close
	end if
end if


set omainpayment = nothing



IF (cardcancelok<>"Y") then
    if (IsAutoScript) then
        response.write "S_ERR|"&cardcancelerrormsg
    else
        response.write cardcancelerrormsg
        response.write "<script>alert('"&TRIM(cardcancelerrormsg)&"');</script>"
        response.write "<script>window.close();</script>"
    end if
    dbget.close()	:	response.End
ENd IF


''''���ŷ�ID,��ұݾ�    , ����αݾ�,  ������email
dim ioldtid, iCancelPrice, iconfirm_price

ioldtid        = orefund.FOneItem.FpaygateTid
iCancelPrice   = orefund.FOneItem.Frefundrequire
iconfirm_price = (mainpaymentorg-cardcancelsum) - orefund.FOneItem.Frefundrequire

'RW iCancelPrice
'RW iconfirm_price
'RW mainpaymentorg
'RW cardcancelsum


'''��� ��.
dim Tid
dim OldTid, CancelPrice, RepayPrice, CntRepay


IF (ioldtid="") or (iCancelPrice<1) or (iconfirm_price<0) THEN
    if (IsAutoScript) then
        response.write "S_ERR|�κ� ��ұݾ� ���� �Ǵ� TID����"
    else
        response.write "<script>alert('�κ� ��ұݾ� ���� �Ǵ� TID����"&iCancelPrice&"."&iconfirm_price&"');</script>"
        ''response.write "<script>window.close();</script>"
    end if
    dbget.close()	:	response.End
END IF


''response.end

''����� ����======================================================================
dim ResultCode, ResultMsg
dim CancelDate, CancelTime
Dim clogIdx

sqlStr = "select * from [db_academy].[dbo].tbl_academy_card_cancel_log where 1=0"
rsget.Open sqlStr,dbget,1,3
    rsget.AddNew
    rsget("orderserial") = orgOrderserial
    rsget("orgtid")      = ioldtid
    rsget("cancelprice") = iCancelPrice
    rsget("repayprice")  = iconfirm_price
    rsget("usermail")    = "" ''buyemail

    rsget.update
	clogIdx = rsget("clogIdx")
rsget.close


''KCPPAY
Dim IsKCPPya   : IsKCPPya = (pggubun = "KP")
''īī������
Dim IsKakaoPay : IsKakaoPay = (pggubun = "KA")
''���̹�����
Dim iprimaryPayRestAmount, inpointRestAmount, itotalRestAmount
Dim IsNaverPay : IsNaverPay = (pggubun = "NP")

if (IsKakaoPay) then
    response.write "<script>alert('�������� �ʴ°ŷ� pggubun:"&pggubun&"."&iCancelPrice&"."&iconfirm_price&"');</script>"
    dbget.close()	:	response.End

'	CALL PartialCancelKakaoPay(orefund.FOneItem.FpaygateTid, (mainpaymentorg-cardcancelsum),orefund.FOneItem.Frefundrequire,"",retval,ResultCode,ResultMsg,CancelDate,CancelTime, Tid)
'
'	'// ���ϰ��� CcPartCl �ϴ� ����(0:�κ���ҺҰ�, 1:�κ���Ұ���)
'	CntRepay = "2"
'
'	RepayPrice = iconfirm_price
'	CancelPrice = orefund.FOneItem.Frefundrequire
'
'	OldTid = ioldtid
elseif (IsNaverPay) then
    response.write "<script>alert('�������� �ʴ°ŷ� pggubun:"&pggubun&"."&iCancelPrice&"."&iconfirm_price&"');</script>"
    dbget.close()	:	response.End
'    dim nPayCancelRequester : nPayCancelRequester = fnGetNPayCancelRequester(id)
'    CALL PartialCancelNaverPay(orefund.FOneItem.FpaygateTid, orefund.FOneItem.Frefundrequire,"",retval,ResultCode,ResultMsg,CancelDate,CancelTime, Tid, iprimaryPayRestAmount, inpointRestAmount, itotalRestAmount, nPayCancelRequester)
'
'	'// ���ϰ��� CcPartCl �ϴ� ����(0:�κ���ҺҰ�, 1:�κ���Ұ���)
'	CntRepay = "2"
'
'	RepayPrice = iconfirm_price
'	CancelPrice = orefund.FOneItem.Frefundrequire
'
'	OldTid = ioldtid
elseif (IsKCPPya) then
    dim ret_amount, ret_panc_mod_mny, ret_panc_rem_mny
    Call fnKCPCancelProc(False ,ioldtid,msg, iCancelPrice,(mainpaymentorg-cardcancelsum),ResultCode,ResultMsg,CancelDate,CancelTime,ret_amount, ret_panc_mod_mny, ret_panc_rem_mny)

    CancelPrice = ret_panc_mod_mny
    RepayPrice = ret_panc_rem_mny
    OldTid = ioldtid
else
    response.write "<script>alert('�������� �ʴ°ŷ� pggubun:"&pggubun&"."&iCancelPrice&"."&iconfirm_price&"');</script>"
    dbget.close()	:	response.End
end if


if (CntRepay="") then CntRepay=1 ''2013/06/27 �߰� :: ''Ű���� 'where' ��ó�� ������ �߸��Ǿ����ϴ�.

sqlStr = "update [db_academy].[dbo].tbl_academy_card_cancel_log"&VbCRLF
sqlStr = sqlStr & " set newtid='"&Tid&"'"&VbCRLF
sqlStr = sqlStr & " ,resultcode='"&ResultCode&"'"&VbCRLF
sqlStr = sqlStr & " ,resultmsg='"&Replace(ResultMsg,"'","")&"'"&VbCRLF
sqlStr = sqlStr & " ,cancelrequestcount="&CntRepay&VbCRLF
sqlStr = sqlStr & " where clogIdx="&clogIdx&""

dbget.Execute sqlStr


''�κ���� ����ó����.tbl_order_paymentETc ������Ʈ.
if (ResultCode="00") or (ResultCode="0000") or (IsKakaoPay and (resultCode="2001")) then
    IF (IsNaverPay) then
        sqlStr = " update db_academy.dbo.tbl_academy_order_paymentETc"&VbCRLF
        sqlStr = sqlStr & " set realPayedSum="&iprimaryPayRestAmount&VbCRLF
        sqlStr = sqlStr & " where orderserial='"&orgOrderserial&"'"&VbCRLF
        sqlStr = sqlStr & " and acctdiv='"&accountdiv&"'"&VbCRLF

        dbget.Execute sqlStr

        sqlStr = " update db_academy.dbo.tbl_academy_order_paymentETc"&VbCRLF
        sqlStr = sqlStr & " set realPayedSum="&inpointRestAmount&VbCRLF
        sqlStr = sqlStr & " where orderserial='"&orgOrderserial&"'"&VbCRLF
        sqlStr = sqlStr & " and acctdiv='120'"&VbCRLF

        dbget.Execute sqlStr

    ELSE
        sqlStr = " update db_academy.dbo.tbl_academy_order_paymentETc"&VbCRLF
        sqlStr = sqlStr & " set realPayedSum=realPayedSum-"&CancelPrice&VbCRLF
        sqlStr = sqlStr & " where orderserial='"&orgOrderserial&"'"&VbCRLF
        sqlStr = sqlStr & " and acctdiv='"&accountdiv&"'"&VbCRLF

        dbget.Execute sqlStr
    END IF
end if

rw "Tid="&Tid  ''�κ���� TID
rw "ResultCode="&ResultCode
rw "ResultMsg="&ResultMsg
rw "OldTid="&OldTid
rw "CancelPrice="&CancelPrice
rw "RepayPrice="&RepayPrice
rw "CntRepay="&CntRepay

set oordermaster = Nothing
set ocsaslist = Nothing
set orefund = Nothing



dim refundrequire,refundresult,userid
dim iorderserial, ibuyhp
dim contents_finish

contents_finish = "��� " & "[" & ResultCode & "]" & ResultMsg & VbCrlf
if (ResultCode="00") or (ResultCode="0000") or (IsKakaoPay and (resultCode="2001")) then
    IF (IsNaverPay) then
        contents_finish = contents_finish & "��ұݾ� : " &FormatNumber(CancelPrice,0) & VbCrlf
        contents_finish = contents_finish & "�����ݾ� : " &FormatNumber(itotalRestAmount,0) & VbCrlf
        ''contents_finish = contents_finish & "����Ͻ� : " & CancelDate & " " & CancelTime & VbCrlf
        contents_finish = contents_finish & "����� ID " & finishuserid
    else
        contents_finish = contents_finish & "��ұݾ� : " &FormatNumber(CancelPrice,0) & VbCrlf
        contents_finish = contents_finish & "�����ݾ� : " &FormatNumber(RepayPrice,0) & VbCrlf
        ''contents_finish = contents_finish & "����Ͻ� : " & CancelDate & " " & CancelTime & VbCrlf
        contents_finish = contents_finish & "����� ID " & finishuserid
    end if
end if

if (ResultCode="00") or (ResultCode="0000") or (IsKakaoPay and (resultCode="2001")) then
    sqlStr = "select r.*, a.userid, m.orderserial, m.buyhp from "
    sqlStr = sqlStr + " [db_academy].[dbo].tbl_academy_as_refund_info r,"
    sqlStr = sqlStr + " [db_academy].dbo.tbl_academy_as_list a"
    sqlStr = sqlStr + "     left join db_academy.dbo.tbl_academy_order_master m "
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


    sqlStr = " update [db_academy].[dbo].tbl_academy_as_refund_info"
    sqlStr = sqlStr + " set refundresult=" + CStr(refundrequire)
    sqlStr = sqlStr + " where asid=" + CStr(id)
    dbget.Execute sqlStr

    Call AddCustomerOpenContents(id, "�κ� ��� �Ϸ�: " & FormatNumber(CancelPrice,0) & VbCRLF & "���� ���� �ݾ�: "& FormatNumber(RepayPrice,0)) '''CStr(refundrequire))


    Call FinishCSMaster(id, finishuserid, contents_finish)

    ''���� ��� ��û SMS �߼�
    if (iorderserial<>"") and (ibuyhp<>"") then
        'sqlStr = "Insert into [db_sms].[ismsuser].em_tran(tran_phone, tran_callback, tran_status, tran_date, tran_msg ) "
    	'sqlStr = sqlStr + " values('" + ibuyhp + "',"
    	'sqlStr = sqlStr + " '1644-6030',"
    	'sqlStr = sqlStr + " '1',"
    	'sqlStr = sqlStr + " getdate(),"
    	'sqlStr = sqlStr + " '[�ٹ�����]���� �κ� ��� �Ǿ����ϴ�. ������Ҿ� : " + FormatNumber(CancelPrice,0) + " �ֹ���ȣ : " + iorderserial + "')"

    	''sqlStr = " exec [db_sms].[dbo].[usp_SendSMS] '"+ibuyhp+"','1644-1557','[�� �ΰŽ�]���� �κ� ��� �Ǿ����ϴ�. ������Ҿ� : " + FormatNumber(CancelPrice,0) + " �ֹ���ȣ : " + iorderserial + "'"
    	
    	sqlStr = "insert into [SMSDB].[db_infoSMS].dbo.em_smt_tran (date_client_req, content, callback, service_type, broadcast_yn, msg_status,recipient_num) "
	    sqlStr = sqlStr + " values(getdate(),'[�� �ΰŽ�]���� �κ� ��� �Ǿ����ϴ�. �ֹ���ȣ : " + iorderserial + "','1644-1557','0','N','1','"+ibuyhp+"')"
        dbget_CS.Execute sqlStr
    end if

    ''����
    Call SendCsActionMail(id)

    if (IsAutoScript) then
        response.write "S_OK"
    else
        response.write "<script>alert('" & TRIM(ResultMsg) & "');</script>"
        'response.write "<script>opener.location.reload();</script>"
        'response.write "<script>window.close();</script>"
    end if
    dbget.close()	:	response.End

else
    if (IsAutoScript) then
        response.write "S_ERR|"&ResultMsg
    else
        response.write ResultCode & "<br>"
        response.write ResultMsg & "<br>"

		response.write "<br><br>* <font color=red>�ݺ������� �κ���ҽ��� ����</font>�� �߻��ϴ� ��� �ý����� ���ǿ��<br>(�ߺ� ����� �� �ֽ��ϴ�.)"

    end if
end if

%>

<!-- #include virtual="/cscenterv2/lib/db/dbclose.asp" -->
