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

''�ֹ� ����Ÿ
dim oordermaster
set oordermaster = new COrderMaster

if (ocsaslist.FResultCount>0) then
    oordermaster.FRectOrderSerial = ocsaslist.FOneItem.FOrderSerial
    oordermaster.QuickSearchOrderMaster
end if

if (oordermaster.FResultCount>0) then
    if (oordermaster.FOneItem.FCancelYn="N") and (oordermaster.FOneItem.Fjumundiv<>"9")  then
        response.write "<script>alert('��ǰ�ֹ��� �̰ų�, ��ҵ� �ŷ��� ��� �����մϴ�.');</script>"
        response.write "<script>window.close();</script>"
        dbget.close()	:	response.End
    end if
else
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

''=============��ü��Ҹ� ������.. �κ���ҵ� ��Ҿȵ�..=============
dim sqlStr, isSameMoney
dim t_refundrequire, t_MaybeOrgPayPrice
isSameMoney = false
sqlStr = " select r.refundrequire, "
''sqlStr = sqlStr & " (select sum(d.itemcost*d.itemno) from "&TABLE_ORDERDETAIL&" d where d.orderserial=l.orderserial and d.cancelyn<>'A')-(select tencardspend+miletotalprice+allatdiscountprice from "&TABLE_ORDERMASTER&" m where m.orderserial=l.orderserial) as MaybeOrgPayPrice"
sqlStr = sqlStr & " (select sum(d.itemcost*d.itemno) from "&TABLE_ORDERDETAIL&" d where d.orderserial=l.orderserial and d.cancelyn<>'A')-(select (CASE WHEN m.bcpnIDX is Not NULL THEN isNULL(tencardspend,0) ELSE 0 END)+miletotalprice+allatdiscountprice from "&TABLE_ORDERMASTER&" m where m.orderserial=l.orderserial) as MaybeOrgPayPrice"
sqlStr = sqlStr & " from "&TABLE_CSMASTER&" l"
sqlStr = sqlStr & " 	Join "&TABLE_CS_REFUND&" r"
sqlStr = sqlStr & " 	on l.id=r.asid"
sqlStr = sqlStr & " 	and r.returnmethod  in ('R100','R020','R400')"
sqlStr = sqlStr & " where l.id="&id
sqlStr = sqlStr & " and l.divcd='A007'"

rsget.Open sqlStr,dbget,1
if Not rsget.Eof then
    t_refundrequire=rsget("refundrequire")
    t_MaybeOrgPayPrice=rsget("MaybeOrgPayPrice")
    isSameMoney    = (t_refundrequire=Abs(t_MaybeOrgPayPrice))
end if
rsget.Close

IF  (Not isSameMoney) THEN
    IF (force="on") then
        response.write "��ұݾװ� ���ݾ� ����<br><br>"
    ELSE
        if (IsAutoScript) then
            response.write "S_ERR|��ұݾװ� ���ݾ� ����"
        else
            response.write "<script>alert('��ұݾװ� ���ݾ� ���� - ������ ���� ���."&t_refundrequire&":"&t_MaybeOrgPayPrice&"');</script>"
            response.write "<script>window.close();</script>"
        end if
        dbget.close() : dbget_CS.close()	:	response.End
    End IF
END IF
'''=================================================================


'' Pg_Mid
dim MctID
MctID = Mid(orefund.FOneItem.FpaygateTid,11,10)
'' response.write MctID

dim INIpay, PInst
dim ResultCode, ResultMsg, CancelDate, CancelTime, Rcash_cancel_noappl


'############################################################## �ڵ��� ���� ��� ##############################################################
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
		ResultMsg	= "����ó��"
	Else
		ResultCode = retval
		Select Case ResultCode
			Case "14"
				ResultMsg = "����"
			Case "20"
				ResultMsg = "�޴��� ������� ����(PG��) (LGT�� ��� ������������濡 ���� ��������)"
			Case "41"
				ResultMsg = "�ŷ����� ������"
			Case "42"
				ResultMsg = "��ұⰣ���"
			Case "43"
				ResultMsg = "���γ������� ( ������������ ����ġ, ���ι�ȣ ��ȿ�ð� �ʰ�( 3�� ) )"
			Case "44"
				ResultMsg = "�ߺ� ��� ��û"
			Case "45"
				ResultMsg = "��� ��û �� ��� ���� ����ġ"
			Case "97"
				ResultMsg = "��û�ڷ� ����"
			Case "98"
				ResultMsg = "��Ż� ��ſ���"
			Case "99"
				ResultMsg = "��Ÿ"
			Case Else
				ResultMsg = ""
		End Select
	End If

	CancelDate	= year(now) & "�� " & month(now) & "�� " & day(now) & "��"
	CancelTime	= hour(now) & "�� " & minute(now) & "�� " & second(now) & "��"
ELSEIF (oordermaster.FoneItem.FPgGubun="KP") then
    dim ret_amount, ret_panc_mod_mny, ret_panc_rem_mny ''�κ���� ���� �Ķ�.
    Call fnKCPCancelProc(True ,orefund.FOneItem.FpaygateTid,msg, "","",ResultCode,ResultMsg,CancelDate,CancelTime,ret_amount, ret_panc_mod_mny, ret_panc_rem_mny)

Else
'############################################################## ī��, �ǽð� ���� ��� ##############################################################
		'###############################################################################
		'# 1. ��ü ���� #
		'################
		Set INIpay = Server.CreateObject("INItx41.INItx41.1")


		'###############################################################################
		'# 2. �ν��Ͻ� �ʱ�ȭ #
		'######################
		PInst = INIpay.Initialize("")

		'###############################################################################
		'# 3. �ŷ� ���� ���� #
		'#####################
		INIpay.SetActionType CLng(PInst), "CANCEL"

		'###############################################################################
		'# 4. ���� ���� #
		'################
		INIpay.SetField CLng(PInst), "pgid", "IniTechPG_" 'PG ID (����)
		INIpay.SetField CLng(PInst), "spgip", "203.238.3.10" '���� PG IP (����)
		INIpay.SetField CLng(PInst), "mid", MctID '�������̵�
		INIpay.SetField CLng(PInst), "admin", "1111" 'Ű�н�����(�������̵� ���� ����)
		INIpay.SetField CLng(PInst), "tid", Request("tid") '����� �ŷ���ȣ(TID)
		INIpay.SetField CLng(PInst), "msg", msg '��� ����
		INIpay.SetField CLng(PInst), "uip", Request.ServerVariables("REMOTE_ADDR") 'IP
		INIpay.SetField CLng(PInst), "debug", "false" '�α׸��("true"�� �����ϸ� ���� �α׸� ����)
		INIpay.SetField CLng(PInst), "merchantreserved", "����" '����

		'###############################################################################
		'# 5. ��� ��û #
		'################
		INIpay.StartAction(CLng(PInst))

		'###############################################################################
		'# 6. ��� ��� #
		'################
		ResultCode = INIpay.GetResult(CLng(PInst), "resultcode") '����ڵ� ("00"�̸� ��Ҽ���)
		ResultMsg  = INIpay.GetResult(CLng(PInst), "resultmsg") '�������
		CancelDate = INIpay.GetResult(CLng(PInst), "pgcanceldate") '�̴Ͻý� ��ҳ�¥
		CancelTime = INIpay.GetResult(CLng(PInst), "pgcanceltime") '�̴Ͻý� ��ҽð�
		Rcash_cancel_noappl = INIpay.GetResult(CLng(PInst), "rcash_cancel_noappl") '���ݿ����� ��� ���ι�ȣ

		'###############################################################################
		'# 7. �ν��Ͻ� ���� #
		'####################
		INIpay.Destroy CLng(PInst)
End If




dim returnmethod,refundrequire,refundresult,userid
dim iorderserial, ibuyhp
dim contents_finish

contents_finish = "��� " & "[" & ResultCode & "]" & ResultMsg & VbCrlf
contents_finish = contents_finish & "����Ͻ� : " & CancelDate & " " & CancelTime & VbCrlf
contents_finish = contents_finish & "����� ID " & finishuserid

if (ResultCode="00") then

    sqlStr = "select r.*, a.userid, m.orderserial, m.buyhp from "
    sqlStr = sqlStr + " "&TABLE_CS_REFUND&" r,"
    sqlStr = sqlStr + " "&TABLE_CSMASTER&" a"
    sqlStr = sqlStr + "     left join "&TABLE_ORDERMASTER&" m "
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


    sqlStr = " update "&TABLE_CS_REFUND&""
    sqlStr = sqlStr + " set refundresult=" + CStr(refundrequire)
    sqlStr = sqlStr + " where asid=" + CStr(id)
    dbget.Execute sqlStr

    Call AddCustomerOpenContents(id, "ȯ��(���) �Ϸ�: " & CStr(refundrequire))


    Call FinishCSMaster(id, finishuserid, contents_finish)

    ''���� ��� ��û SMS �߼�
    if (iorderserial<>"") and (ibuyhp<>"") then
'        sqlStr = "Insert into [db_sms].[ismsuser].em_tran(tran_phone, tran_callback, tran_status, tran_date, tran_msg ) "
'    	sqlStr = sqlStr + " values('" + ibuyhp + "',"
'    	sqlStr = sqlStr + " '1644-6030',"
'    	sqlStr = sqlStr + " '1',"
'    	sqlStr = sqlStr + " getdate(),"
'    	sqlStr = sqlStr + " '[�ΰŽ�]���� ��� �Ǿ����ϴ�. �ֹ���ȣ : " + iorderserial + "')"

        ''2015/10/16
    	sqlStr = "insert into [SMSDB].[db_infoSMS].dbo.em_smt_tran (date_client_req, content, callback, service_type, broadcast_yn, msg_status,recipient_num) "
	    sqlStr = sqlStr + " values(getdate(),'[�ΰŽ�]���� ��� �Ǿ����ϴ�. �ֹ���ȣ : " + iorderserial + "','027419070','0','N','1','"+ibuyhp+"')"

        dbget_CS.Execute sqlStr
    end if

    ''���� :: �ϴ��糦.
    Call SendCsActionMail(id)

    if (IsAutoScript) then
        response.write "S_OK"
    else
        response.write "<script>alert('" & ResultMsg & "');</script>"
        response.write "<script>opener.location.reload();</script>"
        response.write "<script>window.close();</script>"
    end if
    dbget.close() : dbget_CS.close()	:	response.End

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
set oordermaster = Nothing
set ocsaslist = Nothing
set orefund = Nothing
%>
<!-- #include virtual="/cscenterv2/lib/db/dbclose.asp" -->
