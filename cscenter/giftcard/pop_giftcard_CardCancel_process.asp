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

<%
'''�ſ�ī�� �κ���� R120 => �ٸ� ���������� ���� ó��.

dim id, finishuserid, msg, force
dim orgOrderSerial
id           = RequestCheckVar(request("id"),10)
msg          = RequestCheckVar(request("msg"),50)
finishuserid = session("ssBctID")
force = RequestCheckVar(request("force"),10)

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
    dbget.close()	:	response.End
end if

if (ocsaslist.FOneItem.FCurrstate<>"B001") then
    if (IsAutoScript) then
        response.write "S_ERR|���� ���°� �ƴմϴ�."
    else
        response.write "<script>alert('���� ���°� �ƴմϴ�.');</script>"
        response.write "<script>window.close();</script>"
    end if
    dbget.close()	:	response.End
end if

'' �ſ�ī�� ��Ҹ� ����
'if (orefund.FOneItem.Freturnmethod<>"R100") then
'    response.write "<script>alert('���� �ſ�ī�� �ŷ��� ��� �����մϴ�.');</script>"
'    response.write "<script>window.close();</script>"
'    dbget.close()	:	response.End
'end if

Dim returnmethod, IsCardPartialCancel
returnmethod = orefund.FOneItem.Freturnmethod

if Not ((returnmethod="R100") or (returnmethod="R020") or (returnmethod="R400")) Then
    if (IsAutoScript) then
        response.write "S_ERR|�ſ�ī�� ��ü���, �ǽð���ü ���, �޴��� ��ü ��� �� ����."
    else
        response.write "<script>alert('�ſ�ī�� ��ü���, �ǽð���ü ���, �޴��� ��ü ��� �� �����մϴ�.');</script>"
        response.write "<script>window.close();</script>"
    end if
    dbget.close()	:	response.End
end if


'' IniPay �� ��Ҹ� ����
if (Left(orefund.FOneItem.FpaygateTid,10)<>"IniTechPG_") AND Left(orefund.FOneItem.FpaygateTid,6)<>"Stdpay" AND Left(orefund.FOneItem.FpaygateTid,6)<>"INIMX_" AND Left(orefund.FOneItem.FpaygateTid,10)<>"INIAPICARD" AND orefund.FOneItem.Freturnmethod<>"R400" then
    if (IsAutoScript) then
        response.write "S_ERR|�̴Ͻý� �ŷ��� ��� �����մϴ�."
    else
        response.write "<script>alert('�̴Ͻý� �ŷ��� ��� �����մϴ�.');</script>"
        response.write "<script>window.close();</script>"
    end if
    dbget.close()	:	response.End
end if

''=============��ü��Ҹ� ������.. �κ���ҵ� ��Ҿȵ�..=============
dim sqlStr, isSameMoney
dim t_refundrequire, t_MaybeOrgPayPrice
isSameMoney = false

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




dim refundrequire,refundresult,userid
dim iorderserial, ibuyhp
dim contents_finish

contents_finish = "��� " & "[" & ResultCode & "]" & ResultMsg & VbCrlf
contents_finish = contents_finish & "����Ͻ� : " & CancelDate & " " & CancelTime & VbCrlf
contents_finish = contents_finish & "����� ID " & finishuserid

if (ResultCode="00") then

    sqlStr = "select r.*, a.userid, m.giftorderserial, m.buyhp from "
    sqlStr = sqlStr + " [db_cs].[dbo].tbl_as_refund_info r,"
    sqlStr = sqlStr + " [db_cs].dbo.tbl_new_as_list a"
    sqlStr = sqlStr + "     left join [db_order].[dbo].tbl_giftcard_order m "
	sqlStr = sqlStr + "     on a.orderserial=m.giftorderserial"
    sqlStr = sqlStr + " where r.asid=" + CStr(id)
    sqlStr = sqlStr + " and r.asid=a.id"

    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        returnmethod    = rsget("returnmethod")
        refundrequire   = rsget("refundrequire")
        refundresult    = rsget("refundresult")
        userid          = rsget("userid")
        iorderserial    = rsget("giftorderserial")
        ibuyhp          = rsget("buyhp")
    end if
    rsget.Close


    sqlStr = " update [db_cs].[dbo].tbl_as_refund_info"
    sqlStr = sqlStr + " set refundresult=" + CStr(refundrequire)
    sqlStr = sqlStr + " where asid=" + CStr(id)
    dbget.Execute sqlStr

    Call AddCustomerOpenContents(id, "ȯ��(���) �Ϸ�: " & CStr(refundrequire))


    Call FinishCSMaster(id, finishuserid, contents_finish)

    ''���� ��� ��û SMS �߼�
    if (iorderserial<>"") and (ibuyhp<>"") then
        sqlStr = "Insert into [db_sms].[ismsuser].em_tran(tran_phone, tran_callback, tran_status, tran_date, tran_msg ) "
    	sqlStr = sqlStr + " values('" + ibuyhp + "',"
    	sqlStr = sqlStr + " '1644-6030',"
    	sqlStr = sqlStr + " '1',"
    	sqlStr = sqlStr + " getdate(),"
    	sqlStr = sqlStr + " '[�ٹ�����]���� ��� �Ǿ����ϴ�. �ֹ���ȣ : " + iorderserial + "')"
        dbget.Execute sqlStr
    end if

    ''����
    Call SendCsActionMail_GiftCard(id)

    if (IsAutoScript) then
        response.write "S_OK"
    else
        response.write "<script>alert('" & ResultMsg & "');</script>"
        response.write "<script>opener.location.reload();</script>"
        response.write "<script>window.close();</script>"
    end if
    dbget.close()	:	response.End

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
set ocsaslist = Nothing
set orefund = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
