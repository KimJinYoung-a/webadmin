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

'// ===========================================================================
dim id
id = requestCheckVar(request("id"),10)

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



'// ===========================================================================
dim finishuserid, msg, force
dim orgOrderSerial
id           = RequestCheckVar(request("id"),10)
msg          = RequestCheckVar(request("msg"),50)
finishuserid = session("ssBctID")
force = RequestCheckVar(request("force"),10)

if (msg="") and (IsAutoScript) then msg="��������"



'// ===========================================================================
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

'' ����Ƽ�� �� ��Ҹ� ����
if (IsNumeric(orefund.FOneItem.FpaygateTid)<>True) or orefund.FOneItem.Freturnmethod<>"R560" then
    if (IsAutoScript) then
        response.write "S_ERR|����Ƽ�� �� ��� �����մϴ�."
    else
	    response.write "<script>alert('����Ƽ�� �� ��� �����մϴ�.');</script>"
	    response.write "<script>window.close();</script>"
	end if
    dbget.close()	:	response.End
end if



'// ===========================================================================
''orefund.FOneItem.FpaygateTid
dim ResultCode, ResultMsg
dim CancelDate, CancelTime


If orefund.FOneItem.Freturnmethod = "R560" Then

	dim oGicon
	dim ret, bufStr

	set oGicon = new CGiftiCon
	ret = oGicon.reqCouponCancel(CStr(orefund.FOneItem.FpaygateTid), "100100", orefund.FOneItem.Frefundrequire) ''������ȣ, ������ȣ, ��ǰ ��ȯ��

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

End If



'// ===========================================================================
dim returnmethod, refundrequire,refundresult,userid
dim iorderserial, ibuyhp
dim contents_finish
dim sqlStr

if (ResultCode="0000") then
	ResultMsg = "������Ұ� ���������� ó���Ǿ����ϴ�."
end if

contents_finish = "��� " & "[" & ResultCode & "]" & ResultMsg & VbCrlf
contents_finish = contents_finish & "����Ͻ� : " & CancelDate & " " & CancelTime & VbCrlf
contents_finish = contents_finish & "����� ID " & finishuserid

if (ResultCode="0000") then

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

    ''TODO : ������ �ϴ� ����.
    ''Call SendCsActionMail(id)

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
        response.write orefund.FOneItem.FpaygateTid & "<br>"
    end if
end if

%>
<%
set ocsaslist = Nothing
set orefund = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
