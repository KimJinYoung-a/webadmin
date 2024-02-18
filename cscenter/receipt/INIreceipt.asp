<% option Explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<% Response.Addheader "P3P","policyref='/w3c/p3p.xml', CP='NOI DSP LAW NID PSA ADM OUR IND NAV COM'"%>
<%
'###########################################################
' Description : ������ ���ݿ�����,���ݰ�꼭 ����
' History : �̻� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_cashreceiptcls.asp"-->
<%
dim INIpay, PInst, Tid
dim ResultCode, ResultMsg, AuthCode
dim PGAuthDate, PGAuthTime
dim ResultpCRPice, ResultSupplyPrice, ResultTax
dim ResultServicePrice, ResultUseOpt, ResultCashNoAppl
dim AckResult
dim goodname, cr_price, sup_price, tax, srvc_price, buyername
dim buyeremail, buyertel, reg_num, useopt, orderserial, userid, sitename, paymethod
dim sqlStr
dim iidx

goodname = html2db(request.Form("goodname"))
cr_price = request.Form("cr_price")
sup_price = request.Form("sup_price")
tax = request.Form("tax")
srvc_price = request.Form("srvc_price")
buyername = html2db(request.Form("buyername"))
buyeremail = html2db(request.Form("buyeremail"))
buyertel = request.Form("buyertel")
reg_num = request.Form("reg_num")
useopt = request.Form("useopt")
orderserial = request.Form("orderserial")
userid = request.Form("userid")
sitename = request.Form("sitename")
paymethod = request.Form("paymethod")

on Error resume next
sqlStr = " select * from [db_log].[dbo].tbl_cash_receipt where 1=0"
rsget.Open sqlStr,dbget,1,3
rsget.AddNew
rsget("orderserial") = orderserial
rsget("userid") = userid
rsget("sitename") = sitename
rsget("goodname") = goodname
rsget("cr_price") = cr_price
rsget("sup_price") = sup_price
rsget("tax") = tax
rsget("srvc_price") = srvc_price
rsget("buyername") = buyername
rsget("buyeremail") = buyeremail
rsget("buyertel") = buyertel
rsget("reg_num") = reg_num
rsget("useopt") = useopt
rsget("paymethod") = paymethod
rsget("cancelyn") = "N"

rsget.update
iidx = rsget("idx")
rsget.close

if Err then
	response.write "<script>alert('Error - " + Err.description + "');</script>"
	response.write "<script>history.back();</script>"
	response.end
end if

on error goto 0

'*******************************************************************************
'* INIreceipt.asp
'* ���ݰ���(�ǽð� ���������ü, �������Ա�)�� ���� ���ݰ��� ������ ���� ��û�Ѵ�.
'*
'* Date : 2004/12
'* Project : INIpay V4.11 for Unix
'*
'* http://www.inicis.com
'* http://support.inicis.com
'* Copyright (C) 2002 Inicis, Co. All rights reserved.
'*******************************************************************************

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
INIpay.SetActionType CLng(PInst), "receipt"

'###############################################################################
'# 4. �߱� ���� ���� #
'###############################################################################
INIpay.SetField CLng(PInst), "pgid","INIpayRECP"	'PG ID (����)
INIpay.SetField CLng(PInst), "paymethod","CASH"		'���ҹ��
INIpay.SetField CLng(PInst), "spgip", "203.238.3.10" '���� PG IP (����)
INIpay.SetField CLng(PInst), "currency", Request("currency") 'ȭ�����
INIpay.SetField CLng(PInst), "admin", "1111"
INIpay.SetField CLng(PInst), "mid", Request("mid") '�������̵�
INIpay.SetField CLng(PInst), "uip", Request.ServerVariables("REMOTE_ADDR") '��IP
INIpay.SetField CLng(PInst), "goodname", Request("goodname") '��ǰ��
INIpay.SetField CLng(PInst), "cr_price", Request("cr_price") '�� ���� ���� �ݾ�
INIpay.SetField CLng(PInst), "sup_price", Request("sup_price") '���ް���
INIpay.SetField CLng(PInst), "tax", Request("tax") '�ΰ���
INIpay.SetField CLng(PInst), "srvc_price", Request("srvc_price") '�����
INIpay.SetField CLng(PInst), "buyername", Request("buyername") '����
INIpay.SetField CLng(PInst), "buyertel", Request("buyertel") '�̵���ȭ
INIpay.SetField CLng(PInst), "buyeremail", Request("buyeremail") '�̸���
INIpay.SetField CLng(PInst), "reg_num", Request("reg_num") '���ݰ����� �ֹε�Ϲ�ȣ
INIpay.SetField CLng(PInst), "useopt", Request("useopt") '���ݿ����� ����뵵 ("0" - �Һ��� �ҵ������, "1" - ����� ����������)
INIpay.SetField CLng(PInst), "debug", "false" '�α׸��("true"�� �����ϸ� ���� �α׸� ����)

'###############################################################################
'# 5. ���� ��û #
'################
INIpay.StartAction(CLng(PInst))

'###############################################################################
'6. �߱� ��� #
'###############################################################################
'-------------------------------------------------------------------------------
' ��.��� ���� ���ܿ� ����Ǵ� ���� ��� ����
'-------------------------------------------------------------------------------
Tid = INIpay.GetResult(CLng(PInst), "tid") '�ŷ���ȣ
ResultCode = INIpay.GetResult(CLng(PInst), "resultcode") '����ڵ� ("00"�̸� ���Ҽ���)
ResultMsg = INIpay.GetResult(CLng(PInst), "resultmsg") '�������
AuthCode = INIpay.GetResult(CLng(PInst), "authcode") '���ݿ����� �߻� ���ι�ȣ
PGAuthDate = INIpay.GetResult(CLng(PInst), "pgauthdate") '�̴Ͻý� ���γ�¥
PGAuthTime = INIpay.GetResult(CLng(PInst), "pgauthtime") '�̴Ͻý� ���νð�

ResultpCRPice = INIpay.GetResult(CLng(PInst), "ResultpCRPice") '���� �Ǵ� �ݾ�
ResultSupplyPrice = INIpay.GetResult(CLng(PInst), "ResultSupplyPrice") '���ް���
ResultTax = INIpay.GetResult(CLng(PInst), "ResultTax") '�ΰ���
ResultServicePrice = INIpay.GetResult(CLng(PInst), "ResultServicePrice") '�����
ResultUseOpt = INIpay.GetResult(CLng(PInst), "ResultUseOpt") '���౸��
ResultCashNoAppl = INIpay.GetResult(CLng(PInst), "ResultCashNoAppl") '���ι�ȣ

''��� ����
AuthCode = ResultCashNoAppl   ''�̰��� ���ι�ȣ;;


''��� ����
sqlStr = "update [db_log].[dbo].tbl_cash_receipt" + VbCrlf
sqlStr = sqlStr + " set tid='" + CStr(Tid) + "'" + VbCrlf
sqlStr = sqlStr + " , resultcode='" + CStr(ResultCode) + "'" + VbCrlf
sqlStr = sqlStr + " , resultmsg='" + html2db(LeftB(CStr(ResultMsg),200)) + "'" + VbCrlf
sqlStr = sqlStr + " , authcode='" + CStr(AuthCode) + "'" + VbCrlf
sqlStr = sqlStr + " , resultcashnoappl='" + CStr(ResultCashNoAppl) + "'" + VbCrlf
sqlStr = sqlStr + " where idx=" + CStr(iidx)

'response.write sqlStr
dbget.Execute sqlStr

''2016/06/30 �߰�. ������
sqlStr = "update [db_log].[dbo].tbl_cash_receipt" + VbCrlf
sqlStr = sqlStr + " SET evalDt='"&LEFT(PGAuthDate,4)&"-"&MID(PGAuthDate,5,2)&"-"&MID(PGAuthDate,7,2)&" "&LEFT(PGAuthTime,2)&":"&MID(PGAuthTime,3,2)&":"&MID(PGAuthTime,5,2)&"'" + VbCrlf
sqlStr = sqlStr + " where idx=" + CStr(iidx)
dbget.Execute sqlStr
        
''2009�߰�
Dim assignedRow
IF ResultCode = "00" THEN
    sqlStr = " update [db_order].[dbo].tbl_order_master" & VbCrlf
    sqlStr = sqlStr & " set authcode='" & AuthCode & "'" & VbCrlf
    sqlStr = sqlStr + " , cashreceiptreq = (case when accountdiv in ('7', '20') then 'R' else 'S' end) " + VbCrlf
    sqlStr = sqlStr & " where orderserial='" & orderserial & "'"

    sqlStr = " update [db_order].[dbo].tbl_order_master" & VbCrlf
    sqlStr = sqlStr & " set " & VbCrlf
    sqlStr = sqlStr & " authcode = (case when accountdiv in ('7', '20') then '" & AuthCode & "' else authcode end) " + VbCrlf
    sqlStr = sqlStr & " , cashreceiptreq = (case when (accountdiv in ('7', '20')) or (pggubun='NP') then 'R' else 'S' end) " + VbCrlf
    sqlStr = sqlStr & " where orderserial='" & orderserial & "'"

    dbget.Execute sqlStr,assignedRow
    
    IF (assignedRow<1) then
        sqlStr = " update [db_log].[dbo].tbl_old_order_master_2003" & VbCrlf
        sqlStr = sqlStr & " set " & VbCrlf
        sqlStr = sqlStr & " authcode = (case when accountdiv in ('7', '20') then '" & AuthCode & "' else authcode end) " + VbCrlf
        sqlStr = sqlStr & " , cashreceiptreq = (case when accountdiv in ('7', '20') then 'R' else 'S' end) " + VbCrlf
        sqlStr = sqlStr & " where orderserial='" & orderserial & "'"
    
        dbget.Execute sqlStr,assignedRow
    END IF
end if

'###############################################################################
'# 7. ��� ���� Ȯ�� #
'#####################
'���Ұ���� �� �����Ͽ����� �̴Ͻý��� �뺸.
'[����] �� ������ �����Ǹ� ��� �ŷ��� �ڵ���ҵ˴ϴ�.
IF ResultCode = "00" THEN
	AckResult = INIpay.Ack(CLng(PInst))
	IF AckResult <> "SUCCESS" THEN '(����)
		'=================================================================
		' ������� �뺸 ������ ��� �� ������ �̴Ͻý����� �ڵ� ��ҵǹǷ�
		' ���Ұ���� �ٽ� �޾ƿɴϴ�(���� -> ����).
		'=================================================================
		ResultCode = INIpay.GetResult(CLng(PInst), "resultcode")
		ResultMsg = INIpay.GetResult(CLng(PInst), "resultmsg")
	END IF
END IF

'###############################################################################
'# 8. �ν��Ͻ� ���� #
'####################
INIpay.Destroy CLng(PInst)

Set INIpay = Nothing
%>

<%
session("lastreceiptidx") = iidx
%>

<script type="text/javascript">
location.replace('displayreceipt.asp');
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
