<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : cs���� �ڵ� ó��
' History : �̻� ����
'			2020.11.10 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/xmlhttpUtil.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_taxsheetcls.asp"-->
<!-- #include virtual="/lib/email/maillib.asp" -->
<!-- #include virtual="/lib/email/mailLib2.asp"-->
<!-- #include virtual="/cscenter/lib/cs_action_mail_Function.asp"-->
<!-- #include virtual="/lib/classes/smscls.asp" -->
<!-- #include virtual="/lib/incJenkinsCommon.asp" -->

<%
function CheckVaildIP(ref)
    CheckVaildIP = false

    dim VaildIP
    if (application("Svr_Info") = "Dev") then
        VaildIP = Array("110.93.128.107","61.252.133.2","61.252.133.70","61.252.133.10","61.252.133.80","110.93.128.114","110.93.128.113","192.168.1.70", "192.168.1.67")
    else
        VaildIP = Array("110.93.128.107","61.252.133.2","61.252.133.70","61.252.133.10","61.252.133.80","110.93.128.114","110.93.128.113","61.252.133.67", "192.168.1.67")
    end if
    dim i
    for i=0 to UBound(VaildIP)
        if (VaildIP(i)=ref) then
            CheckVaildIP = true
            exit function
        end if
    next
end function

function OneReceiptCancel(orgtid,cancelCause, iResultCode, iResultMsg, iAuthCode)
    dim INIpay, PInst
    dim ResultCode,ResultMsg, CancelDate, CancelTime, Rcash_cancel_noappl

    '###############################################################################
    '# 1. ��ü ���� #
    '################

    ''Set INIpay = Server.CreateObject("INIreceipt41.INIreceiptTX41.1")
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
    INIpay.SetField CLng(PInst), "mid", "teenxteen4" '�������̵�
    INIpay.SetField CLng(PInst), "admin", "1111" 'Ű�н�����(�������̵� ���� ����)
    INIpay.SetField CLng(PInst), "tid", orgtid '����� �ŷ���ȣ(TID)
    INIpay.SetField CLng(PInst), "msg", cancelCause '��� ����
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
    ResultMsg = INIpay.GetResult(CLng(PInst), "resultmsg") '�������
    CancelDate = INIpay.GetResult(CLng(PInst), "pgcanceldate") '�̴Ͻý� ��ҳ�¥
    CancelTime = INIpay.GetResult(CLng(PInst), "pgcanceltime") '�̴Ͻý� ��ҽð�
    Rcash_cancel_noappl = INIpay.GetResult(CLng(PInst), "rcash_cancel_noappl") '���ݿ����� ��� ���ι�ȣ

    '###############################################################################
    '# 7. �ν��Ͻ� ���� #
    '####################
    INIpay.Destroy CLng(PInst)


    iResultCode = ResultCode
    iResultMsg  = ResultMsg
    iAuthCode   = Rcash_cancel_noappl  '' Not AuthCode

    OneReceiptCancel = (iResultCode="00")
end function

function OneReceiptReq(idx,byref iResultCode,byref iResultMsg, byref iAuthCode)
    dim INIpay, PInst

    dim Tid, ResultCode, ResultMsg, AuthCode, PGAuthDate, PGAuthTime
    dim ResultpCRPice, ResultSupplyPrice, ResultTax, ResultServicePrice, ResultUseOpt, ResultCashNoAppl
    dim AckResult

    dim sqlStr
    dim goodname, cr_price, sup_price, tax, srvc_price, buyername, buyertel, buyeremail, reg_num, useopt
    dim reqresultcode

    sqlStr = " select * from [db_log].[dbo].tbl_cash_receipt"
    sqlStr = sqlStr + " where idx=" & idx
    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        goodname    = db2html(rsget("goodname"))
        cr_price    = rsget("cr_price")
        sup_price   = rsget("sup_price")
        tax         = rsget("tax")
        srvc_price  = rsget("srvc_price")
        buyername   = db2html(rsget("buyername"))
        buyertel    = rsget("buyertel")
        buyeremail  = db2html(rsget("buyeremail"))

        reg_num     = rsget("reg_num")
        useopt      = rsget("useopt")
        reqresultcode  = rsget("resultcode")
    end if
    rsget.close

    if (useopt="0") and ((Len(reg_num)<>13) and (Len(reg_num)<>10) and (Len(reg_num)<>11) and (Len(reg_num)<>18)) then
        OneReceiptReq = False
        iResultMsg    = "�ֹι�ȣ/�ڵ��� �ڸ� ����"
        Exit Function
    end if

    if (useopt="1") and ((Len(reg_num)<>13) and (Len(reg_num)<>10) and (Len(reg_num)<>11) and (Len(reg_num)<>18)) then
        OneReceiptReq = False
        iResultMsg    = "����ڹ�ȣ/ �ֹι�ȣ /�ڵ��� �ڸ� ����"
        Exit Function
    end if

    if (reqresultcode<>"R") then
        OneReceiptReq = False
        iResultMsg    = "����� Ȯ��"
        Exit Function
    end if

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
    INIpay.SetField CLng(PInst), "currency", "WON" 'ȭ�����
    INIpay.SetField CLng(PInst), "admin", "1111"
    INIpay.SetField CLng(PInst), "mid", "teenxteen4" '�������̵�
    INIpay.SetField CLng(PInst), "uip", Request.ServerVariables("REMOTE_ADDR") '��IP
    INIpay.SetField CLng(PInst), "goodname", goodname '��ǰ��
    INIpay.SetField CLng(PInst), "cr_price", cr_price '�� ���� ���� �ݾ�
    INIpay.SetField CLng(PInst), "sup_price", sup_price '���ް���
    INIpay.SetField CLng(PInst), "tax", tax         '�ΰ���
    INIpay.SetField CLng(PInst), "srvc_price", srvc_price '�����
    INIpay.SetField CLng(PInst), "buyername", buyername '����
    INIpay.SetField CLng(PInst), "buyertel", buyertel '�̵���ȭ
    INIpay.SetField CLng(PInst), "buyeremail", buyeremail '�̸���
    INIpay.SetField CLng(PInst), "reg_num", reg_num '���ݰ����� �ֹε�Ϲ�ȣ
    INIpay.SetField CLng(PInst), "useopt", useopt '���ݿ����� ����뵵 ("0" - �Һ��� �ҵ������, "1" - ����� ����������)
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
    Tid                 = INIpay.GetResult(CLng(PInst), "tid") '�ŷ���ȣ
    ResultCode          = INIpay.GetResult(CLng(PInst), "resultcode") '����ڵ� ("00"�̸� ���Ҽ���)
    ResultMsg           = INIpay.GetResult(CLng(PInst), "resultmsg") '�������
    AuthCode            = INIpay.GetResult(CLng(PInst), "authcode") '���ݿ����� �߻� ���ι�ȣ
    PGAuthDate          = INIpay.GetResult(CLng(PInst), "pgauthdate") '�̴Ͻý� ���γ�¥
    PGAuthTime          = INIpay.GetResult(CLng(PInst), "pgauthtime") '�̴Ͻý� ���νð�

    ResultpCRPice       = INIpay.GetResult(CLng(PInst), "ResultpCRPice") '���� �Ǵ� �ݾ�
    ResultSupplyPrice   = INIpay.GetResult(CLng(PInst), "ResultSupplyPrice") '���ް���
    ResultTax           = INIpay.GetResult(CLng(PInst), "ResultTax") '�ΰ���
    ResultServicePrice  = INIpay.GetResult(CLng(PInst), "ResultServicePrice") '�����
    ResultUseOpt        = INIpay.GetResult(CLng(PInst), "ResultUseOpt") '���౸��
    ResultCashNoAppl    = INIpay.GetResult(CLng(PInst), "ResultCashNoAppl") '���ι�ȣ

'    response.write Tid & "<br>"
'    response.write ResultCode & "<br>"
'    response.write ResultMsg & "<br>"
'    response.write AuthCode & "<br>"
'    response.write PGAuthDate & "<br>"
'    response.write PGAuthTime & "<br>"
'    response.write ResultpCRPice & "<br>"
'    response.write ResultSupplyPrice & "<br>"
'    response.write ResultTax & "<br>"
'    response.write ResultServicePrice & "<br>"
'    response.write ResultUseOpt & "<br>"
'    response.write ResultCashNoAppl & "<br>"



    iResultCode = ResultCode
    iResultMsg  = ResultMsg
    iAuthCode   = ResultCashNoAppl  '' Not AuthCode

    ''��� ���� - ������ ����� �����ΰ�츸 ����.
    IF ResultCode = "00" THEN
        sqlStr = "update [db_log].[dbo].tbl_cash_receipt" + VbCrlf
        sqlStr = sqlStr + " set tid='" + CStr(Tid) + "'" + VbCrlf
        sqlStr = sqlStr + " , resultcode='" + CStr(ResultCode) + "'" + VbCrlf
        sqlStr = sqlStr + " , resultmsg='" + LeftB(CStr(Replace(ResultMsg,"'","")),100) + "'" + VbCrlf
        sqlStr = sqlStr + " , authcode='" + CStr(AuthCode) + "'" + VbCrlf
        sqlStr = sqlStr + " , resultcashnoappl='" + CStr(ResultCashNoAppl) + "'" + VbCrlf
        sqlStr = sqlStr + " where idx=" + CStr(idx)
        dbget.Execute sqlStr

         ''2016/06/30 �߰�. ������
        sqlStr = "update [db_log].[dbo].tbl_cash_receipt" + VbCrlf
        sqlStr = sqlStr + " SET evalDt='"&LEFT(PGAuthDate,4)&"-"&MID(PGAuthDate,5,2)&"-"&MID(PGAuthDate,7,2)&" "&LEFT(PGAuthTime,2)&":"&MID(PGAuthTime,3,2)&":"&MID(PGAuthTime,5,2)&"'" + VbCrlf
        sqlStr = sqlStr + " where idx=" + CStr(idx)
        dbget.Execute sqlStr

    ELSE
        if (ResultCode="01") and ((Left(iResultMsg,Len("[269051]"))="[269051]") or (Left(iResultMsg,Len("[269050]"))="[269050]") or (Left(iResultMsg,Len("[505658]"))="[505658]")) then
            sqlStr = "update [db_log].[dbo].tbl_cash_receipt" + VbCrlf
            sqlStr = sqlStr + " set cancelyn='F'"
            sqlStr = sqlStr + " , resultmsg='" + LeftB(CStr(Replace(ResultMsg,"'","")),100) + "'" + VbCrlf
            sqlStr = sqlStr + " where idx=" + CStr(idx)

            dbget.Execute sqlStr
        else
            rw iResultMsg
        end if
    End IF

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

    OneReceiptReq = (ResultCode = "00")
end function

sub confirmInsurePayment(InsureCd,orderserial)

	dim objUsafe, result, result_code, result_msg
    On Error Resume Next
	if InsureCd="0" then	'�� tbl_order_master > InsureCd(��� �ڵ�;0-����, 1-����)
		Set objUsafe = CreateObject( "USafeCom.guarantee.1"  )

	'	' Test�� ��
	'	objUsafe.Port = 80
	'	objUsafe.Url = "gateway2.usafe.co.kr"
	'	objUsafe.CallForm = "/esafe/guartrn.asp"

	    ' Real�� ��
	    objUsafe.Port = 80
	    objUsafe.Url = "gateway.usafe.co.kr"
	    objUsafe.CallForm = "/esafe/guartrn.asp"

		objUsafe.gubun	= "C0"				'// �������� (A0:�űԹ߱�, B0:���������, C0:�Ա�Ȯ��)
		objUsafe.EncKey	= ""			'�ΰ��� ��� ��ȣȭ �ȵ�
		objUsafe.mallId	= "ZZcube1010"		'// ���θ�ID
		objUsafe.oId	= CStr(orderserial)	'// �ֹ���ȣ

		'Ȯ��ó�� ����!
		result = objUsafe.confirmPayment

		result_code	= Left( result , 1 )
		result_msg	= Mid( result , 3 )

		'ó����� (��Ȳ�� �°� ���� ���)
		Select Case result_code
			Case "0"
				'response.write "����" & "<BR>" & vbcrlf
				'response.write "�ֹ���ȣ:" & result_msg & "" & vbcrlf
			Case "1"
				'response.write "ó������:" & result_msg & "" & vbcrlf
			Case Else
				'response.write "���ܿ���:" & result_msg & "" & vbcrlf
		End Select

		Set objUsafe = Nothing
	end if
    On Error Goto 0
end sub


dim ref : ref = Request.ServerVariables("REMOTE_ADDR")

if (Not CheckVaildIP(ref)) and (Not CheckJenkinsServerIP(ref)) then
    response.write ref
    dbget.Close()
    response.end
end if

dim IsJenkinsScript : IsJenkinsScript = False
dim jenkinsResponseStatus : jenkinsResponseStatus = "0000"
dim jenkinsResponseText : jenkinsResponseText = ""
if CheckJenkinsServerIP(ref) then
	IsJenkinsScript = True
end if

dim act     : act = requestCheckVar(request("act"),32)
dim param1  : param1 = requestCheckVar(request("param1"),32)
dim sqlStr, i, paramData
dim retCnt : retCnt = 0

dim chkPrint, infoMsg, idx, orderserial, resultcode, cancelyn, preIssuedExists, preIssuedTaxExists, iResultCode, iResultMsg, iAuthCode
dim paramInfo, retParamInfo, RetErr, retval, reg_num
dim userid, buyname, buyhp, buyEmail,InsureCd, vRdSite
dim idArr, paygateTidArr, cnt, divcdArr, orderserialArr
dim osms
dim arrRows, bufStr

select Case act

    Case "cashreceipt"
        ''���꼭������ ���� �ϸ� �ȵ�..
        if application("Svr_Info") = "Dev" then
            response.write "S_ERR|Dev Svr"
            response.end
        end if

        chkPrint = ""
        infoMsg = ""

        ''���� �ݾ� ����ġ ����. 2012/12�߰�
        sqlStr = " update R"
        sqlStr = sqlStr + " set cr_price= m.subtotalPrice"
        sqlStr = sqlStr + " ,tax= convert(int ,m.subtotalPrice*1 /11)"
        sqlStr = sqlStr + " ,sup_price= m.subtotalPrice -convert( int,m.subtotalPrice* 1/11 )"
        sqlStr = sqlStr + "  from db_log.dbo. tbl_cash_receipt R"
        sqlStr = sqlStr + "        Join db_order.dbo.tbl_order_master m"
        sqlStr = sqlStr + "        on R.orderserial=M.orderserial"
        sqlStr = sqlStr + " where R.resultCode='R'"
        sqlStr = sqlStr + " and R.cancelyn='N'"
        sqlStr = sqlStr + " and M.ipkumdiv>7"
        sqlStr = sqlStr + " and M.cancelyn='N' and M.cashreceiptreq='R'"  '' M.cashreceiptreq='R' �����߰� 2015/04/29
        sqlStr = sqlStr + " and M.accountdiv in ('7', '20')"
        sqlStr = sqlStr + " and R.cr_price<>m.subtotalprice"
        sqlStr = sqlStr + " and isNULL(M.pggubun,'')<>'NP'"        ''2016/08/09 �߰�

        dbget.Execute sqlStr

        sqlStr = " select top 10 c.idx, c.orderserial, c.resultcode, c.cancelyn "       ''���� 10
        sqlStr = sqlStr + " from db_order.dbo.tbl_order_master m"
        sqlStr = sqlStr + " 	Join [db_log].[dbo].tbl_cash_receipt c"
        sqlStr = sqlStr + " 	on c.orderserial=m.orderserial"
        sqlStr = sqlStr + " 	and c.resultcode='R'"
        sqlStr = sqlStr + " 	and c.cancelyn='N'"
        '''sqlStr = sqlStr + " 	and c.useopt='1'"
        sqlStr = sqlStr + " where  m.ipkumdiv>='7'"     ''�Ϻ�����̻�. => ���Ϸ��̻�
        sqlStr = sqlStr + " and m.cashreceiptreq='R'"
        sqlStr = sqlStr + " and m.authcode is NULL"
        sqlStr = sqlStr + " and m.accountdiv in ('7','20')"       ''2011 ���� �ǽð� ��ü��..
        sqlStr = sqlStr + " and m.cancelyn='N'"
        sqlStr = sqlStr + " and m.subtotalPrice>0"
        sqlStr = sqlStr + " and m.subtotalPrice=C.cr_price"  '' �κ���ҷ� �ݾ׺��� �߻� ����. ''�ݾ� �����͸�.
        sqlStr = sqlStr + " and c.useopt in ('0','1')"  'useopt �� üũ 2013/10 �������� �� �߻�
        sqlStr = sqlStr + " and c.reg_num<>''"          'reg_num �� üũ 2013/10
        sqlStr = sqlStr + " and isNULL(M.pggubun,'')<>'NP'"
'''sqlStr = sqlStr + " and 1=0"
        ''rsget.Open sqlStr,dbget,1
        rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
        if Not rsget.Eof then
            do until rsget.eof
            chkPrint = chkPrint & rsget("idx") & ","
            rsget.MoveNext
    		loop
        end if
        rsget.close

        if Right(chkPrint,1)="," then chkPrint=Left(chkPrint,Len(chkPrint)-1)
        chkPrint = split(chkPrint,",")

        if UBound(chkPrint)>-1 then
            for i=0 to UBound(chkPrint)

                idx = 0
                sqlStr = " select idx, orderserial, resultcode, cancelyn, reg_num from [db_log].[dbo].tbl_cash_receipt"
                sqlStr = sqlStr + " where idx=" & chkPrint(i)

                rsget.Open sqlStr,dbget,1
                if Not rsget.Eof then
                    idx         = rsget("idx")
                    orderserial = rsget("orderserial")
                    resultcode  = rsget("resultcode")
                    cancelyn    = rsget("cancelyn")
                    reg_num     = rsget("reg_num")
                end if
                rsget.close

                'infoMsg = infoMsg & "[" & idx & "," & orderserial & "]"
                if (idx<>0) then
                    ''����� ���� ���� üũ
                    if (orderserial<>"") then

                       preIssuedExists = False
                       preIssuedTaxExists = False

                       preIssuedTaxExists = chkRegTax(orderserial)

                       sqlStr = " select count(idx) as cnt from  [db_log].[dbo].tbl_cash_receipt"
                       sqlStr = sqlStr + " where orderserial='" + orderserial + "'"
                       sqlStr = sqlStr + " and resultcode='00'"
                       sqlStr = sqlStr + " and cancelyn='N'"
                       sqlStr = sqlStr + " and idx<>"&idx

                       rsget.Open sqlStr,dbget,1
                            preIssuedExists = rsget("cnt")>0
                       rsget.close

                       if (preIssuedExists) then
                            ''infoMsg = infoMsg & " <font color='red'>����� ���� ���� - ����:" & orderserial & "[" & idx & "]" & "</font><br>" & VbCrlf
                            sqlStr = " update [db_log].[dbo].tbl_cash_receipt"
                            sqlStr = sqlStr + " set cancelyn='D'"
                            sqlStr = sqlStr + " where idx=" & CStr(idx)
                            dbget.Execute sqlStr
                       elseif (preIssuedTaxExists<>"none") then
                            ''infoMsg = infoMsg & " <font color='red'>���ݰ�꼭 ���� ���� ���� - ����:" & orderserial & "[" & idx & "]" & "</font><br>" & VbCrlf
                            sqlStr = " update [db_log].[dbo].tbl_cash_receipt"
                            sqlStr = sqlStr + " set cancelyn='D'"
                            sqlStr = sqlStr + " where idx=" & CStr(idx)
                            dbget.Execute sqlStr
                       else
                            iResultCode = ""
                            iResultMsg  = ""
                            if (Not OneReceiptReq(idx, iResultCode, iResultMsg, iAuthCode)) then
                                infoMsg = infoMsg & " <font color='red'>���� ���� :" & "[" & iResultCode & "]" & iResultMsg & "</font><br>" & VbCrlf
                            else
                                ''infoMsg = infoMsg & " ���� ���� :" & "[" & iResultCode & "]" & iResultMsg & "<br>" & VbCrlf

                                sqlStr = " update [db_order].[dbo].tbl_order_master" & VbCrlf
                                sqlStr = sqlStr & " set authcode='" & iAuthCode & "'" & VbCrlf
                                if (reg_num="0100001234") then
                                    sqlStr = sqlStr & " ,cashreceiptreq='J'" & VbCrlf   '' �����߱� 2016/06/22
                                else
                                    sqlStr = sqlStr & " ,cashreceiptreq='S'" & VbCrlf                               ''''2011-04-17
                                end if
                                sqlStr = sqlStr & " where orderserial='" & orderserial & "'"

                                dbget.Execute sqlStr

                                retCnt = retCnt +1
                            end if
                       end if


                    end if
                else
                    infoMsg = "S_ERR|No idx"
                end if
            next

            infoMsg = "S_OK|"&retCnt


        else
            infoMsg = "S_NONE"
        end if

        response.write infoMsg
    Case "cashreceiptDIY"
    ''���꼭������ ���� �ϸ� �ȵ�..
        if application("Svr_Info") = "Dev" then
            response.write "S_ERR|Dev Svr"
            response.end
        end if

        chkPrint = ""
        infoMsg = ""

        sqlStr = " select top 1 c.idx, c.orderserial, c.resultcode, c.cancelyn "
        sqlStr = sqlStr + " from [ACADEMYDB].db_academy.dbo.tbl_academy_order_master m"
       sqlStr = sqlStr + " 	Join [db_log].[dbo].tbl_cash_receipt c"
        sqlStr = sqlStr + " 	on c.orderserial=m.orderserial"
        sqlStr = sqlStr + " 	and c.resultcode='R'"
        sqlStr = sqlStr + " 	and c.cancelyn='N'"
        sqlStr = sqlStr + " 	and LEFT(c.orderserial,1) in ('Y')"
        sqlStr = sqlStr + " where  m.ipkumdiv>='7'"     ''�Ϻ�����̻�.
        sqlStr = sqlStr + " and m.cashreceiptreq='R'"
        sqlStr = sqlStr + " and m.authcode is NULL"
        sqlStr = sqlStr + " and m.accountdiv in ('7','20')"       ''2011 ���� �ǽð� ��ü��..
        sqlStr = sqlStr + " and m.cancelyn='N'"
        sqlStr = sqlStr + " and m.subtotalPrice>0"
        sqlStr = sqlStr + " and c.idx>1100000"
        sqlStr = sqlStr + " and 1=0" ''2016/09/12
        sqlStr = sqlStr + " order by m.idx desc"

        rsget.Open sqlStr,dbget,1
        if Not rsget.Eof then
            do until rsget.eof
            chkPrint = chkPrint & rsget("idx") & ","
            rsget.MoveNext
    		loop
        end if
        rsget.close

        if Right(chkPrint,1)="," then chkPrint=Left(chkPrint,Len(chkPrint)-1)
        chkPrint = split(chkPrint,",")

        if UBound(chkPrint)>-1 then
            for i=0 to UBound(chkPrint)

                idx = 0
                sqlStr = " select idx, orderserial, resultcode, cancelyn from [db_log].[dbo].tbl_cash_receipt"
                sqlStr = sqlStr + " where idx=" & chkPrint(i)

                rsget.Open sqlStr,dbget,1
                if Not rsget.Eof then
                    idx         = rsget("idx")
                    orderserial = rsget("orderserial")
                    resultcode  = rsget("resultcode")
                    cancelyn    = rsget("cancelyn")
                end if
                rsget.close

                'infoMsg = infoMsg & "[" & idx & "," & orderserial & "]"
                if (idx<>0) then
                    ''����� ���� ���� üũ
                    if (orderserial<>"") then

                       preIssuedExists = False
                       preIssuedTaxExists = False

                       preIssuedTaxExists = chkRegTax(orderserial)

                       sqlStr = " select count(idx) as cnt from  [db_log].[dbo].tbl_cash_receipt"
                       sqlStr = sqlStr + " where orderserial='" + orderserial + "'"
                       sqlStr = sqlStr + " and resultcode='00'"
                       sqlStr = sqlStr + " and cancelyn='N'"
                       sqlStr = sqlStr + " and idx<>"&idx

                       rsget.Open sqlStr,dbget,1
                            preIssuedExists = rsget("cnt")>0
                       rsget.close

                       if (preIssuedExists) then
                            ''infoMsg = infoMsg & " <font color='red'>����� ���� ���� - ����:" & orderserial & "[" & idx & "]" & "</font><br>" & VbCrlf
                            sqlStr = " update [db_log].[dbo].tbl_cash_receipt"
                            sqlStr = sqlStr + " set cancelyn='D'"
                            sqlStr = sqlStr + " where idx=" & CStr(idx)
                            dbget.Execute sqlStr
                       elseif (preIssuedTaxExists<>"none") then
                            ''infoMsg = infoMsg & " <font color='red'>���ݰ�꼭 ���� ���� ���� - ����:" & orderserial & "[" & idx & "]" & "</font><br>" & VbCrlf
                            sqlStr = " update [db_log].[dbo].tbl_cash_receipt"
                            sqlStr = sqlStr + " set cancelyn='D'"
                            sqlStr = sqlStr + " where idx=" & CStr(idx)
                            dbget.Execute sqlStr
                       else
                            iResultCode = ""
                            iResultMsg  = ""
                            if (Not OneReceiptReq(idx, iResultCode, iResultMsg, iAuthCode)) then
                                ''infoMsg = infoMsg & " <font color='red'>���� ���� :" & "[" & iResultCode & "]" & iResultMsg & "</font><br>" & VbCrlf
                            else
                                ''infoMsg = infoMsg & " ���� ���� :" & "[" & iResultCode & "]" & iResultMsg & "<br>" & VbCrlf

                                sqlStr = " update [ACADEMYDB].db_academy.dbo.tbl_academy_order_master" & VbCrlf
                                sqlStr = sqlStr & " set authcode='" & iAuthCode & "'" & VbCrlf
                                sqlStr = sqlStr & " ,cashreceiptreq='S'" & VbCrlf                               ''''2011-04-17
                                sqlStr = sqlStr & " where orderserial='" & orderserial & "'"

                                dbget.Execute sqlStr

                                retCnt = retCnt +1
                            end if
                       end if


                    end if
                else
                    infoMsg = "S_ERR|No idx"
                end if
            next

            infoMsg = "S_OK|"&retCnt


        else
            infoMsg = "S_NONE"
        end if

        response.write infoMsg
    Case "cashReEvalList"
        sqlStr = " db_cs.[dbo].[sp_Ten_Get_CashReceiptReEvalList] "&param1
        rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
        if Not rsget.Eof then
            arrRows = rsget.getRows()
        end if
        rsget.close

        if isArray(arrRows) then
            For i=0 To UBound(ArrRows,2)
    			bufStr = bufStr&ArrRows(0,i)&"||"&ArrRows(1,i)&"||"&ArrRows(2,i)&"||"&ArrRows(3,i)&"||"&ArrRows(4,i)&"||"&ArrRows(5,i)&"||"&ArrRows(6,i)&"||"&ArrRows(7,i)&"||"&ArrRows(8,i)&"||"&ArrRows(8,i)&vBCRLF
    		Next
    		bufStr = "S_OK"&vBCRLF&bufStr

    		response.write bufStr
        else
            response.Write "S_NONE"
        end if
    Case "ipkumconfirm"
        paramInfo = Array(Array("@RETURN_VALUE",adInteger,adParamReturnValue,,0) _
            ,Array("@idx"       , adInteger	, adParamInput,   , 0)	_
            ,Array("@BackUserID", adVarchar	, adParamInput, 32, "system")	_
			,Array("@RetVal"	, adInteger  , adParamOutput,, 0) _
			,Array("@MatchOrderSerial", adVarchar	, adParamOutput,11,"") _
		)

        sqlStr = "db_order.dbo.sp_Ten_IpkumConfirm_Proc"
        retParamInfo = fnExecSPOutput(sqlStr,paramInfo)

        RetErr      = GetValue(retParamInfo, "@RETURN_VALUE")   ' ��������
        retval      = GetValue(retParamInfo, "@RetVal")         '
        orderserial  = GetValue(retParamInfo, "@MatchOrderSerial") ' ��Ī�� �ֹ���ȣ

        if (RetErr=0) then
            if (retval=-9) then
                response.Write "S_NONE"
            elseif (retval=-1) then
                response.Write "S_NO"
            elseif (retval=-2) then
                response.Write "S_MANY"
            elseif (retval=1) then
                ''����..
                ''sms�� email �߼�.
                if(orderserial<>"") then
                    sqlStr = "select top 1 userid, buyname, buyhp, buyemail, InsureCd from [db_order].[dbo].tbl_order_master"
                    sqlStr = sqlStr + " where orderserial='" + orderserial + "'"
                    sqlStr = sqlStr + " and cancelyn='N'"

                    rsget.Open sqlStr,dbget,1
                    if Not rsget.Eof then
                        userid  = rsget("userid")
                    	buyname = db2html(rsget("buyname"))
                    	buyhp = db2html(rsget("buyhp"))
                    	buyemail = db2html(rsget("buyemail"))

                    	InsureCd = rsget("InsureCd")
                    end if
                    rsget.close

                    ''SMS �߼�
                    set osms = new CSMSClass
                    osms.SendAcctIpkumOkMsg buyhp,orderserial
                    set osms = Nothing

                    ''Email�߼�
                    call sendmailbankok(buyemail,buyname,orderserial)

                    ''����Ʈ�� �޼���.
''                    dim oXML
''                    If (userid<>"") then
''                        On Error resume Next
''                    		'// POST�� ����
''                    		'�Ǽ����� �˸����� ó�� �������� ���� ����
''                    		set oXML = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")	'xmlHTTP���۳�Ʈ ����
''                            if (application("Svr_Info")<>"Dev") then
''                    			oXML.open "POST", "http://www1.10x10.co.kr/apps/nateon/interface/check_alarmSend.asp", false
''                    		else
''                    			oXML.open "POST", "http://2009www.10x10.co.kr/apps/nateon/interface/check_alarmSend.asp", false
''                    		end if
''                    		oXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
''                    		oXML.send "arid=166&ordsn=" & orderserial	'�Ķ���� ����
''                    		Set oXML = Nothing	'���۳�Ʈ ����
''                        on Error Goto 0
''                    End If

                    if (not IsNULL(InsureCd)) and (InsureCd="0") then
            			call confirmInsurePayment(InsureCd,orderserial)
            		end if
                end if
                response.Write "S_OK"
            else
                response.Write "S_UNKNOWN|"&retval
            end if
        else
            response.Write "S_ERR|Err No - "&RetErr
        end if
    Case "cardCancel"
        ''���꼭������ ��� �ϸ� �ȵ�..
        if application("Svr_Info") = "Dev" then
			if IsJenkinsScript then
				Call WriteJenkinsJsonResponse("1001", "Dev Svr")
			else
				response.write "S_ERR|Dev Svr"
			end if
            dbget.Close() : response.end
        end if

        if application("Svr_Info") = "Dev" then
            sqlStr = " select top 1 a.id, m.orderserial, m.paygateTid "
            sqlStr = sqlStr & " from "
            sqlStr = sqlStr & "[db_cs].dbo.tbl_new_as_list a"
            sqlStr = sqlStr & "	Join db_order.dbo.tbl_order_master m"
            sqlStr = sqlStr & "	on a.orderserial=m.orderserial"
            sqlStr = sqlStr & " where a.currstate='B001'"
            sqlStr = sqlStr & " and a.deleteyn='N'"
            sqlStr = sqlStr & " and a.divcd='A007'"
            sqlStr = sqlStr & " and m.ipkumdiv >='4'"
			sqlStr = sqlStr & " and (Left(m.paygateTid,9) in ('IniTechPG','INIMX_CAR','INIMX_ISP','KCTEN0001','StdpayCAR','StdpayISP','INIMX_RTP','StdpayRTPY', 'StdpayDBN') or (m.pggubun='NP'))"  ''���̹� ���� �߰���.
        else
            '' N(5) ȸ �ݺ� ����ð����� asp�� ���μ����� ����־� �ٸ� ť�� ������ �߻��ϴµ�, �ɰ��� ���.
            if (param1="0") then  ''ù��°�� �ð��� �� �ɸ�.
                sqlStr = " select top 1 "
            else
                if (hour(now())=8) then ''2016/12/19
                    sqlStr = " select top 8 "
                else
                    sqlStr = " select top 4 "
                end if
            end if
            sqlStr = sqlStr & " a.id, m.orderserial, m.paygateTid "
            sqlStr = sqlStr & " from "
            sqlStr = sqlStr & "[db_cs].dbo.tbl_new_as_list a"
            sqlStr = sqlStr & "	    Join db_order.dbo.tbl_order_master m"
            sqlStr = sqlStr & "	    on a.orderserial=m.orderserial"
            sqlStr = sqlStr & "	    Join [db_cs].dbo.tbl_as_refund_info f"
            sqlStr = sqlStr & "	    on a.id=f.asid"
            sqlStr = sqlStr & "	    and f.returnmethod not in ('R120','R022')"                     ''2011-07-25 �߰�
            sqlStr = sqlStr & " where a.currstate='B001'"
            sqlStr = sqlStr & " and a.deleteyn='N'"
            sqlStr = sqlStr & " and a.divcd='A007'"
            sqlStr = sqlStr & " and m.ipkumdiv >='4'"
            sqlStr = sqlStr & " and (Left(m.paygateTid,9) in ('IniTechPG','INIMX_CAR','INIMX_ISP','KCTEN0001','StdpayCAR','StdpayISP','INIMX_RTP','StdpayRTPY', 'StdpayDBN') or (m.pggubun in ('NP','PY','KK','TS','CH')))"  ''���̹� ���� �߰���.
            sqlStr = sqlStr & " and datediff(n, a.regdate,getdate())>10" ''������ 15�� ������ ''2016/04/04 => 15�� =>10 2016/12/14
            if (hour(now())>=23) or (hour(now())<2) then  ''2016/12/05 �߰�
                sqlStr = sqlStr & " and (isNULL(m.pggubun,'')<>'NP')"
            end if
			sqlStr = sqlStr & " and a.id not in (5259617) "		'// ���� CS���� ó��
            sqlStr = sqlStr & " order by a.id " ''2016/03/14 �߰�

			if (hour(now()) = 20) then
				'// ����Ʈī�� ī����Ұ�(�Ϸ� �ѽð�, 2��)
				sqlStr = " select top 5 a.id, m.giftOrderSerial AS orderserial, m.paydateid AS paygateTid "
				sqlStr = sqlStr & " from "
				sqlStr = sqlStr & " [db_cs].dbo.tbl_new_as_list a "
				sqlStr = sqlStr & " 	    Join [db_order].[dbo].[tbl_giftcard_order] m "
				sqlStr = sqlStr & " 	    on a.orderserial=m.giftOrderSerial "
				sqlStr = sqlStr & " 	    Join [db_cs].dbo.tbl_as_refund_info f "
				sqlStr = sqlStr & " 	    on a.id=f.asid "
				sqlStr = sqlStr & " 	    and f.returnmethod not in ('R120','R022') "
				sqlStr = sqlStr & " where a.currstate='B001' "
				sqlStr = sqlStr & " and a.deleteyn='N' "
				sqlStr = sqlStr & " and a.divcd='A007' "
				sqlStr = sqlStr & " and m.ipkumdiv >='4' "
				sqlStr = sqlStr & " and (Left(m.paydateid,9) in ('IniTechPG','INIMX_CAR','INIMX_ISP','KCTEN0001','StdpayCAR','StdpayISP','INIMX_RTP','StdpayRTPY', 'StdpayDBN')) "
				sqlStr = sqlStr & " and datediff(n, a.regdate,getdate())>10 "
				sqlStr = sqlStr & " order by a.id "
			end if
        end if
        ''rsget.Open sqlStr,dbget,1
        rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

        cnt = rsget.RecordCount
        ReDim idArr(cnt)
        ReDim paygateTidArr(cnt)
        i = 0
        if Not rsget.Eof then
            do until rsget.eof
            idArr(i) = rsget("id")
            paygateTidArr(i) = rsget("paygateTid")
            i=i+1
            rsget.MoveNext
    		loop
        end if
        rsget.close

        if (param1="0") then
			if IsJenkinsScript then
				Call WriteJenkinsJsonResponse("0000", "SKIP")
			else
				response.Write "S_SKIP"
			end if
            dbget.Close() : response.end
        elseif (cnt<1) then
			if IsJenkinsScript then
				Call WriteJenkinsJsonResponse("0000", "NONE")
			else
				response.Write "S_NONE"
			end if
            dbget.Close() : response.end
        else
            for i=LBound(idArr) to UBound(idArr)
                if (idArr(i)<>"") then
                    paramData = "redSsnKey=system&id="&idArr(i)&"&tid="&paygateTidArr(i)&"&msg="

                    ''response.write paramData&"<br>"
                    if (application("Svr_Info")<>"Dev") then
						if (hour(now()) = 20) then
							'// ����Ʈī�� ī����Ұ�(�Ϸ� �ѽð�, 2��)
							retVal = SendReq("http://stscm.10x10.co.kr/cscenter/giftcard/pop_giftcard_CardCancel_process.asp",paramData)
						else
							retVal = SendReq("http://stscm.10x10.co.kr/cscenter/action/pop_CardCancel_process.asp",paramData)
						end if
                    else
                         retVal = SendReq("http://testwebadmin.10x10.co.kr/cscenter/action/pop_CardCancel_process.asp",paramData)
                    end if

					if IsJenkinsScript then
						if retVal <> "S_OK" then
							if jenkinsResponseStatus = "0000" then
								jenkinsResponseStatus = "2000"
							end if
							if jenkinsResponseText = "" then
								jenkinsResponseText = "CS ASID : "
							end if
							jenkinsResponseText = jenkinsResponseText & idArr(i) & ", "
							''jenkinsResponseText = jenkinsResponseText & retVal
						end if
					else
						response.write retVal&VbCRLF
					end if
                end if
            next

			if IsJenkinsScript then
				if (jenkinsResponseText = "") then
					jenkinsResponseText = "OK"
				end if
				Call WriteJenkinsJsonResponse(jenkinsResponseStatus, jenkinsResponseText)
			end if
        end if
    Case "cardPartialCancel"
        ''���꼭������ ��� �ϸ� �ȵ�..
        if application("Svr_Info") = "Dev" then
            response.write "S_ERR|Dev Svr"
            response.end
        end if

        if application("Svr_Info") = "Dev" then
            sqlStr = " select top 1 a.id, m.orderserial, m.paygateTid "
            sqlStr = sqlStr & " from "
            sqlStr = sqlStr & "[db_cs].dbo.tbl_new_as_list a"
            sqlStr = sqlStr & "	Join db_order.dbo.tbl_order_master m"
            sqlStr = sqlStr & "	on a.orderserial=m.orderserial"
			sqlStr = sqlStr & " 	Join [db_cs].dbo.tbl_as_refund_info f "
            sqlStr = sqlStr & " 	on a.id=f.asid "
            sqlStr = sqlStr & " 	and f.returnmethod in ('R120') "
            sqlStr = sqlStr & " where a.currstate='B001'"
            sqlStr = sqlStr & " and a.deleteyn='N'"
            sqlStr = sqlStr & " and a.divcd='A007'"
            sqlStr = sqlStr & " and m.ipkumdiv >='4' "
			sqlStr = sqlStr & " and (Left(m.paygateTid,9) in ('IniTechPG','INIMX_CAR','INIMX_ISP','KCTEN0001','StdpayCAR','StdpayISP', 'StdpayDBN') or (m.pggubun in ('NP','PY','KK','TS','CH'))) "
			''sqlStr = sqlStr & " and datediff(n, a.regdate,getdate())>30 "			'// ���� �� 30�� ���� ��
        else
            '' N(5) ȸ �ݺ� ����ð����� asp�� ���μ����� ����־� �ٸ� ť�� ������ �߻��ϴµ�, �ɰ��� ���.
            if (param1="0") then  ''ù��°�� �ð��� �� �ɸ�.
                sqlStr = " select top 1 "
            else
                sqlStr = " select top 5 "  ''���� ���� 3=>5 2016/12/19
            end if
            sqlStr = sqlStr & " a.id, m.orderserial, m.paygateTid "
            sqlStr = sqlStr & " from "
            sqlStr = sqlStr & " 	[db_cs].dbo.tbl_new_as_list a "
            sqlStr = sqlStr & " 	Join db_order.dbo.tbl_order_master m "
            sqlStr = sqlStr & " 	on a.orderserial=m.orderserial "
            sqlStr = sqlStr & " 	Join [db_cs].dbo.tbl_as_refund_info f "
            sqlStr = sqlStr & " 	on a.id=f.asid "
            sqlStr = sqlStr & " 	and f.returnmethod in ('R120','R022') "  ''R022 �߰� 2016/12/06, ,R400 �߰� 2016/12/21
            sqlStr = sqlStr & " where a.currstate='B001' "
            sqlStr = sqlStr & " and a.deleteyn='N' "
            sqlStr = sqlStr & " and a.divcd='A007' "
            sqlStr = sqlStr & " and m.ipkumdiv >='4' "
            sqlStr = sqlStr & " and (Left(m.paygateTid,9) in ('IniTechPG','INIMX_CAR','INIMX_ISP','KCTEN0001','StdpayCAR','StdpayISP', 'StdpayDBN') or (m.pggubun in ('NP','PY','KK','TS','CH'))) " '' kakaopay �κ���� �ϴ�����(20181210 ���� �ٽ� ����)
            sqlStr = sqlStr & " and datediff(n, a.regdate,getdate())>20 "			'// ���� �� 30�� ���� �� >20��
			sqlStr = sqlStr & " and datediff(hour, a.regdate,getdate())<=12 "			'// 12�ð� �ʰ��Ǹ� ��ŵ(�κ���� Ƚ�� �ʰ��Ǽ� �������� : �κ���� Ƚ�� �ʰ� 12���̻��̸� ���� �κ���� �ȵ�)
            if (hour(now())>=23) or (hour(now())<2) then  ''2016/12/05 �߰�
                sqlStr = sqlStr & " and (isNULL(m.pggubun,'')<>'NP')"
            end if
            sqlStr = sqlStr & " and m.orderserial not in ('17061919271','17062655010','17070614323')" ''2017/07/14
            sqlStr = sqlStr & " order by a.id "
        end if
        ''rsget.Open sqlStr,dbget,1
        rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

        cnt = rsget.RecordCount
        ReDim idArr(cnt)
        ReDim paygateTidArr(cnt)
        i = 0
        if Not rsget.Eof then
            do until rsget.eof
				idArr(i) = rsget("id")
				paygateTidArr(i) = rsget("paygateTid")
				i=i+1
				rsget.MoveNext
    		loop
        end if
        rsget.close

        if (cnt<1) then
            response.Write "S_NONE"
            dbget.Close() : response.end
        else
            for i=LBound(idArr) to UBound(idArr)
                if (idArr(i)<>"") then
                    paramData = "redSsnKey=system&id="&idArr(i)&"&tid="&paygateTidArr(i)&"&msg="

                    ''response.write paramData&"<br>"
                    if (application("Svr_Info")<>"Dev") then
                        retVal = SendReq("http://stscm.10x10.co.kr/cscenter/action/pop_PartialCardCancel_process.asp",paramData)
                    else
                        retVal = SendReq("http://testwebadmin.10x10.co.kr/cscenter/action/pop_PartialCardCancel_process.asp",paramData)
                    end if

                    response.write retVal&VbCRLF
                end if
            next
        end if
    Case "cardCancelAcademy"   ''dailyAutoJob_ACA
        ''���꼭������ ��� �ϸ� �ȵ�..
        if application("Svr_Info") = "Dev" then
            response.write "S_ERR|Dev Svr"
            response.end
        end if

        if application("Svr_Info") = "Dev" then
            sqlStr = " select top 1 a.id, m.orderserial, m.paygateTid  "
            sqlStr = sqlStr & " from "
            sqlStr = sqlStr & " [db_cs].dbo.tbl_new_as_list a"
            sqlStr = sqlStr & " Join [ACADEMYDB].db_academy.dbo.tbl_academy_order_master m"
            sqlStr = sqlStr & " on a.orderserial=m.orderserial"
            sqlStr = sqlStr & " where a.currstate='B001'"
            sqlStr = sqlStr & " and a.deleteyn='N'"
            sqlStr = sqlStr & " and a.divcd='A007'"
            sqlStr = sqlStr & " and m.cancelyn='Y'"
            sqlStr = sqlStr & " and m.ipkumdiv >=4"
            sqlStr = sqlStr & " and m.ipkumdiv <7"
            sqlStr = sqlStr & " and Left(m.paygateTid,9)='IniTechPG'"
        else
            sqlStr = " select top 5 a.id, m.orderserial, m.paygateTid "
            sqlStr = sqlStr & " from "
            sqlStr = sqlStr & " [db_academy].dbo.tbl_academy_as_list a"
            sqlStr = sqlStr & " Join db_academy.dbo.tbl_academy_order_master m"
            sqlStr = sqlStr & " on a.orderserial=m.orderserial"
            sqlStr = sqlStr & " where a.currstate='B001'"
            sqlStr = sqlStr & " and a.deleteyn='N'"
            sqlStr = sqlStr & " and a.divcd='A007'"
            sqlStr = sqlStr & " and m.cancelyn='Y'"
            sqlStr = sqlStr & " and m.ipkumdiv >=4"
            sqlStr = sqlStr & " and m.ipkumdiv <7"
            sqlStr = sqlStr & " and Left(m.paygateTid,9)='IniTechPG'"
        end if
        rsACADEMYget.Open sqlStr,dbACADEMYget,1
        cnt = rsACADEMYget.RecordCount
        ReDim idArr(cnt)
        ReDim paygateTidArr(cnt)
        i = 0
        if Not rsACADEMYget.Eof then
            do until rsACADEMYget.eof
            idArr(i) = rsACADEMYget("id")
            paygateTidArr(i) = rsACADEMYget("paygateTid")
            i=i+1
            rsACADEMYget.MoveNext
    		loop
        end if
        rsACADEMYget.close

        if (cnt<1) then
            response.Write "S_NONE"
            dbget.Close() : dbACADEMYget.Close() :  response.end
        else
            for i=LBound(idArr) to UBound(idArr)
                if (idArr(i)<>"") then
                    paramData = "redSsnKey=system&id="&idArr(i)&"&tid="&paygateTidArr(i)&"&msg="

                    ''response.write paramData&"<br>"
                    if (application("Svr_Info")<>"Dev") then
                         retVal = SendReq("http://stscm.10x10.co.kr/cscenterv2/cs/pop_CardCancel_process.asp",paramData)
                    else
                         retVal = SendReq("http://testwebadmin.10x10.co.kr/cscenterv2/cs/pop_CardCancel_process.asp",paramData)
                    end if

                    response.write retVal&VbCRLF
                end if
            next
        end if
    Case "bankmail"
        sqlStr = " select top " + CStr(param1) + " idx,orderserial"
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master"
		sqlStr = sqlStr + " where 1=1"
		sqlStr = sqlStr + " and datediff(day,regdate,getdate())>" + CStr(5)
		sqlStr = sqlStr + " and datediff(day,regdate,getdate())<=" + CStr(10)
		sqlStr = sqlStr + " and cancelyn='N'"
		sqlStr = sqlStr + " and accountdiv='7'"
		sqlStr = sqlStr + " and ipkumdiv='2'"
		sqlStr = sqlStr + " and sitename='10x10'"
		sqlStr = sqlStr + " and jumundiv not in ('6','7','9')"  '' ��ȯ�ֹ�,���̳ʽ��ֹ�,������� ����
'		sqlStr = sqlStr + " and accountno in ("
'        sqlStr = sqlStr + " '���� 470301-01-014754'"
'        sqlStr = sqlStr + " ,'���� 100-016-523130'"
'        sqlStr = sqlStr + " ,'�츮 092-275495-13-001'"
'        sqlStr = sqlStr + " ,'�ϳ� 146-910009-28804'"
'        sqlStr = sqlStr + " ,'��� 277-028182-01-046'"
'        sqlStr = sqlStr + " ,'���� 029-01-246118'"
'        sqlStr = sqlStr + " )"
		sqlStr = sqlStr + " and orderserial not in ("
		sqlStr = sqlStr + "     select orderserial from [db_temp].[dbo].[tbl_bankmail_sendlist]"
		sqlStr = sqlStr + " )"
		sqlStr = sqlStr + " order by idx "

		rsget.Open sqlStr,dbget,1
		idArr = ""
		if Not (rsget.Eof) then
		    do until rsget.eof
		        idArr = idArr & rsget("orderserial") & ","
		        rsget.MoveNext
    		loop
		end if
		rsget.Close

		idArr = Replace(idArr," ","")
		if Right(idArr,1)="," then idArr=Left(idArr,Len(idArr)-1)

		if (idArr="") then
		    response.Write "S_NONE"
            dbget.Close() : response.end
		ELSE
    		paramData = "redSsnKey=system&mode=mail&orderSerialArray="&idArr&"&msg="

    		if (application("Svr_Info")<>"Dev") then
                 retVal = SendReq("http://webadmin.10x10.co.kr/admin/ordermaster/dobankacct.asp",paramData)
            else
                 retVal = SendReq("http://testwebadmin.10x10.co.kr/admin/ordermaster/dobankacct.asp",paramData)
            end if

            response.write retVal&VbCRLF
        END IF
    CASE "bankdel"
        dim searchEnddate
        dim searchStartdate

        dim searchEnddateTicket
        dim searchStartdateTicket

        '======================================================================
        '�Ϲ��ֹ� : Ƽ���ֹ� ����
        searchEnddate = CStr(dateAdd("d",3*-1,now()))
        searchEnddate = Left(searchEnddate,10)

        searchStartdate = CStr(dateAdd("d",-61,now()))
        searchStartdate = Left(searchStartdate,10)

        'Ƽ���ֹ� : �ֹ� ������ �� 12�ñ��� �Աݵ��� ������ ��ҵȴ�.
        searchEnddateTicket = CStr(dateAdd("d",2*-1,now()))
        searchEnddateTicket = Left(searchEnddateTicket,10)

        searchStartdateTicket = CStr(dateAdd("d",-30,now()))
        searchStartdateTicket = Left(searchStartdateTicket,10)

		'Ƽ���ֹ� : (m.jumundiv =  '4')
        sqlStr = " select top " + CStr(param1) + " m.idx,m.orderserial"
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m"
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + " 	and m.cancelyn='N'"
		sqlStr = sqlStr + " 	and m.accountdiv='7'"
		sqlStr = sqlStr + " 	and m.ipkumdiv='2'"
		sqlStr = sqlStr + "     and m.jumundiv not in ('6','9')"  '' ��ȯ�ֹ�,���̳ʽ��ֹ� ����
		sqlStr = sqlStr + " 	and "
		sqlStr = sqlStr + " 		( "
		sqlStr = sqlStr + " 			(m.jumundiv <> '4' and m.regdate>'" & searchStartdate & "' and m.regdate<'" & searchEnddate & "') "
		sqlStr = sqlStr + " 			or "
		sqlStr = sqlStr + " 			(m.jumundiv =  '4' and m.regdate>'" & searchStartdateTicket & "' and m.regdate<'" & searchEnddateTicket & "') "
		sqlStr = sqlStr + " 		) "
		sqlStr = sqlStr + "     and m.orderserial not in ("  '' db_order.dbo.tbl_order_CyberAccountLog ''�Ⱓ ������ ��� �粸�� �ҵ� //2014/06/05
        sqlStr = sqlStr + "         select orderserial from db_order.dbo.tbl_order_CyberAccountLog "
        sqlStr = sqlStr + "         where isMatched='N' "
        sqlStr = sqlStr + "         and isDELETE='N' "
        sqlStr = sqlStr + "         and CLOSEDATE>getdate() "
        'sqlStr = sqlStr + "         and differencekey>0 " '// �Ⱓ������ �ƴ� closedate �������� ���� �Աݱ����� ���� ������ ������ ��� ó�� ���� ����(2021.12.01)
        sqlStr = sqlStr + "     ) "
		'sqlStr = sqlStr + " and accountno in ("
        'sqlStr = sqlStr + " '���� 470301-01-014754'"
        'sqlStr = sqlStr + " ,'���� 100-016-523130'"
        'sqlStr = sqlStr + " ,'�츮 092-275495-13-001'"
        'sqlStr = sqlStr + " ,'�ϳ� 146-910009-28804'"
        'sqlStr = sqlStr + " ,'��� 277-028182-01-046'"
        'sqlStr = sqlStr + " ,'���� 029-01-246118'"
        'sqlStr = sqlStr + " )"
		sqlStr = sqlStr + " order by m.idx"

		rsget.Open sqlStr,dbget,1
		idArr = ""
		if Not (rsget.Eof) then
		    do until rsget.eof
		        idArr = idArr & rsget("idx") & ","
		        rsget.MoveNext
    		loop
		end if
		rsget.Close

		idArr = Replace(idArr," ","")
		if Right(idArr,1)="," then idArr=Left(idArr,Len(idArr)-1)

		if (idArr="") then
		    response.Write "S_NONE"
            dbget.Close() : response.end
		ELSE
    		paramData = "redSsnKey=system&mode=del&orderidx="&idArr&"&msg="

    		if (application("Svr_Info")<>"Dev") then
                 retVal = SendReq("http://webadmin.10x10.co.kr/admin/ordermaster/dobankacct.asp",paramData)
            else
                 retVal = SendReq("http://testwebadmin.10x10.co.kr/admin/ordermaster/dobankacct.asp",paramData)
            end if

            response.write retVal&VbCRLF
        END IF
    Case "mobileCancel"
        ''���꼭������ ��� �ϸ� �ȵ�..
        if application("Svr_Info") = "Dev" then
            'response.write "S_ERR|Dev Svr"
            'response.end
        end if

        if application("Svr_Info") = "Dev" then
            sqlStr = " select top 3 a.id, m.orderserial, m.paygateTid "
            sqlStr = sqlStr & " from "
            sqlStr = sqlStr & "[db_cs].dbo.tbl_new_as_list a"
            sqlStr = sqlStr & "	Join db_order.dbo.tbl_order_master m"
            sqlStr = sqlStr & "	on a.orderserial=m.orderserial"
            sqlStr = sqlStr & " where a.currstate='B001'"
            sqlStr = sqlStr & " and a.deleteyn='N'"
            sqlStr = sqlStr & " and a.divcd='A007'"
            sqlStr = sqlStr & " and m.cancelyn='Y'"
            sqlStr = sqlStr & " and m.ipkumdiv in (4,5)"
            sqlStr = sqlStr & " and Left(m.paygateTid,9)='IniTechPG'"
        else
            sqlStr = " select top 3 a.id, m.orderserial, m.paygateTid, isNull(m.rdsite,'') AS rdsite "
            sqlStr = sqlStr & " from "
            sqlStr = sqlStr & "[db_cs].dbo.tbl_new_as_list a"
            sqlStr = sqlStr & "	Join db_order.dbo.tbl_order_master m"
            sqlStr = sqlStr & "	on a.orderserial=m.orderserial"
            sqlStr = sqlStr & "	Join [db_cs].dbo.tbl_as_refund_info f"
            sqlStr = sqlStr & " 	on a.id=f.asid "
            sqlStr = sqlStr & " 	and f.returnmethod in ('R400')"
            sqlStr = sqlStr & " where a.currstate='B001'"
            sqlStr = sqlStr & " and a.deleteyn='N'"
            sqlStr = sqlStr & " and a.divcd='A007'"
            sqlStr = sqlStr & " and m.ipkumdiv >=4"
            sqlStr = sqlStr & " and m.accountdiv = '400'"
            sqlStr = sqlStr & " and datediff(n, a.regdate,getdate())>20 "			'// ���� �� 20�� ���� ��
        end if
        rsget.Open sqlStr,dbget,1
        cnt = rsget.RecordCount
        ReDim idArr(cnt)
        ReDim paygateTidArr(cnt)
        ReDim vRdSite(cnt)
        i = 0
        if Not rsget.Eof then
            do until rsget.eof
            idArr(i) = rsget("id")
            paygateTidArr(i) = rsget("paygateTid")
			vRdSite(i) = rsget("rdsite")
            i=i+1
            rsget.MoveNext
    		loop
        end if
        rsget.close

        if (cnt<1) then
            response.Write "S_NONE"
            dbget.Close() : response.end
        else
            for i=LBound(idArr) to UBound(idArr)
                if (idArr(i)<>"") then
                	If LEFT(vRdSite(i),6) = "mobile" Then
                    	paramData = "redSsnKey=system&id="&idArr(i)&"&tid="&paygateTidArr(i)&"&msg=&rdsite=mobile"
                    Else
                		paramData = "redSsnKey=system&id="&idArr(i)&"&tid="&paygateTidArr(i)&"&msg="
                	End If

                    ''response.write paramData&"<br>"
                    if (application("Svr_Info")<>"Dev") then
                         retVal = SendReq("http://stscm.10x10.co.kr/cscenter/action/pop_CardCancel_process.asp",paramData)
                    else
                         retVal = SendReq("http://testwebadmin.10x10.co.kr/cscenter/action/pop_CardCancel_process.asp",paramData)
                    end if

                    response.write retVal&VbCRLF
                end if
            next
        end if
	Case "actFinishCsB006"
		'// CS ��üó���Ϸ�� �ڵ� �Ϸ�ó��
		'// ��ǰ, ������߼�
		'// ��ȯȸ��, ��ȯ���, ��ȯȸ��(��ǰ����), ��ȯ���(��ǰ����) : ��ȯ���� ��ȯȸ�� �Ϸ��, ��Ÿȸ�� : �����ȣ �Է¾ȵǸ� ó������
		'// ��� �� ��ǰ �� ��ġ��ȯ��
		'// ���ϸ��� ����
		sqlStr = " update a "
		sqlStr = sqlStr + " set a.needChkYN = 'N', a.currstate = 'B006' "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	[db_cs].[dbo].[tbl_new_as_list] a "
		sqlStr = sqlStr + " 	join [db_cs].[dbo].[tbl_as_refund_info] r on a.id = r.asid "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + " 	and a.divcd = 'A003' "
		sqlStr = sqlStr + " 	and ((r.returnmethod = 'R910' and a.refasid is not NULL) or (r.returnmethod = 'R900')) "
		sqlStr = sqlStr + " 	and a.currstate = 'B001' "
		sqlStr = sqlStr + " 	and a.deleteyn = 'N' "
		sqlStr = sqlStr + " 	and a.regdate >= convert(varchar(10), DateAdd(day, -1, getdate()), 121) "
		sqlStr = sqlStr + " 	and DateDiff(MINUTE, a.regdate, getdate()) >= 30 "
		dbget.Execute sqlStr

		'// ��ȯ��� ��üó���Ϸ� ��, ��ȯȸ�� �Ϸ���ϸ� 1���� �� �ڵ��Ϸ�ó��
		sqlStr = " update b "
		sqlStr = sqlStr + " set b.currstate = 'B007', b.finishuser = 'system', b.finishdate = getdate() "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	[db_cs].[dbo].[tbl_new_as_list] a "
		sqlStr = sqlStr + " 	join [db_cs].[dbo].[tbl_new_as_list] b on b.refasid = a.id "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + " 	and a.divcd = 'A000' "
		sqlStr = sqlStr + " 	and a.currstate = 'B006' "
		sqlStr = sqlStr + " 	and b.currstate < 'B007' "
		sqlStr = sqlStr + " 	and a.finishdate < DateAdd(day, -7, getdate()) "
		sqlStr = sqlStr + " 	and a.deleteyn = 'N' "
		sqlStr = sqlStr + " 	and b.deleteyn = 'N' "
		sqlStr = sqlStr + " 	and a.needChkYN = 'N' "
		sqlStr = sqlStr + " 	and IsNull(b.needChkYN, 'N') = 'N' "
		dbget.Execute sqlStr

		sqlStr = " select top " + CStr(param1) + " a.id as asid, a.divcd, a.orderserial "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	[db_cs].[dbo].[tbl_new_as_list] a "
		sqlStr = sqlStr + " 	left join [db_cs].[dbo].[tbl_new_as_list] a2 on a.id = a2.refasid "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + " 	and a.currstate = 'B006' "
		sqlStr = sqlStr + " 	and ( "
		sqlStr = sqlStr + " 		(a.divcd in ('A008', 'A004', 'A010', 'A001', 'A003')) "
		sqlStr = sqlStr + " 		or "
		sqlStr = sqlStr + " 		(a.divcd in ('A012', 'A112', 'A011', 'A111', 'A200') and a.songjangno <> '') "
		sqlStr = sqlStr + " 		or "
		sqlStr = sqlStr + " 		(a.divcd in ('A000', 'A100') and a2.currstate = 'B007' and a.songjangno <> '') "
		sqlStr = sqlStr + " 	) "
		sqlStr = sqlStr + " 	and a.deleteyn = 'N' "
		sqlStr = sqlStr + " 	and a.needChkYn = 'N' "
		sqlStr = sqlStr + " order by "
		sqlStr = sqlStr + " 	a.finishdate desc "
		''response.Write sqlStr & "<br>"
		rsget.Open sqlStr,dbget,1
        cnt = rsget.RecordCount
        ReDim idArr(cnt)
        ReDim divcdArr(cnt)
		ReDim orderserialArr(cnt)

        i = 0
        if Not rsget.Eof then
            do until rsget.eof
				idArr(i) = rsget("asid")
				divcdArr(i) = rsget("divcd")
				orderserialArr(i) = rsget("orderserial")
				i=i+1
				rsget.MoveNext
    		loop
        end if
        rsget.close

        if (cnt<1) then
            response.Write "S_NONE"
            dbget.Close() : response.end
        else
			response.Write "S_OK" & vbCrLf
			for i=LBound(idArr) to UBound(idArr)
				response.Write idArr(i) & "|" & divcdArr(i) & "|" & orderserialArr(i) & vbCrLf
			next
        end if

	Case "actFinishCsautocancel"
        'dbget.Close() : response.end
		'// CS ǰ����ǰ ��ü���,�κ���� �ڵ� �Ϸ�ó��

        sqlStr = "update a set a.needChkYN='N'" & vbcrlf
        sqlStr = sqlStr & " from db_cs.dbo.tbl_new_as_list a with (nolock)" & vbcrlf
        sqlStr = sqlStr & " join db_cs.dbo.tbl_new_as_detail b with (nolock)" & vbcrlf
        sqlStr = sqlStr & "     on a.id = b.masterid" & vbcrlf
        sqlStr = sqlStr & " join db_temp.dbo.tbl_mibeasong_list l with (nolock)" & vbcrlf
        sqlStr = sqlStr & "     on a.orderserial=l.orderserial" & vbcrlf
        sqlStr = sqlStr & "     and b.itemid = l.itemid" & vbcrlf
        sqlStr = sqlStr & "     and b.itemoption = l.itemoption" & vbcrlf
        'sqlStr = sqlStr & "     and l.code in ('05','06')" & vbcrlf
        sqlStr = sqlStr & " join db_order.dbo.tbl_order_master m with (nolock)" & vbcrlf
        sqlStr = sqlStr & " 	on l.orderserial = m.orderserial" & vbcrlf
        sqlStr = sqlStr & " 	and m.cancelyn='N'" & vbcrlf
        sqlStr = sqlStr & " join db_order.dbo.tbl_order_detail d with (nolock)" & vbcrlf
        sqlStr = sqlStr & "     on a.orderserial=d.orderserial" & vbcrlf
        sqlStr = sqlStr & "     and b.itemid = d.itemid" & vbcrlf
        sqlStr = sqlStr & "     and b.itemoption = d.itemoption" & vbcrlf
        sqlStr = sqlStr & " 	and d.cancelyn<>'Y'" & vbcrlf
        sqlStr = sqlStr & " where a.divcd='A008'" & vbcrlf
        sqlStr = sqlStr & " and a.writeuser='system'" & vbcrlf
        sqlStr = sqlStr & " and a.deleteyn='N'" & vbcrlf
        sqlStr = sqlStr & " and a.currstate='B001'" & vbcrlf
        sqlStr = sqlStr & " and a.needChkYN is null" & vbcrlf
        sqlStr = sqlStr & " and l.isautocanceldate is not null" & vbcrlf

        'response.write sqlStr & "<Br>"
		dbget.Execute sqlStr

        sqlStr = "select top "& CStr(param1) &" a.id as asid, a.divcd, a.orderserial" & vbcrlf
        sqlStr = sqlStr & " from db_cs.dbo.tbl_new_as_list a with (nolock)" & vbcrlf
        sqlStr = sqlStr & " join db_cs.dbo.tbl_new_as_detail b with (nolock)" & vbcrlf
        sqlStr = sqlStr & "     on a.id = b.masterid" & vbcrlf
        sqlStr = sqlStr & " join db_temp.dbo.tbl_mibeasong_list l with (nolock)" & vbcrlf
        sqlStr = sqlStr & "     on a.orderserial=l.orderserial" & vbcrlf
        sqlStr = sqlStr & "     and b.itemid = l.itemid" & vbcrlf
        sqlStr = sqlStr & "     and b.itemoption = l.itemoption" & vbcrlf
        'sqlStr = sqlStr & "     and l.code in ('05','06')" & vbcrlf
        sqlStr = sqlStr & " join db_order.dbo.tbl_order_master m with (nolock)" & vbcrlf
        sqlStr = sqlStr & " 	on l.orderserial = m.orderserial" & vbcrlf
        sqlStr = sqlStr & " 	and m.cancelyn='N'" & vbcrlf
        sqlStr = sqlStr & " join db_order.dbo.tbl_order_detail d with (nolock)" & vbcrlf
        sqlStr = sqlStr & "     on a.orderserial=d.orderserial" & vbcrlf
        sqlStr = sqlStr & "     and b.itemid = d.itemid" & vbcrlf
        sqlStr = sqlStr & "     and b.itemoption = d.itemoption" & vbcrlf
        sqlStr = sqlStr & " 	and d.cancelyn<>'Y'" & vbcrlf
        sqlStr = sqlStr & " where a.divcd='A008'" & vbcrlf
        sqlStr = sqlStr & " and a.writeuser='system'" & vbcrlf
        sqlStr = sqlStr & " and a.deleteyn='N'" & vbcrlf
        sqlStr = sqlStr & " and a.currstate='B001'" & vbcrlf
        sqlStr = sqlStr & " and a.needChkYN='N'" & vbcrlf
        sqlStr = sqlStr & " and l.isautocanceldate is not null" & vbcrlf
        sqlStr = sqlStr & " and a.regdate >= convert(varchar(10), DateAdd(day, -3, getdate()), 121)" & vbcrlf   ' ��ϵ��� 3�� �������� ������� �ʴ´�.
        sqlStr = sqlStr & " group by a.id, a.divcd, a.orderserial" & vbcrlf
        sqlStr = sqlStr & " order by a.orderserial asc" & vbcrlf

		'response.Write sqlStr & "<br>"
		rsget.Open sqlStr,dbget,1
        cnt = rsget.RecordCount
        ReDim idArr(cnt)
        ReDim divcdArr(cnt)
		ReDim orderserialArr(cnt)

        i = 0
        if Not rsget.Eof then
            do until rsget.eof
				idArr(i) = rsget("asid")
				divcdArr(i) = rsget("divcd")
				orderserialArr(i) = rsget("orderserial")
				i=i+1
				rsget.MoveNext
    		loop
        end if
        rsget.close

        if (cnt<1) then
            response.Write "S_NONE"
            dbget.Close() : response.end
        else
			response.Write "S_OK" & vbCrLf
			for i=LBound(idArr) to UBound(idArr)
				response.Write idArr(i) & "|" & divcdArr(i) & "|" & orderserialArr(i) & vbCrLf
			next
        end if

    Case ELSE
        response.Write "S_ERR|Not Valid - "&act
End Select
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
