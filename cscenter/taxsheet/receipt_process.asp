<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbACADEMYopen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_taxsheetcls.asp"-->
<!-- #include virtual="/cscenter/action/incNaverpayCommon.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_cashreceiptcls.asp"-->
<%
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

    if (application("Svr_Info")	= "Dev") then
    	INIpay.SetField CLng(PInst), "mid", "INIpayTest" '�������̵�
    else
    	INIpay.SetField CLng(PInst), "mid", "teenxteen4" '�������̵�
	end if

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
    dim subtotalprice, dataExists
    dim reqresultcode
    dim pggubun, sumpaymentEtc, orgpaygatetid, orgaccountdiv
    
    dataExists = false
    sqlStr = " select c.*, m.subtotalprice, isNULL(m.pggubun,'') as pggubun, isNULL(m.sumpaymentEtc,0) as sumpaymentEtc"
    sqlStr = sqlStr + " , isNULL(m.paygatetid,'') as orgpaygatetid "
    sqlStr = sqlStr + " , isNULL(m.accountdiv,'') as orgaccountdiv"
    sqlStr = sqlStr + "     from [db_log].[dbo].tbl_cash_receipt c"
    sqlStr = sqlStr + "     Join db_order.dbo.tbl_order_master m"
    sqlStr = sqlStr + "     on c.orderserial=m.orderserial"
    sqlStr = sqlStr + " where c.idx=" & idx
    
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
    if Not rsget.Eof then
        dataExists = true
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
        subtotalprice = rsget("subtotalprice")
        reqresultcode  = rsget("resultcode")
        pggubun = rsget("pggubun")
        sumpaymentEtc = rsget("sumpaymentEtc")
        orgpaygatetid = rsget("orgpaygatetid")
        orgaccountdiv = TRIM(rsget("orgaccountdiv"))
    end if
    rsget.close
    
    if (not dataExists) then
        sqlStr = " select c.*  from [db_log].[dbo].tbl_cash_receipt c"
        sqlStr = sqlStr + " where c.idx=" & idx
        rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
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
            subtotalprice = cr_price
            reqresultcode  = rsget("resultcode")
            
            sumpaymentEtc = 0
        end if
        rsget.close
    end if
    
    Dim NPay_Result, NpayCashAmt, NpaySuplyAmt
    if (pggubun="NP") then ''���̹� ������ ��� (2016/08/12)
        Set NPay_Result = fnCallNaverPayCashAmt(orgpaygatetid)
        NpayCashAmt    = CLng(NPay_Result.body.totalCashAmount) + sumpaymentEtc
        NpaySuplyAmt   = CLng(NPay_Result.body.supplyCashAmount) + CLng(sumpaymentEtc*10/11)	'// �� ���ް�
        Set NPay_Result = Nothing
        
        if (NpayCashAmt<>cr_price) or (sup_price<>NpaySuplyAmt) then
            ' sqlStr = " update C"
            ' sqlStr = sqlStr & " SET cr_price="&NpayCashAmt
            ' sqlStr = sqlStr & " from [db_log].[dbo].tbl_cash_receipt c"
            ' sqlStr = sqlStr & " where c.idx=" & idx
            ' dbget.Execute sqlStr
            
            sqlStr = " update C "
            sqlStr = sqlStr & " SET cr_price="&NpayCashAmt&vbCRLF
            sqlStr = sqlStr & " ,sup_price="&NpaySuplyAmt&vbCRLF   '''cr_price-convert(int,cr_price*1/11)"
            sqlStr = sqlStr & " ,tax=("&NpayCashAmt&"-"&NpaySuplyAmt&")"&vbCRLF ''convert(int,cr_price*1/11)"
            sqlStr = sqlStr & " from [db_log].[dbo].tbl_cash_receipt c"&vbCRLF
            sqlStr = sqlStr & " where c.idx=" & idx &vbCRLF
            dbget.Execute sqlStr
            
            OneReceiptReq = False
            iResultMsg    = "NPAY �ݾ� ���� ���ۼ�.["&cr_price&"::"&NpayCashAmt&"]"
            Exit Function
        end if
    else
        if ((orgaccountdiv="20") or (orgaccountdiv="7")) then
                
        else
            subtotalprice = sumpaymentEtc
        end if
        
        subtotalPrice = subtotalPrice+GetReceiptMinusOrderSUM(orderserial) ''��ǰ�ݾ� �߰�
        
        if (subtotalprice<>cr_price) then
            sqlStr = " update C"
            sqlStr = sqlStr & " SET cr_price="&subtotalprice
            sqlStr = sqlStr & " from [db_log].[dbo].tbl_cash_receipt c"
            sqlStr = sqlStr & " where c.idx=" & idx
            dbget.Execute sqlStr
            
            sqlStr = " update C"
            sqlStr = sqlStr & " SET tax=convert(int,cr_price*1/11)"
            sqlStr = sqlStr & " ,sup_price=cr_price-convert(int,cr_price*1/11)"
            sqlStr = sqlStr & " from [db_log].[dbo].tbl_cash_receipt c"
            sqlStr = sqlStr & " where c.idx=" & idx
            dbget.Execute sqlStr
            
            OneReceiptReq = False
            iResultMsg    = "�ݾ� ���� ���ۼ�.["&cr_price&"::"&subtotalprice&"]"
            Exit Function
        end if
    end if

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

    if (application("Svr_Info")	= "Dev") then
    	INIpay.SetField CLng(PInst), "mid", "INIpayTest" '�������̵�
    else
    	INIpay.SetField CLng(PInst), "mid", "teenxteen4" '�������̵�
	end if

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


dim chkPrint, i, Atype
dim pggubun, sumpaymentEtc, subtotalPrice, accountdiv, orgpaygatetid

chkPrint = request("chkPrint")
Atype    = RequestCheckVar(request("Atype"),9)
pggubun  = RequestCheckVar(request("pggubun"),10)

if Right(chkPrint,1)="," then chkPrint=Left(chkPrint,Len(chkPrint)-1)

response.write chkPrint & "<br>"

chkPrint = split(chkPrint,",")

dim sqlStr
dim idx, orderserial, resultcode, cancelyn, reg_num
dim preIssuedExists, infoMsg, iResultCode, iResultMsg, iAuthCode
dim preIssuedTaxExists
dim orgtid, canceltid
dim icancelCause



if (Atype="R") or (Atype="RA") then
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
                    infoMsg = infoMsg & " <font color='red'>����� ���� ���� - ����:" & orderserial & "[" & idx & "]" & "</font><br>" & VbCrlf
                    sqlStr = " update [db_log].[dbo].tbl_cash_receipt"
                    sqlStr = sqlStr + " set cancelyn='D'"
                    sqlStr = sqlStr + " where idx=" & CStr(idx)
                    dbget.Execute sqlStr
               elseif (preIssuedTaxExists<>"none") then
                    infoMsg = infoMsg & " <font color='red'>���ݰ�꼭 ���� ���� ���� - ����:" & orderserial & "[" & idx & "]" & "</font><br>" & VbCrlf
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
                        infoMsg = infoMsg & " ���� ���� :" & "[" & iResultCode & "]" & iResultMsg & "<br>" & VbCrlf

                        IF (Atype="RA") THEN
                            sqlStr = " update [db_academy].[dbo].tbl_academy_order_master" & VbCrlf
                            sqlStr = sqlStr & " set authcode = (case when accountdiv in ('7', '20') then '" & iAuthCode & "' else authcode end) " & VbCrlf
                            if (reg_num="0100001234") then
                                sqlStr = sqlStr & " ,cashreceiptreq='J'" & VbCrlf   '' �����߱� 2016/06/22
                            else
                                sqlStr = sqlStr & " ,cashreceiptreq='S'" & VbCrlf   
                            end if
                            sqlStr = sqlStr & " where orderserial='" & orderserial & "'"

                            dbACADEMYget.Execute sqlStr
                        ELSE
                            sqlStr = " update [db_order].[dbo].tbl_order_master" & VbCrlf
                            sqlStr = sqlStr & " set authcode = (case when accountdiv in ('7', '20') then '" & iAuthCode & "' else authcode end) " & VbCrlf
                            if (reg_num="0100001234") then
                                sqlStr = sqlStr & " ,cashreceiptreq='J'" & VbCrlf   '' �����߱� 2016/06/22
                            else
                                sqlStr = sqlStr & " ,cashreceiptreq='S'" & VbCrlf   
                            end if
                            sqlStr = sqlStr & " where orderserial='" & orderserial & "'"

                            dbget.Execute sqlStr
                        END IF
                    end if
               end if


            end if
        else
            infoMsg = infoMsg & "���� �ڵ� ���� ���� " & "[" & idx & "]" & VbCrlf
        end if

        response.flush
    next
elseif (Atype="C1") or (Atype="C2") or (Atype="CA") then
    for i=0 to UBound(chkPrint)
        idx = 0
        sqlStr = " select idx, orderserial, resultcode, cancelyn, tid from [db_log].[dbo].tbl_cash_receipt"
        sqlStr = sqlStr + " where idx=" & chkPrint(i)

        rsget.Open sqlStr,dbget,1
        if Not rsget.Eof then
            idx         = rsget("idx")
            orderserial = rsget("orderserial")
            resultcode  = rsget("resultcode")
            cancelyn    = rsget("cancelyn")
            orgtid    = rsget("tid")
        end if
        rsget.close

        icancelCause = "�ֹ� ���"
        if (Atype="C2") then icancelCause = "��� ��û"

        if (idx<>0) then
            iResultCode = ""
            iResultMsg  = ""
            if (Not OneReceiptCancel(orgtid,icancelCause, iResultCode, iResultMsg, iAuthCode)) then
                IF (IsAutoScript) then
                    infoMsg = infoMsg & iResultCode&"||"&orderserial&"||"&"[" & iResultCode & "]" & iResultMsg
                else
                    infoMsg = infoMsg & " <font color='red'>��� ���� :" & "[" & iResultCode & "]" & iResultMsg & "</font><br>" & VbCrlf
                    infoMsg = infoMsg & " orgtid :"&orgtid& VbCrlf
                end if
            else
                IF (IsAutoScript) then
                    infoMsg = infoMsg & iResultCode&"||"&orderserial&"||"&"[" & iResultCode & "]" & iResultMsg
                else
                    infoMsg = infoMsg & " ��� ���� :" & "[" & iResultCode & "]" & iResultMsg & "<br>" & VbCrlf
                end if

                sqlStr = " update [db_log].[dbo].tbl_cash_receipt" & VbCrlf
                sqlStr = sqlStr & " set canceltid='" & iAuthCode & "'" & VbCrlf
                sqlStr = sqlStr & " ,cancelyn='Y'" & VbCrlf
                sqlStr = sqlStr & " where idx=" & idx & ""

                dbget.Execute sqlStr

                ''�����Ϳ��� ����� ���
                if (Atype="C2") then
                    sqlStr = " update db_order.dbo.tbl_order_master" & VbCrlf
                    sqlStr = sqlStr & " set cashreceiptreq=NULL" & VbCrlf
                    sqlStr = sqlStr & " ,authcode = (case when accountdiv in ('7', '20') then NULL else authcode end) " & VbCrlf
                    sqlStr = sqlStr & " where orderserial='"&orderserial&"'"

                    dbget.Execute sqlStr
                elseif (Atype="CA") then
                    sqlStr = " update db_academy.dbo.tbl_academy_order_master" & VbCrlf
                    sqlStr = sqlStr & " set cashreceiptreq=NULL" & VbCrlf
                    sqlStr = sqlStr & " ,authcode = (case when accountdiv in ('7', '20') then NULL else authcode end) " & VbCrlf
                    sqlStr = sqlStr & " where orderserial='"&orderserial&"'"

                    dbACADEMYget.Execute sqlStr
                end if
            end if
        else
            IF (IsAutoScript) then
                infoMsg = infoMsg & "FAIL||"&chkPrint(i)&"||"&"[���� �ڵ� ���� ����]"
            else
                infoMsg = infoMsg & "���� �ڵ� ���� ���� " & "[" & idx & "]" & VbCrlf
            end if
        end if

        response.flush
    next
elseif (Atype="RNC") then ''����� �� ���.
    '' ����ݾ�.  // pggubun
    dim reEvalIDX
    
    idx = 0
    sqlStr = " select C.idx, C.orderserial, C.resultcode, C.cancelyn, C.tid"
    sqlStr = sqlStr + " , isNULL(m.sumpaymentEtc,0) as sumpaymentEtc, isNULL(m.subtotalPrice,0) as subtotalPrice"
    sqlStr = sqlStr + " , isNULL(m.pggubun,'') as pggubun, isNULL(m.accountdiv,'') as accountdiv, isNULL(m.paygatetid,'') as paygatetid "
    sqlStr = sqlStr + " from [db_log].[dbo].tbl_cash_receipt C"
    sqlStr = sqlStr + " join db_order.dbo.tbl_order_master m"
    sqlStr = sqlStr + " on c.orderserial=m.orderserial"
    sqlStr = sqlStr + " where C.idx=" & chkPrint(0)

    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
    if Not rsget.Eof then
        idx         = rsget("idx")
        orderserial = rsget("orderserial")
        resultcode  = rsget("resultcode")
        cancelyn    = rsget("cancelyn")
        orgtid      = rsget("tid")
        
        pggubun       = rsget("pggubun")
        sumpaymentEtc = rsget("sumpaymentEtc")
        subtotalPrice = rsget("subtotalPrice")
        accountdiv    = TRIM(rsget("accountdiv"))
        orgpaygatetid = rsget("paygatetid")
    end if
    rsget.close
    
    if (orderserial="") or (idx="") then
        IF (IsAutoScript) then
            infoMsg = infoMsg & "ERR||"&orderserial&"||"&"[�ֹ���ȣ,�ε��� üũ ����]" 
        else
            infoMsg = infoMsg & "�ֹ���ȣ,�ε��� üũ ���� " & "[" & orderserial & "]" & VbCrlf
        end if
        
        response.write infoMsg
        dbget.Close() : response.end
    end if
    
    ''����� üũ
    dim duppEvalIDX : duppEvalIDX=0
    sqlStr = " select top 1 idx from [db_log].[dbo].tbl_cash_receipt C" & VbCrlf
    sqlStr = sqlStr + " where C.orderserial='"&orderserial&"'"& VbCrlf
    sqlStr = sqlStr + " and C.idx<>"&idx& VbCrlf
    sqlStr = sqlStr + " and C.resultcode='00'"& VbCrlf
    sqlStr = sqlStr + " and C.cancelyn='N'"& VbCrlf
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
    if Not rsget.Eof then
        duppEvalIDX = rsget("idx")
    end if
    rsget.close    
    
    if (duppEvalIDX<>0) then
        IF (IsAutoScript) then
            infoMsg = infoMsg & "ERR||"&duppEvalIDX&"||"&"[Ÿ ���� ���� ����]" 
        ELSE
            infoMsg = infoMsg & "Ÿ ���� ���� ���� " & "[" & duppEvalIDX & "]" & VbCrlf
        END IF
        
        response.write infoMsg
        dbget.Close() : response.end
    end if
    
    if (NOT ((resultcode="00") and (cancelyn="N"))) then
        IF (IsAutoScript) then
            infoMsg = infoMsg & "ERR||"&idx&"||"&"[����� ���� �ƴ�]" 
        ELSE
            infoMsg = infoMsg & "����� ���� �ƴ� " & "[" & idx & "]" & VbCrlf
        END IF
        
        response.write infoMsg
        dbget.Close() : response.end
    end if
    
    '' ���� ��� �ݾ���ȸ
    dim NPay_Result, ReEvalCashAmt, ReEvalCashSupp
    if (pggubun="NP") then
        Set NPay_Result = fnCallNaverPayCashAmt(orgpaygatetid)
        ReEvalCashAmt    = CLng(NPay_Result.body.totalCashAmount) + sumpaymentEtc
        ReEvalCashSupp   = CLng(NPay_Result.body.supplyCashAmount) + CLng(sumpaymentEtc*10/11)
        Set NPay_Result = Nothing
    else
        if ((accountdiv="20") or (accountdiv="7")) then
            
        else
            subtotalPrice = sumpaymentEtc
        end if
        ReEvalCashAmt = subtotalPrice+GetReceiptMinusOrderSUM(orderserial)
        ReEvalCashSupp = CLng(ReEvalCashAmt*10/11)
    end if
    
    if (CStr(ReEvalCashAmt)<>request("mayPrc")) then
        IF (IsAutoScript) then
            infoMsg = infoMsg & "ERR||"&ReEvalCashAmt&"||"&"[���� �ݾ� Ȯ�� �ʿ�"&request("mayPrc")&"]" 
        ELSE
            infoMsg = infoMsg & "���� �ݾ� Ȯ�� �ʿ� " & "[" & ReEvalCashAmt & "<>"&request("mayPrc")&"]" & VbCrlf
        END IF
        
        response.write infoMsg
        dbget.Close() : response.end
    end if
    
    ''infoMsg = infoMsg & ReEvalCashAmt &"|"&ReEvalCashSupp & VbCrlf
    
    '' ���� ���� �ɾ� ����
    sqlStr = " select * from [db_log].[dbo].tbl_cash_receipt where 1=0"
    rsget.Open sqlStr,dbget,1,3
    rsget.AddNew
    rsget("orderserial") = orderserial
    ''rsget("userid") = userid
    ''rsget("sitename") = sitename
    ''rsget("goodname") = goodname
    rsget("cr_price") = ReEvalCashAmt
    rsget("sup_price") = ReEvalCashSupp
    rsget("tax") = (ReEvalCashAmt-ReEvalCashSupp)
    rsget("srvc_price") = 0
    'rsget("buyername") = buyername
    'rsget("buyeremail") = buyeremail
    'rsget("buyertel") = buyertel
    'rsget("reg_num") = reg_num
    'rsget("useopt") = useopt
    'rsget("paymethod") = paymethod
    rsget("cancelyn") = "N"
    
    rsget.update
    reEvalIDX = rsget("idx")
    rsget.close
    
    sqlStr = " update N" &VBCRLF
    sqlStr = sqlStr&" set userid=P.userid"&VBCRLF
    sqlStr = sqlStr&" , sitename=P.sitename"&VBCRLF
    sqlStr = sqlStr&" , goodname=P.goodname"&VBCRLF
    sqlStr = sqlStr&" , buyername=P.buyername"&VBCRLF
    sqlStr = sqlStr&" , buyeremail=P.buyeremail"&VBCRLF
    sqlStr = sqlStr&" , buyertel=P.buyertel"&VBCRLF
    sqlStr = sqlStr&" , reg_num=P.reg_num"&VBCRLF
    sqlStr = sqlStr&" , useopt=P.useopt"&VBCRLF
    sqlStr = sqlStr&" , paymethod=P.paymethod"&VBCRLF
    sqlStr = sqlStr&" from [db_log].[dbo].tbl_cash_receipt N"&VBCRLF
    sqlStr = sqlStr&"     JOin [db_log].[dbo].tbl_cash_receipt P"&VBCRLF
    sqlStr = sqlStr&"     on 1=1"&VBCRLF
    sqlStr = sqlStr&"     and P.idx="&idx&VBCRLF
    sqlStr = sqlStr&" where N.idx="&reEvalIDX&VBCRLF
    dbget.Execute sqlStr
    
    ''���� ����
    iResultCode = ""
    iResultMsg  = ""
    iAuthCode   = ""
    
    if (Not OneReceiptReq(reEvalIDX, iResultCode, iResultMsg, iAuthCode)) then
        IF (IsAutoScript) then
            infoMsg = infoMsg & iResultCode&"||"&reEvalIDX&"("&orderserial&")"&"||"&"["&iResultMsg&"]" 
        ELSE
            infoMsg = infoMsg & " <font color='red'>���� ���� :" & "[" & iResultCode & "]" & iResultMsg & "</font><br>" & VbCrlf
        END IF
    else
        IF (IsAutoScript) then
            infoMsg = infoMsg & iResultCode&"||"&reEvalIDX&"("&orderserial&")"&"||"&"["&iResultMsg&"]" 
        ELSE
            infoMsg = infoMsg & " ���� ���� :" & "[" & iResultCode & "]" & iResultMsg & "<br>" & VbCrlf
        END IF
        
        IF (Atype="RNCA") THEN
            sqlStr = " update [db_academy].[dbo].tbl_academy_order_master" & VbCrlf
            sqlStr = sqlStr & " set authcode = (case when accountdiv in ('7', '20') then '" & iAuthCode & "' else authcode end) " & VbCrlf
            if (reg_num="0100001234") then
                sqlStr = sqlStr & " ,cashreceiptreq='J'" & VbCrlf   '' �����߱� 2016/06/22
            else
                sqlStr = sqlStr & " ,cashreceiptreq='S'" & VbCrlf   
            end if
            sqlStr = sqlStr & " where orderserial='" & orderserial & "'"

            dbACADEMYget.Execute sqlStr
        ELSE
            sqlStr = " update [db_order].[dbo].tbl_order_master" & VbCrlf
            sqlStr = sqlStr & " set authcode = (case when accountdiv in ('7', '20') then '" & iAuthCode & "' else authcode end) " & VbCrlf
            if (reg_num="0100001234") then
                sqlStr = sqlStr & " ,cashreceiptreq='J'" & VbCrlf   '' �����߱� 2016/06/22
            else
                sqlStr = sqlStr & " ,cashreceiptreq='S'" & VbCrlf   
            end if
            sqlStr = sqlStr & " where orderserial='" & orderserial & "'"

            dbget.Execute sqlStr
        END IF
        
        ''���
        iResultCode = ""
        iResultMsg  = ""
        icancelCause = "�����"
        iAuthCode   = ""
        
        if (Not OneReceiptCancel(orgtid,icancelCause, iResultCode, iResultMsg, iAuthCode)) then
            IF (IsAutoScript) then
                infoMsg = infoMsg & iResultCode&"||��� ����||"&"["&iResultMsg&"]" 
            ELSE
                infoMsg = infoMsg & " <font color='red'>��� ���� :" & "[" & iResultCode & "]" & iResultMsg & "</font><br>" & VbCrlf
                infoMsg = infoMsg & " orgtid :"&orgtid& VbCrlf
            END IF
        else
            IF (IsAutoScript) then
                infoMsg = infoMsg & iResultCode&"||"&orgtid&"||"&"["&iResultMsg&"]" 
            ELSE
                infoMsg = infoMsg & " ��� ���� :" & "[" & iResultCode & "]" & iResultMsg & "<br>" & VbCrlf
            END IF
            
            sqlStr = " update [db_log].[dbo].tbl_cash_receipt" & VbCrlf
            sqlStr = sqlStr & " set canceltid='" & iAuthCode & "'" & VbCrlf
            sqlStr = sqlStr & " ,cancelyn='Y'" & VbCrlf
            sqlStr = sqlStr & " where idx=" & idx & ""

            dbget.Execute sqlStr

            ''�����Ϳ��� ����� ���
            'if (Atype="RNC") then
            '    sqlStr = " update db_order.dbo.tbl_order_master" & VbCrlf
            '    sqlStr = sqlStr & " set cashreceiptreq=NULL" & VbCrlf
            '    sqlStr = sqlStr & " ,authcode = (case when accountdiv in ('7', '20') then NULL else authcode end) " & VbCrlf
            '    sqlStr = sqlStr & " where orderserial='"&orderserial&"'"
            '    dbget.Execute sqlStr
            'elseif (Atype="RNCA") then
            '    sqlStr = " update db_academy.dbo.tbl_academy_order_master" & VbCrlf
            '    sqlStr = sqlStr & " set cashreceiptreq=NULL" & VbCrlf
            '    sqlStr = sqlStr & " ,authcode = (case when accountdiv in ('7', '20') then NULL else authcode end) " & VbCrlf
            '    sqlStr = sqlStr & " where orderserial='"&orderserial&"'"
            '    dbACADEMYget.Execute sqlStr
            'end if
        end if
    end if
    
    
    
    
elseif (Atype="CH") then
    orgtid = request("tid")
    icancelCause ="������"
    if (Not OneReceiptCancel(orgtid,icancelCause, iResultCode, iResultMsg, iAuthCode)) then
        rw " <font color='red'>��� ���� :" & "[" & iResultCode & "]" & iResultMsg & "</font><br>" & VbCrlf
    else
        rw iResultMsg
    end if
elseif (Atype="AUTO1") then

    chkPrint = ""
    infoMsg = ""

    sqlStr = " select top 5 c.idx, c.orderserial, c.resultcode, c.cancelyn "
    sqlStr = sqlStr + " from db_order.dbo.tbl_order_master m"
    sqlStr = sqlStr + " 	Join [db_log].[dbo].tbl_cash_receipt c"
    sqlStr = sqlStr + " 	on c.orderserial=m.orderserial"
    sqlStr = sqlStr + " 	and c.resultcode='R'"
    sqlStr = sqlStr + " 	and c.cancelyn='N'"
    sqlStr = sqlStr + " where  m.ipkumdiv>6"
    sqlStr = sqlStr + " and m.cashreceiptreq='R'"
    sqlStr = sqlStr + " and m.authcode is NULL"
    sqlStr = sqlStr + " and m.accountdiv='7'"
    sqlStr = sqlStr + " and m.cancelyn='N'"
    sqlStr = sqlStr + " and m.subtotalPrice>0"
    sqlStr = sqlStr + " and m.subtotalPrice=C.cr_price"  '' �κ���ҷ� �ݾ׺��� �߻� ����.
        
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

            infoMsg = infoMsg & "[" & idx & "," & orderserial & "]"
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
                        infoMsg = infoMsg & " <font color='red'>����� ���� ���� - ����:" & orderserial & "[" & idx & "]" & "</font><br>" & VbCrlf
                        sqlStr = " update [db_log].[dbo].tbl_cash_receipt"
                        sqlStr = sqlStr + " set cancelyn='D'"
                        sqlStr = sqlStr + " where idx=" & CStr(idx)
                        dbget.Execute sqlStr
                   elseif (preIssuedTaxExists<>"none") then
                        infoMsg = infoMsg & " <font color='red'>���ݰ�꼭 ���� ���� ���� - ����:" & orderserial & "[" & idx & "]" & "</font><br>" & VbCrlf
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
                            infoMsg = infoMsg & " ���� ���� :" & "[" & iResultCode & "]" & iResultMsg & "<br>" & VbCrlf

                            sqlStr = " update [db_order].[dbo].tbl_order_master" & VbCrlf
                            sqlStr = sqlStr & " set authcode='" & iAuthCode & "'" & VbCrlf
                            sqlStr = sqlStr & " where orderserial='" & orderserial & "'"

                            dbget.Execute sqlStr
                        end if
                   end if


                end if
            else
                infoMsg = infoMsg & "���� �ڵ� ���� ���� " & "[" & idx & "]" & VbCrlf
            end if
        next
    else
        infoMsg = infoMsg & "������ ���� ����." & VbCrlf
    end if
else
    response.write "�������� �ʾҽ��ϴ�. - " & Atype & "<br>"
end if
response.write infoMsg

%>

<% IF (NOT IsAutoScript) then %>
<br>
<a href="javascript:history.back();">&lt;&lt;Back</a>

<% if (Atype="C2") then %>
&nbsp;
<a href="javascript:window.close();">&lt;&lt;Close</a>
<% end if %>

<% end if %>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbACADEMYclose.asp" -->