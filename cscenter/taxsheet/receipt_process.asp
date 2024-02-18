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
    '# 1. 객체 생성 #
    '################

    ''Set INIpay = Server.CreateObject("INIreceipt41.INIreceiptTX41.1")
    Set INIpay = Server.CreateObject("INItx41.INItx41.1")

    '###############################################################################
    '# 2. 인스턴스 초기화 #
    '######################
    PInst = INIpay.Initialize("")

    '###############################################################################
    '# 3. 거래 유형 설정 #
    '#####################
    INIpay.SetActionType CLng(PInst), "CANCEL"

    '###############################################################################
    '# 4. 정보 설정 #
    '################
    INIpay.SetField CLng(PInst), "pgid", "IniTechPG_" 'PG ID (고정)
    INIpay.SetField CLng(PInst), "spgip", "203.238.3.10" '예비 PG IP (고정)

    if (application("Svr_Info")	= "Dev") then
    	INIpay.SetField CLng(PInst), "mid", "INIpayTest" '상점아이디
    else
    	INIpay.SetField CLng(PInst), "mid", "teenxteen4" '상점아이디
	end if

    INIpay.SetField CLng(PInst), "admin", "1111" '키패스워드(상점아이디에 따라 변경)
    INIpay.SetField CLng(PInst), "tid", orgtid '취소할 거래번호(TID)
    INIpay.SetField CLng(PInst), "msg", cancelCause '취소 사유
    INIpay.SetField CLng(PInst), "uip", Request.ServerVariables("REMOTE_ADDR") 'IP
    INIpay.SetField CLng(PInst), "debug", "false" '로그모드("true"로 설정하면 상세한 로그를 남김)
    INIpay.SetField CLng(PInst), "merchantreserved", "예비" '예비

    '###############################################################################
    '# 5. 취소 요청 #
    '################
    INIpay.StartAction(CLng(PInst))

    '###############################################################################
    '# 6. 취소 결과 #
    '################
    ResultCode = INIpay.GetResult(CLng(PInst), "resultcode") '결과코드 ("00"이면 취소성공)
    ResultMsg = INIpay.GetResult(CLng(PInst), "resultmsg") '결과내용
    CancelDate = INIpay.GetResult(CLng(PInst), "pgcanceldate") '이니시스 취소날짜
    CancelTime = INIpay.GetResult(CLng(PInst), "pgcanceltime") '이니시스 취소시각
    Rcash_cancel_noappl = INIpay.GetResult(CLng(PInst), "rcash_cancel_noappl") '현금영수증 취소 승인번호

    '###############################################################################
    '# 7. 인스턴스 해제 #
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
    if (pggubun="NP") then ''네이버 페이의 경우 (2016/08/12)
        Set NPay_Result = fnCallNaverPayCashAmt(orgpaygatetid)
        NpayCashAmt    = CLng(NPay_Result.body.totalCashAmount) + sumpaymentEtc
        NpaySuplyAmt   = CLng(NPay_Result.body.supplyCashAmount) + CLng(sumpaymentEtc*10/11)	'// 총 공급가
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
            iResultMsg    = "NPAY 금액 오류 재작성.["&cr_price&"::"&NpayCashAmt&"]"
            Exit Function
        end if
    else
        if ((orgaccountdiv="20") or (orgaccountdiv="7")) then
                
        else
            subtotalprice = sumpaymentEtc
        end if
        
        subtotalPrice = subtotalPrice+GetReceiptMinusOrderSUM(orderserial) ''반품금액 추가
        
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
            iResultMsg    = "금액 오류 재작성.["&cr_price&"::"&subtotalprice&"]"
            Exit Function
        end if
    end if

    if (useopt="0") and ((Len(reg_num)<>13) and (Len(reg_num)<>10) and (Len(reg_num)<>11) and (Len(reg_num)<>18)) then
        OneReceiptReq = False
        iResultMsg    = "주민번호/핸드폰 자리 오류"
        Exit Function
    end if

    if (useopt="1") and ((Len(reg_num)<>13) and (Len(reg_num)<>10) and (Len(reg_num)<>11) and (Len(reg_num)<>18)) then
        OneReceiptReq = False
        iResultMsg    = "사업자번호/ 주민번호 /핸드폰 자리 오류"
        Exit Function
    end if

    if (reqresultcode<>"R") then
        OneReceiptReq = False
        iResultMsg    = "기발행 확인"
        Exit Function
    end if
    
    '*******************************************************************************
    '* INIreceipt.asp
    '* 현금결제(실시간 은행계좌이체, 무통장입금)에 대한 현금결제 영수증 발행 요청한다.
    '*
    '* Date : 2004/12
    '* Project : INIpay V4.11 for Unix
    '*
    '* http://www.inicis.com
    '* http://support.inicis.com
    '* Copyright (C) 2002 Inicis, Co. All rights reserved.
    '*******************************************************************************

    '###############################################################################
    '# 1. 객체 생성 #
    '################
    Set INIpay = Server.CreateObject("INItx41.INItx41.1")

    '###############################################################################
    '# 2. 인스턴스 초기화 #
    '######################
    PInst = INIpay.Initialize("")

    '###############################################################################
    '# 3. 거래 유형 설정 #
    '#####################
    INIpay.SetActionType CLng(PInst), "receipt"

    '###############################################################################
    '# 4. 발급 정보 설정 #
    '###############################################################################
    INIpay.SetField CLng(PInst), "pgid","INIpayRECP"	'PG ID (고정)
    INIpay.SetField CLng(PInst), "paymethod","CASH"		'지불방법
    INIpay.SetField CLng(PInst), "spgip", "203.238.3.10" '예비 PG IP (고정)
    INIpay.SetField CLng(PInst), "currency", "WON" '화폐단위
    INIpay.SetField CLng(PInst), "admin", "1111"

    if (application("Svr_Info")	= "Dev") then
    	INIpay.SetField CLng(PInst), "mid", "INIpayTest" '상점아이디
    else
    	INIpay.SetField CLng(PInst), "mid", "teenxteen4" '상점아이디
	end if

    INIpay.SetField CLng(PInst), "uip", Request.ServerVariables("REMOTE_ADDR") '고객IP
    INIpay.SetField CLng(PInst), "goodname", goodname '상품명
    INIpay.SetField CLng(PInst), "cr_price", cr_price '총 현금 결제 금액
    INIpay.SetField CLng(PInst), "sup_price", sup_price '공급가액
    INIpay.SetField CLng(PInst), "tax", tax         '부가세
    INIpay.SetField CLng(PInst), "srvc_price", srvc_price '봉사료
    INIpay.SetField CLng(PInst), "buyername", buyername '성명
    INIpay.SetField CLng(PInst), "buyertel", buyertel '이동전화
    INIpay.SetField CLng(PInst), "buyeremail", buyeremail '이메일
    INIpay.SetField CLng(PInst), "reg_num", reg_num '현금결제자 주민등록번호
    INIpay.SetField CLng(PInst), "useopt", useopt '현금영수증 발행용도 ("0" - 소비자 소득공제용, "1" - 사업자 지출증빙용)
    INIpay.SetField CLng(PInst), "debug", "false" '로그모드("true"로 설정하면 상세한 로그를 남김)

    '###############################################################################
    '# 5. 지불 요청 #
    '################
    INIpay.StartAction(CLng(PInst))

    '###############################################################################
    '6. 발급 결과 #
    '###############################################################################
    '-------------------------------------------------------------------------------
    ' 가.모든 결제 수단에 공통되는 결제 결과 내용
    '-------------------------------------------------------------------------------
    Tid                 = INIpay.GetResult(CLng(PInst), "tid") '거래번호
    ResultCode          = INIpay.GetResult(CLng(PInst), "resultcode") '결과코드 ("00"이면 지불성공)
    ResultMsg           = INIpay.GetResult(CLng(PInst), "resultmsg") '결과내용
    AuthCode            = INIpay.GetResult(CLng(PInst), "authcode") '현금영수증 발생 승인번호
    PGAuthDate          = INIpay.GetResult(CLng(PInst), "pgauthdate") '이니시스 승인날짜
    PGAuthTime          = INIpay.GetResult(CLng(PInst), "pgauthtime") '이니시스 승인시각

    ResultpCRPice       = INIpay.GetResult(CLng(PInst), "ResultpCRPice") '결제 되는 금액
    ResultSupplyPrice   = INIpay.GetResult(CLng(PInst), "ResultSupplyPrice") '공급가액
    ResultTax           = INIpay.GetResult(CLng(PInst), "ResultTax") '부가세
    ResultServicePrice  = INIpay.GetResult(CLng(PInst), "ResultServicePrice") '봉사료
    ResultUseOpt        = INIpay.GetResult(CLng(PInst), "ResultUseOpt") '발행구분
    ResultCashNoAppl    = INIpay.GetResult(CLng(PInst), "ResultCashNoAppl") '승인번호

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

    ''결과 저장 - 관리자 발행시 성공인경우만 저장.
    IF ResultCode = "00" THEN
        sqlStr = "update [db_log].[dbo].tbl_cash_receipt" + VbCrlf
        sqlStr = sqlStr + " set tid='" + CStr(Tid) + "'" + VbCrlf
        sqlStr = sqlStr + " , resultcode='" + CStr(ResultCode) + "'" + VbCrlf
        sqlStr = sqlStr + " , resultmsg='" + LeftB(CStr(Replace(ResultMsg,"'","")),100) + "'" + VbCrlf
        sqlStr = sqlStr + " , authcode='" + CStr(AuthCode) + "'" + VbCrlf
        sqlStr = sqlStr + " , resultcashnoappl='" + CStr(ResultCashNoAppl) + "'" + VbCrlf
        sqlStr = sqlStr + " where idx=" + CStr(idx)

        dbget.Execute sqlStr
        
         ''2016/06/30 추가. 승인일
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
    '# 7. 결과 수신 확인 #
    '#####################
    '지불결과를 잘 수신하였음을 이니시스에 통보.
    '[주의] 이 과정이 누락되면 모든 거래가 자동취소됩니다.
    IF ResultCode = "00" THEN
    	AckResult = INIpay.Ack(CLng(PInst))
    	IF AckResult <> "SUCCESS" THEN '(실패)
    		'=================================================================
    		' 정상수신 통보 실패인 경우 이 승인은 이니시스에서 자동 취소되므로
    		' 지불결과를 다시 받아옵니다(성공 -> 실패).
    		'=================================================================
    		ResultCode = INIpay.GetResult(CLng(PInst), "resultcode")
    		ResultMsg = INIpay.GetResult(CLng(PInst), "resultmsg")
    	END IF
    END IF

    '###############################################################################
    '# 8. 인스턴스 해제 #
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
            ''기발행 성공 내역 체크
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
                    infoMsg = infoMsg & " <font color='red'>기발행 내역 존재 - 삭제:" & orderserial & "[" & idx & "]" & "</font><br>" & VbCrlf
                    sqlStr = " update [db_log].[dbo].tbl_cash_receipt"
                    sqlStr = sqlStr + " set cancelyn='D'"
                    sqlStr = sqlStr + " where idx=" & CStr(idx)
                    dbget.Execute sqlStr
               elseif (preIssuedTaxExists<>"none") then
                    infoMsg = infoMsg & " <font color='red'>세금계산서 발행 내역 존재 - 삭제:" & orderserial & "[" & idx & "]" & "</font><br>" & VbCrlf
                    sqlStr = " update [db_log].[dbo].tbl_cash_receipt"
                    sqlStr = sqlStr + " set cancelyn='D'"
                   sqlStr = sqlStr + " where idx=" & CStr(idx)
                    dbget.Execute sqlStr
               else
                    iResultCode = ""
                    iResultMsg  = ""
                    if (Not OneReceiptReq(idx, iResultCode, iResultMsg, iAuthCode)) then
                        infoMsg = infoMsg & " <font color='red'>발행 실패 :" & "[" & iResultCode & "]" & iResultMsg & "</font><br>" & VbCrlf
                    else
                        infoMsg = infoMsg & " 발행 성공 :" & "[" & iResultCode & "]" & iResultMsg & "<br>" & VbCrlf

                        IF (Atype="RA") THEN
                            sqlStr = " update [db_academy].[dbo].tbl_academy_order_master" & VbCrlf
                            sqlStr = sqlStr & " set authcode = (case when accountdiv in ('7', '20') then '" & iAuthCode & "' else authcode end) " & VbCrlf
                            if (reg_num="0100001234") then
                                sqlStr = sqlStr & " ,cashreceiptreq='J'" & VbCrlf   '' 자진발급 2016/06/22
                            else
                                sqlStr = sqlStr & " ,cashreceiptreq='S'" & VbCrlf   
                            end if
                            sqlStr = sqlStr & " where orderserial='" & orderserial & "'"

                            dbACADEMYget.Execute sqlStr
                        ELSE
                            sqlStr = " update [db_order].[dbo].tbl_order_master" & VbCrlf
                            sqlStr = sqlStr & " set authcode = (case when accountdiv in ('7', '20') then '" & iAuthCode & "' else authcode end) " & VbCrlf
                            if (reg_num="0100001234") then
                                sqlStr = sqlStr & " ,cashreceiptreq='J'" & VbCrlf   '' 자진발급 2016/06/22
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
            infoMsg = infoMsg & "발행 코드 존재 안함 " & "[" & idx & "]" & VbCrlf
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

        icancelCause = "주문 취소"
        if (Atype="C2") then icancelCause = "취소 요청"

        if (idx<>0) then
            iResultCode = ""
            iResultMsg  = ""
            if (Not OneReceiptCancel(orgtid,icancelCause, iResultCode, iResultMsg, iAuthCode)) then
                IF (IsAutoScript) then
                    infoMsg = infoMsg & iResultCode&"||"&orderserial&"||"&"[" & iResultCode & "]" & iResultMsg
                else
                    infoMsg = infoMsg & " <font color='red'>취소 실패 :" & "[" & iResultCode & "]" & iResultMsg & "</font><br>" & VbCrlf
                    infoMsg = infoMsg & " orgtid :"&orgtid& VbCrlf
                end if
            else
                IF (IsAutoScript) then
                    infoMsg = infoMsg & iResultCode&"||"&orderserial&"||"&"[" & iResultCode & "]" & iResultMsg
                else
                    infoMsg = infoMsg & " 취소 성공 :" & "[" & iResultCode & "]" & iResultMsg & "<br>" & VbCrlf
                end if

                sqlStr = " update [db_log].[dbo].tbl_cash_receipt" & VbCrlf
                sqlStr = sqlStr & " set canceltid='" & iAuthCode & "'" & VbCrlf
                sqlStr = sqlStr & " ,cancelyn='Y'" & VbCrlf
                sqlStr = sqlStr & " where idx=" & idx & ""

                dbget.Execute sqlStr

                ''고객센터에서 취소할 경우
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
                infoMsg = infoMsg & "FAIL||"&chkPrint(i)&"||"&"[발행 코드 존재 안함]"
            else
                infoMsg = infoMsg & "발행 코드 존재 안함 " & "[" & idx & "]" & VbCrlf
            end if
        end if

        response.flush
    next
elseif (Atype="RNC") then ''재발행 후 취소.
    '' 발행금액.  // pggubun
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
            infoMsg = infoMsg & "ERR||"&orderserial&"||"&"[주문번호,인덱스 체크 오류]" 
        else
            infoMsg = infoMsg & "주문번호,인덱스 체크 오류 " & "[" & orderserial & "]" & VbCrlf
        end if
        
        response.write infoMsg
        dbget.Close() : response.end
    end if
    
    ''기발행 체크
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
            infoMsg = infoMsg & "ERR||"&duppEvalIDX&"||"&"[타 발행 내역 존재]" 
        ELSE
            infoMsg = infoMsg & "타 발행 내역 존재 " & "[" & duppEvalIDX & "]" & VbCrlf
        END IF
        
        response.write infoMsg
        dbget.Close() : response.end
    end if
    
    if (NOT ((resultcode="00") and (cancelyn="N"))) then
        IF (IsAutoScript) then
            infoMsg = infoMsg & "ERR||"&idx&"||"&"[기발행 내역 아님]" 
        ELSE
            infoMsg = infoMsg & "기발행 내역 아님 " & "[" & idx & "]" & VbCrlf
        END IF
        
        response.write infoMsg
        dbget.Close() : response.end
    end if
    
    '' 발행 대상 금액조회
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
            infoMsg = infoMsg & "ERR||"&ReEvalCashAmt&"||"&"[발행 금액 확인 필요"&request("mayPrc")&"]" 
        ELSE
            infoMsg = infoMsg & "발행 금액 확인 필요 " & "[" & ReEvalCashAmt & "<>"&request("mayPrc")&"]" & VbCrlf
        END IF
        
        response.write infoMsg
        dbget.Close() : response.end
    end if
    
    ''infoMsg = infoMsg & ReEvalCashAmt &"|"&ReEvalCashSupp & VbCrlf
    
    '' 발행 한줄 꽃아 넣음
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
    
    ''발행 먼저
    iResultCode = ""
    iResultMsg  = ""
    iAuthCode   = ""
    
    if (Not OneReceiptReq(reEvalIDX, iResultCode, iResultMsg, iAuthCode)) then
        IF (IsAutoScript) then
            infoMsg = infoMsg & iResultCode&"||"&reEvalIDX&"("&orderserial&")"&"||"&"["&iResultMsg&"]" 
        ELSE
            infoMsg = infoMsg & " <font color='red'>발행 실패 :" & "[" & iResultCode & "]" & iResultMsg & "</font><br>" & VbCrlf
        END IF
    else
        IF (IsAutoScript) then
            infoMsg = infoMsg & iResultCode&"||"&reEvalIDX&"("&orderserial&")"&"||"&"["&iResultMsg&"]" 
        ELSE
            infoMsg = infoMsg & " 발행 성공 :" & "[" & iResultCode & "]" & iResultMsg & "<br>" & VbCrlf
        END IF
        
        IF (Atype="RNCA") THEN
            sqlStr = " update [db_academy].[dbo].tbl_academy_order_master" & VbCrlf
            sqlStr = sqlStr & " set authcode = (case when accountdiv in ('7', '20') then '" & iAuthCode & "' else authcode end) " & VbCrlf
            if (reg_num="0100001234") then
                sqlStr = sqlStr & " ,cashreceiptreq='J'" & VbCrlf   '' 자진발급 2016/06/22
            else
                sqlStr = sqlStr & " ,cashreceiptreq='S'" & VbCrlf   
            end if
            sqlStr = sqlStr & " where orderserial='" & orderserial & "'"

            dbACADEMYget.Execute sqlStr
        ELSE
            sqlStr = " update [db_order].[dbo].tbl_order_master" & VbCrlf
            sqlStr = sqlStr & " set authcode = (case when accountdiv in ('7', '20') then '" & iAuthCode & "' else authcode end) " & VbCrlf
            if (reg_num="0100001234") then
                sqlStr = sqlStr & " ,cashreceiptreq='J'" & VbCrlf   '' 자진발급 2016/06/22
            else
                sqlStr = sqlStr & " ,cashreceiptreq='S'" & VbCrlf   
            end if
            sqlStr = sqlStr & " where orderserial='" & orderserial & "'"

            dbget.Execute sqlStr
        END IF
        
        ''취소
        iResultCode = ""
        iResultMsg  = ""
        icancelCause = "재발행"
        iAuthCode   = ""
        
        if (Not OneReceiptCancel(orgtid,icancelCause, iResultCode, iResultMsg, iAuthCode)) then
            IF (IsAutoScript) then
                infoMsg = infoMsg & iResultCode&"||취소 실패||"&"["&iResultMsg&"]" 
            ELSE
                infoMsg = infoMsg & " <font color='red'>취소 실패 :" & "[" & iResultCode & "]" & iResultMsg & "</font><br>" & VbCrlf
                infoMsg = infoMsg & " orgtid :"&orgtid& VbCrlf
            END IF
        else
            IF (IsAutoScript) then
                infoMsg = infoMsg & iResultCode&"||"&orgtid&"||"&"["&iResultMsg&"]" 
            ELSE
                infoMsg = infoMsg & " 취소 성공 :" & "[" & iResultCode & "]" & iResultMsg & "<br>" & VbCrlf
            END IF
            
            sqlStr = " update [db_log].[dbo].tbl_cash_receipt" & VbCrlf
            sqlStr = sqlStr & " set canceltid='" & iAuthCode & "'" & VbCrlf
            sqlStr = sqlStr & " ,cancelyn='Y'" & VbCrlf
            sqlStr = sqlStr & " where idx=" & idx & ""

            dbget.Execute sqlStr

            ''고객센터에서 취소할 경우
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
    icancelCause ="오발행"
    if (Not OneReceiptCancel(orgtid,icancelCause, iResultCode, iResultMsg, iAuthCode)) then
        rw " <font color='red'>취소 실패 :" & "[" & iResultCode & "]" & iResultMsg & "</font><br>" & VbCrlf
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
    sqlStr = sqlStr + " and m.subtotalPrice=C.cr_price"  '' 부분취소로 금액변동 발생 가능.
        
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
                ''기발행 성공 내역 체크
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
                        infoMsg = infoMsg & " <font color='red'>기발행 내역 존재 - 삭제:" & orderserial & "[" & idx & "]" & "</font><br>" & VbCrlf
                        sqlStr = " update [db_log].[dbo].tbl_cash_receipt"
                        sqlStr = sqlStr + " set cancelyn='D'"
                        sqlStr = sqlStr + " where idx=" & CStr(idx)
                        dbget.Execute sqlStr
                   elseif (preIssuedTaxExists<>"none") then
                        infoMsg = infoMsg & " <font color='red'>세금계산서 발행 내역 존재 - 삭제:" & orderserial & "[" & idx & "]" & "</font><br>" & VbCrlf
                        sqlStr = " update [db_log].[dbo].tbl_cash_receipt"
                        sqlStr = sqlStr + " set cancelyn='D'"
                        sqlStr = sqlStr + " where idx=" & CStr(idx)
                        dbget.Execute sqlStr
                   else
                        iResultCode = ""
                        iResultMsg  = ""
                        if (Not OneReceiptReq(idx, iResultCode, iResultMsg, iAuthCode)) then
                            infoMsg = infoMsg & " <font color='red'>발행 실패 :" & "[" & iResultCode & "]" & iResultMsg & "</font><br>" & VbCrlf
                        else
                            infoMsg = infoMsg & " 발행 성공 :" & "[" & iResultCode & "]" & iResultMsg & "<br>" & VbCrlf

                            sqlStr = " update [db_order].[dbo].tbl_order_master" & VbCrlf
                            sqlStr = sqlStr & " set authcode='" & iAuthCode & "'" & VbCrlf
                            sqlStr = sqlStr & " where orderserial='" & orderserial & "'"

                            dbget.Execute sqlStr
                        end if
                   end if


                end if
            else
                infoMsg = infoMsg & "발행 코드 존재 안함 " & "[" & idx & "]" & VbCrlf
            end if
        next
    else
        infoMsg = infoMsg & "발행할 내역 없음." & VbCrlf
    end if
else
    response.write "지정되지 않았습니다. - " & Atype & "<br>"
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