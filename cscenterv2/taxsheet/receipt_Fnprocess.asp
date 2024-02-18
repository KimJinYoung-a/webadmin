<% option Explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/cscenterv2/lib/incSessionAdminCS.asp" -->
<!-- #include virtual="/cscenterv2/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/cscenterv2/lib/classes/order/ordercls.asp"-->
<!-- #include virtual="/cscenterv2/lib/classes/cs/cs_aslistcls.asp"-->
<!-- #include virtual="/cscenterv2/lib/classes/cs/cs_cashreceiptcls.asp"-->
<!-- #include virtual="/cscenterv2/lib/KCP/site_conf_inc.asp" -->
<!-- #include virtual="/cscenterv2/lib/KCP/pp_cli_hub_lib_CASH.asp" -->
<%
'' 원 거래 시각이 2일이 지난 경우 현재시각으로 수정해서 요청 해 주시기 바랍니다.(20090501011030)
function getKCPCashEval_trad_time()
    Dim retVal
    Dim curDate
    curDate = now()
    
    retVal = replace(LEFT(curDate,10),"-","")&right("0" & hour(curDate), 2)&right("0" & minute(curDate), 2)&right("0" & second(curDate), 2)
    getKCPCashEval_trad_time = retVal
end function


function OneReceiptCancel(orderserial, idx, orgtid,cancelCause, iResultCode, iResultMsg, iCancelTid)
    ''KCP 현금영수증 발행 취소
    Dim ordr_idxx : ordr_idxx = orderserial&"_"&idx
    Dim cust_ip : cust_ip = Request.ServerVariables("REMOTE_ADDR")
    Dim tran_cd : tran_cd = "07020000"		' 현금영수증 취소 요청
    Dim mod_type : mod_type = "STSC"        '' 취소요청(STSC), 부분취소요청(STPC), 조회요청 : STSQ
    Dim mod_value : mod_value = orgtid
    Dim mod_gubn  : mod_gubn = "MG01"       '' 영수증 거래번호(TID)
    Dim trad_time : trad_time = getKCPCashEval_trad_time()         '' 원거래시각.  *현금영수증 등록 시 입력했던 trad_time을 정확히 입력해 주시기 바랍니다. (현재시각도 무바한듯?)
    Dim mod_mny                             '' 변경요청금액
    Dim rem_mny                             '' 변경처리 이전금액.
    
    Dim Rcash_cancel_noappl
    Dim Cancel_app_time
    
    Dim c_Mesg, c_Payplus
    Dim mod_data
    Dim resp_mesg, res_cd, res_msg
   
''리얼로 테스트. (임시)
'if (session("ssBctId")="icommang") then
'g_conf_gw_url    = "paygw.kcp.co.kr"
'g_conf_js_url    = "http://pay.kcp.co.kr/plugin/payplus.js"
'g_conf_js_url_ssl    = "https://pay.kcp.co.kr/plugin/payplus.js"
'
'g_conf_site_cd   = "R5523"
'g_conf_site_key  = "3uL-Drm.tpQ9q1yRqwWLSQF__"
'end if
        
    Set c_Mesg = New c_PayPlusData             ' 전문처리용 Class (library에서 정의됨)
    c_Mesg.InitialTX

    Set c_Payplus = Server.CreateObject( "pp_cli_com.KCP" )
    c_Payplus.lf_PP_CLI_LIB__init g_conf_key_dir, g_conf_log_dir, "03", g_conf_gw_url, g_conf_gw_port
    
    
    mod_data = c_Mesg.mf_set_modx_data( "mod_type",       mod_type     )
	mod_data = c_Mesg.mf_set_modx_data( "mod_value",      mod_value    )
	mod_data = c_Mesg.mf_set_modx_data( "mod_gubn",       mod_gubn     )
	mod_data = c_Mesg.mf_set_modx_data( "trad_time",      trad_time     )


    if (mod_type = "STPC") then   ' 부분취소
		mod_data = c_Mesg.mf_set_modx_data( "mod_mny",      mod_mny     )
		mod_data = c_Mesg.mf_set_modx_data( "rem_mny",      rem_mny     )
    end if
    
	c_PayPlus.lf_PP_CLI_LIB__set_plan_data mod_data
	
	
	
	'' 실행
	c_Payplus.lf_PP_CLI_LIB__do_tx g_conf_site_cd, g_conf_site_key, tran_cd, cust_ip, ordr_idxx
    resp_mesg = c_Payplus.lf_PP_CLI_LIB__get_data
    c_Mesg.mf_set_res_data(resp_mesg)                               ' 응답 전문 처리

    res_cd  = c_Mesg.mf_get_data ("res_cd")                         ' 결과 코드
    res_msg = c_Mesg.mf_get_data ("res_msg")                        ' 결과 메시지
    
    iResultCode = res_cd
    iResultMsg  = res_msg
    
    
    if res_cd = "0000" then
        iResultCode = "00" ''기존과 맞추기 위함.
        
        iCancelTid           = c_Mesg.mf_get_data( "cash_no"    ) ' 현금영수증 거래번호
        Rcash_cancel_noappl  = c_Mesg.mf_get_data( "receipt_no" ) ' 현금영수증 승인번호
        Cancel_app_time     = c_Mesg.mf_get_data( "app_time"   ) ' 승인시간(YYYYMMDDhhmmss)
        ''reg_stat   = c_Mesg.mf_get_data( "reg_stat"   ) ' 등록 상태 코드
        ''reg_desc   = c_Mesg.mf_get_data( "reg_desc"   )	' 등록 상태 설명

    end if
    
    rw iResultCode
    rw iResultMsg
    rw iCancelTid
    rw Rcash_cancel_noappl
    rw Cancel_app_time
    
    OneReceiptCancel = (iResultCode="00")
end function

function OneReceiptReq(idx,byref iResultCode,byref iResultMsg, byref iAuthCode)

    dim Tid, ResultCode, ResultMsg, AuthCode, PGAuthDate, PGAuthTime
    dim ResultpCRPice, ResultSupplyPrice, ResultTax, ResultServicePrice, ResultUseOpt, ResultCashNoAppl
    dim AckResult

    dim sqlStr
    dim goodname, cr_price, sup_price, tax, srvc_price, buyername, buyertel, buyeremail, reg_num, useopt
    dim subtotalprice, dataExists
    dim reqresultcode
    dim pggubun, sumpaymentEtc, orgpaygatetid, orgaccountdiv
    dim orderserial
    
    dataExists = false
    sqlStr = " select c.*, m.subtotalprice, isNULL(m.pggubun,'') as pggubun, isNULL(m.sumpaymentEtc,0) as sumpaymentEtc"
    sqlStr = sqlStr + " , isNULL(m.paygatetid,'') as orgpaygatetid "
    sqlStr = sqlStr + " , isNULL(m.accountdiv,'') as orgaccountdiv"
    sqlStr = sqlStr + "     from [db_academy].[dbo].tbl_academy_cash_receipt c"
    sqlStr = sqlStr + "     Join db_academy.dbo.tbl_academy_order_master m"
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
        orderserial   = rsget("orderserial")
    end if
    rsget.close
    
    if (not dataExists) then
        sqlStr = " select c.* from [db_academy].[dbo].tbl_academy_cash_receipt c"
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
    
    Dim NPay_Result, NpayCashAmt
    if (pggubun="NP") then ''네이버 페이의 경우 (2016/08/12)
        Set NPay_Result = fnCallNaverPayCashAmt(orgpaygatetid)
        NpayCashAmt    = CLng(NPay_Result.body.totalCashAmount) + sumpaymentEtc
        Set NPay_Result = Nothing
        
        if (NpayCashAmt<>cr_price) then
            sqlStr = " update C"
            sqlStr = sqlStr & " SET cr_price="&NpayCashAmt
            sqlStr = sqlStr & " from [db_academy].[dbo].tbl_academy_cash_receipt c"
            sqlStr = sqlStr & " where c.idx=" & idx
            dbget.Execute sqlStr
            
            sqlStr = " update C"
            sqlStr = sqlStr & " SET tax=convert(int,cr_price*1/11)"
            sqlStr = sqlStr & " ,sup_price=cr_price-convert(int,cr_price*1/11)"
            sqlStr = sqlStr & " from [db_academy].[dbo].tbl_academy_cash_receipt c"
            sqlStr = sqlStr & " where c.idx=" & idx
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
            sqlStr = sqlStr & " from [db_academy].[dbo].tbl_academy_cash_receipt c"
            sqlStr = sqlStr & " where c.idx=" & idx
            dbget.Execute sqlStr
            
            sqlStr = " update C"
            sqlStr = sqlStr & " SET tax=convert(int,cr_price*1/11)"
            sqlStr = sqlStr & " ,sup_price=cr_price-convert(int,cr_price*1/11)"
            sqlStr = sqlStr & " from [db_academy].[dbo].tbl_academy_cash_receipt c"
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
    
  ''KCP 현금영수증 발행.
    Dim ordr_idxx : ordr_idxx = orderserial&"_"&idx
    Dim cust_ip : cust_ip = Request.ServerVariables("REMOTE_ADDR")
    Dim tran_cd : tran_cd = "07010000" ' 현금영수증 등록 요청
    Dim corp_type : corp_type = "0"     '' 0직접판매, 1:입점몰판매.
    
    Dim c_Mesg, c_Payplus
    Dim rcpt_data_set, ordr_data_set, corp_data_set, tx_req_data_set
    Dim resp_mesg, res_cd, res_msg
   
''리얼로 테스트. (임시)
'if (session("ssBctId")="icommang") then
'g_conf_gw_url    = "paygw.kcp.co.kr"
'g_conf_js_url    = "http://pay.kcp.co.kr/plugin/payplus.js"
'g_conf_js_url_ssl    = "https://pay.kcp.co.kr/plugin/payplus.js"
'
'g_conf_site_cd   = "R5523"
'g_conf_site_key  = "3uL-Drm.tpQ9q1yRqwWLSQF__"
'end if
        
    Set c_Mesg = New c_PayPlusData             ' 전문처리용 Class (library에서 정의됨)
    c_Mesg.InitialTX

    Set c_Payplus = Server.CreateObject( "pp_cli_com.KCP" )
    c_Payplus.lf_PP_CLI_LIB__init g_conf_key_dir, g_conf_log_dir, "03", g_conf_gw_url, g_conf_gw_port
    
    
    
    ' 현금영수증 정보
    rcpt_data_set = c_Mesg.mf_set_data( "rcpt_data", "user_type",      "PGNW")           '' g_conf_user_type V6 가맹점의 경우 PGNW (변경불가)
    rcpt_data_set = c_Mesg.mf_set_data( "rcpt_data", "trad_time",      getKCPCashEval_trad_time()    )    '' 원거래시각 (원 거래 시각이 2일이 지난 경우 현재시각으로 수정해서 요청 해 주시기 바랍니다.)
    rcpt_data_set = c_Mesg.mf_set_data( "rcpt_data", "tr_code",        useopt      )    ''발행구분 (0/1)소득공제용/지출증빙용
    rcpt_data_set = c_Mesg.mf_set_data( "rcpt_data", "id_info",        reg_num      )
    rcpt_data_set = c_Mesg.mf_set_data( "rcpt_data", "amt_tot",        cr_price      )
    rcpt_data_set = c_Mesg.mf_set_data( "rcpt_data", "amt_sup",        sup_price      )
    rcpt_data_set = c_Mesg.mf_set_data( "rcpt_data", "amt_svc",        0      )
    rcpt_data_set = c_Mesg.mf_set_data( "rcpt_data", "amt_tax",        tax      )
    rcpt_data_set = c_Mesg.mf_set_data( "rcpt_data", "pay_type",       "PAXX"       )
    
    c_Mesg.InitialTX

    ' 주문 정보
    ordr_data_set = c_Mesg.mf_set_data( "ordr_data", "ordr_idxx",      ordr_idxx    )
    ordr_data_set = c_Mesg.mf_set_data( "ordr_data", "good_name",      goodname    )
    ordr_data_set = c_Mesg.mf_set_data( "ordr_data", "buyr_name",      buyername    )
    ordr_data_set = c_Mesg.mf_set_data( "ordr_data", "buyr_tel1",      buyertel    )
    ordr_data_set = c_Mesg.mf_set_data( "ordr_data", "buyr_mail",      buyeremail    )
    ordr_data_set = c_Mesg.mf_set_data( "ordr_data", "comment",        ""      )                ''비고
    c_Mesg.InitialTX

    ' 가맹점 정보
    corp_data_set = c_Mesg.mf_set_data( "corp_data", "corp_type",      corp_type    )           '' 0직접판매, 1:입점몰판매.

    if corp_type = "1" then ' 입점몰인 경우 판매상점 DATA 전문 생성
'        corp_data_set = c_Mesg.mf_set_data( "corp_data", "corp_tax_type",   corp_tax_type)
'        corp_data_set = c_Mesg.mf_set_data( "corp_data", "corp_tax_no",     corp_tax_no  )
'        corp_data_set = c_Mesg.mf_set_data( "corp_data", "corp_sell_tax_no",corp_tax_no  )
'        corp_data_set = c_Mesg.mf_set_data( "corp_data", "corp_nm",         corp_nm      )
'        corp_data_set = c_Mesg.mf_set_data( "corp_data", "corp_owner_nm",   corp_owner_nm)
'        corp_data_set = c_Mesg.mf_set_data( "corp_data", "corp_addr",       corp_addr    )
'        corp_data_set = c_Mesg.mf_set_data( "corp_data", "corp_telno",      corp_telno   )
    end if
    c_Mesg.InitialTX

	tx_req_data_set = c_Mesg.mf_set_req_data( rcpt_data_set )
    tx_req_data_set = c_Mesg.mf_set_req_data( ordr_data_set )
    tx_req_data_set = c_Mesg.mf_set_req_data( corp_data_set )
    c_Mesg.InitialTX

    c_PayPlus.lf_PP_CLI_LIB__set_plan_data tx_req_data_set
	c_Mesg.InitialTX
	
	'' 실행
	c_Payplus.lf_PP_CLI_LIB__do_tx g_conf_site_cd, g_conf_site_key, tran_cd, cust_ip, ordr_idxx
    resp_mesg = c_Payplus.lf_PP_CLI_LIB__get_data
    c_Mesg.mf_set_res_data(resp_mesg)                               ' 응답 전문 처리

    res_cd  = c_Mesg.mf_get_data ("res_cd")                         ' 결과 코드
    res_msg = c_Mesg.mf_get_data ("res_msg")                        ' 결과 메시지
    
    iResultCode = res_cd
    iResultMsg  = res_msg
    
    if res_cd = "0000" then
        iResultCode = "00" ''기존과 맞추기 위함.
        
    	Tid             = c_Mesg.mf_get_data( "cash_no"    ) ' 현금영수증 거래번호 TID
        iAuthCode       = c_Mesg.mf_get_data( "receipt_no" ) ' 현금영수증 승인번호
        ResultCashNoAppl = iAuthCode
        PGAuthDate   = c_Mesg.mf_get_data( "app_time"   ) ' 승인시간(YYYYMMDDhhmmss)
        PGAuthTime   = MID(PGAuthDate,9,6)
        'reg_stat   = c_Mesg.mf_get_data( "reg_stat"   ) ' 등록 상태 코드
        'reg_desc   = c_Mesg.mf_get_data( "reg_desc"   )	' 등록 상태 설명

    end if

'rw iResultCode
'rw iResultMsg
'rw dumicash_no
'rw iAuthCode
'rw PGAuthDate
'rw PGAuthTime
'
'response.end

    ''결과 저장 - 관리자 발행시 성공인경우만 저장.
    IF iResultCode = "00" THEN
        sqlStr = "update [db_academy].[dbo].tbl_academy_cash_receipt" + VbCrlf
        sqlStr = sqlStr + " set tid='" + CStr(Tid) + "'" + VbCrlf
        sqlStr = sqlStr + " , resultcode='" + CStr(iResultCode) + "'" + VbCrlf
        sqlStr = sqlStr + " , resultmsg='" + LeftB(CStr(Replace(iResultMsg,"'","")),100) + "'" + VbCrlf
        sqlStr = sqlStr + " , authcode='" + CStr(iAuthCode) + "'" + VbCrlf
        sqlStr = sqlStr + " , resultcashnoappl='" + CStr(ResultCashNoAppl) + "'" + VbCrlf
        sqlStr = sqlStr + " where idx=" + CStr(idx)

        dbget.Execute sqlStr
        
         ''2016/06/30 추가. 승인일
        sqlStr = "update [db_academy].[dbo].tbl_academy_cash_receipt" + VbCrlf
        sqlStr = sqlStr + " SET evalDt='"&LEFT(PGAuthDate,4)&"-"&MID(PGAuthDate,5,2)&"-"&MID(PGAuthDate,7,2)&" "&LEFT(PGAuthTime,2)&":"&MID(PGAuthTime,3,2)&":"&MID(PGAuthTime,5,2)&"'" + VbCrlf
        sqlStr = sqlStr + " where idx=" + CStr(idx)
        dbget.Execute sqlStr
    ELSE
        if (iResultCode<>"00") then ''and ((Left(iResultMsg,Len("[269051]"))="[269051]") or (Left(iResultMsg,Len("[269050]"))="[269050]") or (Left(iResultMsg,Len("[505658]"))="[505658]")) then
            sqlStr = "update [db_academy].[dbo].tbl_academy_cash_receipt" + VbCrlf
            sqlStr = sqlStr + " set cancelyn='F'"
            sqlStr = sqlStr + " , resultmsg='" + LeftB(CStr(Replace(ResultMsg,"'","")),100) + "'" + VbCrlf
            sqlStr = sqlStr + " where idx=" + CStr(idx)
            dbget.Execute sqlStr
        end if
    End IF

    set c_Payplus = nothing
    set c_Mesg    = nothing

    OneReceiptReq = (iResultCode = "00")
end function


dim chkPrint, i, Atype
dim pggubun, sumpaymentEtc, subtotalPrice, accountdiv, orgpaygatetid

chkPrint = request("chkPrint")
Atype    = RequestCheckVar(request("Atype"),9)
pggubun  = RequestCheckVar(request("pggubun"),10)
if chkPrint <> "" then
	if checkNotValidHTML(chkPrint) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end if
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
        sqlStr = " select idx, orderserial, resultcode, cancelyn, reg_num from [db_academy].[dbo].tbl_academy_cash_receipt"
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

               ''preIssuedTaxExists = chkRegTax(orderserial)

               sqlStr = " select count(idx) as cnt from  [db_academy].[dbo].tbl_academy_cash_receipt"
               sqlStr = sqlStr + " where orderserial='" + orderserial + "'"
               sqlStr = sqlStr + " and resultcode='00'"
               sqlStr = sqlStr + " and cancelyn='N'"
               sqlStr = sqlStr + " and idx<>"&idx

               rsget.Open sqlStr,dbget,1
                    preIssuedExists = rsget("cnt")>0
               rsget.close

               if (preIssuedExists) then
                    infoMsg = infoMsg & " <font color='red'>기발행 내역 존재 - 삭제:" & orderserial & "[" & idx & "]" & "</font><br>" & VbCrlf
                    sqlStr = " update [db_academy].[dbo].tbl_academy_cash_receipt"
                    sqlStr = sqlStr + " set cancelyn='D'"
                    sqlStr = sqlStr + " where idx=" & CStr(idx)
                    dbget.Execute sqlStr
               'elseif (preIssuedTaxExists<>"none") then
               '     infoMsg = infoMsg & " <font color='red'>세금계산서 발행 내역 존재 - 삭제:" & orderserial & "[" & idx & "]" & "</font><br>" & VbCrlf
               '     sqlStr = " update [db_academy].[dbo].tbl_academy_cash_receipt"
               '     sqlStr = sqlStr + " set cancelyn='D'"
               '    sqlStr = sqlStr + " where idx=" & CStr(idx)
               '     dbget.Execute sqlStr
               else
                    iResultCode = ""
                    iResultMsg  = ""
                    if (Not OneReceiptReq(idx, iResultCode, iResultMsg, iAuthCode)) then
                        infoMsg = infoMsg & " <font color='red'>발행 실패 :" & "[" & iResultCode & "]" & iResultMsg & "</font><br>" & VbCrlf
                    else
                        infoMsg = infoMsg & " 발행 성공 :" & "[" & iResultCode & "]" & iResultMsg & "<br>" & VbCrlf

                        
                            sqlStr = " update [db_academy].[dbo].tbl_academy_order_master" & VbCrlf
                            sqlStr = sqlStr & " set authcode = (case when accountdiv in ('7', '20') then '" & iAuthCode & "' else authcode end) " & VbCrlf
                            if (reg_num="0100001234") then
                                sqlStr = sqlStr & " ,cashreceiptreq='J'" & VbCrlf   '' 자진발급 2016/06/22
                            else
                                sqlStr = sqlStr & " ,cashreceiptreq='S'" & VbCrlf   
                            end if
                            sqlStr = sqlStr & " where orderserial='" & orderserial & "'"

                            dbget.Execute sqlStr
                        
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
        sqlStr = " select idx, orderserial, resultcode, cancelyn, tid from [db_academy].[dbo].tbl_academy_cash_receipt"
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
            if (Not OneReceiptCancel(orderserial,idx,orgtid,icancelCause, iResultCode, iResultMsg, iAuthCode)) then
                infoMsg = infoMsg & " <font color='red'>취소 실패 :" & "[" & iResultCode & "]" & iResultMsg & "</font><br>" & VbCrlf
                infoMsg = infoMsg & " orgtid :"&orgtid& VbCrlf
            else
                infoMsg = infoMsg & " 취소 성공 :" & "[" & iResultCode & "]" & iResultMsg & "<br>" & VbCrlf

                sqlStr = " update [db_academy].[dbo].tbl_academy_cash_receipt" & VbCrlf
                sqlStr = sqlStr & " set canceltid='" & iAuthCode & "'" & VbCrlf
                sqlStr = sqlStr & " ,cancelyn='Y'" & VbCrlf
                sqlStr = sqlStr & " where idx=" & idx & ""

                dbget.Execute sqlStr

                ''고객센터에서 취소할 경우
                
                    sqlStr = " update db_academy.dbo.tbl_academy_order_master" & VbCrlf
                    sqlStr = sqlStr & " set cashreceiptreq=NULL" & VbCrlf
                    sqlStr = sqlStr & " ,authcode = (case when accountdiv in ('7', '20') then NULL else authcode end) " & VbCrlf
                    sqlStr = sqlStr & " where orderserial='"&orderserial&"'"

                    dbget.Execute sqlStr
            end if
        else
            infoMsg = infoMsg & "발행 코드 존재 안함 " & "[" & idx & "]" & VbCrlf
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
    sqlStr = sqlStr + " from [db_academy].[dbo].tbl_academy_cash_receipt C"
    sqlStr = sqlStr + " join db_academy.dbo.tbl_academy_order_master m"
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
        infoMsg = infoMsg & "주문번호,인덱스 체크 오류 " & "[" & orderserial & "]" & VbCrlf
        
        response.write infoMsg
        dbget.Close() : response.end
    end if
    
    ''기발행 체크
    dim duppEvalIDX : duppEvalIDX=0
    sqlStr = " select top 1 idx from [db_academy].[dbo].tbl_academy_cash_receipt C" & VbCrlf
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
        infoMsg = infoMsg & "타 발행 내역 존재 " & "[" & duppEvalIDX & "]" & VbCrlf
        
        response.write infoMsg
        dbget.Close() : response.end
    end if
    
    if (NOT ((resultcode="00") and (cancelyn="N"))) then
        infoMsg = infoMsg & "기발행 내역 아님 " & "[" & idx & "]" & VbCrlf
        
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
        infoMsg = infoMsg & "발행 금액 확인 필요 " & "[" & ReEvalCashAmt & "<>"&request("mayPrc")&"]" & VbCrlf
        
        response.write infoMsg
        dbget.Close() : response.end
    end if
    
    ''infoMsg = infoMsg & ReEvalCashAmt &"|"&ReEvalCashSupp & VbCrlf
    
    '' 발행 한줄 꽃아 넣음
    sqlStr = " select * from [db_academy].[dbo].tbl_academy_cash_receipt where 1=0"
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
    sqlStr = sqlStr&" from [db_academy].[dbo].tbl_academy_cash_receipt N"&VBCRLF
    sqlStr = sqlStr&"     JOin [db_academy].[dbo].tbl_academy_cash_receipt P"&VBCRLF
    sqlStr = sqlStr&"     on 1=1"&VBCRLF
    sqlStr = sqlStr&"     and P.idx="&idx&VBCRLF
    sqlStr = sqlStr&" where N.idx="&reEvalIDX&VBCRLF
    dbget.Execute sqlStr
    
    ''발행 먼저
    iResultCode = ""
    iResultMsg  = ""
    iAuthCode   = ""
    
    if (Not OneReceiptReq(reEvalIDX, iResultCode, iResultMsg, iAuthCode)) then
        infoMsg = infoMsg & " <font color='red'>발행 실패 :" & "[" & iResultCode & "]" & iResultMsg & "</font><br>" & VbCrlf
    else
        infoMsg = infoMsg & " 발행 성공 :" & "[" & iResultCode & "]" & iResultMsg & "<br>" & VbCrlf

        
            sqlStr = " update [db_academy].[dbo].tbl_academy_order_master" & VbCrlf
            sqlStr = sqlStr & " set authcode = (case when accountdiv in ('7', '20') then '" & iAuthCode & "' else authcode end) " & VbCrlf
            if (reg_num="0100001234") then
                sqlStr = sqlStr & " ,cashreceiptreq='J'" & VbCrlf   '' 자진발급 2016/06/22
            else
                sqlStr = sqlStr & " ,cashreceiptreq='S'" & VbCrlf   
            end if
            sqlStr = sqlStr & " where orderserial='" & orderserial & "'"

            dbget.Execute sqlStr

        
        ''취소
        iResultCode = ""
        iResultMsg  = ""
        icancelCause = "재발행"
        iAuthCode   = ""
        
        if (Not OneReceiptCancel(orderserial,idx,orgtid,icancelCause, iResultCode, iResultMsg, iAuthCode)) then
            infoMsg = infoMsg & " <font color='red'>취소 실패 :" & "[" & iResultCode & "]" & iResultMsg & "</font><br>" & VbCrlf
            infoMsg = infoMsg & " orgtid :"&orgtid& VbCrlf
        else
            infoMsg = infoMsg & " 취소 성공 :" & "[" & iResultCode & "]" & iResultMsg & "<br>" & VbCrlf

            sqlStr = " update [db_academy].[dbo].tbl_academy_cash_receipt" & VbCrlf
            sqlStr = sqlStr & " set canceltid='" & iAuthCode & "'" & VbCrlf
            sqlStr = sqlStr & " ,cancelyn='Y'" & VbCrlf
            sqlStr = sqlStr & " where idx=" & idx & ""

            dbget.Execute sqlStr

            
        end if
    end if
    
    
    
    
elseif (Atype="CH") then
    response.write "지정되지 않았습니다. - " & Atype & "<br>"
'    orgtid = request("tid")
'    icancelCause ="오발행"
'    if (Not OneReceiptCancel(orderserial,idx,orgtid,icancelCause, iResultCode, iResultMsg, iAuthCode)) then
'        rw " <font color='red'>취소 실패 :" & "[" & iResultCode & "]" & iResultMsg & "</font><br>" & VbCrlf
'    else
'        rw iResultMsg
'    end if
elseif (Atype="AUTO1") then
    response.write "지정되지 않았습니다. - " & Atype & "<br>"
else
    response.write "지정되지 않았습니다. - " & Atype & "<br>"
end if
response.write infoMsg

%>
<br>
<a href="javascript:history.back();">&lt;&lt;Back</a>

<% if (Atype="C2") then %>
&nbsp;
<a href="javascript:window.close();">&lt;&lt;Close</a>
<% end if %>
<!-- #include virtual="/cscenterv2/lib/db/dbclose.asp" -->
