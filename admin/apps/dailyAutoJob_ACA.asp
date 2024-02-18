<%@ language=vbscript %>
<% option explicit %>

<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/xmlhttpUtil.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_taxsheetcls.asp"-->
<!-- #include virtual="/lib/email/maillib.asp" -->
<!-- #include virtual="/lib/classes/smscls.asp" -->
<!-- #include virtual="/cscenterv2/lib/classes/cs/cs_aslistcls.asp"-->
<!-- #include virtual="/cscenterv2/lib/classes/cs/cs_cashreceiptcls.asp"-->
<!-- #include virtual="/cscenterv2/lib/KCP/site_conf_inc.asp" -->
<!-- #include virtual="/cscenterv2/lib/KCP/pp_cli_hub_lib_CASH.asp" -->
<%
function CheckVaildIP(ref)
    CheckVaildIP = false

    dim VaildIP : VaildIP = Array("110.93.128.107","61.252.133.2","61.252.133.9","61.252.133.10","61.252.133.80","110.93.128.114","110.93.128.113","61.252.133.70")
    if application("Svr_Info") = "Dev" then
        VaildIP = Array("192.168.1.70")
    end if
    dim i
    for i=0 to UBound(VaildIP)
        if (VaildIP(i)=ref) then
            CheckVaildIP = true
            exit function
        end if
    next
end function

function getKCPCashEval_trad_time()
    Dim retVal
    Dim curDate
    curDate = now()
    
    retVal = replace(LEFT(curDate,10),"-","")&right("0" & hour(curDate), 2)&right("0" & minute(curDate), 2)&right("0" & second(curDate), 2)
    getKCPCashEval_trad_time = retVal
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

function OneReceiptReqACA_KCP(idx,byref iResultCode,byref iResultMsg, byref iAuthCode)

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
    
    rsACADEMYget.CursorLocation = adUseClient
    rsACADEMYget.Open sqlStr,dbACADEMYget,adOpenForwardOnly, adLockReadOnly
    if Not rsACADEMYget.Eof then
        dataExists = true
        goodname    = db2html(rsACADEMYget("goodname"))
        cr_price    = rsACADEMYget("cr_price")
        sup_price   = rsACADEMYget("sup_price")
        tax         = rsACADEMYget("tax")
        srvc_price  = rsACADEMYget("srvc_price")
        buyername   = db2html(rsACADEMYget("buyername"))
        buyertel    = rsACADEMYget("buyertel")
        buyeremail  = db2html(rsACADEMYget("buyeremail"))

        reg_num     = rsACADEMYget("reg_num")
        useopt      = rsACADEMYget("useopt")
        subtotalprice = rsACADEMYget("subtotalprice")
        reqresultcode  = rsACADEMYget("resultcode")
        pggubun = rsACADEMYget("pggubun")
        sumpaymentEtc = rsACADEMYget("sumpaymentEtc")
        orgpaygatetid = rsACADEMYget("orgpaygatetid")
        orgaccountdiv = TRIM(rsACADEMYget("orgaccountdiv"))
        orderserial   = rsACADEMYget("orderserial")
    end if
    rsACADEMYget.close
    
    if (not dataExists) then
        sqlStr = " select c.* from [db_academy].[dbo].tbl_academy_cash_receipt c"
        sqlStr = sqlStr + " where c.idx=" & idx
        rsACADEMYget.CursorLocation = adUseClient
        rsACADEMYget.Open sqlStr,dbACADEMYget,adOpenForwardOnly, adLockReadOnly
        if Not rsACADEMYget.Eof then
            goodname    = db2html(rsACADEMYget("goodname"))
            cr_price    = rsACADEMYget("cr_price")
            sup_price   = rsACADEMYget("sup_price")
            tax         = rsACADEMYget("tax")
            srvc_price  = rsACADEMYget("srvc_price")
            buyername   = db2html(rsACADEMYget("buyername"))
            buyertel    = rsACADEMYget("buyertel")
            buyeremail  = db2html(rsACADEMYget("buyeremail"))
    
            reg_num     = rsACADEMYget("reg_num")
            useopt      = rsACADEMYget("useopt")
            subtotalprice = cr_price
            reqresultcode  = rsACADEMYget("resultcode")
            
            sumpaymentEtc = 0
        end if
        rsACADEMYget.close
    end if
    
    Dim NPay_Result, NpayCashAmt
    if (pggubun="NP") then ''���̹� ������ ��� (2016/08/12)
        Set NPay_Result = fnCallNaverPayCashAmt(orgpaygatetid)
        NpayCashAmt    = CLng(NPay_Result.body.totalCashAmount) + sumpaymentEtc
        Set NPay_Result = Nothing
        
        if (NpayCashAmt<>cr_price) then
            sqlStr = " update C"
            sqlStr = sqlStr & " SET cr_price="&NpayCashAmt
            sqlStr = sqlStr & " from [db_academy].[dbo].tbl_academy_cash_receipt c"
            sqlStr = sqlStr & " where c.idx=" & idx
            dbACADEMYget.Execute sqlStr
            
            sqlStr = " update C"
            sqlStr = sqlStr & " SET tax=convert(int,cr_price*1/11)"
            sqlStr = sqlStr & " ,sup_price=cr_price-convert(int,cr_price*1/11)"
            sqlStr = sqlStr & " from [db_academy].[dbo].tbl_academy_cash_receipt c"
            sqlStr = sqlStr & " where c.idx=" & idx
            dbACADEMYget.Execute sqlStr
            
            OneReceiptReqACA_KCP = False
            iResultMsg    = "NPAY �ݾ� ���� ���ۼ�.["&cr_price&"::"&NpayCashAmt&"]"
            Exit Function
        end if
    else
        if ((orgaccountdiv="20") or (orgaccountdiv="7")) then
                
        else
            subtotalprice = sumpaymentEtc
        end if
        
        ''subtotalPrice = subtotalPrice+GetReceiptMinusOrderSUM(orderserial) ''��ǰ�ݾ� �߰� 
        subtotalPrice = subtotalPrice+0 ''���� �ǹ̾���..
        
        if (subtotalprice<>cr_price) then
            sqlStr = " update C"
            sqlStr = sqlStr & " SET cr_price="&subtotalprice
            sqlStr = sqlStr & " from [db_academy].[dbo].tbl_academy_cash_receipt c"
            sqlStr = sqlStr & " where c.idx=" & idx
            dbACADEMYget.Execute sqlStr
            
            sqlStr = " update C"
            sqlStr = sqlStr & " SET tax=convert(int,cr_price*1/11)"
            sqlStr = sqlStr & " ,sup_price=cr_price-convert(int,cr_price*1/11)"
            sqlStr = sqlStr & " from [db_academy].[dbo].tbl_academy_cash_receipt c"
            sqlStr = sqlStr & " where c.idx=" & idx
            dbACADEMYget.Execute sqlStr
            
            OneReceiptReqACA_KCP = False
            iResultMsg    = "�ݾ� ���� ���ۼ�.["&cr_price&"::"&subtotalprice&"]"
            Exit Function
        end if
    end if
    
    if (useopt="0") and ((Len(reg_num)<>13) and (Len(reg_num)<>10) and (Len(reg_num)<>11) and (Len(reg_num)<>18)) then
        OneReceiptReqACA_KCP = False
        iResultMsg    = "�ֹι�ȣ/�ڵ��� �ڸ� ����"
        Exit Function
    end if

    if (useopt="1") and ((Len(reg_num)<>13) and (Len(reg_num)<>10) and (Len(reg_num)<>11) and (Len(reg_num)<>18)) then
        OneReceiptReqACA_KCP = False
        iResultMsg    = "����ڹ�ȣ/ �ֹι�ȣ /�ڵ��� �ڸ� ����"
        Exit Function
    end if

    if (reqresultcode<>"R") then
        OneReceiptReqACA_KCP = False
        iResultMsg    = "����� Ȯ��"
        Exit Function
    end if

   
  ''KCP ���ݿ����� ����.
    Dim ordr_idxx : ordr_idxx = orderserial&"_"&idx
    Dim cust_ip : cust_ip = Request.ServerVariables("REMOTE_ADDR")
    Dim tran_cd : tran_cd = "07010000" ' ���ݿ����� ��� ��û
    Dim corp_type : corp_type = "0"     '' 0�����Ǹ�, 1:�������Ǹ�.
    
    Dim c_Mesg, c_Payplus
    Dim rcpt_data_set, ordr_data_set, corp_data_set, tx_req_data_set
    Dim resp_mesg, res_cd, res_msg
   
''����� �׽�Ʈ. (�ӽ�)
'g_conf_gw_url    = "paygw.kcp.co.kr"
'g_conf_js_url    = "http://pay.kcp.co.kr/plugin/payplus.js"
'g_conf_js_url_ssl    = "https://pay.kcp.co.kr/plugin/payplus.js"
'
'g_conf_site_cd   = "R5523"
'g_conf_site_key  = "3uL-Drm.tpQ9q1yRqwWLSQF__"
        
    Set c_Mesg = New c_PayPlusData             ' ����ó���� Class (library���� ���ǵ�)
    c_Mesg.InitialTX

    Set c_Payplus = Server.CreateObject( "pp_cli_com.KCP" )
    c_Payplus.lf_PP_CLI_LIB__init g_conf_key_dir, g_conf_log_dir, "03", g_conf_gw_url, g_conf_gw_port
    
    
    
    ' ���ݿ����� ����
    rcpt_data_set = c_Mesg.mf_set_data( "rcpt_data", "user_type",      "PGNW")           '' g_conf_user_type V6 �������� ��� PGNW (����Ұ�)
    rcpt_data_set = c_Mesg.mf_set_data( "rcpt_data", "trad_time",      getKCPCashEval_trad_time()    )    '' ���ŷ��ð� (�� �ŷ� �ð��� 2���� ���� ��� ����ð����� �����ؼ� ��û �� �ֽñ� �ٶ��ϴ�.)
    rcpt_data_set = c_Mesg.mf_set_data( "rcpt_data", "tr_code",        useopt      )    ''���౸�� (0/1)�ҵ������/����������
    rcpt_data_set = c_Mesg.mf_set_data( "rcpt_data", "id_info",        reg_num      )
    rcpt_data_set = c_Mesg.mf_set_data( "rcpt_data", "amt_tot",        cr_price      )
    rcpt_data_set = c_Mesg.mf_set_data( "rcpt_data", "amt_sup",        sup_price      )
    rcpt_data_set = c_Mesg.mf_set_data( "rcpt_data", "amt_svc",        0      )
    rcpt_data_set = c_Mesg.mf_set_data( "rcpt_data", "amt_tax",        tax      )
    rcpt_data_set = c_Mesg.mf_set_data( "rcpt_data", "pay_type",       "PAXX"       )
    
    c_Mesg.InitialTX

    ' �ֹ� ����
    ordr_data_set = c_Mesg.mf_set_data( "ordr_data", "ordr_idxx",      ordr_idxx    )
    ordr_data_set = c_Mesg.mf_set_data( "ordr_data", "good_name",      goodname    )
    ordr_data_set = c_Mesg.mf_set_data( "ordr_data", "buyr_name",      buyername    )
    ordr_data_set = c_Mesg.mf_set_data( "ordr_data", "buyr_tel1",      buyertel    )
    ordr_data_set = c_Mesg.mf_set_data( "ordr_data", "buyr_mail",      buyeremail    )
    ordr_data_set = c_Mesg.mf_set_data( "ordr_data", "comment",        ""      )                ''���
    c_Mesg.InitialTX

    ' ������ ����
    corp_data_set = c_Mesg.mf_set_data( "corp_data", "corp_type",      corp_type    )           '' 0�����Ǹ�, 1:�������Ǹ�.

    if corp_type = "1" then ' �������� ��� �ǸŻ��� DATA ���� ����
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
	
	'' ����
	c_Payplus.lf_PP_CLI_LIB__do_tx g_conf_site_cd, g_conf_site_key, tran_cd, cust_ip, ordr_idxx
    resp_mesg = c_Payplus.lf_PP_CLI_LIB__get_data
    c_Mesg.mf_set_res_data(resp_mesg)                               ' ���� ���� ó��

    res_cd  = c_Mesg.mf_get_data ("res_cd")                         ' ��� �ڵ�
    res_msg = c_Mesg.mf_get_data ("res_msg")                        ' ��� �޽���
    
    iResultCode = res_cd
    iResultMsg  = res_msg
    
    if res_cd = "0000" then
        iResultCode = "00" ''������ ���߱� ����.
        
    	Tid             = c_Mesg.mf_get_data( "cash_no"    ) ' ���ݿ����� �ŷ���ȣ TID
        iAuthCode       = c_Mesg.mf_get_data( "receipt_no" ) ' ���ݿ����� ���ι�ȣ
        ResultCashNoAppl = iAuthCode
        PGAuthDate   = c_Mesg.mf_get_data( "app_time"   ) ' ���νð�(YYYYMMDDhhmmss)
        PGAuthTime   = MID(PGAuthDate,9,6)
        'reg_stat   = c_Mesg.mf_get_data( "reg_stat"   ) ' ��� ���� �ڵ�
        'reg_desc   = c_Mesg.mf_get_data( "reg_desc"   )	' ��� ���� ����

    end if

'rw iResultCode
'rw iResultMsg
'rw dumicash_no
'rw iAuthCode
'rw PGAuthDate
'rw PGAuthTime
'
'response.end

    ''��� ���� - ������ ����� �����ΰ�츸 ����.
    IF iResultCode = "00" THEN
        sqlStr = "update [db_academy].[dbo].tbl_academy_cash_receipt" + VbCrlf
        sqlStr = sqlStr + " set tid='" + CStr(Tid) + "'" + VbCrlf
        sqlStr = sqlStr + " , resultcode='" + CStr(iResultCode) + "'" + VbCrlf
        sqlStr = sqlStr + " , resultmsg='" + LeftB(CStr(Replace(iResultMsg,"'","")),100) + "'" + VbCrlf
        sqlStr = sqlStr + " , authcode='" + CStr(iAuthCode) + "'" + VbCrlf
        sqlStr = sqlStr + " , resultcashnoappl='" + CStr(ResultCashNoAppl) + "'" + VbCrlf
        sqlStr = sqlStr + " where idx=" + CStr(idx)

        dbACADEMYget.Execute sqlStr
        
         ''2016/06/30 �߰�. ������
        sqlStr = "update [db_academy].[dbo].tbl_academy_cash_receipt" + VbCrlf
        sqlStr = sqlStr + " SET evalDt='"&LEFT(PGAuthDate,4)&"-"&MID(PGAuthDate,5,2)&"-"&MID(PGAuthDate,7,2)&" "&LEFT(PGAuthTime,2)&":"&MID(PGAuthTime,3,2)&":"&MID(PGAuthTime,5,2)&"'" + VbCrlf
        sqlStr = sqlStr + " where idx=" + CStr(idx)
        dbACADEMYget.Execute sqlStr
    ELSE
        if (iResultCode<>"00") then ''and ((Left(iResultMsg,Len("[269051]"))="[269051]") or (Left(iResultMsg,Len("[269050]"))="[269050]") or (Left(iResultMsg,Len("[505658]"))="[505658]")) then
            sqlStr = "update [db_academy].[dbo].tbl_academy_cash_receipt" + VbCrlf
            sqlStr = sqlStr + " set cancelyn='F'"
            sqlStr = sqlStr + " , resultmsg='" + LeftB(CStr(Replace(ResultMsg,"'","")),100) + "'" + VbCrlf
            sqlStr = sqlStr + " where idx=" + CStr(idx)
            dbACADEMYget.Execute sqlStr
        end if
    End IF

    set c_Payplus = nothing
    set c_Mesg    = nothing

    OneReceiptReqACA_KCP = (iResultCode = "00")
end function


dim ref : ref = Request.ServerVariables("REMOTE_ADDR")

if (Not CheckVaildIP(ref)) then
    dbACADEMYget.Close()
    response.write ref
    response.end
end if

dim act     : act = requestCheckVar(request("act"),32)
dim param1  : param1 = requestCheckVar(request("param1"),32)
dim sqlStr, i, paramData
dim retCnt : retCnt = 0

dim chkPrint, infoMsg, idx, orderserial, resultcode, cancelyn, preIssuedExists, preIssuedTaxExists, iResultCode, iResultMsg, iAuthCode, reg_num
dim paramInfo, retParamInfo, RetErr, retval
dim userid, buyname, buyhp, buyEmail,InsureCd, vRdSite
dim idArr, paygateTidArr, cnt
dim osms
select Case act


    Case "cardCancelAcademy"
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
            sqlStr = " select top 3 a.id, m.orderserial, m.paygateTid "
            sqlStr = sqlStr & " from "
            sqlStr = sqlStr & " [db_academy].dbo.tbl_academy_as_list a"
            sqlStr = sqlStr & " Join db_academy.dbo.tbl_academy_order_master m"
            sqlStr = sqlStr & " on a.orderserial=m.orderserial"
            sqlStr = sqlStr & " Join [db_academy].dbo.tbl_academy_as_refund_info f"
            sqlStr = sqlStr & " on a.id=f.asid"
            sqlStr = sqlStr & " and f.returnmethod not in ('R120','R022')"
            sqlStr = sqlStr & " where a.currstate='B001'"
            sqlStr = sqlStr & " and a.deleteyn='N'"
            sqlStr = sqlStr & " and a.divcd='A007'"
            sqlStr = sqlStr & " and m.cancelyn='Y'"
            sqlStr = sqlStr & " and m.ipkumdiv >=4"
            sqlStr = sqlStr & " and m.ipkumdiv <7"
            sqlStr = sqlStr & " and (Left(m.paygateTid,9)='IniTechPG' or isNULL(m.pggubun,'')='KP')"
            ''sqlStr = sqlStr & " and Left(m.paygateTid,9)='IniTechPG'"
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
            dbACADEMYget.Close() :  response.end
        else
            for i=LBound(idArr) to UBound(idArr)
                if (idArr(i)<>"") then
                    paramData = "redSsnKey=system&id="&idArr(i)&"&tid="&paygateTidArr(i)&"&msg="

                    ''response.write paramData&"<br>"
                    if (application("Svr_Info")<>"Dev") then
                         retVal = SendReq("http://webadmin.10x10.co.kr/cscenterv2/cs/pop_CardCancel_process.asp",paramData)
                    else
                         retVal = SendReq("http://testwebadmin.10x10.co.kr/cscenterv2/cs/pop_CardCancel_process.asp",paramData)
                    end if

                    response.write retVal&VbCRLF
                end if
            next
        end if
    Case "cashreceiptACA"
        if application("Svr_Info") = "Dev" then
            response.write "S_ERR|Dev Svr"
            response.end
        end if
        
        chkPrint = ""
        infoMsg = ""

        sqlStr = " select top 1 c.idx, c.orderserial, c.resultcode, c.cancelyn "
        sqlStr = sqlStr + " from db_academy.dbo.tbl_academy_order_master m"
        sqlStr = sqlStr + " 	Join [db_academy].[dbo].[tbl_academy_cash_receipt] c"
        sqlStr = sqlStr + " 	on c.orderserial=m.orderserial"
        sqlStr = sqlStr + " 	and c.resultcode='R'"
        sqlStr = sqlStr + " 	and c.cancelyn='N'"
        if application("Svr_Info") = "Dev" then
            sqlStr = sqlStr + " where  m.ipkumdiv>='4'"     ''�Ϻ�����̻�.
        else
            sqlStr = sqlStr + " where  m.ipkumdiv>='7'"     ''�Ϻ�����̻�.
        end if
        sqlStr = sqlStr + " and m.cashreceiptreq='R'"
        sqlStr = sqlStr + " and m.authcode is NULL"
        sqlStr = sqlStr + " and m.accountdiv in ('7','20')"       ''2011 ���� �ǽð� ��ü��..
        sqlStr = sqlStr + " and m.cancelyn='N'"
        sqlStr = sqlStr + " and m.subtotalPrice>0"
        sqlStr = sqlStr + " order by m.idx desc"
        
        rsACADEMYget.CursorLocation = adUseClient
        rsACADEMYget.Open sqlStr,dbACADEMYget,adOpenForwardOnly, adLockReadOnly
        if Not rsACADEMYget.Eof then
            do until rsACADEMYget.eof
            chkPrint = chkPrint & rsACADEMYget("idx") & ","
            rsACADEMYget.MoveNext
    		loop
        end if
        rsACADEMYget.close
        
        if Right(chkPrint,1)="," then chkPrint=Left(chkPrint,Len(chkPrint)-1)
        chkPrint = split(chkPrint,",")
        
        if UBound(chkPrint)>-1 then
            for i=0 to UBound(chkPrint)
                idx = 0
                sqlStr = " select idx, orderserial, resultcode, cancelyn, reg_num from [db_academy].[dbo].tbl_academy_cash_receipt"
                sqlStr = sqlStr + " where idx=" & chkPrint(i)
                
                rsACADEMYget.CursorLocation = adUseClient
                rsACADEMYget.Open sqlStr,dbACADEMYget,adOpenForwardOnly, adLockReadOnly
                if Not rsACADEMYget.Eof then
                    idx         = rsACADEMYget("idx")
                    orderserial = rsACADEMYget("orderserial")
                    resultcode  = rsACADEMYget("resultcode")
                    cancelyn    = rsACADEMYget("cancelyn")
                    reg_num     = rsACADEMYget("reg_num")
                end if
                rsACADEMYget.close
        
                if (idx<>0) then
                    ''����� ���� ���� üũ
                    if (orderserial<>"") then
        
                       preIssuedExists = False
                       preIssuedTaxExists = False
        
                       ''preIssuedTaxExists = chkRegTax(orderserial)
        
                       sqlStr = " select count(idx) as cnt from  [db_academy].[dbo].tbl_academy_cash_receipt"
                       sqlStr = sqlStr + " where orderserial='" + orderserial + "'"
                       sqlStr = sqlStr + " and resultcode='00'"
                       sqlStr = sqlStr + " and cancelyn='N'"
                       sqlStr = sqlStr + " and idx<>"&idx
        
                       rsACADEMYget.CursorLocation = adUseClient
                       rsACADEMYget.Open sqlStr,dbACADEMYget,adOpenForwardOnly, adLockReadOnly
                            preIssuedExists = rsACADEMYget("cnt")>0
                       rsACADEMYget.close
        
                       if (preIssuedExists) then
                            infoMsg = infoMsg & " <font color='red'>����� ���� ���� - ����:" & orderserial & "[" & idx & "]" & "</font><br>" & VbCrlf
                            sqlStr = " update [db_academy].[dbo].tbl_academy_cash_receipt"
                            sqlStr = sqlStr + " set cancelyn='D'"
                            sqlStr = sqlStr + " where idx=" & CStr(idx)
                            dbACADEMYget.Execute sqlStr
                       'elseif (preIssuedTaxExists<>"none") then
                       '     infoMsg = infoMsg & " <font color='red'>���ݰ�꼭 ���� ���� ���� - ����:" & orderserial & "[" & idx & "]" & "</font><br>" & VbCrlf
                       '     sqlStr = " update [db_academy].[dbo].tbl_academy_cash_receipt"
                       '     sqlStr = sqlStr + " set cancelyn='D'"
                       '    sqlStr = sqlStr + " where idx=" & CStr(idx)
                       '     dbACADEMYget.Execute sqlStr
                       else
                            iResultCode = ""
                            iResultMsg  = ""
                            if (Not OneReceiptReqACA_KCP(idx, iResultCode, iResultMsg, iAuthCode)) then
                                infoMsg = infoMsg & " <font color='red'>���� ���� :" & "[" & iResultCode & "]" & iResultMsg & "</font><br>" & VbCrlf
                            else
                                infoMsg = infoMsg & " ���� ���� :" & "[" & iResultCode & "]" & iResultMsg & "<br>" & VbCrlf
        
                                
                                    sqlStr = " update [db_academy].[dbo].tbl_academy_order_master" & VbCrlf
                                    sqlStr = sqlStr & " set authcode = (case when accountdiv in ('7', '20') then '" & iAuthCode & "' else authcode end) " & VbCrlf
                                    if (reg_num="0100001234") then
                                        sqlStr = sqlStr & " ,cashreceiptreq='J'" & VbCrlf   '' �����߱� 2016/06/22
                                    else
                                        sqlStr = sqlStr & " ,cashreceiptreq='S'" & VbCrlf   
                                    end if
                                    sqlStr = sqlStr & " where orderserial='" & orderserial & "'"
        
                                    dbACADEMYget.Execute sqlStr
                                    
                                    retCnt = retCnt +1
                            end if
                       end if
                    end if
                else
                    infoMsg = infoMsg&"S_ERR|No idx"
                end if
            next
            infoMsg = infoMsg&"S_OK|"&retCnt
        else
            infoMsg = "S_NONE"
        end if
        
        response.write infoMsg
    Case ELSE
        response.Write "S_ERR|Not Valid - "&act
End Select
%>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
