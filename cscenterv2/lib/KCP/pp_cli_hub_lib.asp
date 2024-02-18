<%

function fnKCPCancelProc(byVal isAllCancel,byVal tno,byVal mod_desc, byVal mod_mny,byVal rem_mny,byRef res_cd,byRef res_msg,byRef CancelDate,byRef CancelTime,byRef ret_amount,byRef ret_panc_mod_mny,byRef ret_panc_rem_mny)

    Dim c_Mesg, c_Payplus, tran_cd
    Dim mod_data, mod_type, resp_mesg, canc_time
    Dim cust_ip : cust_ip = Request.ServerVariables( "REMOTE_ADDR" )
    
    if (isAllCancel) then
        mod_type = "STSC"
    else
        mod_type = "STPC"
    end if
    
    Set c_Mesg = New c_PayPlusData             ' 전문처리용 Class (library에서 정의됨)
    c_Mesg.InitialTX

    Set c_Payplus = Server.CreateObject( "pp_cli_com.KCP" )
    c_Payplus.lf_PP_CLI_LIB__init g_conf_key_dir, g_conf_log_dir, "03", g_conf_gw_url, g_conf_gw_port
    
    tran_cd = "00200000" ''업무구분

    mod_data = c_Mesg.mf_set_modx_data( "tno",        tno          )  ' KCP 원거래 거래번호
    mod_data = c_Mesg.mf_set_modx_data( "mod_type",   mod_type    )  ' 전체취소 STSC / 부분취소 STPC 
    mod_data = c_Mesg.mf_set_modx_data( "mod_ip",     cust_ip     )  ' 변경 요청자 IP
    mod_data = c_Mesg.mf_set_modx_data( "mod_desc",   mod_desc          )  ' 변경 사유

    if (mod_type = "STPC") then ' 부분취소의 경우
        mod_data = c_Mesg.mf_set_modx_data( "mod_mny", mod_mny )  ' 취소요청금액
        mod_data = c_Mesg.mf_set_modx_data( "rem_mny", rem_mny )  ' 취소가능잔액
        '복합거래 부분 취소시 주석을 풀어 주시기 바랍니다.
        'mod_data = c_Mesg.mf_set_modx_data( "tax_flag",     "TG03"          )  ' 복합과세 구분
        'mod_data = c_Mesg.mf_set_modx_data( "mod_tax_mny",  mod_tax_mny     )  ' 공급가 부분 취소 요청 금액
        'mod_data = c_Mesg.mf_set_modx_data( "mod_vat_mny",  mod_vat_mny     )  ' 부과세 부분 취소 요청 금액
        'mod_data = c_Mesg.mf_set_modx_data( "mod_free_mny", mod_free_mny    )  ' 비과세 부분 취소 요청 금액
    End If

    c_PayPlus.lf_PP_CLI_LIB__set_plan_data mod_data
    
    '' 실행
    c_Payplus.lf_PP_CLI_LIB__do_tx g_conf_site_cd, g_conf_site_key, tran_cd, "", ""
    resp_mesg = c_Payplus.lf_PP_CLI_LIB__get_data
    c_Mesg.mf_set_res_data(resp_mesg)                               ' 응답 전문 처리
    
    res_cd  = c_Mesg.mf_get_data ("res_cd")                         ' 결과 코드
    res_msg = c_Mesg.mf_get_data ("res_msg")                        ' 결과 메시지
    canc_time = c_Mesg.mf_get_data ("app_time")                     ' 취소일시 canc_time=>app_time
    
    if (res_cd="0000") then res_cd="00"               ''기존 코드와 맞추기 위함.
    
    if (mod_type = "STPC") then ' 부분취소의 경우
        ret_amount          = c_Mesg.mf_get_data ("amount")             '' 원거래결제금액
        ret_panc_mod_mny    = c_Mesg.mf_get_data ("panc_mod_mny")       '' 부분취소된금액
        ret_panc_rem_mny    = c_Mesg.mf_get_data ("panc_rem_mny")       '' 부분취소후잔액
    end if
    
    CancelDate	= LEFT(canc_time,4) & "년 " & MID(canc_time,5,2) & "월 " & MID(canc_time,7,2) & "일"
	CancelTime	= MID(canc_time,9,2) & "시 " & MID(canc_time,11,2) & "분 " & MID(canc_time,13,2) & "초"
	
	
	IF (application("Svr_Info")	= "Dev") then
	    response.write resp_mesg
	end if
	''response.write resp_mesg  ''신용카드 전체
	''res_cd=0000|res_msg=정상처리|res_en_msg=processing completed|trace_no=T00008ILHhFmoqMp|pay_method=PACA|
	''order_no=Y6081878263|card_cd=CCSG|card_name=씨티카드|acqu_cd=CCBC|acqu_name=BC카드|mcht_taxno=1138521083|
	''mall_taxno=1138521083|ca_order_id=Y6081878263|tno=16861900617603|amount=5000|card_mny=5000|coupon_mny=0|
	''escw_yn=N|canc_gubn=B|van_cd=VNKC|app_time=20160818211753|van_apptime=20160818211753|canc_time=20160818211753|
	''app_no=97722442|bizx_numb=725479023|quota=00|noinf=N|pg_txid=0818211753MP01ACGMT7YJ00000000500000977224420000
    
    ''신용카드 부분취소
    ''res_cd=0000|res_msg=정상처리|res_en_msg=processing completed
    ''|tno=16869900661952|amount=27900|card_mny=27900|coupon_mny=0|panc_mod_mny=22500|panc_card_mod_mny=22500|panc_coupon_mod_mny=0
    ''|panc_rem_mny=5400|mod_seq_no=201608262116108|mod_pcan_seq_no=16082600006458|escw_yn=N|van_cd=VNKC|app_time=20160826112321|
    ''van_apptime=20160826112321|canc_time=20160826133728|app_no=24434303|bizx_numb=880085676|quota=00|noinf=N|
    ''pg_txid=0826112321MP34AES4ALZE0000000279000024434303|card_cd=CCDI|card_name=현대카드|acqu_cd=CCDI|acqu_name=현대카드|mcht_taxno=1138521083|mall_taxno=1138521083|ca_order_id=Y6082678301


    '' response.write resp_mesg  ''실시간이체 전체
    ''res_cd=0000|res_msg=정상처리|res_en_msg=processing completed|trace_no=|pay_method=PABK|order_no=Y6081878271|
    ''bank_code=BK03|bank_name=기업은행|bank_com_type=0|bank_com_code=03|tno=20160818926266|bk_tid=0367385|amount=9500|bk_mny=9500|coupon_mny=0|
    ''app_time=20160818213856|mod_seq_no=20160818101431|mod_time=20160818213925 
    
    

end function

   
  '/* ============================================================================== */
  '/* =   PAGE : 라이브버리 PAGE                                                   = */
  '/* = -------------------------------------------------------------------------- = */
  '/* =   Copyright (c)  2016  NHN KCP Inc.   All Rights Reserved.                 = */
  '/* ============================================================================== */

  '/* ============================================================================== */
  '/* =   지불 연동 CLASS                                                          = */
  '/* ============================================================================== */

    Class c_PayPlusData

    '/* -------------------------------------------------------------------- */
    '/* -   처리 결과 값                                                   - */
    '/* -------------------------------------------------------------------- */
        Dim m_retData
        Dim arrData
        Dim arrRetData
        Dim arrDataList()

    '/* -------------------------------------------------------------------- */
    '/* -   초기화                                                         - */
    '/* -------------------------------------------------------------------- */
        Function InitialTX()

            m_retData   = ""
            arrData     = ""
            arrRetData  = ""

        End Function

    '/* -------------------------------------------------------------------- */
    '/* -   MOD DATA 전문 구성                                             - */
    '/* -------------------------------------------------------------------- */
        Function mf_set_modx_data(name,value)

            if isnull(m_retData) or m_retData = "" then
                m_retData = "mod_data="
                
            end if

            if value <> "" then
                m_retData = m_retData & name & "=" & value & chr(31)

                mf_set_modx_data = m_retData
                
            end if

            mf_set_modx_data = m_retData

        End Function

	'/* -------------------------------------------------------------------- */
    '/* -   ORDER DATA 전문 구성                                           - */
    '/* -------------------------------------------------------------------- */
        Function mf_set_ordr_data( name, value )

            if m_retData = "" and value <> "" then
                m_retData = "ordr_data="
                m_retData = m_retData & name & "=" & value
                mf_set_ordr_data = m_retData
            else
                if value <> "" then
                    m_retData = m_retData & chr(31) & name & "=" & value
                    mf_set_ordr_data = m_retData
                end if
            end if

            mf_set_ordr_data = m_retData

        End Function

	'/* -------------------------------------------------------------------- */
    '/* -   REQUEST DATA 전문 구성                                         - */
    '/* -------------------------------------------------------------------- */
        Function mf_set_req_data( value )

            if m_retData = "" and value <> "" then
                m_retData = value
                mf_set_req_data = m_retData
            else
                if value <> "" then
                    m_retData = m_retData & chr(28) & value
                    mf_set_req_data = m_retData
                end if
            end if

            mf_set_req_data = m_retData

        End Function

    '/* -------------------------------------------------------------------- */
    '/* -   RESULT DATA PARSING                                            - */
    '/* -------------------------------------------------------------------- */
        Function mf_set_res_data(name)
            Dim k
            Dim i,j
            
            k = 0
            Redim arrDataList(k+1)
            arrData = Split(name,chr(31))
            
            for i=0 to Ubound(arrData)
                arrRetData = Split(arrData(i),"=")

                for j=0 to Ubound(arrRetData)
                    Redim preserve arrDataList(k+1)
                    arrDataList(k) = Trim(arrRetData(j))
                    k = k+1
                next

            next

        End Function

    '/* -------------------------------------------------------------------- */
    '/* -   RESULT DATA 전문 구성                                          - */
    '/* -------------------------------------------------------------------- */
        Function mf_get_data(name)
            Dim i
            for i=0 to Ubound(arrDataList)
                if StrComp(name,arrDataList(i)) = 0 then
                    mf_get_data = arrDataList(i+1)
                end if
            next

        End Function

    End Class
%>