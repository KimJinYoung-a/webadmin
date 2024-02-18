<%
    '/* ============================================================================== */
    '/* =   PAGE : 결제 정보 환경 설정 PAGE                                          = */
    '/* = -------------------------------------------------------------------------- = */
    '/* =   연동시 오류가 발생하는 경우 아래의 주소로 접속하셔서 확인하시기 바랍니다.= */
    '/* =   접속 주소 : http://testpay.kcp.co.kr/pgsample/FAQ/search_error.jsp       = */
    '/* = -------------------------------------------------------------------------- = */
    '/* =   Copyright (c)  2009   KCP Inc.   All Rights Reserved.                    = */
    '/* ============================================================================== */


    '/* ============================================================================== */
    '/* =   01. 지불 데이터 셋업 (업체에 맞게 수정)                                  = */
    '/* = -------------------------------------------------------------------------- = */
    '/* = ※ 주의 ※                                                                 = */
    '/* = * g_conf_key_dir 변수 설정                                                 = */
    '/* = pub.key 파일의 절대 경로 설정(파일명을 포함한 경로로 설정)                 = */
    '/* =                                                                            = */
    '/* = * g_conf_log_dir 변수 설정                                                 = */
    '/* = log 디렉토리 설정                                                          = */
    '/* = -------------------------------------------------------------------------- = */

    DIM g_conf_key_dir : g_conf_key_dir   = "C:\KCP_AX_HUB\bin\pub.key"
    DIM g_conf_log_dir : g_conf_log_dir   = "C:\KCP_AX_HUB\log"


    '/* ============================================================================== */
    '/* =   02. 쇼핑몰 지불 정보 설정                                                = */
    '/* = -------------------------------------------------------------------------- = */

    '/* = -------------------------------------------------------------------------- = */
    '/* =     01-1. 쇼핑몰 지불 필수 정보 설정(업체에 맞게 수정)                     = */
    '/* = -------------------------------------------------------------------------- = */
    '/* = ※ 주의 ※                                                                 = */
    '/* = * g_conf_js_path 설정                                                      = */
	'/* = 테스트 시 : src="http://pay.kcp.co.kr/plugin/payplus_test.js"              = */
	'/* =             src="https://pay.kcp.co.kr/plugin/payplus_test.js"             = */
    '/* = 실결제 시 : src="http://pay.kcp.co.kr/plugin/payplus.js"                   = */
	'/* =             src="https://pay.kcp.co.kr/plugin/payplus.js"                  = */
    '/* =                                                                            = */
	'/* = 테스트 시(UTF-8) : src="http://pay.kcp.co.kr/plugin/payplus_test_un.js"    = */
	'/* =                    src="https://pay.kcp.co.kr/plugin/payplus_test_un.js"   = */
    '/* = 실결제 시(UTF-8) : src="http://pay.kcp.co.kr/plugin/payplus_un.js"         = */
	'/* =                    src="https://pay.kcp.co.kr/plugin/payplus_un.js"        = */
    '/* =                                                                            = */
	'/* =                                                                            = */
    '/* = * g_conf_site_cd, g_conf_site_key 설정                                     = */
    '/* = 실결제시 KCP에서 발급한 사이트코드(site_cd), 사이트키(site_key)를 반드시   = */
    '/* =   변경해 주셔야 결제가 정상적으로 진행됩니다.                              = */
    '/* =                                                                            = */
    '/* = 테스트 시 : 사이트코드(T0000)와 사이트키(3grptw1.zW0GSo4PQdaGvsF__)로      = */
    '/* =            설정해 주십시오.                                                = */
    '/* = 실결제 시 : 반드시 KCP에서 발급한 사이트코드(site_cd)와 사이트키(site_key) = */
    '/* =            로 설정해 주십시오.                                             = */
    '/* =                                                                            = */
    '/* =                                                                            = */
    '/* = * g_conf_site_name 설정                                                    = */
    '/* = 사이트명 설정(한글 불가) : Payplus Plugin에서 상점명 및 오른쪽 상단에      = */
	'/* =                            표기되는 값입니다.                              = */
    '/* =                            반드시 영문자로 설정하여 주시기 바랍니다.       = */
    '/* = -------------------------------------------------------------------------- = */
    
    ''''테스트
    DIM g_conf_gw_url,g_conf_js_url,g_conf_js_url_ssl
    DIM g_conf_site_cd,g_conf_site_key,g_conf_site_name
    DIM g_conf_log_level, g_conf_module_type            ''2016/08/16 추가
    
    IF (application("Svr_Info")="Dev") THEN
        g_conf_gw_url    = "testpaygw.kcp.co.kr"
        g_conf_js_url    = "http://pay.kcp.co.kr/plugin/payplus_test.js"
        g_conf_js_url_ssl = "https://pay.kcp.co.kr/plugin/payplus_test.js"
    
        g_conf_site_cd   = "T0000"
        g_conf_site_key  = "3grptw1.zW0GSo4PQdaGvsF__"
        g_conf_site_name = "KCP TEST SHOP"
        
        g_conf_log_level    = "3"
        g_conf_module_type  = "01"
    ELSE
    '''REAL
        g_conf_gw_url    = "paygw.kcp.co.kr"
        g_conf_js_url    = "http://pay.kcp.co.kr/plugin/payplus.js"
        g_conf_js_url_ssl    = "https://pay.kcp.co.kr/plugin/payplus.js"
    
        g_conf_site_cd   = "R5523"
        g_conf_site_key  = "3uL-Drm.tpQ9q1yRqwWLSQF__"
        g_conf_site_name = "핑거스"
        
        g_conf_log_level    = "3"
        g_conf_module_type  = "01"
    END IF
    '/* ============================================================================== */


    '/* = -------------------------------------------------------------------------- = */
    '/* =     01-2. 지불 데이터 셋업 (변경 불가)                                     = */
    '/* = -------------------------------------------------------------------------- = */

    DIM g_conf_gw_port : g_conf_gw_port   = "8090"        ' 포트번호(변경불가)

    '/* ============================================================================== */
%>
