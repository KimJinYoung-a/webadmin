<%
    '/* ============================================================================== */
    '/* =   PAGE : ���� ���� ȯ�� ���� PAGE                                          = */
    '/* = -------------------------------------------------------------------------- = */
    '/* =   ������ ������ �߻��ϴ� ��� �Ʒ��� �ּҷ� �����ϼż� Ȯ���Ͻñ� �ٶ��ϴ�.= */
    '/* =   ���� �ּ� : http://testpay.kcp.co.kr/pgsample/FAQ/search_error.jsp       = */
    '/* = -------------------------------------------------------------------------- = */
    '/* =   Copyright (c)  2009   KCP Inc.   All Rights Reserved.                    = */
    '/* ============================================================================== */


    '/* ============================================================================== */
    '/* =   01. ���� ������ �¾� (��ü�� �°� ����)                                  = */
    '/* = -------------------------------------------------------------------------- = */
    '/* = �� ���� ��                                                                 = */
    '/* = * g_conf_key_dir ���� ����                                                 = */
    '/* = pub.key ������ ���� ��� ����(���ϸ��� ������ ��η� ����)                 = */
    '/* =                                                                            = */
    '/* = * g_conf_log_dir ���� ����                                                 = */
    '/* = log ���丮 ����                                                          = */
    '/* = -------------------------------------------------------------------------- = */

    DIM g_conf_key_dir : g_conf_key_dir   = "C:\KCP_AX_HUB\bin\pub.key"
    DIM g_conf_log_dir : g_conf_log_dir   = "C:\KCP_AX_HUB\log"


    '/* ============================================================================== */
    '/* =   02. ���θ� ���� ���� ����                                                = */
    '/* = -------------------------------------------------------------------------- = */

    '/* = -------------------------------------------------------------------------- = */
    '/* =     01-1. ���θ� ���� �ʼ� ���� ����(��ü�� �°� ����)                     = */
    '/* = -------------------------------------------------------------------------- = */
    '/* = �� ���� ��                                                                 = */
    '/* = * g_conf_js_path ����                                                      = */
	'/* = �׽�Ʈ �� : src="http://pay.kcp.co.kr/plugin/payplus_test.js"              = */
	'/* =             src="https://pay.kcp.co.kr/plugin/payplus_test.js"             = */
    '/* = �ǰ��� �� : src="http://pay.kcp.co.kr/plugin/payplus.js"                   = */
	'/* =             src="https://pay.kcp.co.kr/plugin/payplus.js"                  = */
    '/* =                                                                            = */
	'/* = �׽�Ʈ ��(UTF-8) : src="http://pay.kcp.co.kr/plugin/payplus_test_un.js"    = */
	'/* =                    src="https://pay.kcp.co.kr/plugin/payplus_test_un.js"   = */
    '/* = �ǰ��� ��(UTF-8) : src="http://pay.kcp.co.kr/plugin/payplus_un.js"         = */
	'/* =                    src="https://pay.kcp.co.kr/plugin/payplus_un.js"        = */
    '/* =                                                                            = */
	'/* =                                                                            = */
    '/* = * g_conf_site_cd, g_conf_site_key ����                                     = */
    '/* = �ǰ����� KCP���� �߱��� ����Ʈ�ڵ�(site_cd), ����ƮŰ(site_key)�� �ݵ��   = */
    '/* =   ������ �ּž� ������ ���������� ����˴ϴ�.                              = */
    '/* =                                                                            = */
    '/* = �׽�Ʈ �� : ����Ʈ�ڵ�(T0000)�� ����ƮŰ(3grptw1.zW0GSo4PQdaGvsF__)��      = */
    '/* =            ������ �ֽʽÿ�.                                                = */
    '/* = �ǰ��� �� : �ݵ�� KCP���� �߱��� ����Ʈ�ڵ�(site_cd)�� ����ƮŰ(site_key) = */
    '/* =            �� ������ �ֽʽÿ�.                                             = */
    '/* =                                                                            = */
    '/* =                                                                            = */
    '/* = * g_conf_site_name ����                                                    = */
    '/* = ����Ʈ�� ����(�ѱ� �Ұ�) : Payplus Plugin���� ������ �� ������ ��ܿ�      = */
	'/* =                            ǥ��Ǵ� ���Դϴ�.                              = */
    '/* =                            �ݵ�� �����ڷ� �����Ͽ� �ֽñ� �ٶ��ϴ�.       = */
    '/* = -------------------------------------------------------------------------- = */
    
    ''''�׽�Ʈ
    DIM g_conf_gw_url,g_conf_js_url,g_conf_js_url_ssl
    DIM g_conf_site_cd,g_conf_site_key,g_conf_site_name
    DIM g_conf_log_level, g_conf_module_type            ''2016/08/16 �߰�
    
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
        g_conf_site_name = "�ΰŽ�"
        
        g_conf_log_level    = "3"
        g_conf_module_type  = "01"
    END IF
    '/* ============================================================================== */


    '/* = -------------------------------------------------------------------------- = */
    '/* =     01-2. ���� ������ �¾� (���� �Ұ�)                                     = */
    '/* = -------------------------------------------------------------------------- = */

    DIM g_conf_gw_port : g_conf_gw_port   = "8090"        ' ��Ʈ��ȣ(����Ұ�)

    '/* ============================================================================== */
%>
