<%
'����,���� �� �� ���
Const KMPAY_WEB_SERVER_URL = "https://kmpay.lgcns.com:8443"
Const KMPAY_CERT_SERVER_URL = "https://kmpay.lgcns.com:8443"
Const KMPAY_CERT_SERVER_PAGE = "/merchant/requestDealApprove.dev"
Const CNSPAY_DEAL_REQUEST_URL = "https://pg.cnspay.co.kr:443"

'TXN_ID ȣ������ Ű��
DIM KMPAY_MERCHANT_ID : KMPAY_MERCHANT_ID= "KCTEN0001m"
DIM KMPAY_MERCHANT_ENCKEY : KMPAY_MERCHANT_ENCKEY= "689dcff525811766"
DIM KMPAY_MERCHANT_HASHKEY : KMPAY_MERCHANT_HASHKEY= "7195bb5277817133"

'����������Ű (�� �ش� ������Ű�� �ٲ��ּ���)
DIM KMPAY_MERCHANT_KEY : KMPAY_MERCHANT_KEY= "+f41Tt/PW/EtEuUEf0kYBpK4D41Jtpt3l9LP+Fke1wYS8nv1KL9+N0b8qnHecxMmyIMCasagGApBNAGSCuCuHw=="

'��� ��й�ȣ
Dim KMPAY_CANCEL_PWD : KMPAY_CANCEL_PWD= "KCTEN0001" ''�ٸ��ɷ� �Ϲ� ��ҽ� �ȵ�?..

if (application("Svr_Info")="Dev") then
    KMPAY_MERCHANT_ID = "cnstest25m"
    KMPAY_MERCHANT_ENCKEY = "10a3189211e1dfc6"
    KMPAY_MERCHANT_HASHKEY = "10a3189211e1dfc6"
    
    '����������Ű (�� �ش� ������Ű�� �ٲ��ּ���)
    KMPAY_MERCHANT_KEY = "33F49GnCMS1mFYlGXisbUDzVf2ATWCl9k3R++d5hDd3Frmuos/XLx8XhXpe+LDYAbpGKZYSwtlyyLOtS/8aD7A=="
    
    '��� ��й�ȣ
    KMPAY_CANCEL_PWD = "123456" ''�ٸ��ɷ� �Ϲ� ��ҽ� �ȵ�?..
end if

'�α�
Const KMPAY_LOG_DIR = "C:/KMPay/Log/"
Const KMPAY_LOG_LEVEL = 2   '-1:�α� ��� ����, 0:Error, 1:Info, 2:Debug  //
%>