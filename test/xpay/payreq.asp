<!-- #include file="md5.asp" -->
<%
    '/*
    ' * [���� ������û ������(STEP2-1)]
    ' *
    ' * ���������������� �⺻ �Ķ���͸� ���õǾ� ������, ������ �ʿ��Ͻ� �Ķ���ʹ� �����޴����� �����Ͻþ� �߰� �Ͻñ� �ٶ��ϴ�.
    ' */

    '/*
    ' * 1. �⺻���� ������û ���� ����
    ' *
    ' * �⺻������ �����Ͽ� �ֽñ� �ٶ��ϴ�.(�Ķ���� ���޽� POST�� ����ϼ���)
    ' */
    CST_PLATFORM               = trim(request("CST_PLATFORM"))       'LG������ ���� ���� ����(test:�׽�Ʈ, service:����)
    CST_MID                    = trim(request("CST_MID"))            '�������̵�(LG���������� ���� �߱޹����� �������̵� �Է��ϼ���)
                                                                     '�׽�Ʈ ���̵�� 't'�� �ݵ�� �����ϰ� �Է��ϼ���.
    if CST_PLATFORM = "test" then                                    '�������̵�(�ڵ�����)
        LGD_MID = "t" & CST_MID
    else
        LGD_MID = CST_MID
    end if
    LGD_OID                    = trim(request("LGD_OID"))            '�ֹ���ȣ(�������� ����ũ�� �ֹ���ȣ�� �Է��ϼ���)
    LGD_AMOUNT                 = trim(request("LGD_AMOUNT"))         '�����ݾ�("," �� ������ �����ݾ��� �Է��ϼ���)
    LGD_MERTKEY                = trim(request("LGD_MERTKEY"))        '[�ݵ�� ����]����MertKey(mertkey�� ���������� -> ������� -> ���������������� Ȯ���ϽǼ� �ֽ��ϴ�')
    LGD_BUYER                  = trim(request("LGD_BUYER"))          '�����ڸ�
    LGD_PRODUCTINFO            = trim(request("LGD_PRODUCTINFO"))    '��ǰ��
    LGD_BUYEREMAIL             = trim(request("LGD_BUYEREMAIL"))     '������ �̸���
    LGD_TIMESTAMP              = year(now) & right("0" & month(now),2) & right("0" & day(now),2) & right("0" & hour(now),2) & right("0" & minute(now),2) & right("0" & second(now),2) 'Ÿ�ӽ�����
    LGD_CUSTOM_SKIN            = "red"                               '�������� ����â ��Ų (red, blue, cyan, green, yellow)
 
    '/*
    ' * �������(������) ���� ������ �Ͻô� ��� �Ʒ� LGD_CASNOTEURL �� �����Ͽ� �ֽñ� �ٶ��ϴ�.
    ' */
    ''LGD_CASNOTEURL             = "http://����URL/cas_noteurl.asp"

    '/*
    ' *************************************************
    ' * 2. MD5 �ؽ���ȣȭ (�������� ������) - BEGIN
    ' *
    ' * MD5 �ؽ���ȣȭ�� �ŷ� �������� �������� ����Դϴ�.
    ' *************************************************
    ' *
    ' * �ؽ� ��ȣȭ ����( LGD_MID + LGD_OID + LGD_AMOUNT + LGD_TIMESTAMP + LGD_MERTKEY )
    ' * LGD_MID          : �������̵�
    ' * LGD_OID          : �ֹ���ȣ
    ' * LGD_AMOUNT       : �ݾ�
    ' * LGD_TIMESTAMP    : Ÿ�ӽ�����
    ' * LGD_MERTKEY      : ����MertKey (mertkey�� ���������� -> ������� -> ���������������� Ȯ���ϽǼ� �ֽ��ϴ�)
    ' *
    ' * MD5 �ؽ������� ��ȣȭ ������ ����
    ' * LG�����޿��� �߱��� ����Ű(MertKey)�� ȯ�漳�� ����(lgdacom/conf/mall.conf)�� �ݵ�� �Է��Ͽ� �ֽñ� �ٶ��ϴ�.
    ' */
    LGD_HASHDATA = md5( LGD_MID & LGD_OID & LGD_AMOUNT & LGD_TIMESTAMP & LGD_MERTKEY )
    LGD_CUSTOM_PROCESSTYPE = "TWOTR"
    '/*
    ' *************************************************
    ' * 2. MD5 �ؽ���ȣȭ (�������� ������) - END
    ' *************************************************
    ' */
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<title>������ eCredit���� �����׽�Ʈ</title>

<script type="text/javascript">
<!--
/*
 * �������� ������û�� PAYKEY�� �޾Ƽ� �������� ��û.
 */
function doPay_ActiveX(){
    ret = xpay_check(document.getElementById('LGD_PAYINFO'), '<%= CST_PLATFORM %>');

    if (ret=="00"){     //ActiveX �ε� ����
        var LGD_RESPCODE        = dpop.getData('LGD_RESPCODE');       //����ڵ�
        var LGD_RESPMSG         = dpop.getData('LGD_RESPMSG');        //����޼���

        if( "0000" == LGD_RESPCODE ) { //��������
            var LGD_PAYKEY      = dpop.getData('LGD_PAYKEY');         //LG������ ����KEY
            var msg = "������� : " + LGD_RESPMSG + "\n";
            msg += "LGD_PAYKEY : " + LGD_PAYKEY +"\n\n";
            document.getElementById('LGD_PAYKEY').value = LGD_PAYKEY;
            alert(msg);
            document.getElementById('LGD_PAYINFO').submit();
        } else { //��������
            alert("������ �����Ͽ����ϴ�. " + LGD_RESPMSG);
            /*
             * �������� ȭ�� ó��
             */
        }
    } else {
        alert("LG������ ���ڰ����� ���� ActiveX ��ġ ����");
        /*
         * �������� ȭ�� ó��
         */
    }
}

function isActiveXOK(){
	if(lgdacom_atx_flag == true){
    	document.getElementById('LGD_BUTTON1').style.display='none';
        document.getElementById('LGD_BUTTON2').style.display='';
	}else{
		document.getElementById('LGD_BUTTON1').style.display='';
        document.getElementById('LGD_BUTTON2').style.display='none';	
	}
}
-->
</script>
</head>

</head>
<body onload="isActiveXOK();">
<div id="LGD_ACTIVEX_DIV"/> <!-- ActiveX ��ġ �ȳ� Layer �Դϴ�. �������� ������. -->
<form method="post" id="LGD_PAYINFO" action="payres.asp">
<table>
    <tr>
        <td>������ �̸� </td>
        <td><%= LGD_BUYER %></td>
    </tr>
    <tr>
        <td>��ǰ���� </td>
        <td><%= LGD_PRODUCTINFO %></td>
    </tr>
    <tr>
        <td>�����ݾ� </td>
        <td><%= LGD_AMOUNT %></td>
    </tr>
    <tr>
        <td>������ �̸��� </td>
        <td><%= LGD_BUYEREMAIL %></td>
    </tr>
    <tr>
        <td>�ֹ���ȣ </td>
        <td><%= LGD_OID %></td>
    </tr>
    <tr>
        <td colspan="2">* �߰� �� ������û �Ķ���ʹ� �޴����� ������ �ֽñ� �ٶ��ϴ�.</td>
    </tr>
    <tr>
        <td colspan="2"></td>
    </tr>    
    <tr>
        <td colspan="2">
		<div id="LGD_BUTTON1">������ ���� ����� �ٿ� ���̰ų�, ����� ��ġ���� �ʾҽ��ϴ�. </div>
		<div id="LGD_BUTTON2" style="display:none"><input type="button" value="������û" onclick="doPay_ActiveX();"/> </div>        
        </td>
    </tr>    
</table>
<br>

<br>
<input type="hidden" name="CST_PLATFORM"                value="<%= CST_PLATFORM %>">                   <!-- �׽�Ʈ, ���� ���� -->
<input type="hidden" name="CST_MID"                     value="<%= CST_MID %>">                        <!-- �������̵� -->
<input type="hidden" name="LGD_MID"                     value="<%= LGD_MID %>">                        <!-- �������̵� -->
<input type="hidden" name="LGD_OID"                     value="<%= LGD_OID %>">                        <!-- �ֹ���ȣ -->
<input type="hidden" name="LGD_BUYER"                   value="<%= LGD_BUYER %>">                      <!-- ������ -->
<input type="hidden" name="LGD_PRODUCTINFO"             value="<%= LGD_PRODUCTINFO %>">                <!-- ��ǰ���� -->
<input type="hidden" name="LGD_AMOUNT"                  value="<%= LGD_AMOUNT %>">                     <!-- �����ݾ� -->
<input type="hidden" name="LGD_BUYEREMAIL"              value="<%= LGD_BUYEREMAIL %>">                 <!-- ������ �̸��� -->
<input type="hidden" name="LGD_CUSTOM_SKIN"             value="<%= LGD_CUSTOM_SKIN %>">                <!-- ����â SKIN -->
<input type="hidden" name="LGD_CUSTOM_PROCESSTYPE"      value="<%= LGD_CUSTOM_PROCESSTYPE %>">         <!-- Ʈ����� ó����� -->
<input type="hidden" name="LGD_TIMESTAMP"               value="<%= LGD_TIMESTAMP %>">                  <!-- Ÿ�ӽ����� -->
<input type="hidden" name="LGD_HASHDATA"                value="<%= LGD_HASHDATA %>">                   <!-- MD5 �ؽ���ȣ�� -->
<input type="hidden" name="LGD_PAYKEY"                  value="">                                      <!-- LG������ PAYKEY(������ �ڵ�����)-->
<input type="hidden" name="LGD_VERSION"         		value="ASP_XPay_1.0">						   <!-- �������� (�������� ������) -->
<!-- �������(������) ���������� �Ͻô� ���  �Ҵ�/�Ա� ����� �뺸�ޱ� ���� �ݵ�� LGD_CASNOTEURL ������ LG �����޿� �����ؾ� �մϴ� . -->
<!--input type="hidden" name="LGD_CASNOTEURL"           value="<%= LGD_CASNOTEURL %>"-->                 <!-- ������� NOTEURL -->

<!--
<input type="hidden" name="LGD_CUSTOM_USABLEPAY"        value="SC0010">	
-->
</form>
</body>
<!--  xpay.js�� �ݵ��  body �ؿ� �νñ� �ٶ��ϴ�. -->
<!--  UTF-8 ���ڵ� ��� �ô� xpay.js ��� xpay_utf-8.js ��  ȣ���Ͻñ� �ٶ��ϴ�.-->
<%
     protocol = "http"
     If request.serverVariables("SERVER_PORT") = "443" Then protocol = "https"

     if CST_PLATFORM = "test" then
     	port = "7080"
     	If request.serverVariables("SERVER_PORT") = "443" Then port = "7443"
        Response.Write "<script language='javascript' src='"& protocol &"://xpay.lgdacom.net:" & port & "/xpay/js/xpay.js' type='text/javascript'>"
     else
        Response.Write "<script language='javascript' src='"& protocol &"://xpay.lgdacom.net/xpay/js/xpay.js' type='text/javascript'>"
     end if
%>
</script>
</html>
