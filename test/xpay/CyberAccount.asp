<%
    '/*
    ' * [������� �߱�/�����û ������]
    ' *
    ' * ������� �߱� ����(CHANGE)�� �ݾװ� �����ϸ� ���� �Ҽ� �ֽ��ϴ�. 
    ' */
    CST_PLATFORM         = trim(request("CST_PLATFORM"))         ' LG�ڷ��� �������� ����(test:�׽�Ʈ, service:����)
    CST_MID              = trim(request("CST_MID"))              ' LG�ڷ������� ���� �߱޹����� �������̵� �Է��ϼ���.
                                                                 ' �׽�Ʈ ���̵�� 't'�� �����ϰ� �Է��ϼ���.
    if CST_PLATFORM = "test" then                                ' �������̵�(�ڵ�����)
        LGD_MID = "t" & CST_MID
    else
        LGD_MID = CST_MID
    end if
    LGD_METHOD           = trim(request("LGD_METHOD"))           ' ASSIGN:�Ҵ�, CHANGE:����
    LGD_OID     		 = trim(request("LGD_OID"))    			 ' �ֹ���ȣ(�������� ����ũ�� �ֹ���ȣ�� �Է��ϼ���)
    LGD_AMOUNT      	 = trim(request("LGD_AMOUNT"))      	 ' �ݾ�("," �� ������ �ݾ��� �Է��ϼ���)
    LGD_PRODUCTINFO   	 = trim(request("LGD_PRODUCTINFO"))  	 ' ��ǰ����
    LGD_BUYER          	 = trim(request("LGD_BUYER"))         	 ' �����ڸ�
	LGD_ACCOUNTOWNER     = trim(request("LGD_ACCOUNTOWNER"))  	 ' �Ա��ڸ�
	LGD_ACCOUNTPID       = trim(request("LGD_ACCOUNTPID"))       ' �Ա����ֹι�ȣ(�ɼ�)
	LGD_BUYERPHONE       = trim(request("LGD_BUYERPHONE"))       ' �������޴�����ȣ
	LGD_BUYEREMAIL       = trim(request("LGD_BUYEREMAIL"))       ' �������̸���(�ɼ�)
	LGD_BANKCODE         = trim(request("LGD_BANKCODE"))         ' �Աݰ��������ڵ�
	LGD_CASHRECEIPTUSE   = trim(request("LGD_CASHRECEIPTUSE"))   ' ���ݿ����� ���౸��('1':�ҵ����, '2':��������)
	LGD_CASHCARDNUM      = trim(request("LGD_CASHCARDNUM"))      ' ���ݿ����� ī���ȣ
	LGD_CLOSEDATE        = trim(request("LGD_CLOSEDATE"))        ' �Ա� ������
	LGD_TAXFREEAMOUNT    = trim(request("LGD_TAXFREEAMOUNT"))    ' �鼼�ݾ�
	LGD_CASNOTEURL       = "http://webadmin.10x10.co.kr/admin/apps/DC_CA_noteurl.asp" ''"http://����URL/cas_noteurl.asp"       ' �Աݰ�� ó���� ���� ������������ �ݵ�� ������ �ּ���
	

    'configPath           = "C:/lgdacom"         				 ' LG�ڷ��޿��� ������ ȯ������("/conf/lgdacom.conf") ��ġ ����.
    configPath				   = "C:/lgdacom/conf/" & CST_MID

    Set xpay = server.CreateObject("XPayClientCOM.XPayClient")
    xpay.Init configPath, CST_PLATFORM
    xpay.Init_TX(LGD_MID)

    xpay.Set "LGD_TXNAME", "CyberAccount"
    xpay.Set "LGD_METHOD", LGD_METHOD
    xpay.Set "LGD_OID", LGD_OID
    xpay.Set "LGD_AMOUNT", LGD_AMOUNT
    xpay.Set "LGD_PRODUCTINFO", LGD_PRODUCTINFO
    xpay.Set "LGD_BUYER", LGD_BUYER
    xpay.Set "LGD_ACCOUNTOWNER", LGD_ACCOUNTOWNER
    xpay.Set "LGD_ACCOUNTPID", LGD_ACCOUNTPID
    xpay.Set "LGD_BUYERPHONE", LGD_BUYERPHONE
    xpay.Set "LGD_BUYEREMAIL", LGD_BUYEREMAIL
    xpay.Set "LGD_BANKCODE", LGD_BANKCODE
    xpay.Set "LGD_CASHRECEIPTUSE", LGD_CASHRECEIPTUSE
    xpay.Set "LGD_CASHCARDNUM", LGD_CASHCARDNUM
    xpay.Set "LGD_CLOSEDATE", LGD_CLOSEDATE
    xpay.Set "LGD_TAXFREEAMOUNT", LGD_TAXFREEAMOUNT
    xpay.Set "LGD_CASNOTEURL", LGD_CASNOTEURL
    

    '/*
    ' * 1. ������� �߱�/���� ��û ���ó��
    ' *
    ' * ��� ���� �Ķ���ʹ� �����޴����� �����Ͻñ� �ٶ��ϴ�.
    ' */
    if xpay.TX() then
        if LGD_METHOD = "ASSIGN" then      '������� �߱��� ���
        
        	'1)������� �߱ް�� ȭ��ó��(����,���� ��� ó���� �Ͻñ� �ٶ��ϴ�.)
        	Response.Write("������� �߱� ��ûó���� �Ϸ�Ǿ����ϴ�. <br>")
        	Response.Write("TX Response_code = " & xpay.resCode & "<br>")
        	Response.Write("TX Response_msg = " & xpay.resMsg & "<p>")
			
			Response.Write("����ڵ� : " & xpay.Response("LGD_RESPCODE", 0) & "<br>")
	    	Response.Write("�ŷ���ȣ : " & xpay.Response("LGD_TID", 0) & "<p>")
        	
        	'�Ʒ��� ������û ��� �Ķ���͸� ��� ��� �ݴϴ�.
        	Dim itemCount
        	Dim resCount
        	itemCount = xpay.resNameCount
        	resCount = xpay.resCount

        	For i = 0 To itemCount - 1
            	itemName = xpay.ResponseName(i)
            	Response.Write(itemName & "&nbsp:&nbsp")
            	For j = 0 To resCount - 1
                	Response.Write(xpay.Response(itemName, j) & "<br>")
            	Next
        	Next
        
        else		'������� ������ ���
        	'1)������� ������ ȭ��ó��(����,���� ��� ó���� �Ͻñ� �ٶ��ϴ�.)
        	Response.Write("������� ���� ��ûó���� �Ϸ�Ǿ����ϴ�. <br>")
        	Response.Write("����ڵ� : " & xpay.Response("LGD_RESPCODE", 0) & "<br>")
            Response.Write("�ֹ���ȣ : " & LGD_OID & "<br>")
            Response.Write("�Աݾ� : " & LGD_AMOUNT & "<br>")
        	Response.Write("�Աݸ����� : " & LGD_CLOSEDATE & "<p>")
            
            
        	itemCount = xpay.resNameCount
        	resCount = xpay.resCount

        	For i = 0 To itemCount - 1
            	itemName = xpay.ResponseName(i)
            	Response.Write(itemName & "&nbsp:&nbsp")
            	For j = 0 To resCount - 1
                	Response.Write(xpay.Response(itemName, j) & "<br>")
            	Next
        	Next
        	
        end if    
        
        Response.Write("<p>")
            
    else
        '2)API ��û ���� ȭ��ó��
        Response.Write("������� �߱�/���� ��ûó���� ���еǾ����ϴ�. <br>")
        Response.Write("TX Response_code = " & xpay.resCode & "<br>")
        Response.Write("TX Response_msg = " & xpay.resMsg & "<p>")
    end if
%>
