<%
	'��������
    CST_PLATFORM               = trim(request("CST_PLATFORM"))       'LG������ ���� ���� ����(test:�׽�Ʈ, service:����)
    CST_MID                    = trim(request("CST_MID"))            '�������̵�(LG���������� ���� �߱޹����� �������̵� �Է��ϼ���)
                                                                     '�׽�Ʈ ���̵�� 't'�� �ݵ�� �����ϰ� �Է��ϼ���.
    if CST_PLATFORM = "test" then                                    '�������̵�(�ڵ�����)
        LGD_MID = "t" & CST_MID
    else
        LGD_MID = CST_MID
    end if
    LGD_OID                 = trim(request("LGD_OID"))            	'�ֹ���ȣ(�������� ����ũ�� �ֹ���ȣ�� �Է��ϼ���)
    LGD_OID                 = "2010010100003"
    LGD_BUYER          		= trim(request("LGD_BUYER"))     		'�����ڸ�
    LGD_PRODUCTINFO         = trim(request("LGD_PRODUCTINFO"))     	'��ǰ����
    LGD_BUYEREMAIL          = trim(request("LGD_BUYEREMAIL"))     	'�̸����ּ�(���������� ���Ϲ߼�)
    LGD_AMOUNT              = trim(request("LGD_AMOUNT"))         	'�����ݾ�("," �� ������ �����ݾ��� �Է��ϼ���)
    LGD_AUTHTYPE			= trim(request("LGD_AUTHTYPE"))		 	'��������(ISP�ΰ�츸  'ISP')
    LGD_CARDTYPE			= trim(request("LGD_CARDTYPE"))			'ī����ڵ�
    
    '�Ƚ�Ŭ�� ���� �Ǵ� �ؿ�ī��
    LGD_PAN                 = trim(request("LGD_PAN"))            	'ī���ȣ    
    LGD_INSTALL             = trim(request("LGD_INSTALL"))        	'�Һΰ�����(���ڸ�����)
    LGD_NOINT				= trim(request("LGD_NOINT"))		    '�������Һο���('1':�����δ㹫�����Һ�,'0':�Ϲ��Һ�)
    LGD_EXPYEAR             = trim(request("LGD_EXPYEAR"))        	'��ȿ�Ⱓ��(YY)
   	LGD_EXPMON              = trim(request("LGD_EXPMON"))         	'��ȿ�Ⱓ��(MM)
    VBV_ECI             	= trim(request("VBV_ECI"))				'�Ƚ�Ŭ��ECI  
 	VBV_CAVV				= trim(request("VBV_CAVV"))			 	'�Ƚ�Ŭ��CAVV
 	VBV_XID				   	= trim(request("VBV_XID"))			 	'�Ƚ�Ŭ��XID    
    
    'ISP����
    KVP_QUOTA				= trim(request("KVP_QUOTA"))			'�Һΰ�����
    KVP_NOINT				= trim(request("KVP_NOINT"))			'�������Һο���('1':�����δ㹫�����Һ�,'0':�Ϲ��Һ�)
	KVP_CARDCODE			= trim(request("KVP_CARDCODE"))			'ISPī���ڵ�
	KVP_SESSIONKEY			= trim(request("KVP_SESSIONKEY"))		'ISP����Ű
	KVP_ENCDATA				= trim(request("KVP_ENCDATA"))		 	'ISP��ȣȭ������
	
    '' configPath				   = "C:/lgdacom"
    configPath				   = "C:/lgdacom/conf/" & CST_MID					 'LG�����޿��� ������ ȯ������(/conf/lgdacom.conf, /conf/mall.conf)�� ��ġ�� ���丮 ���� 
    
	Dim xpay
	Dim i, j
	Dim itemName
	
	Set xpay = server.CreateObject("XPayClientCOM.XPayClient")	
    xpay.Init configPath, CST_PLATFORM    
    xpay.Init_TX(LGD_MID)

    xpay.Set "LGD_TXNAME", "CardAuth"
    xpay.Set "LGD_OID", LGD_OID 
	xpay.Set "LGD_AMOUNT", LGD_AMOUNT
	xpay.Set "LGD_BUYER", LGD_BUYER
	xpay.Set "LGD_PRODUCTINFO", LGD_PRODUCTINFO
	xpay.Set "LGD_BUYEREMAIL", LGD_BUYEREMAIL
	xpay.Set "LGD_AUTHTYPE", LGD_AUTHTYPE
	xpay.Set "LGD_CARDTYPE", LGD_CARDTYPE
	xpay.Set "LGD_BUYERIP", Request.ServerVariables("REMOTE_ADDR")	'�ݵ�� �������� IP�� �Ѱܾ� ��
	
	if LGD_AUTHTYPE = "ISP" then
		xpay.Set "KVP_QUOTA", KVP_QUOTA
		xpay.Set "KVP_NOINT", KVP_NOINT
		xpay.Set "KVP_CARDCODE", KVP_CARDCODE
		xpay.Set "KVP_SESSIONKEY", KVP_SESSIONKEY
		xpay.Set "KVP_ENCDATA", KVP_ENCDATA 
	else
		xpay.Set "LGD_PAN", LGD_PAN
		xpay.Set "LGD_INSTALL", LGD_INSTALL
		xpay.Set "LGD_NOINT", LGD_NOINT
		xpay.Set "LGD_EXPYEAR", LGD_EXPYEAR
		xpay.Set "LGD_EXPMON", LGD_EXPMON
		xpay.Set "VBV_ECI", VBV_ECI
		xpay.Set "VBV_CAVV", VBV_CAVV
		xpay.Set "VBV_XID", VBV_XID
	end if 

    
    if  xpay.TX() then
        '1)������� ó��(����,���� ��� ó���� �Ͻñ� �ٶ��ϴ�.)
        Response.Write("������û�� �Ϸ�Ǿ����ϴ�. <br>")
        Response.Write("TX Response_code = " & xpay.resCode & "<br>")
        Response.Write("TX Response_msg = " & xpay.resMsg & "<p>")

	    Response.Write("�ŷ���ȣ : " & xpay.Response("LGD_TID", 0) & "<br>")
	    Response.Write("�������̵� : " & xpay.Response("LGD_MID", 0) & "<br>")
	    Response.Write("�����ֹ���ȣ : " & xpay.Response("LGD_OID", 0) & "<br>")
	    Response.Write("�����ݾ� : " & xpay.Response("LGD_AMOUNT", 0) & "<br>")
	    Response.Write("����ڵ� : " & xpay.Response("LGD_RESPCODE", 0) & "<br>")
	    Response.Write("����޼��� : " & xpay.Response("LGD_RESPMSG", 0) & "<p>")

        Response.Write("[������û ��� �Ķ����]<br>")

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
        
        Response.Write("<p>")
        
        if xpay.resCode = "0000" then
        	'����������û ��� ���� DBó��
        	Response.Write("����������û ��� ���� DBó���Ͻñ� �ٶ��ϴ�." & "<br>")
        	            	            	
        	'����������û ��� ���� DBó�� ���н� Rollback ó��
        	isDBOK = true 'DBó�� ���н� false�� ������ �ּ���.
        	
        	if isDBOK then
        	else
        		Response.Write("<p>")
        		xpay.Rollback("���� DBó�� ���з� ���Ͽ� Rollback ó�� [TID:" & xpay.Response("LGD_TID",0) & ",MID:" & xpay.Response("LGD_MID",0) & ",OID:" & xpay.Response("LGD_OID",0) & "]")
        		
                Response.Write("TX Rollback Response_code = " & xpay.resCode & "<br>")
                Response.Write("TX Rollback Response_msg = " & xpay.resMsg & "<p>")
        		
                if "0000" = xpay.resCode then
                	Response.Write("�ڵ���Ұ� ���������� �Ϸ� �Ǿ����ϴ�.<br>")
                else
                	Response.Write("�ڵ���Ұ� ���������� ó������ �ʾҽ��ϴ�.<br>")
                end if
        	end if            	
        else
        	'����������û ��� ���� DBó��
        	Response.Write("����������û ��� ���� DBó���Ͻñ� �ٶ��ϴ�." & "<br>")
        end if
    else
        '2)API ��û���� ó��
        Response.Write("������û�� �����Ͽ����ϴ�. <br>")
        Response.Write("TX Response_code = " & xpay.resCode & "<br>")
        Response.Write("TX Response_msg = " & xpay.resMsg & "<p>")
        
        '������û ��� ���� ���� DBó��
        Response.Write("����������û ��� ���� DBó���Ͻñ� �ٶ��ϴ�." & "<br>")
    end if 
%>
