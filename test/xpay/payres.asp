<%
    '/*
    ' * [����������û ������(STEP2-2)]
    ' *
    ' * LG���������� ���� �������� LGD_PAYKEY(����Key)�� ������ ���� ������û.(�Ķ���� ���޽� POST�� ����ϼ���)
    ' */

    'configPath = "C:/lgdacom/conf/" & "thefingers" & "/conf/" 'LG�����޿��� ������ ȯ������("/conf/lgdacom.conf, /conf/mall.conf") ��ġ ����.  
    ''configPath = "C:/lgdacom/conf/thefingers/"
    configPath = "C:/lgdacom/conf/youareagirl/"
    
    '/*
    ' *************************************************
    ' * 1.�������� ��û - BEGIN
    ' *  (��, ���� �ݾ�üũ�� ���Ͻô� ��� �ݾ�üũ �κ� �ּ��� ���� �Ͻø� �˴ϴ�.)
    ' *************************************************
    ' */
response.write "LGD_OID=" + request("LGD_OID")
response.write "LGD_HASHDATA=" + request("LGD_HASHDATA")

    CST_PLATFORM               = trim(request("CST_PLATFORM"))
    CST_MID                    = trim(request("CST_MID"))
    if CST_PLATFORM = "test" then
        LGD_MID = "t" & CST_MID
    else
        LGD_MID = CST_MID
    end if
    LGD_PAYKEY                 = trim(request("LGD_PAYKEY"))

    Dim xpay            '������û API ��ü
    Dim amount_check    '�ݾ׺� ���
    Dim i, j
    Dim itemName

	'�ش� API�� ����ϱ� ���� setup.exe �� ��ġ�ؾ� �մϴ�.
    Set xpay = server.CreateObject("XPayClientCOM.XPayClient")
    xpay.Init configPath, CST_PLATFORM

    xpay.Init_TX(LGD_MID)
    xpay.Set "LGD_TXNAME", "PaymentByKey"
    xpay.Set "LGD_PAYKEY", LGD_PAYKEY
    
response.write "LGD_PAYKEY="&LGD_PAYKEY
    
    '�ݾ��� üũ�Ͻñ� ���ϴ� ��� �Ʒ� �ּ��� Ǯ� �̿��Ͻʽÿ�.
	'DB_AMOUNT = "DB�� ���ǿ��� ������ �ݾ�" 	'�ݵ�� �������� �Ұ����� ��(DB�� ����)���� �ݾ��� �������ʽÿ�.
	''xpay.Set "LGD_AMOUNTCHECKYN", "Y"
	''xpay.Set "LGD_AMOUNT", 1000
	
	''�ֹ���ȣ ���� �׽�Ʈ :: �ȵ�..
    ''xpay.Set "LGD_OID", "2010011100007" xpay.Set "LGD_OID", "2010011100007" 
    
    '/*
    ' *************************************************
    ' * 1.�������� ��û(�������� ������) - END
    ' *************************************************
    ' */

    '/*
    ' * 2. �������� ��û ���ó��
    ' *
    ' * ���� ������û ��� ���� �Ķ���ʹ� �����޴����� �����Ͻñ� �ٶ��ϴ�.
    ' */

    if  xpay.TX() then
        '1)������� ȭ��ó��(����,���� ��� ó���� �Ͻñ� �ٶ��ϴ�.)
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
        '2)API ��û���� ȭ��ó��
        Response.Write("������û�� �����Ͽ����ϴ�. <br>")
        Response.Write("TX Response_code = " & xpay.resCode & "<br>")
        Response.Write("TX Response_msg = " & xpay.resMsg & "<p>")
            
        '������û ��� ���� ���� DBó��
        Response.Write("����������û ��� ���� DBó���Ͻñ� �ٶ��ϴ�." & "<br>")
        
        Response.Write("�ŷ���ȣ : " & xpay.Response("LGD_TID", 0) & "<br>")
    end if
 %>
