<%
    '/*
    ' * [������� ��û ������]
    ' *
    ' * LG���������� ���� �������� �ŷ���ȣ(LGD_TID)�� ������ ��� ��û�� �մϴ�.(�Ķ���� ���޽� POST�� ����ϼ���)
    ' * (���ν� LG���������� ���� �������� PAYKEY�� ȥ������ ������.)
    ' */
    CST_PLATFORM         = trim(request("CST_PLATFORM"))        ' LG������ �������� ����(test:�׽�Ʈ, service:����)
    CST_MID              = trim(request("CST_MID"))             ' LG���������� ���� �߱޹����� �������̵� �Է��ϼ���.
                                                                ' �׽�Ʈ ���̵�� 't'�� �����ϰ� �Է��ϼ���.
    if CST_PLATFORM = "test" then                               ' �������̵�(�ڵ�����)
        LGD_MID = "t" & CST_MID
    else
        LGD_MID = CST_MID
    end if
    LGD_TID              = trim(request("LGD_TID"))             ' LG���������� ���� �������� �ŷ���ȣ(LGD_TID)

    configPath           = "C:/lgdacom"         				' LG�����޿��� ������ ȯ������("/conf/lgdacom.conf") ��ġ ����.


    Set xpay = server.CreateObject("XPayClientCOM.XPayClient")
    xpay.Init configPath, CST_PLATFORM
    xpay.Init_TX(LGD_MID)

    xpay.Set "LGD_TXNAME", "Cancel"
    xpay.Set "LGD_TID", LGD_TID
 

    '/*
    ' * 1. ������� ��û ���ó��
    ' *
    ' * ��Ұ�� ���� �Ķ���ʹ� �����޴����� �����Ͻñ� �ٶ��ϴ�.
    ' */
    if xpay.TX() then
        '1)������Ұ�� ȭ��ó��(����,���� ��� ó���� �Ͻñ� �ٶ��ϴ�.)
        Response.Write("������� ��û�� �Ϸ�Ǿ����ϴ�. <br>")
        Response.Write("TX Response_code = " & xpay.resCode & "<br>")
        Response.Write("TX Response_msg = " & xpay.resMsg & "<p>")
    else
        '2)API ��û ���� ȭ��ó��
        Response.Write("������� ��û�� �����Ͽ����ϴ�. <br>")
        Response.Write("TX Response_code = " & xpay.resCode & "<br>")
        Response.Write("TX Response_msg = " & xpay.resMsg & "<p>")
    end if
%>
