<!-- #include file="./lgdacom/md5.asp" -->
<%
    '/*
    ' * [���� �������ó��(DB) ������]
    ' *
    ' * 1) ������ ������ ���� hashdata�� ������ �ݵ�� �����ϼž� �մϴ�.
    ' *
    ' */
    LGD_RESPCODE            = trim(request("LGD_RESPCODE"))             '// �����ڵ�: 0000(����) �׿� ����
    LGD_RESPMSG             = trim(request("LGD_RESPMSG"))              '// ����޼���
    LGD_MID                 = trim(request("LGD_MID"))                  '// �������̵�
    LGD_OID                 = trim(request("LGD_OID"))                  '// �ֹ���ȣ
    LGD_AMOUNT              = trim(request("LGD_AMOUNT"))               '// �ŷ��ݾ�
    LGD_TID                 = trim(request("LGD_TID"))                  '// �������� �ο��� �ŷ���ȣ
    LGD_PAYTYPE             = trim(request("LGD_PAYTYPE"))              '// ���������ڵ�
    LGD_PAYDATE             = trim(request("LGD_PAYDATE"))              '// �ŷ��Ͻ�(�����Ͻ�/��ü�Ͻ�)
    LGD_HASHDATA            = trim(request("LGD_HASHDATA"))             '// �ؽ���
    LGD_FINANCECODE         = trim(request("LGD_FINANCECODE"))          '// ��������ڵ�(�����ڵ�)
    LGD_FINANCENAME         = trim(request("LGD_FINANCENAME"))          '// ��������̸�(�����̸�)
    LGD_ESCROWYN            = trim(request("LGD_ESCROWYN"))             '// ����ũ�� ���뿩��
    LGD_TIMESTAMP           = trim(request("LGD_TIMESTAMP"))            '// Ÿ�ӽ�����
    LGD_ACCOUNTNUM          = trim(request("LGD_ACCOUNTNUM"))           '// ���¹�ȣ(�������Ա�)
    LGD_CASTAMOUNT          = trim(request("LGD_CASTAMOUNT"))           '// �Ա��Ѿ�(�������Ա�)
    LGD_CASCAMOUNT          = trim(request("LGD_CASCAMOUNT"))           '// ���Աݾ�(�������Ա�)
    LGD_CASFLAG             = trim(request("LGD_CASFLAG"))              '// �������Ա� �÷���(�������Ա�) - 'R':�����Ҵ�, 'I':�Ա�, 'C':�Ա����
    LGD_CASSEQNO            = trim(request("LGD_CASSEQNO"))             '// �Աݼ���(�������Ա�)
    LGD_CASHRECEIPTNUM      = trim(request("LGD_CASHRECEIPTNUM"))       '// ���ݿ����� ���ι�ȣ
    LGD_CASHRECEIPTSELFYN   = trim(request("LGD_CASHRECEIPTSELFYN"))    '// ���ݿ����������߱������� Y: �����߱��� ����, �׿� : ������
    LGD_CASHRECEIPTKIND     = trim(request("LGD_CASHRECEIPTKIND"))      '// ���ݿ����� ���� 0: �ҵ������ , 1: ����������
	LGD_PAYER            	= trim(request("LGD_PAYER"))             	'// �Ա��ڸ�
	
    '/*
    ' * ��������
    ' */
    LGD_BUYER               = trim(request("LGD_BUYER"))                '// ������
    LGD_PRODUCTINFO         = trim(request("LGD_PRODUCTINFO"))          '// ��ǰ��
    LGD_BUYERID             = trim(request("LGD_BUYERID"))              '// ������ ID
    LGD_BUYERADDRESS        = trim(request("LGD_BUYERADDRESS"))         '// ������ �ּ�
    LGD_BUYERPHONE          = trim(request("LGD_BUYERPHONE"))           '// ������ ��ȭ��ȣ
    LGD_BUYEREMAIL          = trim(request("LGD_BUYEREMAIL"))           '// ������ �̸���
    LGD_BUYERSSN            = trim(request("LGD_BUYERSSN"))             '// ������ �ֹι�ȣ
    LGD_PRODUCTCODE         = trim(request("LGD_PRODUCTCODE"))          '// ��ǰ�ڵ�
    LGD_RECEIVER            = trim(request("LGD_RECEIVER"))             '// ������
    LGD_RECEIVERPHONE       = trim(request("LGD_RECEIVERPHONE"))        '// ������ ��ȭ��ȣ
    LGD_DELIVERYINFO        = trim(request("LGD_DELIVERYINFO"))         '// �����
	
	
    '/*
    ' * hashdata ������ ���� mertkey�� ���������� -> ������� -> ���������������� Ȯ���ϽǼ� �ֽ��ϴ�.
    ' * LG�����޿��� �߱��� ����Ű�� �ݵ�ú����� �ֽñ� �ٶ��ϴ�.
    ' */
    LGD_MERTKEY = ""  '//mertkey
    LGD_HASHDATA2 = md5( LGD_MID & LGD_OID & LGD_AMOUNT & LGD_RESPCODE & LGD_TIMESTAMP & LGD_MERTKEY )

    '/*
    ' * ���� ó����� ���ϸ޼���
    ' *
    ' * OK  : ���� ó����� ����
    ' * �׿� : ���� ó����� ����
    ' *
    ' * �� ���ǻ��� : ������ 'OK' �����̿��� �ٸ����ڿ��� ���ԵǸ� ����ó�� �ǿ��� �����Ͻñ� �ٶ��ϴ�.
    ' */
    resultMSG = "������� ���� DBó��(LGD_CASNOTEURL) ������� �Է��� �ֽñ� �ٶ��ϴ�."

    if LGD_HASHDATA2 = LGD_HASHDATA then
        '//�ؽ��� ������ �����̸�
        if LGD_RESPCODE = "0000" then
            '//������ �����̸�
            if LGD_CASFLAG = "R" then
                '/*
                ' * ������ �Ҵ� ���� ��� ���� ó��(DB) �κ�
                ' * ���� ��� ó���� �����̸� "OK"
                ' */
                resultMSG = "OK"
            elseif LGD_CASFLAG = "I" then
                '/*
                ' * ������ �Ա� ���� ��� ���� ó��(DB) �κ�
                ' * ���� ��� ó���� �����̸� "OK"
                ' */
                resultMSG = "OK"
            elseif LGD_CASFLAG = "C" then
                '/*
                ' * ������ �Ա���� ���� ��� ���� ó��(DB) �κ�
                ' * ���� ��� ó���� �����̸� "OK"
                ' */
                resultMSG = "OK"
            end if
        else
            '//������ �����̸�
            '/*
            ' * �ŷ����� ��� ���� ó��(DB) �κ�
            ' * ������� ó���� �����̸� "OK"
            ' */
            resultMSG = "OK"
        end if
    else
        '//�ؽ����� ������ �����̸�
        '/*
        ' * hashdata���� ���� �α׸� ó���Ͻñ� �ٶ��ϴ�.
        ' */
        resultMSG = "������� ���� DBó��(LGD_CASNOTEURL) �ؽ��� ������ �����Ͽ����ϴ�."
    end if

    Response.Write(resultMSG)
%>

