<%
'' include virtual="/admin/etc/incOutMallCommonFunction.asp  �������� ��ħ

function TenDlvCode2InterParkDlvCode(itenCode)
    select Case itenCode
        CASE "1" : TenDlvCode2InterParkDlvCode = "169178"     ''����
        CASE "2" : TenDlvCode2InterParkDlvCode = "169198"     ''����
        CASE "3" : TenDlvCode2InterParkDlvCode = "169177"     ''�������
        CASE "4" : TenDlvCode2InterParkDlvCode = "169168"     ''CJ GLS
        CASE "5" : TenDlvCode2InterParkDlvCode = "169211"     ''��Ŭ����
        CASE "6" : TenDlvCode2InterParkDlvCode = "169181"     ''�Ｚ HTH
        CASE "7" : TenDlvCode2InterParkDlvCode = "231145"     ''����(���ѹ̸�)
        CASE "8" : TenDlvCode2InterParkDlvCode = "169199"     ''��ü���ù�
        CASE "9" : TenDlvCode2InterParkDlvCode = "169187"     ''KGB�ù�
        CASE "10" : TenDlvCode2InterParkDlvCode = "169194"     ''�����ù� / �ο���(�� ����)
        CASE "11" : TenDlvCode2InterParkDlvCode = ""     ''�������ù�
        CASE "12" : TenDlvCode2InterParkDlvCode = ""     ''�ѱ��ù� / ���ѱ��ù蹰��?
        CASE "13" : TenDlvCode2InterParkDlvCode = "169200"     ''���ο�ĸ
        CASE "14" : TenDlvCode2InterParkDlvCode = ""     ''���̽��ù�
        CASE "15" : TenDlvCode2InterParkDlvCode = ""     ''�߾��ù�
        CASE "16" : TenDlvCode2InterParkDlvCode = ""     ''�����ù�
        CASE "17" : TenDlvCode2InterParkDlvCode = ""     ''Ʈ����ù�
        CASE "18" : TenDlvCode2InterParkDlvCode = "169182"     ''�����ù�
        CASE "19" : TenDlvCode2InterParkDlvCode = ""     ''KGBƯ���ù�
        CASE "20" : TenDlvCode2InterParkDlvCode = ""     ''KT������
        CASE "21" : TenDlvCode2InterParkDlvCode = "303978"     ''�浿�ù�
        CASE "22" : TenDlvCode2InterParkDlvCode = "169526"     ''����ù�
        CASE "23" : TenDlvCode2InterParkDlvCode = "236288"     ''�굦���ù� �ż���
        CASE "24" : TenDlvCode2InterParkDlvCode = "231491"     ''�簡���ͽ�������
        CASE "25" : TenDlvCode2InterParkDlvCode = "229381"     ''�ϳ����ù�
        CASE "26" : TenDlvCode2InterParkDlvCode = "263792"     ''�Ͼ��ù�
        CASE "27" : TenDlvCode2InterParkDlvCode = "169194"     ''LOEX�ù�
        CASE "28" : TenDlvCode2InterParkDlvCode = "231145"     ''�����ͽ�������
        CASE "29" : TenDlvCode2InterParkDlvCode = "231194"     ''�ǿ��ù�
        CASE "30" : TenDlvCode2InterParkDlvCode = "266237"     ''�̳�����
        CASE "31" : TenDlvCode2InterParkDlvCode = "230175"     ''õ���ù�
        CASE "33" : TenDlvCode2InterParkDlvCode = "250701"     ''ȣ���ù�
        CASE "34" : TenDlvCode2InterParkDlvCode = "258064"     ''���ȭ���ù�
        CASE "35" : TenDlvCode2InterParkDlvCode = "169172"     ''CVSnet�ù�
        CASE "98" : TenDlvCode2InterParkDlvCode = "169316"     ''������->�����
        CASE "99" : TenDlvCode2InterParkDlvCode = "169167"     ''��Ÿ
        CASE  Else
            TenDlvCode2InterParkDlvCode = ""      ''��Ÿ�߼�(169167)
    end Select
end function

function TenDlvCode2LotteDlvCode(itenCode)
    ''if IsNULL(itenCode) then Exit function
    if IsNULL(itenCode) then itenCode="99"
    
    itenCode = TRIM(CStr(itenCode))
    select Case itenCode
        CASE "1" : TenDlvCode2LotteDlvCode = "27"     ''����
        CASE "2" : TenDlvCode2LotteDlvCode = "1"     ''����v
        CASE "3" : TenDlvCode2LotteDlvCode = "5"     ''�������
        CASE "4" : TenDlvCode2LotteDlvCode = "31"     ''CJ GLS
        CASE "5" : TenDlvCode2LotteDlvCode = "23"     ''��Ŭ����
        CASE "6" : TenDlvCode2LotteDlvCode = "32"     ''�Ｚ HTH
        CASE "7" : TenDlvCode2LotteDlvCode = "56"     ''����(���ѹ̸�) ''Ȯ
        CASE "8" : TenDlvCode2LotteDlvCode = "9339"     ''��ü���ù�
        CASE "9" : TenDlvCode2LotteDlvCode = "39"     ''KGB�ù�
        CASE "10" : TenDlvCode2LotteDlvCode = "34"     ''�����ù� / �ο���(�� ����)
        CASE "11" : TenDlvCode2LotteDlvCode = ""     ''�������ù�
        CASE "12" : TenDlvCode2LotteDlvCode = "29"     ''�ѱ��ù� / �ѱ�Ư��
        CASE "13" : TenDlvCode2LotteDlvCode = "37"     ''���ο�ĸ
        CASE "14" : TenDlvCode2LotteDlvCode = ""     ''���̽��ù�
        CASE "15" : TenDlvCode2LotteDlvCode = "43"     ''�߾��ù�
        CASE "16" : TenDlvCode2LotteDlvCode = ""     ''�����ù�
        CASE "17" : TenDlvCode2LotteDlvCode = "36"     ''Ʈ����ù�
        CASE "18" : TenDlvCode2LotteDlvCode = "41"     ''�����ù�
        CASE "19" : TenDlvCode2LotteDlvCode = "44"     ''KGBƯ���ù�
        CASE "20" : TenDlvCode2LotteDlvCode = "30"     ''KT������
        CASE "21" : TenDlvCode2LotteDlvCode = "52"     ''�浿�ù�
        CASE "22" : TenDlvCode2LotteDlvCode = ""     ''����ù�
        CASE "23" : TenDlvCode2LotteDlvCode = "42"     ''�굦���ù� �ż���
        CASE "24" : TenDlvCode2LotteDlvCode = "51"     ''�簡���ͽ�������
        CASE "25" : TenDlvCode2LotteDlvCode = "3"     ''�ϳ����ù�v
        CASE "26" : TenDlvCode2LotteDlvCode = "47"     ''�Ͼ��ù�
        CASE "27" : TenDlvCode2LotteDlvCode = ""     ''LOEX�ù�
        CASE "28" : TenDlvCode2LotteDlvCode = "35"     ''�����ͽ�������
        CASE "29" : TenDlvCode2LotteDlvCode = "45"     ''�ǿ��ù�
        CASE "30" : TenDlvCode2LotteDlvCode = "57"     ''�̳�����
        CASE "31" : TenDlvCode2LotteDlvCode = "33"     ''õ���ù�
        CASE "33" : TenDlvCode2LotteDlvCode = "99"     ''ȣ���ù�
        CASE "34" : TenDlvCode2LotteDlvCode = "46"     ''���ȭ���ù�
        CASE "35" : TenDlvCode2LotteDlvCode = "99"     ''CVSnet�ù�
        CASE "98" : TenDlvCode2LotteDlvCode = "99"     ''������
        CASE "99" : TenDlvCode2LotteDlvCode = "99"     ''��ü����
        CASE  Else
            TenDlvCode2LotteDlvCode = "99"
    end Select
end function


'''�Ե�iMall ���庯ȯ
function TenDlvCode2LotteiMallDlvCode(itenCode)
    if IsNULL(itenCode) then Exit function
    itenCode = TRIM(CStr(itenCode))
''41	�����ù�
''99	��Ÿ

    select Case itenCode
        CASE "1" : TenDlvCode2LotteiMallDlvCode = "15"     ''����
        CASE "2" : TenDlvCode2LotteiMallDlvCode = "11"     ''����v
        CASE "3" : TenDlvCode2LotteiMallDlvCode = "12"     ''�������
        CASE "4" : TenDlvCode2LotteiMallDlvCode = "16"     ''CJ GLS
        CASE "5" : TenDlvCode2LotteiMallDlvCode = ""     ''��Ŭ����
        CASE "6" : TenDlvCode2LotteiMallDlvCode = "22"     ''�Ｚ HTH
        CASE "7" : TenDlvCode2LotteiMallDlvCode = "26"     ''����(���ѹ̸�) ''Ȯ
        CASE "8" : TenDlvCode2LotteiMallDlvCode = "31"     ''��ü���ù�
        CASE "9" : TenDlvCode2LotteiMallDlvCode = "40"     ''KGB�ù�
        CASE "10" : TenDlvCode2LotteiMallDlvCode = "34"     ''�����ù� / �ο���(�� ����)
        CASE "11" : TenDlvCode2LotteiMallDlvCode = ""     ''�������ù�
        CASE "12" : TenDlvCode2LotteiMallDlvCode = "37"     ''�ѱ��ù� / �ѱ�Ư��
        CASE "13" : TenDlvCode2LotteiMallDlvCode = "32"     ''���ο�ĸ
        CASE "14" : TenDlvCode2LotteiMallDlvCode = ""     ''���̽��ù�
        CASE "15" : TenDlvCode2LotteiMallDlvCode = ""     ''�߾��ù�
        CASE "16" : TenDlvCode2LotteiMallDlvCode = ""     ''�����ù�
        CASE "17" : TenDlvCode2LotteiMallDlvCode = "36"     ''Ʈ����ù�
        CASE "18" : TenDlvCode2LotteiMallDlvCode = "24"     ''�����ù�
        CASE "19" : TenDlvCode2LotteiMallDlvCode = "40"     ''KGBƯ���ù�
        CASE "20" : TenDlvCode2LotteiMallDlvCode = ""     ''KT������
        CASE "21" : TenDlvCode2LotteiMallDlvCode = "49"     ''�浿�ù�
        CASE "22" : TenDlvCode2LotteiMallDlvCode = ""     ''����ù�
        CASE "23" : TenDlvCode2LotteiMallDlvCode = "47"     ''�굦���ù� �ż���
        CASE "24" : TenDlvCode2LotteiMallDlvCode = "43"     ''�簡���ͽ�������
        CASE "25" : TenDlvCode2LotteiMallDlvCode = "46"     ''�ϳ����ù�v
        CASE "26" : TenDlvCode2LotteiMallDlvCode = "18"     ''�Ͼ��ù�
        CASE "27" : TenDlvCode2LotteiMallDlvCode = "48"     ''LOEX�ù�
        CASE "28" : TenDlvCode2LotteiMallDlvCode = "26"     ''�����ͽ�������
        CASE "29" : TenDlvCode2LotteiMallDlvCode = "99"     ''�ǿ��ù�
        CASE "30" : TenDlvCode2LotteiMallDlvCode = "23"     ''�̳�����
        CASE "31" : TenDlvCode2LotteiMallDlvCode = "17"     ''õ���ù�
        CASE "33" : TenDlvCode2LotteiMallDlvCode = ""     ''ȣ���ù�
        CASE "34" : TenDlvCode2LotteiMallDlvCode = "38"     ''���ȭ���ù�
        CASE "35" : TenDlvCode2LotteiMallDlvCode = "99"     ''CVSnet�ù�
        CASE "98" : TenDlvCode2LotteiMallDlvCode = "99"     ''������
        CASE "99" : TenDlvCode2LotteiMallDlvCode = "99"     ''��ü����
        CASE  Else
            TenDlvCode2LotteiMallDlvCode = "99"
    end Select
end function

function LotteiMallDlvCode2Name(iltDlvCode)
    LotteiMallDlvCode2Name = "��Ÿ"
    if IsNULL(iltDlvCode) then Exit function
    iltDlvCode = TRIM(CStr(iltDlvCode))
    
    select Case iltDlvCode
        CASE "11" : LotteiMallDlvCode2Name="�����ù�"
        CASE "12" : LotteiMallDlvCode2Name="�����̴������"
        CASE "15" : LotteiMallDlvCode2Name="�����ù�"
        CASE "16" : LotteiMallDlvCode2Name="CJGLS"
        CASE "17" : LotteiMallDlvCode2Name="õ���ù�"
        CASE "18" : LotteiMallDlvCode2Name="�Ͼ��ù�"
        CASE "19" : LotteiMallDlvCode2Name="��Ÿ�ù�"
        CASE "22" : LotteiMallDlvCode2Name="HTH�ù�"
        CASE "24" : LotteiMallDlvCode2Name="�����ù�"
        CASE "26" : LotteiMallDlvCode2Name="�����ͽ�������"
        CASE "31" : LotteiMallDlvCode2Name="��ü���ù�"
        CASE "32" : LotteiMallDlvCode2Name="���ο�ĸ"
        CASE "34" : LotteiMallDlvCode2Name="�����ù�"
        CASE "36" : LotteiMallDlvCode2Name="Ʈ���"
        CASE "37" : LotteiMallDlvCode2Name="�ѱ��ù�"
        CASE "38" : LotteiMallDlvCode2Name="����ù�"
        CASE "40" : LotteiMallDlvCode2Name="KGB�ù�"
        CASE "41" : LotteiMallDlvCode2Name="�����ù�"
        CASE "43" : LotteiMallDlvCode2Name="�簡���ͽ�������"
        CASE "46" : LotteiMallDlvCode2Name="�ϳ����ù�"
        CASE "47" : LotteiMallDlvCode2Name="�������ù�"
        CASE "48" : LotteiMallDlvCode2Name="�ο����ù�"
        CASE "49" : LotteiMallDlvCode2Name="�浿�ù�"
        CASE "99" : LotteiMallDlvCode2Name="��Ÿ"
        CASE  Else
            LotteiMallDlvCode2Name = "��Ÿ"
    end Select
end function

%>