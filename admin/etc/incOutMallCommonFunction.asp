<%

function DrawApiMallSelect(sitename,selsitename)
    dim buf
    buf = "<select class='select' name='"&sitename&"' >"
    buf = buf&"<option value=''  >����"
	buf = buf&"<option value='lotteCom' "& chkIIF(selsitename="lotteCom","selected","") &" >�Ե�����"
	buf = buf&"<option value='lotteimall' "& chkIIF(selsitename="lotteimall","selected","") &" >�Ե�iMall"
    buf = buf&"<option value='lotteon' "& chkIIF(selsitename="lotteon","selected","") &" >�Ե�On"
    buf = buf&"<option value='shintvshopping' "& chkIIF(selsitename="shintvshopping","selected","") &" >�ż���TV����"
    buf = buf&"<option value='skstoa' "& chkIIF(selsitename="skstoa","selected","") &" >SKSTOA"
    buf = buf&"<option value='wetoo1300k' "& chkIIF(selsitename="wetoo1300k","selected","") &" >1300k"
	buf = buf&"<option value='interpark' "& chkIIF(selsitename="interpark","selected","") &" >������ũ"
	buf = buf&"<option value='cjmall' "& chkIIF(selsitename="cjmall","selected","") &" >cjmall"
	buf = buf&"<option value='gseshop' "& chkIIF(selsitename="gseshop","selected","") &" >gseshop"
	buf = buf&"<option value='ezwel' "& chkIIF(selsitename="ezwel","selected","") &" >ezwel"
    buf = buf&"<option value='benepia1010' "& chkIIF(selsitename="benepia1010","selected","") &" >�����Ǿ�"
	buf = buf&"<option value='auction1010' "& chkIIF(selsitename="auction1010","selected","") &" >����"
	buf = buf&"<option value='gmarket1010' "& chkIIF(selsitename="gmarket1010","selected","") &" >Gmarket"
	buf = buf&"<option value='nvstorefarm' "& chkIIF(selsitename="nvstorefarm","selected","") &" >�������"
    buf = buf&"<option value='nvstoremoonbangu' "& chkIIF(selsitename="nvstoremoonbangu","selected","") &" >������� ���汸"
    buf = buf&"<option value='Mylittlewhoopee' "& chkIIF(selsitename="Mylittlewhoopee","selected","") &" >������� Ĺ�ص�"
	buf = buf&"<option value='11st1010' "& chkIIF(selsitename="11st1010","selected","") &" >11����"
	buf = buf&"<option value='ssg' "& chkIIF(selsitename="ssg","selected","") &" >�ż����(SSG)"
	buf = buf&"<option value='halfclub' "& chkIIF(selsitename="halfclub","selected","") &" >����Ŭ��"
    buf = buf&"<option value='gsisuper' "& chkIIF(selsitename="gsisuper","selected","") &" >GS���̽���"
    buf = buf&"<option value='yes24' "& chkIIF(selsitename="yes24","selected","") &" >YES24"
    buf = buf&"<option value='wconcept1010' "& chkIIF(selsitename="wconcept1010","selected","") &" >����������"
    buf = buf&"<option value='withnature1010' "& chkIIF(selsitename="withnature1010","selected","") &" >�ڿ��̶�"
    buf = buf&"<option value='goodshop1010' "& chkIIF(selsitename="goodshop1010","selected","") &" >�¼�"
    buf = buf&"<option value='alphamall' "& chkIIF(selsitename="alphamall","selected","") &" >���ĸ�"
    buf = buf&"<option value='kakaostore' "& chkIIF(selsitename="kakaostore","selected","") &" >īī���彺���"
    buf = buf&"<option value='boribori1010' "& chkIIF(selsitename="boribori1010","selected","") &" >��������"
    buf = buf&"<option value='ohou1010' "& chkIIF(selsitename="ohou1010","selected","") &" >��������"
    buf = buf&"<option value='wadsmartstore' "& chkIIF(selsitename="wadsmartstore","selected","") &" >�͵彺��Ʈ�����"
    buf = buf&"<option value='casamia_good_com' "& chkIIF(selsitename="casamia_good_com","selected","") &" >���̾�"
    buf = buf&"<option value='lfmall' "& chkIIF(selsitename="lfmall","selected","") &" >LFmall"
    buf = buf&"<option value='coupang' "& chkIIF(selsitename="coupang","selected","") &" >����"
    buf = buf&"<option value='hmall1010' "& chkIIF(selsitename="hmall1010","selected","") &" >HMall"
    buf = buf&"<option value='WMP' "& chkIIF(selsitename="WMP","selected","") &" >������"
    buf = buf&"<option value='wmpfashion' "& chkIIF(selsitename="wmpfashion","selected","") &" >������W�м�"
	buf = buf&"</select>"

	response.write buf
end function

function DrawApiMallSelectSongjangInput(sitename,selsitename)
    dim buf
    buf = "<select name='"&sitename&"' >"
    buf = buf&"<option value=''  >����"
	buf = buf&"<option value='lotteCom' "& chkIIF(selsitename="lotteCom","selected","") &" >�Ե�����"
	buf = buf&"<option value='lotteimall' "& chkIIF(selsitename="lotteimall","selected","") &" >�Ե�iMall"
	buf = buf&"<option value='interpark' "& chkIIF(selsitename="interpark","selected","") &" >������ũ"
	buf = buf&"<option value='cjmall' "& chkIIF(selsitename="cjmall","selected","") &" >cjmall"
	buf = buf&"<option value='shoplinker' "& chkIIF(selsitename="shoplinker","selected","") &" >shoplinker"
	buf = buf&"</select>"

	response.write buf
end function

''��𿡻��?
function DrawApiMallCheck()
    dim buf
    buf = ""
    buf = buf&"<input type='checkbox' name='outmallck' value='interpark'>������ũ"
    buf = buf&"<input type='checkbox' name='outmallck' value='lotteCom'>�Ե�����"
    buf = buf&"<input type='checkbox' name='outmallck' value='lotteimall'>�Ե�iMall"

    response.write buf
end function

function TenDlvCode2AuctionDlvCode(itenCode)
    select Case itenCode
        CASE "1" : TenDlvCode2AuctionDlvCode = "hanjin"     ''����
        CASE "2" : TenDlvCode2AuctionDlvCode = "hyundai"     ''���� -> �Ե�
        CASE "3" : TenDlvCode2AuctionDlvCode = "korex"     ''�������
        CASE "4" : TenDlvCode2AuctionDlvCode = "cjgls"     ''CJ GLS
        CASE "5" : TenDlvCode2AuctionDlvCode = "etc"     ''��Ŭ����
        CASE "6" : TenDlvCode2AuctionDlvCode = "samsung"     ''�Ｚ HTH
        CASE "7" : TenDlvCode2AuctionDlvCode = "dongbu"     ''����(���ѹ̸�)
        CASE "8" : TenDlvCode2AuctionDlvCode = "epost"     ''��ü���ù�
        CASE "9" : TenDlvCode2AuctionDlvCode = "kgbls"     ''KGB�ù�
        CASE "10" : TenDlvCode2AuctionDlvCode = "etc"     ''�����ù� / �ο���(�� ����)
        CASE "11" : TenDlvCode2AuctionDlvCode = "etc"     ''�������ù�
        CASE "12" : TenDlvCode2AuctionDlvCode = "etc"     ''�ѱ��ù� / ���ѱ��ù蹰��?
        CASE "13" : TenDlvCode2AuctionDlvCode = "yellow"     ''���ο�ĸ
        CASE "14" : TenDlvCode2AuctionDlvCode = "etc"     ''���̽��ù�
        CASE "15" : TenDlvCode2AuctionDlvCode = "etc"     ''�߾��ù�
        CASE "16" : TenDlvCode2AuctionDlvCode = "etc"     ''�����ù�
        CASE "17" : TenDlvCode2AuctionDlvCode = "etc"     ''Ʈ����ù�
        CASE "18" : TenDlvCode2AuctionDlvCode = "kgb"     ''�����ù�
        CASE "19" : TenDlvCode2AuctionDlvCode = "kgb"     ''KGBƯ���ù�
        CASE "20" : TenDlvCode2AuctionDlvCode = "etc"     ''KT������
        CASE "21" : TenDlvCode2AuctionDlvCode = "kyungdong"     ''�浿�ù�
        CASE "22" : TenDlvCode2AuctionDlvCode = "etc"     ''�����ù�
        CASE "23" : TenDlvCode2AuctionDlvCode = "etc"     ''�굦���ù� �ż���
        CASE "24" : TenDlvCode2AuctionDlvCode = "sagawa"     ''�簡���ͽ�������
        CASE "25" : TenDlvCode2AuctionDlvCode = "etc"     ''�ϳ����ù�
        CASE "26" : TenDlvCode2AuctionDlvCode = "ilyang"     ''�Ͼ��ù�
        CASE "27" : TenDlvCode2AuctionDlvCode = "etc"     ''LOEX�ù�
        CASE "28" : TenDlvCode2AuctionDlvCode = "dongbu"     ''�����ͽ�������
        CASE "29" : TenDlvCode2AuctionDlvCode = "etc"     ''�ǿ��ù�
        CASE "30" : TenDlvCode2AuctionDlvCode = "etc"     ''�̳�����
        CASE "31" : TenDlvCode2AuctionDlvCode = "chonil"     ''õ���ù�
        CASE "33" : TenDlvCode2AuctionDlvCode = "etc"     ''ȣ���ù�
        CASE "34" : TenDlvCode2AuctionDlvCode = "daesin"     ''���ȭ���ù�
        CASE "35" : TenDlvCode2AuctionDlvCode = "cvsnet"     ''CVSnet�ù�  - 99 ��Ÿ�߼����ù�
        CASE "42" : TenDlvCode2AuctionDlvCode = "cvsnet"     ''CVSnet�ù�  - 99 ��Ÿ�߼����ù�
        CASE "38" : TenDlvCode2AuctionDlvCode = "gtx"     ''GTX������
        CASE "39" : TenDlvCode2AuctionDlvCode = "dongbu"     ''KG������ - �����ͽ�������
        CASE "98" : TenDlvCode2AuctionDlvCode = "etc"     ''������->�����
        CASE "99" : TenDlvCode2AuctionDlvCode = "etc"     ''��Ÿ
        CASE  Else
            TenDlvCode2AuctionDlvCode = "etc"      ''��Ÿ�߼�
    end Select
end function

Function TenDlvCode2LotteonDlvCode(itenCode)
' 0001||�Ե��ù�
' 0002||CJ�������
' 0003||�����ù�
' 0004||��ü���ù�
' 0005||�����ù�
' 0006||�����ù�
' 0007||APEX(ECMS Express)
' 0008||DHL
' 0009||DHL Global Mail
' 0010||EMS
' 0011||Fedex
' 0012||GSI Express
' 0013||GSMNtoN(�ην�)
' 0014||GTX ������ �ù�
' 0015||i-Parcel
' 0016||KGB�ù�
' 0017||KGL��Ʈ����
' 0018||KG������
' 0019||SEDEX
' 0020||TNT Express
' 0021||TPL
' 0022||USPS
' 0023||�ǿ��ù�
' 0024||�浿�ù�
' 0025||�����ù�
' 0026||������
' 0027||�������ȭ���ù�
' 0028||����ù�
' 0029||�������
' 0030||�����ͽ�������
' 0031||�帲�ù�
' 0032||����������Ư��
' 0033||�������佺
' 0034||�ＺHTH
' 0035||�ִ�Ʈ��
' 0036||������ͽ�������
' 0037||���ο�ĸ�ù�
' 0038||��ü��
' 0039||��ü�����
' 0040||�̳������ù�
' 0041||�Ͼ������
' 0042||�Ͼ��ù�
' 0043||õ���ù�
' 0044||�������ù�
' 0045||�������ͽ�������
' 0046||�ϳ����ù�
' 0047||�ѵ���
' 0048||���ǻ���ù�
' 0049||�յ��ù�
' 0050||ȣ���ù�
' 0051||KT������
' 0052||�簡��
' 0053||�츮�ù�
' 0054||���Ͽ�
' 9000||��ü���
' 9999||��Ÿ�ù�
' LE_QUICK||���Ե�����ۻ�

    select Case itenCode
        CASE "1" : TenDlvCode2LotteonDlvCode = "0006"     ''����
        CASE "2" : TenDlvCode2LotteonDlvCode = "0001"     ''���� -> �Ե��ù�� ����� 2017-03-13 ������ ����
        CASE "3" : TenDlvCode2LotteonDlvCode = "0002"     ''������� (CJ�������(CJGLS))
        CASE "4" : TenDlvCode2LotteonDlvCode = "0002"     ''CJ GLS (CJ�������(CJGLS))
        CASE "5" : TenDlvCode2LotteonDlvCode = ""     ''��Ŭ����
        CASE "6" : TenDlvCode2LotteonDlvCode = "0034"     ''�Ｚ HTH
        CASE "7" : TenDlvCode2LotteonDlvCode = "0030"     ''����(���ѹ̸�)
        CASE "8" : TenDlvCode2LotteonDlvCode = "0004"     ''��ü���ù�
        CASE "9" : TenDlvCode2LotteonDlvCode = "0016"     ''KGB�ù�
        CASE "10" : TenDlvCode2LotteonDlvCode = ""     ''�����ù� / �ο���(�� ����)
        CASE "11" : TenDlvCode2LotteonDlvCode = ""     ''�������ù�
        CASE "12" : TenDlvCode2LotteonDlvCode = ""     ''�ѱ��ù� / ���ѱ��ù蹰��?
        CASE "13" : TenDlvCode2LotteonDlvCode = "0016"     ''���ο�ĸ
        CASE "14" : TenDlvCode2LotteonDlvCode = ""     ''���̽��ù�
        CASE "15" : TenDlvCode2LotteonDlvCode = ""     ''�߾��ù�
        CASE "16" : TenDlvCode2LotteonDlvCode = ""     ''�����ù�
        CASE "17" : TenDlvCode2LotteonDlvCode = "0051"     ''Ʈ����ù�
        CASE "18" : TenDlvCode2LotteonDlvCode = "0005"     ''�����ù�
        CASE "19" : TenDlvCode2LotteonDlvCode = "0016"     ''KGBƯ���ù�
        CASE "20" : TenDlvCode2LotteonDlvCode = "0051"     ''KT������
        CASE "21" : TenDlvCode2LotteonDlvCode = "0024"     ''�浿�ù�
        CASE "22" : TenDlvCode2LotteonDlvCode = "0025"     ''�����ù�
        CASE "23" : TenDlvCode2LotteonDlvCode = ""     ''�굦���ù� �ż���
        CASE "24" : TenDlvCode2LotteonDlvCode = "0052"     ''�簡���ͽ�������
        CASE "25" : TenDlvCode2LotteonDlvCode = "0046"     ''�ϳ����ù�
        CASE "26" : TenDlvCode2LotteonDlvCode = "0041"     ''�Ͼ��ù�
        CASE "27" : TenDlvCode2LotteonDlvCode = ""     ''LOEX�ù�
        CASE "28" : TenDlvCode2LotteonDlvCode = "0030"     ''�����ͽ�������
        CASE "29" : TenDlvCode2LotteonDlvCode = "0023"     ''�ǿ��ù�
        CASE "30" : TenDlvCode2LotteonDlvCode = "0040"     ''�̳�����
        CASE "31" : TenDlvCode2LotteonDlvCode = "0043"     ''õ���ù�
        CASE "33" : TenDlvCode2LotteonDlvCode = "0050"     ''ȣ���ù�
        CASE "34" : TenDlvCode2LotteonDlvCode = "0028"     ''���ȭ���ù�
        CASE "35" : TenDlvCode2LotteonDlvCode = "0044"     ''CVSnet�ù�  -
        CASE "37" : TenDlvCode2LotteonDlvCode = "0049"     ''�յ��ù�  -
        CASE "38" : TenDlvCode2LotteonDlvCode = "0014"     ''GTX������
        CASE "39" : TenDlvCode2LotteonDlvCode = "0018"     ''KG������ - �����ͽ�������
        CASE "98" : TenDlvCode2LotteonDlvCode = ""     ''������->�����
        CASE "41" : TenDlvCode2LotteonDlvCode = "0031"     ''�帲�ù�(�����ù�,���ο�ĸ)  ''2018/02/13
        CASE "42" : TenDlvCode2LotteonDlvCode = "0044"     ''CU POST�� �������ù�� �ش޶��..2019-03-08 ������ ����
        CASE "99" : TenDlvCode2LotteonDlvCode = "9999"     ''��Ÿ  0000033028
        CASE  Else
            TenDlvCode2LotteonDlvCode = "9999"      ''��Ÿ�߼�
    end Select
end function

Function TenDlvCode2ShintvshoppingDlvCode(itenCode)
' 10||CJ �������
' 11||�Ե��ù�
' 12||�����ù�
' 13||��ü���ù�
' 14||�����ù�
' 17||�浿�ù�
' 20||�̳�����
' 21||�Ͼ��ù�
' 22||õ���ù�
' 23||�ο���(����)�ù�
' 24||SC�������ù�
' 25||����ù�
' 26||CVS�������ù�
' 27||��ü���
' 28||��ġ��ǰ
' 29||�ＺHTH
' 30||�ѹ̸���
' 31||��Ŭ����
' 32||�����ù�
' 33||ȣ���ù�
' 34||�츮�ù�
' 35||Ʈ���
' 36||�ѱ��ù�
' 37||�յ��ù�
' 38||GTX������
' 39||SLX�ù�
' 40||��üó��
' 41||������
' 42||HI�ù�
' 43||�۷����ù�
' 44||YDH
' 45||ȭ������Ź��
' 60||���۷ι�
' 61||ACI Express
' 62||��������
' 63||���̽�����
' 64||ĳ���ٽ���
' 65||�������ڸ���
' 66||�ٹٹٷ�����
' 67||��������
' 70||�����������
' 71||��Ÿ�ڸ���
' 72||�Ե�ĥ��
' 73||yunda express
' 74||�����ͽ��÷���
' 75||���ڵ��ؿ���
' 76||��������
' 77||Ƽ�ǿ��ڸ���
' 78||�ǿ��ù�
' 79||����
' 90||����
' 99||��Ÿ

    select Case itenCode
        CASE "1" : TenDlvCode2ShintvshoppingDlvCode = "14"     ''����
        CASE "2" : TenDlvCode2ShintvshoppingDlvCode = "11"     '�Ե��ù�
        CASE "3" : TenDlvCode2ShintvshoppingDlvCode = "10"     ''������� (CJ�������(CJGLS))
        CASE "4" : TenDlvCode2ShintvshoppingDlvCode = "10"     ''CJ GLS (CJ�������(CJGLS))
        CASE "5" : TenDlvCode2ShintvshoppingDlvCode = "31"     ''��Ŭ����
        CASE "5" : TenDlvCode2ShintvshoppingDlvCode = "24"     ''SC������
        CASE "8" : TenDlvCode2ShintvshoppingDlvCode = "13"     ''��ü���ù�
        CASE "9" : TenDlvCode2ShintvshoppingDlvCode = ""     ''KGB�ù�
        CASE "10" : TenDlvCode2ShintvshoppingDlvCode = "23"     ''�����ù� / �ο���(�� ����)
        CASE "12" : TenDlvCode2ShintvshoppingDlvCode = "36"     ''�ѱ��ù� / ���ѱ��ù蹰��?
        CASE "13" : TenDlvCode2ShintvshoppingDlvCode = ""     ''���ο�ĸ
        CASE "16" : TenDlvCode2ShintvshoppingDlvCode = "32"     ''�����ù�
        CASE "17" : TenDlvCode2ShintvshoppingDlvCode = "35"     ''Ʈ����ù�
        CASE "18" : TenDlvCode2ShintvshoppingDlvCode = "12"     ''�����ù�
        CASE "20" : TenDlvCode2ShintvshoppingDlvCode = "99"     ''KT������..2023-02-02 ������..��Ī�� �� ����..��Ÿ�� �ϴ� ó��
        CASE "21" : TenDlvCode2ShintvshoppingDlvCode = "17"     ''�浿�ù�
        CASE "22" : TenDlvCode2ShintvshoppingDlvCode = ""     ''�����ù�
        CASE "24" : TenDlvCode2ShintvshoppingDlvCode = "24"     ''SC������
        CASE "25" : TenDlvCode2ShintvshoppingDlvCode = ""     ''�ϳ����ù�
        CASE "26" : TenDlvCode2ShintvshoppingDlvCode = "21"     ''�Ͼ��ù�
        CASE "27" : TenDlvCode2ShintvshoppingDlvCode = "23"     ''LOEX�ù�
        CASE "28" : TenDlvCode2ShintvshoppingDlvCode = ""     ''�����ͽ�������
        CASE "29" : TenDlvCode2ShintvshoppingDlvCode = "78"     ''�ǿ��ù�
        CASE "30" : TenDlvCode2ShintvshoppingDlvCode = "20"     ''�̳�����
        CASE "31" : TenDlvCode2ShintvshoppingDlvCode = "22"     ''õ���ù�
        CASE "33" : TenDlvCode2ShintvshoppingDlvCode = "33"     ''ȣ���ù�
        CASE "34" : TenDlvCode2ShintvshoppingDlvCode = "25"     ''���ȭ���ù�
        CASE "35" : TenDlvCode2ShintvshoppingDlvCode = "26"     ''CVSnet�ù�  -
        CASE "36" : TenDlvCode2ShintvshoppingDlvCode = ""     '��������ȭ��
        CASE "37" : TenDlvCode2ShintvshoppingDlvCode = "37"     ''�յ��ù�  -
        CASE "38" : TenDlvCode2ShintvshoppingDlvCode = "38"     ''GTX������
        CASE "39" : TenDlvCode2ShintvshoppingDlvCode = ""     ''KG������ - �����ͽ�������
        CASE "40" : TenDlvCode2ShintvshoppingDlvCode = ""     ''KG������ - �����ͽ�������
        CASE "41" : TenDlvCode2ShintvshoppingDlvCode = ""     ''�帲�ù�(�����ù�,���ο�ĸ)  ''2018/02/13
        CASE "42" : TenDlvCode2ShintvshoppingDlvCode = "26"     ''CU POST�� �������ù�� �ش޶��..2019-03-08 ������ ����
        CASE "43" : TenDlvCode2ShintvshoppingDlvCode = "42"    'HI�ù�	http:
        CASE "44" : TenDlvCode2ShintvshoppingDlvCode = ""    'Ȩ��	http://ww
        CASE "45" : TenDlvCode2ShintvshoppingDlvCode = "43"    'FLF�۷����ù�	h
        CASE "46" : TenDlvCode2ShintvshoppingDlvCode = ""    'FedEx	https
        CASE "47" : TenDlvCode2ShintvshoppingDlvCode = "77"    'Ƽ�ǿ��ڸ���	h
        CASE "48" : TenDlvCode2ShintvshoppingDlvCode = ""    '�������븮	h
        CASE "49" : TenDlvCode2ShintvshoppingDlvCode = ""    '�������븮�ù�	h
        CASE "90" : TenDlvCode2ShintvshoppingDlvCode = ""    'EMS	http://se
        CASE "91" : TenDlvCode2ShintvshoppingDlvCode = ""    'DHL	http://ww
        CASE "98" : TenDlvCode2ShintvshoppingDlvCode = "99"    '������		Y
        CASE "99" : TenDlvCode2ShintvshoppingDlvCode = "99"    '��Ÿ		Y	N
        CASE "100": TenDlvCode2ShintvshoppingDlvCode = ""     '�ѿ츮����	h
        CASE  Else
            TenDlvCode2ShintvshoppingDlvCode = "99"      ''��Ÿ�߼�
    end Select
end function

Function TenDlvCode2SkstoaDlvCode(itenCode)
' 10||CJ�������
' 11||�����ù�
' 12||�Ե��ù�
' 13||��ü���ù�
' 14||�����ù�
' 18||�Ͼ������
' 26||SBGLS
' 27||����ù�
' 28||�浿�ù�
' 29||�յ��ù�
' 30||�������ù�
' 32||���ǻ���ù�
' 34||õ���ù�
' 35||�ǿ��ù�
' 36||�����ù�
' 37||Ƽ�ǿ��ڸ���
' 38||�þ˷�����
' 39||�������븮
' 40||��üó��
' 47||��ü���
' 48||��ġ��ǰ
' 90||���������ù�
' 91||�Ե��۷ι�
' 92||��������
' 99||��Ÿ
    select Case itenCode
        CASE "1" : TenDlvCode2SkstoaDlvCode = "11"     ''����
        CASE "2" : TenDlvCode2SkstoaDlvCode = "12"     '�Ե��ù�
        CASE "3" : TenDlvCode2SkstoaDlvCode = "10"     ''������� (CJ�������(CJGLS))
        CASE "4" : TenDlvCode2SkstoaDlvCode = "10"     ''CJ GLS (CJ�������(CJGLS))
        CASE "5" : TenDlvCode2SkstoaDlvCode = ""     ''��Ŭ����
        CASE "5" : TenDlvCode2SkstoaDlvCode = ""     ''SC������
        CASE "8" : TenDlvCode2SkstoaDlvCode = "13"     ''��ü���ù�
        CASE "9" : TenDlvCode2SkstoaDlvCode = ""   ''KGB�ù�
        CASE "10" : TenDlvCode2SkstoaDlvCode = ""     ''�����ù� / �ο���(�� ����)
        CASE "12" : TenDlvCode2SkstoaDlvCode = ""     ''�ѱ��ù� / ���ѱ��ù蹰��?
        CASE "13" : TenDlvCode2SkstoaDlvCode = ""   ''���ο�ĸ
        CASE "16" : TenDlvCode2SkstoaDlvCode = ""     ''�����ù�
        CASE "17" : TenDlvCode2SkstoaDlvCode = ""     ''Ʈ����ù�
        CASE "18" : TenDlvCode2SkstoaDlvCode = "14"     ''�����ù�
        CASE "20" : TenDlvCode2SkstoaDlvCode = ""   ''KT������
        CASE "21" : TenDlvCode2SkstoaDlvCode = "28"     ''�浿�ù�
        CASE "22" : TenDlvCode2SkstoaDlvCode = "36"   ''�����ù�
        CASE "24" : TenDlvCode2SkstoaDlvCode = ""     ''SC������
        CASE "25" : TenDlvCode2SkstoaDlvCode = ""   ''�ϳ����ù�
        CASE "26" : TenDlvCode2SkstoaDlvCode = "18"     ''�Ͼ��ù�
        CASE "27" : TenDlvCode2SkstoaDlvCode = ""     ''LOEX�ù�
        CASE "28" : TenDlvCode2SkstoaDlvCode = ""   ''�����ͽ�������
        CASE "29" : TenDlvCode2SkstoaDlvCode = "35"     ''�ǿ��ù�
        CASE "30" : TenDlvCode2SkstoaDlvCode = ""     ''�̳�����
        CASE "31" : TenDlvCode2SkstoaDlvCode = "34"     ''õ���ù�
        CASE "33" : TenDlvCode2SkstoaDlvCode = "40"     ''ȣ���ù� / �ϼҶ���� ��ü������� ��û
        CASE "34" : TenDlvCode2SkstoaDlvCode = "27"     ''���ȭ���ù�
        CASE "35" : TenDlvCode2SkstoaDlvCode = "30"     ''CVSnet�ù�  -
        CASE "36" : TenDlvCode2SkstoaDlvCode = ""   '��������ȭ��
        CASE "37" : TenDlvCode2SkstoaDlvCode = "29"     ''�յ��ù�  -
        CASE "38" : TenDlvCode2SkstoaDlvCode = ""     ''GTX������
        CASE "39" : TenDlvCode2SkstoaDlvCode = ""   ''KG������ - �����ͽ�������
        CASE "40" : TenDlvCode2SkstoaDlvCode = ""   ''KG������ - �����ͽ�������
        CASE "41" : TenDlvCode2SkstoaDlvCode = ""   ''�帲�ù�(�����ù�,���ο�ĸ)  ''2018/02/13
        CASE "42" : TenDlvCode2SkstoaDlvCode = "30"     ''CU POST�� �������ù�� �ش޶��..2019-03-08 ������ ����
        CASE "43" : TenDlvCode2SkstoaDlvCode = ""    'HI�ù�	http:
        CASE "44" : TenDlvCode2SkstoaDlvCode = ""  'Ȩ��	http://ww
        CASE "45" : TenDlvCode2SkstoaDlvCode = ""    'FLF�۷����ù�	h
        CASE "46" : TenDlvCode2SkstoaDlvCode = ""  'FedEx	https
        CASE "47" : TenDlvCode2SkstoaDlvCode = "37"    'Ƽ�ǿ��ڸ���	h
        CASE "48" : TenDlvCode2SkstoaDlvCode = "39"  '�������븮	h
        CASE "49" : TenDlvCode2SkstoaDlvCode = "39"  '�������븮�ù�	h
        CASE "90" : TenDlvCode2SkstoaDlvCode = ""  'EMS	http://se
        CASE "91" : TenDlvCode2SkstoaDlvCode = ""  'DHL	http://ww
        CASE "98" : TenDlvCode2SkstoaDlvCode = "99"    '������		Y
        CASE "99" : TenDlvCode2SkstoaDlvCode = "99"    '��Ÿ		Y	N
        CASE "100": TenDlvCode2SkstoaDlvCode = ""   '�ѿ츮����	h
        CASE  Else
            TenDlvCode2SkstoaDlvCode = "99"      ''��Ÿ�߼�
    end Select
end function

Function TenDlvCode2Wetoo1300kDlvCode(itenCode)
' D001	CJ�������
' D003	�����(�������)
' D004	KGB�ù�
' D005	�Ͼ��ù�
' D008	���ο�ĸ�ù�
' D009	��Ÿ
' D010	�������
' D011	�����ù�
' D015	������
' D016	��ü���ù�
' D020	�����ù�
' D021	�Ե��ù�
' D023	�����ͽ��������ù�
' D025	�̳������ù�
' D026	��ü���
' D027	õ���ù�
' D029	����ù�
' D030	�浿�ù�
' D031	Ƽ��(����߱�)
' D032	EMS(�ؿܹ��)
' D033	�¶��� �ٿ�ε�
' D034	�ǿ��ù�
' D035	�����߱�
' D036	CVSNET(������)
' D037	�յ��ù�
' D038	�ѿ츮����
' D039	GTX������
' D040	�帲�ù�
    select Case itenCode
        CASE "1" : TenDlvCode2Wetoo1300kDlvCode = "D020"     ''����
        CASE "2" : TenDlvCode2Wetoo1300kDlvCode = "D021"     '�Ե��ù�
        CASE "3" : TenDlvCode2Wetoo1300kDlvCode = "D001"     ''������� (CJ�������(CJGLS))
        CASE "4" : TenDlvCode2Wetoo1300kDlvCode = "D001"     ''CJ GLS (CJ�������(CJGLS))
        CASE "8" : TenDlvCode2Wetoo1300kDlvCode = "D016"     ''��ü���ù�
        CASE "9" : TenDlvCode2Wetoo1300kDlvCode = "D004"     ''KGB�ù�
        CASE "13" : TenDlvCode2Wetoo1300kDlvCode = "D008"     ''���ο�ĸ
        CASE "18" : TenDlvCode2Wetoo1300kDlvCode = "D011"     ''�����ù�
        CASE "21" : TenDlvCode2Wetoo1300kDlvCode = "D030"     ''�浿�ù�
        CASE "26" : TenDlvCode2Wetoo1300kDlvCode = "D005"     ''�Ͼ��ù�
        CASE "28" : TenDlvCode2Wetoo1300kDlvCode = "D023"     ''�����ͽ�������
        CASE "29" : TenDlvCode2Wetoo1300kDlvCode = "D034"     ''�ǿ��ù�
        CASE "30" : TenDlvCode2Wetoo1300kDlvCode = "D025"     ''�̳�����
        CASE "31" : TenDlvCode2Wetoo1300kDlvCode = "D027"     ''õ���ù�
        CASE "34" : TenDlvCode2Wetoo1300kDlvCode = "D029"     ''���ȭ���ù�
        CASE "35" : TenDlvCode2Wetoo1300kDlvCode = "D036"     ''CVSnet�ù�  -
        CASE "37" : TenDlvCode2Wetoo1300kDlvCode = "D037"     ''�յ��ù�  -
        CASE "38" : TenDlvCode2Wetoo1300kDlvCode = "D039"     ''GTX������
        CASE "39" : TenDlvCode2Wetoo1300kDlvCode = "D023"     ''KG������ - �����ͽ�������
        CASE "40" : TenDlvCode2Wetoo1300kDlvCode = "D023"     ''KG������ - �����ͽ�������
        CASE "41" : TenDlvCode2Wetoo1300kDlvCode = "D040"     ''�帲�ù�(�����ù�,���ο�ĸ)  ''2018/02/13
        CASE "42" : TenDlvCode2Wetoo1300kDlvCode = "D036"     ''CU POST�� �������ù�� �ش޶��..2019-03-08 ������ ����
        CASE "90" : TenDlvCode2Wetoo1300kDlvCode = "D032"    'EMS	http://se
        CASE "98" : TenDlvCode2Wetoo1300kDlvCode = "D003"    '������		Y
        CASE "99" : TenDlvCode2Wetoo1300kDlvCode = "D009"    '��Ÿ		Y	N
        CASE "100": TenDlvCode2Wetoo1300kDlvCode = "D038"     '�ѿ츮����	h
        CASE  Else
            TenDlvCode2Wetoo1300kDlvCode = "D009"      ''��Ÿ�߼�
    end Select
End Function

Function TenDlvCode2MarketforDlvCode(itenCode)
' korex       CJ �������
' yellow      ���ο�ĸ
' logen       �����ù�
' dongbu      �����ͽ��������ù�
' epost       ��ü���ù�
' hanjin      �����ù�
' hyundai     �Ե��ù�(�� �����ù�)
' kdexp       �浿�ù�
' ETC         ��Ÿ
' pantos      �������佺
' hilogis     HI �ù�
' tnt         TNT
' kgbps       KGB �ù�
' chunil      õ���ù�
' ilyang      �Ͼ������
' fedex       FEDEX
' swgexp      �����۷ι�
' daesin      ����ù�
' ups         UPS
' hdexp       �յ��ù�
' gsmnton     GSM NtoN
' daewoon 	���ѱ۷ι�
' direct 		�������
' korexg 		cj ��������Ư��
' cvsnet 		�������ù�

    select Case itenCode
        CASE "1" : TenDlvCode2MarketforDlvCode = "hanjin"     ''����
        CASE "2" : TenDlvCode2MarketforDlvCode = "hyundai"     ''���� -> �Ե��ù�� ����� 2017-03-13 ������ ����
        CASE "3" : TenDlvCode2MarketforDlvCode = "korex"     ''������� (CJ�������(CJGLS))
        CASE "4" : TenDlvCode2MarketforDlvCode = "korex"     ''CJ GLS (CJ�������(CJGLS))
        CASE "5" : TenDlvCode2MarketforDlvCode = ""     ''��Ŭ����
        CASE "6" : TenDlvCode2MarketforDlvCode = ""     ''�Ｚ HTH
        CASE "7" : TenDlvCode2MarketforDlvCode = "dongbu"     ''����(���ѹ̸�)
        CASE "8" : TenDlvCode2MarketforDlvCode = "epost"     ''��ü���ù�
        CASE "9" : TenDlvCode2MarketforDlvCode = "kgbps"     ''KGB�ù�
        CASE "10" : TenDlvCode2MarketforDlvCode = ""     ''�����ù� / �ο���(�� ����)
        CASE "11" : TenDlvCode2MarketforDlvCode = ""     ''�������ù�
        CASE "12" : TenDlvCode2MarketforDlvCode = ""     ''�ѱ��ù� / ���ѱ��ù蹰��?
        CASE "13" : TenDlvCode2MarketforDlvCode = "yellow"     ''���ο�ĸ
        CASE "14" : TenDlvCode2MarketforDlvCode = ""     ''���̽��ù�
        CASE "15" : TenDlvCode2MarketforDlvCode = ""     ''�߾��ù�
        CASE "16" : TenDlvCode2MarketforDlvCode = ""     ''�����ù�
        CASE "17" : TenDlvCode2MarketforDlvCode = ""     ''Ʈ����ù�
        CASE "18" : TenDlvCode2MarketforDlvCode = "logen"     ''�����ù�
        CASE "19" : TenDlvCode2MarketforDlvCode = ""     ''KGBƯ���ù�
        CASE "20" : TenDlvCode2MarketforDlvCode = ""     ''KT������
        CASE "21" : TenDlvCode2MarketforDlvCode = "kdexp"     ''�浿�ù�
        CASE "22" : TenDlvCode2MarketforDlvCode = ""     ''�����ù�
        CASE "23" : TenDlvCode2MarketforDlvCode = ""     ''�굦���ù� �ż���
        CASE "24" : TenDlvCode2MarketforDlvCode = ""     ''�簡���ͽ�������
        CASE "25" : TenDlvCode2MarketforDlvCode = ""     ''�ϳ����ù�
        CASE "26" : TenDlvCode2MarketforDlvCode = "ilyang"     ''�Ͼ��ù�
        CASE "27" : TenDlvCode2MarketforDlvCode = ""     ''LOEX�ù�
        CASE "28" : TenDlvCode2MarketforDlvCode = "dongbu"     ''�����ͽ�������
        CASE "29" : TenDlvCode2MarketforDlvCode = ""     ''�ǿ��ù�
        CASE "30" : TenDlvCode2MarketforDlvCode = ""     ''�̳�����
        CASE "31" : TenDlvCode2MarketforDlvCode = "chunil"     ''õ���ù�
        CASE "33" : TenDlvCode2MarketforDlvCode = ""     ''ȣ���ù�
        CASE "34" : TenDlvCode2MarketforDlvCode = "daesin"     ''���ȭ���ù�
        CASE "35" : TenDlvCode2MarketforDlvCode = "cvsnet"     ''CVSnet�ù�  -
        CASE "37" : TenDlvCode2MarketforDlvCode = "hdexp"     ''�յ��ù�  -
        CASE "38" : TenDlvCode2MarketforDlvCode = ""     ''GTX������
        CASE "39" : TenDlvCode2MarketforDlvCode = "dongbu"     ''KG������ - �����ͽ�������
        CASE "98" : TenDlvCode2MarketforDlvCode = "direct"     ''������->�����
        CASE "41" : TenDlvCode2MarketforDlvCode = "yellow"     ''�帲�ù�(�����ù�,���ο�ĸ)  ''2018/02/13
        CASE "42" : TenDlvCode2MarketforDlvCode = "cvsnet"     ''CU POST�� �������ù�� �ش޶��..2019-03-08 ������ ����
        CASE "99" : TenDlvCode2MarketforDlvCode = "ETC"     ''��Ÿ  0000033028
        CASE  Else
            TenDlvCode2MarketforDlvCode = "ETC"      ''��Ÿ�߼�
    end Select
end function

function TenDlvCode2GmarketDlvCode(itenCode)
    select Case itenCode
        CASE "1" : TenDlvCode2GmarketDlvCode = "�����ù�"     ''����
        CASE "2" : TenDlvCode2GmarketDlvCode = "�Ե��ù�"     ''���� -> �Ե��ù�� ����� 2017-03-13 ������ ����
        CASE "3" : TenDlvCode2GmarketDlvCode = "�������"     ''�������
        CASE "4" : TenDlvCode2GmarketDlvCode = "�������"     ''CJ GLS
        CASE "5" : TenDlvCode2GmarketDlvCode = "��Ÿ"     ''��Ŭ����
        CASE "6" : TenDlvCode2GmarketDlvCode = "��Ÿ"     ''�Ｚ HTH
        CASE "7" : TenDlvCode2GmarketDlvCode = "��Ÿ"     ''����(���ѹ̸�)
        CASE "8" : TenDlvCode2GmarketDlvCode = "��ü���ù�"     ''��ü���ù�
        CASE "9" : TenDlvCode2GmarketDlvCode = "KGB�ù�"     ''KGB�ù�
        CASE "10" : TenDlvCode2GmarketDlvCode = "��Ÿ"     ''�����ù� / �ο���(�� ����)
        CASE "11" : TenDlvCode2GmarketDlvCode = "��Ÿ"     ''�������ù�
        CASE "12" : TenDlvCode2GmarketDlvCode = "�ѱ��ù�"     ''�ѱ��ù� / ���ѱ��ù蹰��?
        CASE "13" : TenDlvCode2GmarketDlvCode = "���ο�ĸ�ù�"     ''���ο�ĸ
        CASE "14" : TenDlvCode2GmarketDlvCode = "��Ÿ"     ''���̽��ù�
        CASE "15" : TenDlvCode2GmarketDlvCode = "��Ÿ"     ''�߾��ù�
        CASE "16" : TenDlvCode2GmarketDlvCode = "��Ÿ"     ''�����ù�
        CASE "17" : TenDlvCode2GmarketDlvCode = "��Ÿ"     ''Ʈ����ù�
        CASE "18" : TenDlvCode2GmarketDlvCode = "�����ù�"     ''�����ù�
        CASE "19" : TenDlvCode2GmarketDlvCode = "KGB�ù�"     ''KGBƯ���ù�
        CASE "20" : TenDlvCode2GmarketDlvCode = "��Ÿ"     ''KT������
        CASE "21" : TenDlvCode2GmarketDlvCode = "�浿�ù�"     ''�浿�ù�
        CASE "22" : TenDlvCode2GmarketDlvCode = "��Ÿ"     ''�����ù�
        CASE "23" : TenDlvCode2GmarketDlvCode = "��Ÿ"     ''�굦���ù� �ż���
        CASE "24" : TenDlvCode2GmarketDlvCode = "��Ÿ"     ''�簡���ͽ�������
        CASE "25" : TenDlvCode2GmarketDlvCode = "��Ÿ"     ''�ϳ����ù�
        CASE "26" : TenDlvCode2GmarketDlvCode = "�Ͼ��ù�"     ''�Ͼ��ù�
        CASE "27" : TenDlvCode2GmarketDlvCode = "��Ÿ"     ''LOEX�ù�
        CASE "28" : TenDlvCode2GmarketDlvCode = "��Ÿ"     ''�����ͽ�������
        CASE "29" : TenDlvCode2GmarketDlvCode = "�ǿ��ù�"     ''�ǿ��ù�
        CASE "30" : TenDlvCode2GmarketDlvCode = "��Ÿ"     ''�̳�����
        CASE "31" : TenDlvCode2GmarketDlvCode = "õ���ù�"     ''õ���ù�
        CASE "33" : TenDlvCode2GmarketDlvCode = "ȣ���ù�"     ''ȣ���ù�
        CASE "34" : TenDlvCode2GmarketDlvCode = "����ù�"     ''���ȭ���ù�
        CASE "35" : TenDlvCode2GmarketDlvCode = "�������ù�(GS25)" ''"CVSNET(������)"     ''CVSnet�ù�  - 99 ��Ÿ�߼����ù�  ''2019/08/20 ����
        CASE "38" : TenDlvCode2GmarketDlvCode = "GTX������"     ''GTX������
        'CASE "39" : TenDlvCode2GmarketDlvCode = "KG������"     ''KG������ - �����ͽ�������
        CASE "39" : TenDlvCode2GmarketDlvCode = "�帲�ù�"     ''2018-02-23 ���� ����
        CASE "42" : TenDlvCode2GmarketDlvCode = "�������ù�(GS25)"     ''CU�������ù�
        CASE "98" : TenDlvCode2GmarketDlvCode = "������"     ''������->�����
        CASE "99" : TenDlvCode2GmarketDlvCode = "�������"     ''��Ÿ
        CASE "102" : TenDlvCode2GmarketDlvCode = "�������"     ''�����
        CASE  Else
            TenDlvCode2GmarketDlvCode = "��Ÿ"      ''��Ÿ�߼�
    end Select
end function

Function TenDlvCode2NvstorefarmDlvCode(itenCode)
    select Case itenCode
        CASE "1" : TenDlvCode2NvstorefarmDlvCode = "HANJIN"     ''����
        CASE "2" : TenDlvCode2NvstorefarmDlvCode = "HYUNDAI"     ''����
        CASE "3" : TenDlvCode2NvstorefarmDlvCode = "CJGLS"     ''�������
        CASE "4" : TenDlvCode2NvstorefarmDlvCode = "CJGLS"     ''CJ GLS
        CASE "5" : TenDlvCode2NvstorefarmDlvCode = ""     ''��Ŭ����
        CASE "6" : TenDlvCode2NvstorefarmDlvCode = ""     ''�Ｚ HTH
        CASE "7" : TenDlvCode2NvstorefarmDlvCode = "DONGBU"     ''����(���ѹ̸�)
        CASE "8" : TenDlvCode2NvstorefarmDlvCode = "EPOST"     ''��ü���ù�
        CASE "9" : TenDlvCode2NvstorefarmDlvCode = "KGBLS"     ''KGB�ù�
        CASE "10" : TenDlvCode2NvstorefarmDlvCode = ""     ''�����ù� / �ο���(�� ����)
        CASE "11" : TenDlvCode2NvstorefarmDlvCode = ""     ''�������ù�
        CASE "12" : TenDlvCode2NvstorefarmDlvCode = ""     ''�ѱ��ù� / ���ѱ��ù蹰��?
        CASE "13" : TenDlvCode2NvstorefarmDlvCode = "YELLOW"     ''���ο�ĸ
        CASE "14" : TenDlvCode2NvstorefarmDlvCode = ""     ''���̽��ù�
        CASE "15" : TenDlvCode2NvstorefarmDlvCode = ""     ''�߾��ù�
        CASE "16" : TenDlvCode2NvstorefarmDlvCode = ""     ''�����ù�
        CASE "17" : TenDlvCode2NvstorefarmDlvCode = ""     ''Ʈ����ù�
        CASE "18" : TenDlvCode2NvstorefarmDlvCode = "KGB"     ''�����ù�
        CASE "19" : TenDlvCode2NvstorefarmDlvCode = ""     ''KGBƯ���ù�
        CASE "20" : TenDlvCode2NvstorefarmDlvCode = ""     ''KT������
        CASE "21" : TenDlvCode2NvstorefarmDlvCode = "KDEXP"     ''�浿�ù�
        CASE "22" : TenDlvCode2NvstorefarmDlvCode = ""     ''�����ù�
        CASE "23" : TenDlvCode2NvstorefarmDlvCode = ""     ''�굦���ù� �ż���
        CASE "24" : TenDlvCode2NvstorefarmDlvCode = ""     ''�簡���ͽ�������
        CASE "25" : TenDlvCode2NvstorefarmDlvCode = ""     ''�ϳ����ù�
        CASE "26" : TenDlvCode2NvstorefarmDlvCode = "ILYANG"     ''�Ͼ��ù�
        CASE "27" : TenDlvCode2NvstorefarmDlvCode = ""     ''LOEX�ù�
        CASE "28" : TenDlvCode2NvstorefarmDlvCode = "DONGBU"     ''�����ͽ�������
        CASE "29" : TenDlvCode2NvstorefarmDlvCode = ""     ''�ǿ��ù�
        CASE "30" : TenDlvCode2NvstorefarmDlvCode = ""     ''�̳�����
        CASE "31" : TenDlvCode2NvstorefarmDlvCode = "CHUNIL"     ''õ���ù�
        CASE "33" : TenDlvCode2NvstorefarmDlvCode = ""     ''ȣ���ù�
        CASE "34" : TenDlvCode2NvstorefarmDlvCode = "DAESIN"     ''���ȭ���ù�
        CASE "35" : TenDlvCode2NvstorefarmDlvCode = "CVSNET"     ''CVSnet�ù�  - 99 ��Ÿ�߼����ù�
        CASE "37" : TenDlvCode2NvstorefarmDlvCode = "HDEXP"     ''�յ��ù�
        CASE "38" : TenDlvCode2NvstorefarmDlvCode = "INNOGIS"     ''GTX������   ''GTX(��ī�̷�����)::2586778  ''2015/06/29 �߰�
        CASE "42" : TenDlvCode2NvstorefarmDlvCode = "CUPARCEL"     ''CU�������ù�
        CASE "98" : TenDlvCode2NvstorefarmDlvCode = "ETC1"     ''������->����� | 2019-04-11 ������..ETC1 �߰� �� �޴� ó��
        CASE "99" : TenDlvCode2NvstorefarmDlvCode = "ETC2"     ''��Ÿ | 2019-04-11 ������..ETC2 �߰� �� �޴� ó��
        CASE  Else
            TenDlvCode2NvstorefarmDlvCode = "CH1"      ''��Ÿ�߼�
    end Select
end function

Function TenDlvCode211stDlvCode(itenCode)
    select Case itenCode
        CASE "1" : TenDlvCode211stDlvCode = "00011"     ''����
        CASE "2" : TenDlvCode211stDlvCode = "00012"    ''����(�Ե�)�ù�
        CASE "3" : TenDlvCode211stDlvCode = "00034"     ''�������
        CASE "4" : TenDlvCode211stDlvCode = "00034"     ''CJ GLS
        CASE "8" : TenDlvCode211stDlvCode = "00007"     ''��ü���ù�
        CASE "18" : TenDlvCode211stDlvCode = "00002"     ''�����ù�
        CASE "21" : TenDlvCode211stDlvCode = "00026"     ''�浿�ù�
        CASE "26" : TenDlvCode211stDlvCode = "00022"     ''�Ͼ��ù�
        CASE "29" : TenDlvCode211stDlvCode = "00037"     ''�ǿ��ù�
        CASE "31" : TenDlvCode211stDlvCode = "00027"     ''õ���ù�
        CASE "37" : TenDlvCode211stDlvCode = "00035"     ''�յ��ù�
        CASE "38" : TenDlvCode211stDlvCode = "00033"     ''GTX������   ''GTX(��ī�̷�����)::2586778  ''2015/06/29 �߰�
        CASE "39" : TenDlvCode211stDlvCode = "00001"     ''KG������ - �����ͽ�������
        CASE "99" : TenDlvCode211stDlvCode = "00099"     ''��Ÿ

        CASE "34" : TenDlvCode211stDlvCode = "00021"     ''���(ȭ��)�ù�
        CASE "35" : TenDlvCode211stDlvCode = "00060"     ''CVSnet�ù�
        CASE "42" : TenDlvCode211stDlvCode = "00061"     ''CU POST

        CASE  Else
            TenDlvCode211stDlvCode = "00099"      ''��Ÿ�߼�
    end Select
end function

Function TenDlvCode2HalfClubDlvCode(itenCode)
	select Case itenCode
		CASE "1" : TenDlvCode2HalfClubDlvCode = "3"			''����
		CASE "2" : TenDlvCode2HalfClubDlvCode = "6"			''���� -> �Ե�
		CASE "3" : TenDlvCode2HalfClubDlvCode = "7"			''�������
		CASE "4" : TenDlvCode2HalfClubDlvCode = "1"			''CJ GLS
		CASE "8" : TenDlvCode2HalfClubDlvCode = "4"			''��ü���ù�
		CASE "9" : TenDlvCode2HalfClubDlvCode = "23"		''KGB�ù�
		CASE "10" : TenDlvCode2HalfClubDlvCode = "27"		''�����ù� / �ο���(�� ����)
		CASE "13" : TenDlvCode2HalfClubDlvCode = "10"		''���ο�ĸ
		CASE "17" : TenDlvCode2HalfClubDlvCode = "16"		''Ʈ����ù�
		CASE "18" : TenDlvCode2HalfClubDlvCode = "8"		''�����ù�
		CASE "19" : TenDlvCode2HalfClubDlvCode = "23"		''KGBƯ���ù�
		CASE "21" : TenDlvCode2HalfClubDlvCode = "15"		''�浿�ù�
		CASE "22" : TenDlvCode2HalfClubDlvCode = "24"		''�����ù�
		CASE "25" : TenDlvCode2HalfClubDlvCode = "25"		''�ϳ����ù�
		CASE "26" : TenDlvCode2HalfClubDlvCode = "32"		''�Ͼ��ù�
		CASE "27" : TenDlvCode2HalfClubDlvCode = "27"		''LOEX�ù�
		CASE "29" : TenDlvCode2HalfClubDlvCode = "56"		''�ǿ��ù�
		CASE "30" : TenDlvCode2HalfClubDlvCode = "30"		''�̳�����
		CASE "31" : TenDlvCode2HalfClubDlvCode = "40"		''õ���ù�
		CASE "33" : TenDlvCode2HalfClubDlvCode = "54"		''ȣ���ù�
		CASE "34" : TenDlvCode2HalfClubDlvCode = "33"		''���ȭ���ù�
		CASE "37" : TenDlvCode2HalfClubDlvCode = "37"		''�յ��ù�  -
		CASE "38" : TenDlvCode2HalfClubDlvCode = "48"		''GTX������
		CASE "41" : TenDlvCode2HalfClubDlvCode = "26"		''�帲�ù�(�����ù�,���ο�ĸ)  ''2018/02/13
		CASE "98" : TenDlvCode2HalfClubDlvCode = "39"		''������->�����
		CASE  Else
			TenDlvCode2HalfClubDlvCode = "0"      ''���Է�
	end Select
end function

Function TenDlvCode2CoupangDlvCode(itenCode)
	select Case itenCode
		CASE "1" : TenDlvCode2CoupangDlvCode = "HANJIN"			''����
		CASE "2" : TenDlvCode2CoupangDlvCode = "HYUNDAI"			''���� -> �Ե�
		CASE "3" : TenDlvCode2CoupangDlvCode = "CJGLS"			''�������
		CASE "4" : TenDlvCode2CoupangDlvCode = "CJGLS"			''CJ GLS
		CASE "5" : TenDlvCode2CoupangDlvCode = "CSLOGIS"
		CASE "8" : TenDlvCode2CoupangDlvCode = "EPOST"			''��ü���ù�
		CASE "9" : TenDlvCode2CoupangDlvCode = "KGBPS"		''KGB�ù�
		CASE "10" : TenDlvCode2CoupangDlvCode = "AJOU"		''�����ù� / �ο���(�� ����)
		CASE "18" : TenDlvCode2CoupangDlvCode = "KGB"		''�����ù�
		CASE "21" : TenDlvCode2CoupangDlvCode = "KDEXP"		''�浿�ù�
		CASE "24" : TenDlvCode2CoupangDlvCode = "CSLOGIS"
		CASE "26" : TenDlvCode2CoupangDlvCode = "ILYANG"		''�Ͼ��ù�
		CASE "29" : TenDlvCode2CoupangDlvCode = "KUNYOUNG"		''�ǿ��ù�
		CASE "31" : TenDlvCode2CoupangDlvCode = "CHUNIL"		''õ���ù�
		CASE "33" : TenDlvCode2CoupangDlvCode = "HONAM"		''ȣ���ù�
		CASE "34" : TenDlvCode2CoupangDlvCode = "DAESIN"		''���ȭ���ù�
		CASE "35" : TenDlvCode2CoupangDlvCode = "CVS"
		CASE "36" : TenDlvCode2CoupangDlvCode = "HANJIN"
		CASE "37" : TenDlvCode2CoupangDlvCode = "HDEXP"		''�յ��ù�  -
		CASE "39" : TenDlvCode2CoupangDlvCode = "DONGBU"
		CASE "41" : TenDlvCode2CoupangDlvCode = "DONGBU"
        CASE "42" : TenDlvCode2CoupangDlvCode = "BGF"       'CU POST�� BGF����Ʈ�� ����
        CASE "47" : TenDlvCode2CoupangDlvCode = "TPMLOGIS"	''Ƽ�ǿ�������(�����Ư��)
        CASE "54" : TenDlvCode2CoupangDlvCode = "DIRECT"	'NDEX KOREA�� ��ü�������� ��û..20220721 ������B ��û
		CASE "91" : TenDlvCode2CoupangDlvCode = "DHL"
        CASE "98" : TenDlvCode2CoupangDlvCode = "DIRECT"    '�����񽺸� ��ü�������� ��û..20190411 �ϼҶ�� ��û
		CASE "99" : TenDlvCode2CoupangDlvCode = "DIRECT"
	end Select
end function

Function TenDlvCode2HmallDlvCode(itenCode)
'1	11	�Ե��ù�
'2	12	CJ�������
'3	13	�����ù�
'4	16	�ڰ����
'5	24	�½�����Ʈ
'6	25	���븮��Ʈ
'7	29	KGB�ù�
'8	31	�������
'9	33	�����ù�
'10	35	��ü���ù�
'11	38	�Ͼ��ù�
'12	60	������
'13	61	�浿�ù�
'14	63	�Ѽ�ȣ���ù�
'15	64	õ���ù�
'16	65	�������ù�(GS25)
'17	68	�յ��ù�
'18	69	����ù�
'19	70	�ǿ��ù�
'20	71	GTX������
'21	72	���ǻ���ù�
'22	78	�����ù�
'23	79	�����ù�
'24	83	TNT express
'25	84	�������佺
'26	89	DHL

	SELECT Case itenCode
		CASE "1" : TenDlvCode2HmallDlvCode = "13"		''����
		CASE "2" : TenDlvCode2HmallDlvCode = "11"		''���� -> �Ե�
		CASE "3" : TenDlvCode2HmallDlvCode = "12"		''�������
		CASE "4" : TenDlvCode2HmallDlvCode = "12"		''CJ GLS
		CASE "8" : TenDlvCode2HmallDlvCode = "35"		''��ü���ù�
		CASE "9" : TenDlvCode2HmallDlvCode = "29"		''KGB�ù�
		CASE "18" : TenDlvCode2HmallDlvCode = "33"		''�����ù�
		CASE "21" : TenDlvCode2HmallDlvCode = "61"		''�浿�ù�
		CASE "26" : TenDlvCode2HmallDlvCode = "38"		''�Ͼ��ù�
		CASE "29" : TenDlvCode2HmallDlvCode = "70"		''�ǿ��ù�
		CASE "31" : TenDlvCode2HmallDlvCode = "64"		''õ���ù�
		CASE "33" : TenDlvCode2HmallDlvCode = "63"		''ȣ���ù�
		CASE "34" : TenDlvCode2HmallDlvCode = "69"		''���ȭ���ù�
		CASE "35" : TenDlvCode2HmallDlvCode = "65"
		CASE "37" : TenDlvCode2HmallDlvCode = "68"		''�յ��ù�  -
        CASE "38" : TenDlvCode2HmallDlvCode = "71"		''GTX������
        CASE "42" : TenDlvCode2HmallDlvCode = "65"		''CU POST
        CASE "45" : TenDlvCode2HmallDlvCode = "74"		''FLF�۷����ù�
        CASE "47" : TenDlvCode2HmallDlvCode = "93"		''Ƽ�ǿ�������
		CASE "91" : TenDlvCode2HmallDlvCode = "89"
        CASE "98" : TenDlvCode2HmallDlvCode = "60"		''������
        CASE "99" : TenDlvCode2HmallDlvCode = "16"
	End Select
end function

Function TenDlvCode2WMPDlvCode(itenCode)
' D001 : ��ü���ù�
' D002 : CJ�������
' D003 : �����ù�
' D005 : �Ե��ù�
' D004 : �����ù�
' D006 : KGB�ù�
' D011 : GTX������
' D007 : �Ͼ������
' D008 : EMS
' D009 : DHL
' D010 : UPS
' D012 : ���ǻ���ù�
' D013 : õ���ù�
' D014 : �ǿ��ù�
' D015 : �����ù�
' D016 : �ѵ���
' D017 : Fedex
' D018 : ����ù�
' D019 : �浿�ù�
' D020 : CVSnet �������ù�
' D021 : TNT Express
' D040 : CU �������ù�
' D022 : USPS
' D041 : �����ù�
' D023 : TPL
' D042 : ����
' D043 : �Ｚ���� ����
' D024 : GSMNtoN
' D025 : ������ͽ�������
' D026 : KGL��ũ����
' D027 : �յ��ù�
' D028 : DHL Global Mail
' D029 : i-Parcel
' D030 : ������ �ͽ�������
' D031 : �������佺
' D032 : APEX(ECMS Express)
' D034 : ������
' D035 : GSI Express
' D036 : CJ������� ����Ư��
' D037 : SLX
' D038 : ȣ���ù�
' D039 : �����ù� �ؿ�Ư��
' D046 : GPS Logix
' D045 : Ȩ���ù�
' D044 : LG���� ����
	SELECT Case itenCode
		CASE "1" : TenDlvCode2WMPDlvCode = "D003"		''����
		CASE "2" : TenDlvCode2WMPDlvCode = "D005"		''���� -> �Ե�
		CASE "3" : TenDlvCode2WMPDlvCode = "D002"		''�������
		CASE "4" : TenDlvCode2WMPDlvCode = "D002"		''CJ GLS
		CASE "8" : TenDlvCode2WMPDlvCode = "D001"		''��ü���ù�
		CASE "9" : TenDlvCode2WMPDlvCode = "D006"		''KGB�ù�
		CASE "18" : TenDlvCode2WMPDlvCode = "D004"		''�����ù�
		CASE "21" : TenDlvCode2WMPDlvCode = "D019"		''�浿�ù�
        CASE "22" : TenDlvCode2WMPDlvCode = "D015"		''�����ù�
		CASE "26" : TenDlvCode2WMPDlvCode = "D007"		''�Ͼ��ù�
		CASE "29" : TenDlvCode2WMPDlvCode = "D014"		''�ǿ��ù�
		CASE "31" : TenDlvCode2WMPDlvCode = "D013"		''õ���ù�
		CASE "33" : TenDlvCode2WMPDlvCode = "D038"		''ȣ���ù�
		CASE "34" : TenDlvCode2WMPDlvCode = "D018"		''���ȭ���ù�
		CASE "35" : TenDlvCode2WMPDlvCode = "D020"      ''CVSnet�ù�
		CASE "37" : TenDlvCode2WMPDlvCode = "D027"		''�յ��ù�  -
        CASE "38" : TenDlvCode2WMPDlvCode = "D011"		''GTX������
        CASE "42" : TenDlvCode2WMPDlvCode = "D040"		''CU Post => �������
		CASE "91" : TenDlvCode2WMPDlvCode = "D009"      ''DHL
        CASE  Else
            TenDlvCode2WMPDlvCode = "0"     ''���Է�
	End Select
end function

Function TenDlvCode2SabangNetDlvCode(itenCode)
	select Case itenCode
		CASE "1" : TenDlvCode2SabangNetDlvCode = "004"			''����
		CASE "2" : TenDlvCode2SabangNetDlvCode = "002"			''���� -> �Ե�
		CASE "3" : TenDlvCode2SabangNetDlvCode = "001"			''�������
		CASE "4" : TenDlvCode2SabangNetDlvCode = "003"			''CJ GLS
		CASE "8" : TenDlvCode2SabangNetDlvCode = "009"			''��ü���ù�
		CASE "9" : TenDlvCode2SabangNetDlvCode = "005"		''KGB�ù�
		CASE "10" : TenDlvCode2SabangNetDlvCode = "021"		''�����ù� / �ο���(�� ����)
		CASE "13" : TenDlvCode2SabangNetDlvCode = "032"		''���ο�ĸ
		CASE "17" : TenDlvCode2SabangNetDlvCode = "022"		''Ʈ����ù�
		CASE "18" : TenDlvCode2SabangNetDlvCode = "007"		''�����ù�
		CASE "19" : TenDlvCode2SabangNetDlvCode = "035"		''KGBƯ���ù�
		CASE "21" : TenDlvCode2SabangNetDlvCode = "013"		''�浿�ù�
		CASE "22" : TenDlvCode2SabangNetDlvCode = "044"		''�����ù�
		CASE "25" : TenDlvCode2SabangNetDlvCode = "010"		''�ϳ����ù�
		CASE "26" : TenDlvCode2SabangNetDlvCode = "047"		''�Ͼ��ù�
		CASE "27" : TenDlvCode2SabangNetDlvCode = "011"		''LOEX�ù�
		CASE "29" : TenDlvCode2SabangNetDlvCode = "043"		''�ǿ��ù�
		CASE "30" : TenDlvCode2SabangNetDlvCode = "023"		''�̳�����
		CASE "31" : TenDlvCode2SabangNetDlvCode = "016"		''õ���ù�
		CASE "33" : TenDlvCode2SabangNetDlvCode = "033"		''ȣ���ù�
		CASE "34" : TenDlvCode2SabangNetDlvCode = "037"		''���ȭ���ù�
		CASE "37" : TenDlvCode2SabangNetDlvCode = "056"		''�յ��ù�  -
		CASE "38" : TenDlvCode2SabangNetDlvCode = "053"		''GTX������
		CASE "41" : TenDlvCode2SabangNetDlvCode = "104"		''�帲�ù�(�����ù�,���ο�ĸ)  ''2018/02/13
		CASE "98" : TenDlvCode2SabangNetDlvCode = "999"		''������->�����
        CASE "99" : TenDlvCode2SabangNetDlvCode = "999"		''��Ÿ
		CASE  Else
			TenDlvCode2SabangNetDlvCode = "0"      ''���Է�
	end Select
end function

'' ssg���� ��ü�������  �ҽ�����
function TenDlvCode2SSGDlvCode(itenCode)
''<option value="0000033023">SC������</option> ? �簡���ͽ�������
''<option value="0000033026">����ù�</option>
''<option value="0000033029">�׵����ù�</option>
''<option value="0000033032">�������</option> = >CJ�������
''<option value="0000033033">KG������(�����ù�,���ο�ĸ)</option>
''<option value="0000033034">�����ù�</option>
'<option value="0000033050">��ü��EMS</option>
'<option value="0000033051">��ü�����</option>
'<option value="0000033052">��ü���ù�</option>
'<option value="0000033063">�ڵ���</option>
'<option value="0000033064">��/�ݺ�</option>
''<option value="0008369131">�������ù�</option>
    select Case itenCode
        CASE "1" : TenDlvCode2SSGDlvCode = "0000033071"     ''����
        CASE "2" : TenDlvCode2SSGDlvCode = "0000033073"     ''���� -> �Ե��ù�� ����� 2017-03-13 ������ ����
        CASE "3" : TenDlvCode2SSGDlvCode = "0000033011"     ''������� (CJ�������(CJGLS))
        CASE "4" : TenDlvCode2SSGDlvCode = "0000033011"     ''CJ GLS (CJ�������(CJGLS))
        CASE "5" : TenDlvCode2SSGDlvCode = "0000033056"     ''��Ŭ����
        CASE "6" : TenDlvCode2SSGDlvCode = ""     ''�Ｚ HTH
        CASE "7" : TenDlvCode2SSGDlvCode = "0000033033"     ''����(���ѹ̸�)
        CASE "8" : TenDlvCode2SSGDlvCode = "0000033052"     ''��ü���ù�
        CASE "9" : TenDlvCode2SSGDlvCode = "0000033017"     ''KGB�ù�
        CASE "10" : TenDlvCode2SSGDlvCode = "0000033035"     ''�����ù� / �ο���(�� ����)
        CASE "11" : TenDlvCode2SSGDlvCode = ""     ''�������ù�
        CASE "12" : TenDlvCode2SSGDlvCode = "0000033069"     ''�ѱ��ù� / ���ѱ��ù蹰��?
        CASE "13" : TenDlvCode2SSGDlvCode = "0000033033"     ''���ο�ĸ
        CASE "14" : TenDlvCode2SSGDlvCode = ""     ''���̽��ù�
        CASE "15" : TenDlvCode2SSGDlvCode = "0000033061"     ''�߾��ù�
        CASE "16" : TenDlvCode2SSGDlvCode = "0000033060"     ''�����ù�
        CASE "17" : TenDlvCode2SSGDlvCode = "0000033067"     ''Ʈ����ù�
        CASE "18" : TenDlvCode2SSGDlvCode = "0000033036"     ''�����ù�
        CASE "19" : TenDlvCode2SSGDlvCode = "0000033018"     ''KGBƯ���ù�
        CASE "20" : TenDlvCode2SSGDlvCode = "0000033021"     ''KT������
        CASE "21" : TenDlvCode2SSGDlvCode = "0000033027"     ''�浿�ù�
        CASE "22" : TenDlvCode2SSGDlvCode = ""     ''�����ù�
        CASE "23" : TenDlvCode2SSGDlvCode = ""     ''�굦���ù� �ż���
        CASE "24" : TenDlvCode2SSGDlvCode = ""     ''�簡���ͽ�������
        CASE "25" : TenDlvCode2SSGDlvCode = "0000033068"     ''�ϳ����ù�
        CASE "26" : TenDlvCode2SSGDlvCode = "0000033057"     ''�Ͼ��ù�
        CASE "27" : TenDlvCode2SSGDlvCode = ""     ''LOEX�ù�
        CASE "28" : TenDlvCode2SSGDlvCode = "0000033033"     ''�����ͽ�������
        CASE "29" : TenDlvCode2SSGDlvCode = "0000033025"     ''�ǿ��ù�
        CASE "30" : TenDlvCode2SSGDlvCode = "0000033055"     ''�̳�����
        CASE "31" : TenDlvCode2SSGDlvCode = "0000033062"     ''õ���ù�
        CASE "33" : TenDlvCode2SSGDlvCode = "0000033077"     ''ȣ���ù�
        CASE "34" : TenDlvCode2SSGDlvCode = "0000033030"     ''���ȭ���ù�
        CASE "35" : TenDlvCode2SSGDlvCode = "0000033013"     ''CVSnet�ù�  -
        CASE "37" : TenDlvCode2SSGDlvCode = "0000038977"     ''�յ��ù�  -
        CASE "38" : TenDlvCode2SSGDlvCode = "0000033014"     ''GTX������
        CASE "39" : TenDlvCode2SSGDlvCode = "0000033033"     ''KG������ - �����ͽ�������
        CASE "98" : TenDlvCode2SSGDlvCode = "0000033064"     ''������->�����
        CASE "41" : TenDlvCode2SSGDlvCode = "0000033033"     ''�帲�ù�(�����ù�,���ο�ĸ)  ''2018/02/13
        CASE "42" : TenDlvCode2SSGDlvCode = "0008369131"     ''CU POST�� �������ù�� �ش޶��..2019-03-08 ������ ����
        CASE "99" : TenDlvCode2SSGDlvCode = "0000033028"     ''��Ÿ (��Ÿ�ù��) 0000033028
        CASE  Else
            TenDlvCode2SSGDlvCode = "��Ÿ"      ''��Ÿ�߼�
    end Select
end function

function TenDlvCode2HomeplusDlvCode(itenCode)
    select Case itenCode
        CASE "1" : TenDlvCode2HomeplusDlvCode = "�����ù�"     ''����
        CASE "2" : TenDlvCode2HomeplusDlvCode = "�����ù�"     ''����
        CASE "3" : TenDlvCode2HomeplusDlvCode = "�������"     ''�������
        CASE "4" : TenDlvCode2HomeplusDlvCode = "CJGLS"     ''CJ GLS
        CASE "5" : TenDlvCode2HomeplusDlvCode = "��Ŭ�����ù�"     ''��Ŭ����
        CASE "6" : TenDlvCode2HomeplusDlvCode = "CJHTH"     ''�Ｚ HTH
        CASE "7" : TenDlvCode2HomeplusDlvCode = "�ѹ̸��ù�"     ''����(���ѹ̸�)
        CASE "8" : TenDlvCode2HomeplusDlvCode = "��ü���ù�"     ''��ü���ù�
        CASE "9" : TenDlvCode2HomeplusDlvCode = "KGB�ù�"     ''KGB�ù�
        CASE "10" : TenDlvCode2HomeplusDlvCode = "�����ù�"     ''�����ù� / �ο���(�� ����)
        CASE "11" : TenDlvCode2HomeplusDlvCode = ""     ''�������ù�
        CASE "12" : TenDlvCode2HomeplusDlvCode = "�ѱ��ù�"     ''�ѱ��ù� / ���ѱ��ù蹰��?
        CASE "13" : TenDlvCode2HomeplusDlvCode = "���ο�ĸ"     ''���ο�ĸ
        CASE "14" : TenDlvCode2HomeplusDlvCode = ""     ''���̽��ù�
        CASE "15" : TenDlvCode2HomeplusDlvCode = ""     ''�߾��ù�
        CASE "16" : TenDlvCode2HomeplusDlvCode = ""     ''�����ù�
        CASE "17" : TenDlvCode2HomeplusDlvCode = "Ʈ����ù�"     ''Ʈ����ù�
        CASE "18" : TenDlvCode2HomeplusDlvCode = "�����ù�"     ''�����ù�
        CASE "19" : TenDlvCode2HomeplusDlvCode = ""     ''KGBƯ���ù�
        CASE "20" : TenDlvCode2HomeplusDlvCode = "KT�������ù�"     ''KT������
        CASE "21" : TenDlvCode2HomeplusDlvCode = "�浿�ù�"     ''�浿�ù�
        CASE "22" : TenDlvCode2HomeplusDlvCode = ""     ''�����ù�
        CASE "23" : TenDlvCode2HomeplusDlvCode = "�굦��"     ''�굦���ù� �ż���
        CASE "24" : TenDlvCode2HomeplusDlvCode = "�簡���ù�"     ''�簡���ͽ�������
        CASE "25" : TenDlvCode2HomeplusDlvCode = "�ϳ����ù�"     ''�ϳ����ù�
        CASE "26" : TenDlvCode2HomeplusDlvCode = "��Ÿ�ù�"     ''�Ͼ��ù�
        CASE "27" : TenDlvCode2HomeplusDlvCode = ""     ''LOEX�ù�
        CASE "28" : TenDlvCode2HomeplusDlvCode = "�����ù�"     ''�����ͽ�������
        CASE "29" : TenDlvCode2HomeplusDlvCode = ""     ''�ǿ��ù�	'27310
        CASE "30" : TenDlvCode2HomeplusDlvCode = "�̳������ù�"     ''�̳�����
        CASE "31" : TenDlvCode2HomeplusDlvCode = "õ���ù�"     ''õ���ù�
        CASE "33" : TenDlvCode2HomeplusDlvCode = ""     ''ȣ���ù�
        CASE "34" : TenDlvCode2HomeplusDlvCode = "����ù�"     ''���ȭ���ù�
        CASE "35" : TenDlvCode2HomeplusDlvCode = "��Ÿ�ù�"     ''CVSnet�ù�  - 99 ��Ÿ�߼����ù�
        CASE "98" : TenDlvCode2HomeplusDlvCode = "�����"     ''������->�����
        CASE "99" : TenDlvCode2HomeplusDlvCode = "��Ÿ�ù�"     ''��Ÿ
        CASE  Else
            TenDlvCode2HomeplusDlvCode = "��Ÿ�ù�"      ''��Ÿ�߼�
    end Select
end function

function TenDlvCode2EzwelDlvCode(itenCode)
    select Case itenCode
        CASE "1" : TenDlvCode2EzwelDlvCode = "1016"     ''����
        CASE "2" : TenDlvCode2EzwelDlvCode = "1017"     ''����(�Ե�)
        CASE "3" : TenDlvCode2EzwelDlvCode = "1007"     ''�������
        CASE "4" : TenDlvCode2EzwelDlvCode = "1007"     ''CJ GLS
        CASE "8" : TenDlvCode2EzwelDlvCode = "1012"     ''��ü���ù�
        CASE "9" : TenDlvCode2EzwelDlvCode = "1002"     ''KGB�ù�
        CASE "13" : TenDlvCode2EzwelDlvCode = "1011"     ''���ο�ĸ
        CASE "18" : TenDlvCode2EzwelDlvCode = "1008"     ''�����ù�
        CASE "20" : TenDlvCode2EzwelDlvCode = "1082"     ''KT������
        CASE "21" : TenDlvCode2EzwelDlvCode = "1005"     ''�浿�ù�
        CASE "24" : TenDlvCode2EzwelDlvCode = "1160"     ''�簡���ͽ�������
        CASE "26" : TenDlvCode2EzwelDlvCode = "1180"     ''�Ͼ��ù�
        CASE "28" : TenDlvCode2EzwelDlvCode = "1080"     ''�����ͽ�������
        CASE "29" : TenDlvCode2EzwelDlvCode = "1106"     ''�ǿ��ù�
        CASE "30" : TenDlvCode2EzwelDlvCode = "1163"     ''�̳�����
        CASE "31" : TenDlvCode2EzwelDlvCode = "1014"     ''õ���ù�
        CASE "33" : TenDlvCode2EzwelDlvCode = "1107"     ''ȣ���ù�
        CASE "34" : TenDlvCode2EzwelDlvCode = "1200"     ''���ȭ���ù�
        CASE "35" : TenDlvCode2EzwelDlvCode = "1007" ''"1240"     ''CVSnet�ù�  - 99 ��Ÿ�߼����ù� ''2019.08/20 �����ԷºҰ� CJ�� ����(CJ�� ������)
        CASE "37" : TenDlvCode2EzwelDlvCode = "1102"     ''�յ��ù�
        CASE "38" : TenDlvCode2EzwelDlvCode = "1260"     ''GTX������   ''GTX(��ī�̷�����)::2586778  ''2015/06/29 �߰�
        CASE "39" : TenDlvCode2EzwelDlvCode = "1080"     ''KG������
        CASE "41" : TenDlvCode2EzwelDlvCode = "1080"     ''�帲�ù�(�����ù�,���ο�ĸ)  ''2018/02/13
        CASE "91" : TenDlvCode2EzwelDlvCode = "1001"
        CASE "98" : TenDlvCode2EzwelDlvCode = "1081"     ''������->�����
        CASE "99" : TenDlvCode2EzwelDlvCode = "1082"     ''��Ÿ
        CASE  Else
            TenDlvCode2EzwelDlvCode = "1082"      ''��Ÿ�߼�
    end Select
end function

function TenDlvCode2cjMallDlvCode(itenCode)
    select Case itenCode
        CASE "1" : TenDlvCode2cjMallDlvCode = "15"     ''����
        CASE "2" : TenDlvCode2cjMallDlvCode = "11"     ''����
        CASE "3" : TenDlvCode2cjMallDlvCode = "22"     ''�������
        CASE "4" : TenDlvCode2cjMallDlvCode = "22"     ''CJ GLS
        CASE "5" : TenDlvCode2cjMallDlvCode = "21"     ''��Ŭ����
        CASE "6" : TenDlvCode2cjMallDlvCode = "29"     ''�Ｚ HTH
        CASE "7" : TenDlvCode2cjMallDlvCode = "79"     ''����(���ѹ̸�)
        CASE "8" : TenDlvCode2cjMallDlvCode = "16"     ''��ü���ù�
        CASE "9" : TenDlvCode2cjMallDlvCode = "93"     ''KGB�ù�
        CASE "10" : TenDlvCode2cjMallDlvCode = "67"     ''�����ù� / �ο���(�� ����)
        CASE "11" : TenDlvCode2cjMallDlvCode = "17"     ''�������ù�
        CASE "12" : TenDlvCode2cjMallDlvCode = "99"     ''�ѱ��ù� / ���ѱ��ù蹰��?
        CASE "13" : TenDlvCode2cjMallDlvCode = "69"     ''���ο�ĸ
        CASE "14" : TenDlvCode2cjMallDlvCode = "99"     ''���̽��ù�
        CASE "15" : TenDlvCode2cjMallDlvCode = "99"     ''�߾��ù�
        CASE "16" : TenDlvCode2cjMallDlvCode = "99"     ''�����ù�
        CASE "17" : TenDlvCode2cjMallDlvCode = "57"     ''Ʈ����ù�
        CASE "18" : TenDlvCode2cjMallDlvCode = "70"     ''�����ù�
        CASE "19" : TenDlvCode2cjMallDlvCode = "99"     ''KGBƯ���ù�
        CASE "20" : TenDlvCode2cjMallDlvCode = "68"     ''KT������
        CASE "21" : TenDlvCode2cjMallDlvCode = "78"     ''�浿�ù�
        CASE "22" : TenDlvCode2cjMallDlvCode = "99"     ''�����ù�
        CASE "23" : TenDlvCode2cjMallDlvCode = "99"     ''�굦���ù� �ż���
        CASE "24" : TenDlvCode2cjMallDlvCode = "62"     ''�簡���ͽ�������
        CASE "25" : TenDlvCode2cjMallDlvCode = "60"     ''�ϳ����ù�
        CASE "26" : TenDlvCode2cjMallDlvCode = "71"     ''�Ͼ��ù�
        CASE "27" : TenDlvCode2cjMallDlvCode = ""     ''LOEX�ù�
        CASE "28" : TenDlvCode2cjMallDlvCode = "87"     ''�����ͽ�������
        CASE "29" : TenDlvCode2cjMallDlvCode = "65"     ''�ǿ��ù�
        CASE "30" : TenDlvCode2cjMallDlvCode = "88"     ''�̳�����
        CASE "31" : TenDlvCode2cjMallDlvCode = "82"     ''õ���ù�
        CASE "33" : TenDlvCode2cjMallDlvCode = "58"     ''ȣ���ù�
        CASE "34" : TenDlvCode2cjMallDlvCode = "81"     ''���ȭ���ù�
        CASE "35" : TenDlvCode2cjMallDlvCode = "12"     ''CVSnet�ù�  - CJ����������� ������û
        CASE "39" : TenDlvCode2cjMallDlvCode = "87"     ''KG������
        CASE "98" : TenDlvCode2cjMallDlvCode = "32"     ''������->�����
        CASE "99" : TenDlvCode2cjMallDlvCode = "99"     ''��Ÿ
        CASE  Else
            TenDlvCode2cjMallDlvCode = "99"      ''��Ÿ�߼�
    end Select
end function

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
        CASE "20" : TenDlvCode2InterParkDlvCode = "169167"     ''KT������
        CASE "21" : TenDlvCode2InterParkDlvCode = "303978"     ''�浿�ù�
        CASE "22" : TenDlvCode2InterParkDlvCode = "169526"     ''�����ù�
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
        CASE "37" : TenDlvCode2InterParkDlvCode = "2641054"     ''�յ��ù�
        CASE "38" : TenDlvCode2InterParkDlvCode = "2272970"     ''GTX������   ''GTX(��ī�̷�����)::2586778  ''2015/06/29 �߰�
        CASE "39" : TenDlvCode2InterParkDlvCode = "2964976"     ''KG������
        CASE "41" : TenDlvCode2InterParkDlvCode = "2964976"     ''�帲�ù�(�����ù�,���ο�ĸ)  ''2018/02/13
		CASE "42" : TenDlvCode2InterParkDlvCode = "169177"     ''CU Post => �������
        CASE "50" : TenDlvCode2InterParkDlvCode = "4656462"     ''�������Ⱦ�
        CASE "54" : TenDlvCode2InterParkDlvCode = "169167"     ''��Ÿ
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
        CASE "3" : TenDlvCode2LotteDlvCode = "31"     ''�������
        CASE "4" : TenDlvCode2LotteDlvCode = "31"     ''CJ GLS
        CASE "5" : TenDlvCode2LotteDlvCode = "23"     ''��Ŭ����
        CASE "6" : TenDlvCode2LotteDlvCode = "32"     ''�Ｚ HTH
        CASE "7" : TenDlvCode2LotteDlvCode = "56"     ''����(���ѹ̸�) ''Ȯ
        CASE "8" : TenDlvCode2LotteDlvCode = "9339"     ''��ü���ù�
        CASE "9" : TenDlvCode2LotteDlvCode = "39"     ''KGB�ù�
        CASE "10" : TenDlvCode2LotteDlvCode = "34"     ''�����ù� / �ο���(�� ����)
        CASE "11" : TenDlvCode2LotteDlvCode = ""     ''�������ù�
        CASE "12" : TenDlvCode2LotteDlvCode = "29"     ''�ѱ��ù� / �ѱ�Ư��
        CASE "13" : TenDlvCode2LotteDlvCode = "70"     ''���ο�ĸ
        CASE "14" : TenDlvCode2LotteDlvCode = ""     ''���̽��ù�
        CASE "15" : TenDlvCode2LotteDlvCode = "43"     ''�߾��ù�
        CASE "16" : TenDlvCode2LotteDlvCode = ""     ''�����ù�
        CASE "17" : TenDlvCode2LotteDlvCode = "36"     ''Ʈ����ù�
        CASE "18" : TenDlvCode2LotteDlvCode = "41"     ''�����ù�
        CASE "19" : TenDlvCode2LotteDlvCode = "44"     ''KGBƯ���ù�
        CASE "20" : TenDlvCode2LotteDlvCode = "30"     ''KT������
        CASE "21" : TenDlvCode2LotteDlvCode = "52"     ''�浿�ù�
        CASE "22" : TenDlvCode2LotteDlvCode = ""     ''�����ù�
        CASE "23" : TenDlvCode2LotteDlvCode = "42"     ''�굦���ù� �ż���
        CASE "24" : TenDlvCode2LotteDlvCode = "51"     ''�簡���ͽ�������
        CASE "25" : TenDlvCode2LotteDlvCode = "3"     ''�ϳ����ù�v
        CASE "26" : TenDlvCode2LotteDlvCode = "47"     ''�Ͼ��ù�
        CASE "27" : TenDlvCode2LotteDlvCode = ""     ''LOEX�ù�
        CASE "28" : TenDlvCode2LotteDlvCode = "70"     ''�����ͽ�������
        CASE "29" : TenDlvCode2LotteDlvCode = "45"     ''�ǿ��ù�
        CASE "30" : TenDlvCode2LotteDlvCode = "57"     ''�̳�����
        CASE "31" : TenDlvCode2LotteDlvCode = "33"     ''õ���ù�
        CASE "33" : TenDlvCode2LotteDlvCode = "99"     ''ȣ���ù�
        CASE "34" : TenDlvCode2LotteDlvCode = "46"     ''���ȭ���ù�
        CASE "35" : TenDlvCode2LotteDlvCode = "99"     ''CVSnet�ù�
        CASE "39" : TenDlvCode2LotteDlvCode = "70"     ''KG������
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
        CASE "20" : TenDlvCode2LotteiMallDlvCode = "37"     ''KT������
        CASE "21" : TenDlvCode2LotteiMallDlvCode = "49"     ''�浿�ù�
        CASE "22" : TenDlvCode2LotteiMallDlvCode = ""     ''�����ù�
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

'''�Ե�iMall New ���庯ȯ(2015-09-01 ��������Ѵ��� by.������)
function TenDlvCode2LotteiMallNewDlvCode(itenCode)
    if IsNULL(itenCode) then Exit function
    itenCode = TRIM(CStr(itenCode))
''41	�����ù�
''99	��Ÿ
'rw itenCode
'response.end
    select Case itenCode
        CASE "1" : TenDlvCode2LotteiMallNewDlvCode = "15"     ''����
        CASE "2" : TenDlvCode2LotteiMallNewDlvCode = "11"     ''����v
        CASE "3" : TenDlvCode2LotteiMallNewDlvCode = "12"     ''�������
        CASE "4" : TenDlvCode2LotteiMallNewDlvCode = "12"     ''CJ GLS
        CASE "5" : TenDlvCode2LotteiMallNewDlvCode = ""     ''��Ŭ����
        CASE "6" : TenDlvCode2LotteiMallNewDlvCode = "99"     ''�Ｚ HTH
        CASE "7" : TenDlvCode2LotteiMallNewDlvCode = "21"     ''����(���ѹ̸�) ''Ȯ
        CASE "8" : TenDlvCode2LotteiMallNewDlvCode = "31"     ''��ü���ù�
        CASE "9" : TenDlvCode2LotteiMallNewDlvCode = "40"     ''KGB�ù�
        CASE "10" : TenDlvCode2LotteiMallNewDlvCode = "34"     ''�����ù� / �ο���(�� ����)
        CASE "11" : TenDlvCode2LotteiMallNewDlvCode = ""     ''�������ù�
        CASE "12" : TenDlvCode2LotteiMallNewDlvCode = "37"     ''�ѱ��ù� / �ѱ�Ư��
        CASE "13" : TenDlvCode2LotteiMallNewDlvCode = "99"     ''���ο�ĸ
        CASE "14" : TenDlvCode2LotteiMallNewDlvCode = ""     ''���̽��ù�
        CASE "15" : TenDlvCode2LotteiMallNewDlvCode = ""     ''�߾��ù�
        CASE "16" : TenDlvCode2LotteiMallNewDlvCode = ""     ''�����ù�
        CASE "17" : TenDlvCode2LotteiMallNewDlvCode = "36"     ''Ʈ����ù�
        CASE "18" : TenDlvCode2LotteiMallNewDlvCode = "24"     ''�����ù�
        CASE "19" : TenDlvCode2LotteiMallNewDlvCode = "40"     ''KGBƯ���ù�
        CASE "20" : TenDlvCode2LotteiMallNewDlvCode = "37"     ''KT������
        CASE "21" : TenDlvCode2LotteiMallNewDlvCode = "49"     ''�浿�ù�
        CASE "22" : TenDlvCode2LotteiMallNewDlvCode = ""     ''�����ù�
        CASE "23" : TenDlvCode2LotteiMallNewDlvCode = "47"     ''�굦���ù� �ż���
        CASE "24" : TenDlvCode2LotteiMallNewDlvCode = "43"     ''�簡���ͽ�������
        CASE "25" : TenDlvCode2LotteiMallNewDlvCode = "46"     ''�ϳ����ù�v
        CASE "26" : TenDlvCode2LotteiMallNewDlvCode = "18"     ''�Ͼ��ù�
        CASE "27" : TenDlvCode2LotteiMallNewDlvCode = "48"     ''LOEX�ù�
        CASE "28" : TenDlvCode2LotteiMallNewDlvCode = "21"     ''�����ͽ�������
        CASE "29" : TenDlvCode2LotteiMallNewDlvCode = "99"     ''�ǿ��ù�
        CASE "30" : TenDlvCode2LotteiMallNewDlvCode = "23"     ''�̳�����
        CASE "31" : TenDlvCode2LotteiMallNewDlvCode = "17"     ''õ���ù�
        CASE "33" : TenDlvCode2LotteiMallNewDlvCode = ""     ''ȣ���ù�
        CASE "34" : TenDlvCode2LotteiMallNewDlvCode = "38"     ''���ȭ���ù�
        CASE "35" : TenDlvCode2LotteiMallNewDlvCode = "62"     ''CVSnet�ù�
        CASE "39" : TenDlvCode2LotteiMallNewDlvCode = "21"     ''KG������
        CASE "98" : TenDlvCode2LotteiMallNewDlvCode = "99"     ''������
        CASE "99" : TenDlvCode2LotteiMallNewDlvCode = "99"     ''��ü����
        CASE  Else
            TenDlvCode2LotteiMallNewDlvCode = "99"
    end Select
end function

function TenDlvCode2GSShopDlvCode(itenCode)
    ''if IsNULL(itenCode) then Exit function
    if IsNULL(itenCode) then itenCode="ZY"

    itenCode = TRIM(CStr(itenCode))
    select Case itenCode
        CASE "1" : TenDlvCode2GSShopDlvCode = "HJ"     ''����
        CASE "2" : TenDlvCode2GSShopDlvCode = "HD"     ''����v
        CASE "3" : TenDlvCode2GSShopDlvCode = "DH"     ''�������
        CASE "4" : TenDlvCode2GSShopDlvCode = "DH"      ''"CJ"     ''CJ GLS  2017/07/27 CJ=>DH
        CASE "5" : TenDlvCode2GSShopDlvCode = ""     ''��Ŭ����
        CASE "6" : TenDlvCode2GSShopDlvCode = ""     ''�Ｚ HTH
        CASE "7" : TenDlvCode2GSShopDlvCode = "FA"     ''����(���ѹ̸�) ''Ȯ
        CASE "8" : TenDlvCode2GSShopDlvCode = "EP"     ''��ü���ù�
        CASE "9" : TenDlvCode2GSShopDlvCode = "KL"     ''KGB�ù�
        CASE "10" : TenDlvCode2GSShopDlvCode = ""     ''�����ù� / �ο���(�� ����)
        CASE "11" : TenDlvCode2GSShopDlvCode = ""     ''�������ù�
        CASE "12" : TenDlvCode2GSShopDlvCode = ""     ''�ѱ��ù� / �ѱ�Ư��
        CASE "13" : TenDlvCode2GSShopDlvCode = "YC"     ''���ο�ĸ
        CASE "14" : TenDlvCode2GSShopDlvCode = ""     ''���̽��ù�
        CASE "15" : TenDlvCode2GSShopDlvCode = ""     ''�߾��ù�
        CASE "16" : TenDlvCode2GSShopDlvCode = ""     ''�����ù�
        CASE "17" : TenDlvCode2GSShopDlvCode = ""     ''Ʈ����ù�
        CASE "18" : TenDlvCode2GSShopDlvCode = "KG"     ''�����ù�
        CASE "19" : TenDlvCode2GSShopDlvCode = "KL"     ''KGBƯ���ù�
        CASE "20" : TenDlvCode2GSShopDlvCode = ""     ''KT������
        CASE "21" : TenDlvCode2GSShopDlvCode = "KD"     ''�浿�ù�
        CASE "22" : TenDlvCode2GSShopDlvCode = ""     ''�����ù�
        CASE "23" : TenDlvCode2GSShopDlvCode = ""     ''�굦���ù� �ż���
        CASE "24" : TenDlvCode2GSShopDlvCode = ""     ''�簡���ͽ�������
        CASE "25" : TenDlvCode2GSShopDlvCode = ""     ''�ϳ����ù�v
        CASE "26" : TenDlvCode2GSShopDlvCode = "IY"     ''�Ͼ��ù�
        CASE "27" : TenDlvCode2GSShopDlvCode = ""     ''LOEX�ù�
        CASE "28" : TenDlvCode2GSShopDlvCode = "FA"     ''�����ͽ�������
        CASE "29" : TenDlvCode2GSShopDlvCode = "KY"     ''�ǿ��ù�
        CASE "30" : TenDlvCode2GSShopDlvCode = ""     ''�̳�����
        CASE "31" : TenDlvCode2GSShopDlvCode = "CI"     ''õ���ù�
        CASE "33" : TenDlvCode2GSShopDlvCode = "ZY"     ''ȣ���ù�
        CASE "34" : TenDlvCode2GSShopDlvCode = "DS"     ''���ȭ���ù�
        CASE "35" : TenDlvCode2GSShopDlvCode = "CV"     ''CVSnet�ù�
        CASE "37" : TenDlvCode2GSShopDlvCode = "H1"     ''�յ��ù�
        CASE "38" : TenDlvCode2GSShopDlvCode = "IN"     ''GTX������
        CASE "39" : TenDlvCode2GSShopDlvCode = "FA"     ''KG������
        CASE "98" : TenDlvCode2GSShopDlvCode = "ZY"     ''������
        CASE "99" : TenDlvCode2GSShopDlvCode = "ZY"     ''��ü����
        CASE  Else
            TenDlvCode2GSShopDlvCode = "ZY"
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

function Fn_ActOutMall_CateSummary(iMallID)
    dim sqlStr
    sqlStr = "exec db_item.dbo.sp_Ten_OutMall_CateSummary '"&iMallID&"'"

    dbget.Execute sqlStr
end function

Function Fn_AcctFailTouch(iMallID,iitemid,iLastErrStr)
    Dim strSql
    iLastErrStr = html2db(iLastErrStr)

    IF (iMallID="lotteCom") THEN
        strSql = "Update R"&VbCRLF
        strSql = strSql &" SET accFailCnt=accFailCnt+1"&VbCRLF
        strSql = strSql &" ,lastErrStr=convert(varchar(100),'"&iLastErrStr&"')"&VbCRLF
        strSql = strSql &" From db_item.dbo.tbl_lotte_regItem R"&VbCRLF
        strSql = strSql &" where itemid="&iitemid&VbCRLF
        dbget.Execute(strSql)

    ELSEIF (iMallID="lotteimall") THEN
        strSql = "Update R"&VbCRLF
        strSql = strSql &" SET accFailCnt=accFailCnt+1"&VbCRLF
        strSql = strSql &" ,lastErrStr=convert(varchar(100),'"&iLastErrStr&"')"&VbCRLF
        strSql = strSql &" From db_item.dbo.tbl_LTiMall_regItem R"&VbCRLF
        strSql = strSql &" where itemid="&iitemid&VbCRLF
        dbget.Execute(strSql)

    ELSEIF (iMallID="interpark") THEN
        strSql = "Update R"&VbCRLF
        strSql = strSql &" SET accFailCnt=accFailCnt+1"&VbCRLF
        strSql = strSql &" ,lastErrStr=convert(varchar(100),'"&iLastErrStr&"')"&VbCRLF
        strSql = strSql &" From db_item.dbo.tbl_interpark_reg_item R"&VbCRLF
        strSql = strSql &" where itemid="&iitemid&VbCRLF
        dbget.Execute(strSql)
    ELSEIF (iMallID="gsshop") THEN
        strSql = "Update R"&VbCRLF
        strSql = strSql &" SET accFailCnt=accFailCnt+1"&VbCRLF
        strSql = strSql &" ,lastErrStr=convert(varchar(100),'"&iLastErrStr&"')"&VbCRLF
        strSql = strSql &" From db_item.dbo.tbl_gsshop_regitem R"&VbCRLF
        strSql = strSql &" where itemid="&iitemid&VbCRLF
        dbget.Execute(strSql)
  	ElseIf (iMallID = "coupang") Then
		strSql = ""
		strSql = strSql & "UPDATE R "&VbCRLF
		strSql = strSql &" SET accFailCnt = accFailCnt + 1" & VBCRLF
		strSql = strSql &" ,lastErrStr = convert(varchar(100),'"&iLastErrStr&"')" & VBCRLF
		strSql = strSql &" FROM db_etcmall.[dbo].[tbl_coupang_regitem] as R" & VBCRLF
		strSql = strSql &" WHERE itemid = "&iitemid & VBCRLF
		dbget.Execute(strSql)
	End If
end function


function Fn_AcctFailLog(iMallID,iitemid,ErrMsg,ErrCode)
    Dim sqlStr
    ''db_log.dbo.tbl_interparkEdit_log
    IF (iMallID="lotteCom") THEN
        sqlStr = " insert into db_log.dbo.tbl_interparkEdit_log" & VbCrlf
        sqlStr = sqlStr & " (itemid, interparkprdno,sellcash,buycash,sellyn, ErrMsg, errCode, mallid)" & VbCrlf
        sqlStr = sqlStr & " select R.itemid, isNULL(lotteGoodNo,lotteTmpGoodNo), i.sellcash,i.buycash,i.sellyn,convert(varchar(100),'"&html2db(ErrMsg)&"'),'"&ErrCode&"' " & VbCrlf
        sqlStr = sqlStr & " ,'"&iMallID&"'" & VbCrlf
        sqlStr = sqlStr & "  from db_item.dbo.tbl_lotte_regItem R" & VbCrlf
        sqlStr = sqlStr & " 	Join db_item.dbo.tbl_item i" & VbCrlf
        sqlStr = sqlStr & " 	on R.itemid=i.itemid" & VbCrlf
        sqlStr = sqlStr & " where R.itemid=" & iitemid & VbCrlf
        'rw sqlStr
        dbget.execute sqlStr
    ELSEIF (iMallID="lotteimall") THEN
        sqlStr = " insert into db_log.dbo.tbl_interparkEdit_log" & VbCrlf
        sqlStr = sqlStr & " (itemid, interparkprdno,sellcash,buycash,sellyn, ErrMsg, errCode, mallid)" & VbCrlf
        sqlStr = sqlStr & " select R.itemid, isNULL(R.LTimallGoodno,R.LtiMallTmpGoodNo), i.sellcash,i.buycash,i.sellyn,convert(varchar(100),'"&html2db(ErrMsg)&"'),'"&ErrCode&"' " & VbCrlf
        sqlStr = sqlStr & " ,'"&iMallID&"'" & VbCrlf
        sqlStr = sqlStr & "  from db_item.dbo.tbl_ltimall_regItem R" & VbCrlf
        sqlStr = sqlStr & " 	Join db_item.dbo.tbl_item i" & VbCrlf
        sqlStr = sqlStr & " 	on R.itemid=i.itemid" & VbCrlf
        sqlStr = sqlStr & " where R.itemid=" & iitemid & VbCrlf
        'rw sqlStr
        dbget.execute sqlStr

    ELSEIF (iMallID="interpark") THEN
        sqlStr = " insert into db_log.dbo.tbl_interparkEdit_log" & VbCrlf
        sqlStr = sqlStr & " (itemid, interparkprdno,sellcash,buycash,sellyn, ErrMsg, errCode, mallid)" & VbCrlf
        sqlStr = sqlStr & " select R.itemid, isNULL(R.interparkprdno,''), i.sellcash,i.buycash,i.sellyn,convert(varchar(100),'"&html2db(ErrMsg)&"'),'"&ErrCode&"' " & VbCrlf
        sqlStr = sqlStr & " ,'"&iMallID&"'" & VbCrlf
        sqlStr = sqlStr & "  from db_item.dbo.tbl_interpark_reg_item R" & VbCrlf
        sqlStr = sqlStr & " 	Join db_item.dbo.tbl_item i" & VbCrlf
        sqlStr = sqlStr & " 	on R.itemid=i.itemid" & VbCrlf
        sqlStr = sqlStr & " where R.itemid=" & iitemid & VbCrlf
        'rw sqlStr
        dbget.execute sqlStr
    ENd IF
end function

function Fn_AcctFailLogNone(iMallID,iitemid,ioutmallPrdno,ioutmallsellyn,ioutmallsellcash,ioutmallbuycash,ErrMsg,ErrCode)
    Dim sqlStr
    sqlStr = " insert into db_log.dbo.tbl_interparkEdit_log" & VbCrlf
    sqlStr = sqlStr & " (itemid, interparkprdno,sellcash,buycash,sellyn, ErrMsg, errCode, mallid)" & VbCrlf
    sqlStr = sqlStr & " values("&iitemid& VbCrlf
    sqlStr = sqlStr & " ,'"&ioutmallPrdno&"'"& VbCrlf
    sqlStr = sqlStr & " ,"&ioutmallsellcash& VbCrlf
    sqlStr = sqlStr & " ,'"&ioutmallbuycash&"'"& VbCrlf
    sqlStr = sqlStr & " ,'"&ioutmallsellyn&"'"& VbCrlf
    sqlStr = sqlStr & " ,convert(varchar(100),'"&html2db(ErrMsg)&"')"& VbCrlf
    sqlStr = sqlStr & " ,'"&ErrCode&"'"& VbCrlf
    sqlStr = sqlStr & " ,'"&iMallID&"')"& VbCrlf
    dbget.execute sqlStr
end function

Function SugiQueLogInsert(imallid, iapiaction, iitemid, iresultcode, ilastErrMsg, ilastUpdateid)
	Dim strSQL

	ilastErrMsg = replace(ilastErrMsg, "'", "��")
	strSQL = ""
	strSQL = strSQL & " INSERT INTO [db_etcmall].[dbo].[tbl_outmall_API_Que] (mallid, apiAction, itemid, priority, regdate, readdate, findate, resultCode, lastErrMsg, lastUserid) VAlUES " & VBCRLF
	strSQL = strSQL & " ('"& imallid &"', '"& iapiaction &"', '"& iitemid &"', '999999', getdate(), getdate(), getdate(), '"& iresultcode &"', '"& LEFT(ilastErrMsg, 100) &"', '"& ilastUpdateid &"') " & VBCRLF
	dbget.Execute strSQL
	If iresultcode = "OK" Then
		If imallid = "interpark" OR imallid = "ezwel" OR imallid = "halfclub" Then
			If ilastUpdateid = "kjy8517" OR ilastUpdateid = "icommang" Then
				If (iapiaction = "EDIT") OR (iapiaction = "EditSellYn") OR (iapiaction = "CHKSTAT") Then
					strSQL = ""
					strSQL = strSQL & " UPDATE db_etcmall.dbo.tbl_outmall_API_Que " & VBCRLF
					strSQL = strSQL & " SET readdate=getdate() " & VBCRLF
					strSQL = strSQL & " ,findate=getdate() " & VBCRLF
					strSQL = strSQL & " ,resultCode='DUPP' " & VBCRLF
					strSQL = strSQL & " ,lastErrMsg='' " & VBCRLF
					strSQL = strSQL & " WHERE mallid = '"&imallid&"' " & VBCRLF
					strSQL = strSQL & " and itemid = '"&iitemid&"' " & VBCRLF
					If iapiaction = "EDIT" Then
						strSQL = strSQL & " and apiAction in ('EDIT', 'PRICE', 'SOLDOUT') " & VBCRLF
                    ElseIf iapiaction = "CHKSTAT" Then
                        strSQL = strSQL & " and apiAction in ('CHKSTAT') " & VBCRLF
					Else
						strSQL = strSQL & " and apiAction in ('SOLDOUT') " & VBCRLF
					End If
					strSQL = strSQL & " and readdate is null " & VBCRLF
					strSQL = strSQL & " and lastUserid = 'system' "
					dbget.Execute strSQL
				End If
			End If
		ElseIf imallid = "auction1010" Then
			If ilastUpdateid = "kjy8517" OR ilastUpdateid = "icommang" Then
				If (iapiaction = "EditInfo") OR (iapiaction = "EditSellYn") OR (iapiaction = "EDIT") Then
					strSQL = ""
					strSQL = strSQL & " UPDATE db_etcmall.dbo.tbl_outmall_API_Que " & VBCRLF
					strSQL = strSQL & " SET readdate=getdate() " & VBCRLF
					strSQL = strSQL & " ,findate=getdate() " & VBCRLF
					strSQL = strSQL & " ,resultCode='DUPP' " & VBCRLF
					strSQL = strSQL & " ,lastErrMsg='' " & VBCRLF
					strSQL = strSQL & " WHERE mallid = '"&imallid&"' " & VBCRLF
					strSQL = strSQL & " and itemid = '"&iitemid&"' " & VBCRLF
					If iapiaction = "EditInfo" Then
						strSQL = strSQL & " and apiAction in ('PRICE') " & VBCRLF
                    ElseIf iapiaction = "EDIT" Then
                        strSQL = strSQL & " and apiAction in ('EDIT') " & VBCRLF
					Else
						strSQL = strSQL & " and apiAction in ('SOLDOUT') " & VBCRLF
					End If
					strSQL = strSQL & " and readdate is null " & VBCRLF
					strSQL = strSQL & " and lastUserid = 'system' "
					dbget.Execute strSQL
				End If
			End If
		ElseIf (imallid = "lotteCom") OR (imallid = "lotteimall") OR (imallid = "cjmall") OR (imallid = "gsshop") OR (imallid = "nvstorefarm") OR (imallid = "nvstoremoonbangu") OR (imallid = "Mylittlewhoopee") OR (imallid = "nvstorefarmclass") or imallid = ("gmarket1010") OR imallid = ("11st1010") OR imallid = ("ssg") OR imallid = ("coupang") OR imallid = ("hmall1010") OR imallid = ("WMP") OR imallid = ("wmpfashion") OR imallid = ("lfmall") Then
			If ilastUpdateid = "kjy8517" OR ilastUpdateid = "icommang" OR ilastUpdateid = "yhj0613" Then
				If (iapiaction = "EDIT") OR (iapiaction = "EditSellYn") OR (iapiaction = "PRICE") OR (iapiaction = "CHKSTAT") OR (iapiaction = "EDITPOLICY") OR (iapiaction = "EDITINFO") Then
					strSQL = ""
					strSQL = strSQL & " UPDATE db_etcmall.dbo.tbl_outmall_API_Que " & VBCRLF
					strSQL = strSQL & " SET readdate=getdate() " & VBCRLF
					strSQL = strSQL & " ,findate=getdate() " & VBCRLF
					strSQL = strSQL & " ,resultCode='DUPP' " & VBCRLF
					strSQL = strSQL & " ,lastErrMsg='' " & VBCRLF
					strSQL = strSQL & " WHERE mallid = '"&imallid&"' " & VBCRLF
					strSQL = strSQL & " and itemid = '"&iitemid&"' " & VBCRLF
					If iapiaction = "PRICE" Then
						strSQL = strSQL & " and apiAction in ('PRICE') " & VBCRLF
                    ElseIf (imallid = "gsshop") AND (iapiaction = "EDITINFO") Then
						strSQL = strSQL & " and apiAction in ('EDITINFO') " & VBCRLF
                    ElseIf (imallid = "gmarket1010") AND (iapiaction = "EDITPOLICY") Then
						strSQL = strSQL & " and apiAction in ('EDITPOLICY', 'EDITBATCH') " & VBCRLF
					ElseIf iapiaction = "EditSellYn" Then
						strSQL = strSQL & " and apiAction in ('SOLDOUT') " & VBCRLF
                    ElseIf iapiaction = "EDIT" Then
                        If imallid = "coupang" OR imallid = "ssg" OR imallid = "WMP" OR imallid = "wmpfashion" Then
                            strSQL = strSQL & " and apiAction in ('EDIT', 'EDITBATCH') " & VBCRLF
                        ElseIf imallid = "lfmall" Then
                            strSQL = strSQL & " and apiAction in ('PRICE') " & VBCRLF
                        Else
    						strSQL = strSQL & " and apiAction in ('EDIT') " & VBCRLF
                        End If
					ElseIf iapiaction = "CHKSTAT" Then
                        If imallid = "cjmall" Then
                            strSQL = strSQL & " and apiAction in ('CHKSTAT', 'CONFIRM') " & VBCRLF
                        Else
    						strSQL = strSQL & " and apiAction in ('CHKSTAT') " & VBCRLF
                        End If
					Else
						strSQL = strSQL & " and apiAction in ('EDIT', 'PRICE', 'SOLDOUT') " & VBCRLF
					End If
					strSQL = strSQL & " and readdate is null " & VBCRLF
					strSQL = strSQL & " and lastUserid = 'system' "
					dbget.Execute strSQL
				End If
			End If
		End If
	End If
End Function
%>