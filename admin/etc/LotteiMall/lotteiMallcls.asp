<%
'' �����å  3���� ���� 2500
CONST CMAXMARGIN = 14.9			'' MaxMagin��.. '(�Ե�iMall 10%)
CONST CMAXLIMITSELL = 5        '' �� ���� �̻��̾�� �Ǹ���. // �ɼ������� ��������.
CONST CMALLNAME = "lotteimall"
CONST CLTIMALLMARGIN = 11       ''���� 10%	''2013-06-27���� 11%�� ����
CONST CHEADCOPY = "Design Your Life! ���ο� �ϻ��� ����� ������Ȱ�귣�� �ٹ�����" ''��Ȱ ����ä�� �ٹ�����
CONST CPREFIXITEMNAME ="[�ٹ�����]"
CONST CitemGbnKey ="K1099999" ''��ǰ����Ű ''�ϳ��� ����
CONST CUPJODLVVALID = FALSE   ''��ü ���ǹ�� ��� ���ɿ���

class CLotteiMallItem
	public FLastUpdate
	public FisUsing

	'���MD
	public FMDCode
	public FMDName
	public FSellFeeType
	public FNormalSellFee
	public FEventSellFee

	'MD��ǰ��
	public FgroupCode               ''' �Ե�iMall =>LCode. 50000000 : ������
	public FSuperGroupName
	public FGroupName

	'�Ե����� ī�װ�
	public FitemGbnKey
    public FitemGbnNm

	public FDispNo
	public FDispNm

	public FDispLrgNm
	public FDispMidNm
	public FDispSmlNm
	public FDispThnNm

	public FGbnLrgNm
    public FGbnMidNm
    public FGbnSmlNm
    public FGbnThnNm
    public FCateIsUsing

	public FtenCateLarge
	public FtenCateMid
	public FtenCateSmall
	public FtenCDLName
	public FtenCDMName
	public FtenCDSName
	public FtenCateName
    public Fdisptpcd

	'�Ե����� �귣��
	public FlotteBrandCd
	public FlotteBrandName
	public FTenMakerid
	public FTenBrandName

	'�Ե����� ��ǰ���
	public FLTiMallRegdate
	public FLTiMallLastUpdate
	public FLTiMallGoodNo				'�ǻ�ǰ��ȣ
	public FLTiMallTmpGoodNo			'�ӽû�ǰ��ȣ
	public FLTiMallPrice
	public FLTiMallSellYn
	public FregUserid
	public FLotteDispCnt
	public FCateMapCnt
	public FLTiMallStatCd				'��ǰ��ϻ���
    public FregedOptCnt
    public FrctSellCNT
    public FaccFailCNT              '��ϼ��� ���� Ƚ��
    public FlastErrStr              '��������

	'�ٹ����� ��ǰ���
	public Fitemid
	public Fitemname
	public FitemDiv
	public FsmallImage
	public FbasicImage
	public FmainImage
	public FmainImage2
	public Fmakerid
	public Fregdate
	public ForgPrice
	public ForgSuplyCash
	public FSellCash
	public FBuyCash
	public FsellYn
	public FsaleYn
	public FLimitYn
	public FLimitNo
	public FLimitSold
	public Fkeywords
	public ForderComment
	public FoptionCnt
	public Fsourcearea
	public Fmakername
	public Fitemcontent
	public FUsingHTML
    public Fdeliverytype
    public Fvatinclude
    public Fdefaultdeliverytype
    public FdefaultfreeBeasongLimit
    public FinfoDiv

    public FoptAddPrcCnt
    public FoptAddPrcRegType

    public FRectMode    ''??

    function getLimitEa()
        dim ret : ret = (FLimitno-FLimitSold)
        if (ret<1) then ret=0

        getLimitEa = ret
    end function

    function getLimitHtmlStr()
        if IsNULL(FLimityn) then Exit function

        if (FLimityn="Y") then
            getLimitHtmlStr = "<font color=blue>����:"&getLimitEa&"</font>"
        end if
    end function

    function getNOREST_ALLOW_MONTH()
        '1~29���� : �Ͻú�
        '30~59���� : 5����
        '60~99���� ���� : 7����
        '100���� �̻� : 10����
        dim retVal : retVal = ""
        if (FSellCash<300000) then
            exit function
        elseif (FSellCash<600000) then
            getNOREST_ALLOW_MONTH = "5"
        elseif (FSellCash<1000000) then
            getNOREST_ALLOW_MONTH = "7"
        elseif (FSellCash>=1000000) then
            getNOREST_ALLOW_MONTH = "10"
        end if

    end function

    function getItemNameFormat()
        dim buf
        buf = replace(FItemName,"'","")
        buf = replace(buf,"~","-")
        buf = replace(buf,"<","[")
        buf = replace(buf,">","]")
        buf = replace(buf,"%","����")
        buf = replace(buf,"[������]","")
        buf = replace(buf,"[���� ���]","")
        getItemNameFormat = buf
    end function

    ''�ɼǱ��и� - :�ȵ� max20Byte
    function getGOODSDT_NmFormat(idtname)
        dim buf
        buf = Replace(db2Html(idtname),":","")
        buf = Replace(buf,"�������� �������ּ���","������ ����")
        buf = Replace(buf,"�������� ���� �ϼ���","������ ����")
        buf = Replace(buf,"�������� ������ �ּ���","������ ����")
        buf = Replace(buf,"�������� ����ּ���","������ ����")
        buf = Replace(buf,"���̾ �����ϱ�!","���̾ ����")
        getGOODSDT_NmFormat = Trim(buf)
    end function

    function getLTiMallSuplyPrice()
        getLTiMallSuplyPrice = CLNG(FSellCash*(100-CLTIMALLMARGIN)/100)
    end function

    public function getLtiMallStatName
        if IsNULL(FLTiMallStatCd) then FLTiMallStatCd=-1

        Select Case FLTiMallStatCd
            CASE -9 : getLtiMallStatName = "�̵��"
            CASE -2 : getLtiMallStatName = "<font color=red>�ݷ�</font>"
            CASE -1 : getLtiMallStatName = "��Ͻ���"
            CASE 0 : getLtiMallStatName = "<font color=blue>��Ͽ���</font>"
            CASE 1 : getLtiMallStatName = "���۽õ�"
            CASE 3 : getLtiMallStatName = "���δ��"
            CASE 7 : getLtiMallStatName = getLimitHtmlStr ''"" ''��ϿϷ�
            CASE ELSE : getLtiMallStatName = FLTiMallStatCd
        end Select
    end function

    function getDispGubunNm()
        getDispGubunNm = getDisptpcdName
    end function

    public function getDisptpcdName
        if (Fdisptpcd="B") then
            getDisptpcdName = "<font color='blue'>����</font>"
        elseif (Fdisptpcd="D") then
            getDisptpcdName = "�Ϲ�"
        else
            getDisptpcdName = Fdisptpcd
        end if
    end function

    public function getDeliverytypeName
        if (Fdeliverytype="9") then
            getDeliverytypeName = "<font color='blue'>[���� "&FormatNumber(FdefaultfreeBeasongLimit,0)&"]</font>"
        elseif (Fdeliverytype="7") then
            getDeliverytypeName = "<font color='red'>[��ü����]</font>"
        elseif (Fdeliverytype="2") then
            getDeliverytypeName = "<font color='blue'>[��ü]</font>"
        else
            getDeliverytypeName = ""
        end if
    end function

	'// ǰ������
	public function IsSoldOut()
		ISsoldOut = (FSellyn<>"Y") or ((FLimitYn="Y") and (FLimitNo-FLimitSold<1))
	end function

	'// �˻���迭
	public function getItemKeywordArray(sno)
		dim arrRst, arrRst2
		if trim(Fkeywords)="" then exit Function

		arrRst = split(Fkeywords,",")
		if ubound(arrRst)=0 then
			'������ ������ ���
			arrRst2 = split(arrRst(0)," ")
			if ubound(arrRst2)>0 then
				arrRst = split(Fkeywords," ")
			end if
		end if

		if ubound(arrRst)>=sno then
			getItemKeywordArray=trim(arrRst(sno))
		else
			getItemKeywordArray=""
		end if
	end function

	'// ��ǰ��� �Ķ���� ����
	public Function getLTiMallItemRegXML()
		dim strRst
		dim ioriginCode,ioriginname
		ioriginCode = getOriginName2Code(Fsourcearea, ioriginname) '''�������ڵ�

		strRst = "<?xml version=""1.0"" encoding=""utf-8"" ?>" '''
		strRst = strRst & "<GoodsEntry_V01>"
        strRst = strRst & "<MessageHeader>"
    	strRst = strRst & "<SENDER>TENBYTEN</SENDER>"
    	strRst = strRst & "<RECEIVER>LotteH</RECEIVER>"
    	strRst = strRst & "<DATETIME>"&replace(Left(FormatDateTime(now,0),10),"-","")&" "&Left(FormatDateTime(now,4),5)&Right(FormatDateTime(now,0),3)&"</DATETIME>"
    	strRst = strRst & "<DOCUMENTID>GOODSENTRY</DOCUMENTID>"
    	strRst = strRst & "<ERROROCCUR>N</ERROROCCUR>"
    	strRst = strRst & "<ERRORMESSAGE>N</ERRORMESSAGE>"
        strRst = strRst & "</MessageHeader>"
		strRst = strRst & "<MessageBody>"
    	strRst = strRst & "<GoodsEntry>"
    	strRst = strRst & "<ENTP_CODE>"&ENTP_CODE&"</ENTP_CODE>"   ' in incLotteiMallFunction
    	IF (FLTiMallStatCd<1) then
    	    strRst = strRst & "<CUDTYPE>C</CUDTYPE>"                   ' �űԻ���(C)/��������ǰ����(U)
    	ELSE
    	    strRst = strRst & "<CUDTYPE>U</CUDTYPE>"
        end if
    	strRst = strRst & "<GoodsEntryLineItem>"
    	if (FItemID=210499) or (FItemID=724724) or (FItemID=692489) then
    	    strRst = strRst & "<ENTP_GOODS_CODE>"&999&FItemID&"</ENTP_GOODS_CODE>"
        else
    	    strRst = strRst & "<ENTP_GOODS_CODE>"&FItemID&"</ENTP_GOODS_CODE>"
        end if
    	strRst = strRst & "<GOODS_NAME><![CDATA["&CPREFIXITEMNAME&getItemNameFormat()&"]]></GOODS_NAME>"
    	strRst = strRst & "<LGROUP>"&Left(FitemGbnKey,2)&"</LGROUP>"                    '��ǰ ��з�
    	strRst = strRst & "<MGROUP>"&Mid(FitemGbnKey,3,2)&"</MGROUP>"                    '��ǰ �ߺз�
    	strRst = strRst & "<SGROUP>"&Mid(FitemGbnKey,5,2)&"</SGROUP>"                    '��ǰ �Һз�
    	strRst = strRst & "<DGROUP>"&Mid(FitemGbnKey,7,2)&"</DGROUP>"                    '��ǰ ���з�

    	strRst = strRst & "<COMMENT_GB></COMMENT_GB>"                                   ''NULL:�Ϲݻ�ǰ�� 10:�Ƿ� 20:��ǰ
    	strRst = strRst & "<MD_CODE>"&MD_CODE&"</MD_CODE>"                              ''MD_CODE
    	strRst = strRst & "<KEYWORD><![CDATA["&Fkeywords&"]]></KEYWORD>"                            ''keyword
    	strRst = strRst & "<BRAND_CODE>"&BRAND_CODE&"</BRAND_CODE>"                     ''�귣���ڵ�
    	strRst = strRst & "<BRAND_NAME>"&BRAND_NAME&"</BRAND_NAME>"                     ''�귣���ڵ��
    	strRst = strRst & "<MAKECO_CODE>"&MAKECO_CODE&"</MAKECO_CODE>"                  ''������ü�ڵ�
    	strRst = strRst & "<MAKECO_NAME><![CDATA["&Fmakername&"]]></MAKECO_NAME>"                       ''������ü��
    	strRst = strRst & "<ORIGIN_CODE>"&ioriginCode&"</ORIGIN_CODE>"   ''�������ڵ�
    	strRst = strRst & "<ORIGIN_NAME>"&ioriginname&"</ORIGIN_NAME>"                       ''��������
    	strRst = strRst & "<DELY_TYPE>0</DELY_TYPE>"                                    ''��۹�� 0:���¾�ü��� 1:�Ե�Ȩ�������ù�
    	strRst = strRst & "<EXCH_YN>Y</EXCH_YN>"                                        ''��ȯ���� N:�Ұ� Y:����
    	strRst = strRst & "<RETURN_YN>Y</RETURN_YN>"                                    ''��ǰ���� N:�Ұ� Y:����
    	strRst = strRst & "<GIFT_RETURN_YN>N</GIFT_RETURN_YN>"                          ''����ǰȸ���ʼ� N:�Ұ� Y:����
    	if (FALSE) AND ((FitemDiv="06") or (FitemDiv="16")) then
    	    strRst = strRst & "<DELY_SHAP_CODE>301</DELY_SHAP_CODE>"                        ''Ȯ�ο��.. �ֹ�����. 288295 TEST ���(301 �ƴ�)
    	else
    	    strRst = strRst & "<DELY_SHAP_CODE>000</DELY_SHAP_CODE>"                        ''������� 000:�Ϲݹ�� 201:��ġ���
        end if
    	strRst = strRst & "<MAX_ORD_PSBT_CQTY></MAX_ORD_PSBT_CQTY>"                     ''�� ���� �� ���� �ִ� ������ �� �ִ� ���� (�Ϻ�), ���� ������ Default: 20
    	strRst = strRst & "<MIN_ORD_PSBT_CQTY></MIN_ORD_PSBT_CQTY>"                     ''�� ���� �� ���� �ּ� ~�� �̻� �����ؾ� �ϴ� ���� (�ֹ���), ���� ������ Default: 1
    	strRst = strRst & "<MIXPACK_YN>Y</MIXPACK_YN>"                                  ''�����忩�� N:�Ұ� Y:����
    	strRst = strRst & "<DAMT_APLC_ORD></DAMT_APLC_ORD>"                             ''�ֹ���ۺ񱸺� NULL:���»纰�� ���õ� �⺻ ��ۺ� ����(MD�� Ȯ��)  1: ������ 2:���� 3:������
    	strRst = strRst & "<DAMT_APLC_ORD_AMT></DAMT_APLC_ORD_AMT>"                     ''�ֹ���ۺ� ��ۺ񱸺��� 2,3�� ��츸 ��ȿ��. �ݾ� �Է� ��, �� ��ǰ�� �⺻ ��ۺ� �����ϰ� �Էµ� �ݾ����� ��ۺ� ������
    	strRst = strRst & "<DAMT_APLC_REGD></DAMT_APLC_REGD>"                           ''��ǰȸ���񱸺� NULL:���»纰�� ���õ� �⺻ ��ۺ� ����(MD�� Ȯ��)  1: ������ 2:���� 3:������
    	strRst = strRst & "<DAMT_APLC_REGD_AMT></DAMT_APLC_REGD_AMT>"                   ''��ǰȸ����
    	strRst = strRst & "<DAMT_APLC_EXCH></DAMT_APLC_EXCH>"                           ''��ȯ��ۺ񱸺�
    	strRst = strRst & "<DAMT_APLC_EXCH_AMT></DAMT_APLC_EXCH_AMT>"                   ''��ȯ��ۺ�
    	strRst = strRst & "<AS_TERM>1</AS_TERM>"                                        ''AS �����Ⱓ(������)
    	strRst = strRst & "<AS_REPAIR_TERM>1</AS_REPAIR_TERM>"                          ''AS �ҿ�Ⱓ(�ϼ�)
    	strRst = strRst & "<AS_RECEIVE_TYPE>10</AS_RECEIVE_TYPE>"                       ''AS ������ü 10:�Ե�Ȩ���� 20:��ü�ȳ�
    	strRst = strRst & "<AS_DELY_TYPE>20</AS_DELY_TYPE>"                             ''AS �����ü 10:�Ե�Ȩ���� 20:���¾�ü
    	strRst = strRst & "<AS_OUT_COM>1</AS_OUT_COM>"                                  ''AS ����ü �������SEQ //Ȯ��
    	strRst = strRst & "<AS_RETURN_COM>1</AS_RETURN_COM>"                            ''AS ȸ����ü ȸ�������SEQ //Ȯ��
    	strRst = strRst & "<AS_NOTE></AS_NOTE>"                                         ''AS ����
    	strRst = strRst & "<ENTP_MAN_SEQ>1</ENTP_MAN_SEQ>"                              ''�������  ���������SEQ  //Ȯ��
    	strRst = strRst & "<OUTPLACE_SEQ>1</OUTPLACE_SEQ>"                              ''�����  �������SEQ  //Ȯ��
    	strRst = strRst & "<RETURN_SEQ>1</RETURN_SEQ>"                                  ''ȸ�����  ȸ�������SEQ  //Ȯ��
    	strRst = strRst & "<BUY_PRICE>"&getLTiMallSuplyPrice()&"</BUY_PRICE>"           ''���ް�(���԰�)
    	strRst = strRst & "<SALE_PRICE>"&FSellCash&"</SALE_PRICE>"                      ''�ǸŰ�
    	strRst = strRst & "<TAX_YN>"&CHKIIF(FVatInclude="N","2","1")&"</TAX_YN>"        ''1:���� 2:�鼼
    	strRst = strRst & "<NOREST_ALLOW_MONTH>"&getNOREST_ALLOW_MONTH()&"</NOREST_ALLOW_MONTH>"                   ''�����ڰ�����
''��������
		strRst = strRst & getLotteSafeToReg
    	'strRst = strRst & "<SAFETY_TEST_GB></SAFETY_TEST_GB>"                           ''��������
    	'strRst = strRst & "<SAFETY_TEST_CENTER></SAFETY_TEST_CENTER>"                   ''�������
    	'strRst = strRst & "<SAFETY_TEST_NO></SAFETY_TEST_NO>"                           ''������ȣ
    	'strRst = strRst & "<SAFETY_MODEL_NAME></SAFETY_MODEL_NAME>"                     ''�𵨸�
    	'strRst = strRst & "<SAFETY_TEST_DATE></SAFETY_TEST_DATE>"                       ''������
    	strRst = strRst & "<MODEL_NAME><![CDATA["&FItemName&"]]></MODEL_NAME>"          ''��ǰ����� �𵨸�
    	strRst = strRst & "<HEADCOPY>"&CHEADCOPY&"</HEADCOPY>"                 ''��ǰ����� ���ī��
    	strRst = strRst & "<GIFT_GOODS></GIFT_GOODS>"                                   ''��ǰ����� ����ǰ ���� text�Է�
    	strRst = strRst & "<RETURN_CONDITION></RETURN_CONDITION>"                       ''��ǰ����� ��ǰ���� ���� text�Է�
    	strRst = strRst & "<CARE_NOTE><![CDATA["&ForderComment&"]]></CARE_NOTE>"                                     ''��ǰ����� ���ǻ��� ���� text�Է�
    	strRst = strRst & "<DETAIL_HTML><![CDATA["&getLotteItemContParamToReg&"]]></DETAIL_HTML>"  ''html�Է°���, stringŸ������ �Է��ϸ� ���� ��ȯ�ؼ� �ۼ���
    	strRst = strRst & "<MC_NOTE></MC_NOTE>"                                         ''����ϱ����
    	strRst = strRst & getLotteCateParamToReg
    	strRst = strRst & getLotteAddImageParamToReg
    	strRst = strRst & getLotteOptionParamToReg(true,false)
'2012-11-07 ��������' ��ǰǰ����������
		strRst = strRst & getLotteimallItemInfoCdToReg(ioriginCode)


'    	strRst = strRst & "<LISART_INFO>"
'    	strRst = strRst & "<LISART_CODE></LISART_CODE>"
'    	strRst = strRst & "<LISART_CSTN_CODE></LISART_CSTN_CODE>"
'    	strRst = strRst & "<LISART_CSTN_DG1_CNTT></LISART_CSTN_DG1_CNTT>"
'    	strRst = strRst & "<LISART_CSTN_DG2_CNTT></LISART_CSTN_DG2_CNTT>"
'    	strRst = strRst & "<LISART_CSTN_DG3_CNTT></LISART_CSTN_DG3_CNTT>"
'    	strRst = strRst & "<LISART_CSTN_DG4_CNTT></LISART_CSTN_DG4_CNTT>"
'    	strRst = strRst & "<LISART_CSTN_DG5_CNTT></LISART_CSTN_DG5_CNTT>"
'    	strRst = strRst & "</LISART_INFO>"
'2012-11-07 ��������' ��ǰǰ���������� ��

    	strRst = strRst & "</GoodsEntryLineItem>"
    	strRst = strRst & "</GoodsEntry>"
        strRst = strRst & "</MessageBody>"
        strRst = strRst & "</GoodsEntry_V01>"

'    	strRst = strRst & "<CATEGORY_CODE1></CATEGORY_CODE1><CATEGORY_CODE2></CATEGORY_CODE2>~<CATEGORY_CODE10></CATEGORY_CODE10>"
'    	strRst = strRst & "<IMG_L>�̹���URL</IMG_L><IMG_L1></IMG_L1>~<IMG_L8></IMG_L8>"

'    	strRst = strRst & "<GOODSDT_D1>color</GOODSDT_D1>"
'    	strRst = strRst & "<GOODSDT_D2>size</GOODSDT_D2>"
'    	strRst = strRst & "<GOODSDT_D3></GOODSDT_D3>"
'    	strRst = strRst & "<GOODSDT_INFO>"
'    	strRst = strRst & "<ENTP_DT_CODE>20110324_1</ENTP_DT_CODE>"
'    	strRst = strRst & "	<GOODSDT_COLOR>999</GOODSDT_COLOR>"
'    	strRst = strRst & "	<GOODSDT_COLORNAME>����</GOODSDT_COLORNAME>"
'    	strRst = strRst & "	<GOODSDT_SIZE>999</GOODSDT_SIZE>"
'    	strRst = strRst & "	<GOODSDT_SIZENAME> 90</GOODSDT_SIZENAME>"
'    	strRst = strRst & "	<GOODSDT_PATTERN>000</GOODSDT_PATTERN>"
'    	strRst = strRst & "	<GOODSDT_PATTERNNAME />"
'    	strRst = strRst & "	<GOODSDT_FDATE>20110324</GOODSDT_FDATE>"
'    	strRst = strRst & "	<GOODSDT_DAILY_CAPA>4</GOODSDT_DAILY_CAPA>"
'    	strRst = strRst & "	<GOODSDT_MAX_SALE>4</GOODSDT_MAX_SALE>"
'    	strRst = strRst & "	<GOODSDT_SAFE_STOCK>0</GOODSDT_SAFE_STOCK>"
'    	strRst = strRst & "</GOODSDT_INFO>"
'    	strRst = strRst & "<GOODSDT_INFO>"
'    	strRst = strRst & "<ENTP_DT_CODE>20110324_2</ENTP_DT_CODE>"
'    	strRst = strRst & "	<GOODSDT_COLOR>999</GOODSDT_COLOR>"
'    	strRst = strRst & "	<GOODSDT_COLORNAME>��</GOODSDT_COLORNAME>"
'    	strRst = strRst & "	<GOODSDT_SIZE>999</GOODSDT_SIZE>"
'    	strRst = strRst & "	<GOODSDT_SIZENAME> 90</GOODSDT_SIZENAME>"
'    	strRst = strRst & "	<GOODSDT_PATTERN>000</GOODSDT_PATTERN>"
'    	strRst = strRst & "	<GOODSDT_PATTERNNAME />"
'    	strRst = strRst & "	<GOODSDT_FDATE>20110324</GOODSDT_FDATE>"
'    	strRst = strRst & "	<GOODSDT_DAILY_CAPA>1</GOODSDT_DAILY_CAPA>"
'    	strRst = strRst & "	<GOODSDT_MAX_SALE>1</GOODSDT_MAX_SALE>"
'    	strRst = strRst & "	<GOODSDT_SAFE_STOCK>0</GOODSDT_SAFE_STOCK>"
'    	strRst = strRst & "</GOODSDT_INFO>"
''=========================================================================================

		'��� ��ȯ
		getLTiMallItemRegXML = strRst
	end Function

	'// ��ǰ���� �Ķ���� ����
	public Function getLTiMallItemModXML()
		dim strRst
        dim iORIGIN_CODE
        dim iORIGIN_NAME
        iORIGIN_CODE= getOriginName2Code(Fsourcearea,iORIGIN_NAME)
''rw iORIGIN_CODE&","&iORIGIN_NAME
'        if (iORIGIN_CODE="9996") then
'            iORIGIN_NAME="���� �� ���Ի�"
'        else
'            ''iORIGIN_NAME=getOriginName2EditName(Fsourcearea)
'            iORIGIN_NAME = getOriginCode2EditName(iORIGIN_CODE)
'        end if

        strRst = "<?xml version=""1.0"" encoding=""utf-8"" ?>" '''
		strRst = strRst & "<ModifyGoods_V01>"
        strRst = strRst & "<MessageHeader>"
    	strRst = strRst & "<SENDER>TENBYTEN</SENDER>"
    	strRst = strRst & "<RECEIVER>LotteH</RECEIVER>"
    	strRst = strRst & "<DATETIME>"&replace(Left(FormatDateTime(now,0),10),"-","")&" "&Left(FormatDateTime(now,4),5)&Right(FormatDateTime(now,0),3)&"</DATETIME>"
    	strRst = strRst & "<DOCUMENTID>MODIFYGOODS</DOCUMENTID>"
    	strRst = strRst & "<ERROROCCUR>N</ERROROCCUR>"
    	strRst = strRst & "<ERRORMESSAGE>N</ERRORMESSAGE>"
        strRst = strRst & "</MessageHeader>"
		strRst = strRst & "<MessageBody>"
    	strRst = strRst & "<ModifyGoods>"
    	strRst = strRst & "<ENTP_CODE>"&ENTP_CODE&"</ENTP_CODE>"   ' in incLotteiMallFunction
    	strRst = strRst & "<ModifyGoodsLineItem>"
    	if (FItemID=210499) or (FItemID=724724) or (FItemID=692489) then
    	    strRst = strRst & "<ENTP_GOODS_CODE>"&999&FItemID&"</ENTP_GOODS_CODE>"
        else
    	    strRst = strRst & "<ENTP_GOODS_CODE>"&FItemID&"</ENTP_GOODS_CODE>"
        end if
    	strRst = strRst & "<GOODS_NAME><![CDATA["&CPREFIXITEMNAME&getItemNameFormat()&"]]></GOODS_NAME>"
    	''K1099999 ��Ȱ/��ȭ ��Ÿ��ȭ ��Ÿ��ȭ ��Ÿ ��Ÿ ���� ����
    	'strRst = strRst & "<LGROUP>K1</LGROUP>"                    '��ǰ ��з�
    	'strRst = strRst & "<MGROUP>09</MGROUP>"                    '��ǰ �ߺз�
    	'strRst = strRst & "<SGROUP>99</SGROUP>"                    '��ǰ �Һз�
    	'strRst = strRst & "<DGROUP>99</DGROUP>"                    '��ǰ ���з�
    	'strRst = strRst & "<COMMENT_GB></COMMENT_GB>"                                   ''NULL:�Ϲݻ�ǰ�� 10:�Ƿ� 20:��ǰ
    	'strRst = strRst & "<MD_CODE>"&MD_CODE&"</MD_CODE>"                              ''MD_CODE
    	'strRst = strRst & "<KEYWORD>"&Fkeywords&"</KEYWORD>"                            ''keyword
    	'strRst = strRst & "<BRAND_CODE>"&BRAND_CODE&"</BRAND_CODE>"                     ''�귣���ڵ�
    	'strRst = strRst & "<BRAND_NAME>"&BRAND_NAME&"</BRAND_NAME>"                     ''�귣���ڵ��
    	'strRst = strRst & "<MAKECO_CODE>"&MAKECO_CODE&"</MAKECO_CODE>"                  ''������ü�ڵ�
    	'strRst = strRst & "<MAKECO_NAME>"&Fmakername&"</MAKECO_NAME>"                       ''������ü��
    	strRst = strRst & "<ORIGIN_CODE>"&iORIGIN_CODE&"</ORIGIN_CODE>"   ''�������ڵ�
    	strRst = strRst & "<ORIGIN_NAME>"&iORIGIN_NAME&"</ORIGIN_NAME>"                       ''��������
    	'strRst = strRst & "<DELY_TYPE>0</DELY_TYPE>"                                    ''��۹�� 0:���¾�ü��� 1:�Ե�Ȩ�������ù�
    	'strRst = strRst & "<EXCH_YN>Y</EXCH_YN>"                                        ''��ȯ���� N:�Ұ� Y:����
    	'strRst = strRst & "<RETURN_YN>Y</RETURN_YN>"                                    ''��ǰ���� N:�Ұ� Y:����
    	'strRst = strRst & "<GIFT_RETURN_YN>N</GIFT_RETURN_YN>"                          ''����ǰȸ���ʼ� N:�Ұ� Y:����
    	'strRst = strRst & "<DELY_SHAP_CODE>000</DELY_SHAP_CODE>"                        ''������� 000:�Ϲݹ�� 201:��ġ���
    	'strRst = strRst & "<MAX_ORD_PSBT_CQTY></MAX_ORD_PSBT_CQTY>"                     ''�� ���� �� ���� �ִ� ������ �� �ִ� ���� (�Ϻ�), ���� ������ Default: 20
    	'strRst = strRst & "<MIN_ORD_PSBT_CQTY></MIN_ORD_PSBT_CQTY>"                     ''�� ���� �� ���� �ּ� ~�� �̻� �����ؾ� �ϴ� ���� (�ֹ���), ���� ������ Default: 1
    	'strRst = strRst & "<MIXPACK_YN>Y</MIXPACK_YN>"                                  ''�����忩�� N:�Ұ� Y:����
    	'strRst = strRst & "<DAMT_APLC_ORD></DAMT_APLC_ORD>"                             ''�ֹ���ۺ񱸺� NULL:���»纰�� ���õ� �⺻ ��ۺ� ����(MD�� Ȯ��)  1: ������ 2:���� 3:������
    	'strRst = strRst & "<DAMT_APLC_ORD_AMT></DAMT_APLC_ORD_AMT>"                     ''�ֹ���ۺ� ��ۺ񱸺��� 2,3�� ��츸 ��ȿ��. �ݾ� �Է� ��, �� ��ǰ�� �⺻ ��ۺ� �����ϰ� �Էµ� �ݾ����� ��ۺ� ������
    	'strRst = strRst & "<DAMT_APLC_REGD></DAMT_APLC_REGD>"                           ''��ǰȸ���񱸺� NULL:���»纰�� ���õ� �⺻ ��ۺ� ����(MD�� Ȯ��)  1: ������ 2:���� 3:������
    	'strRst = strRst & "<DAMT_APLC_REGD_AMT></DAMT_APLC_REGD_AMT>"                   ''��ǰȸ����
    	'strRst = strRst & "<DAMT_APLC_EXCH></DAMT_APLC_EXCH>"                           ''��ȯ��ۺ񱸺�
    	'strRst = strRst & "<DAMT_APLC_EXCH_AMT></DAMT_APLC_EXCH_AMT>"                   ''��ȯ��ۺ�
    	'strRst = strRst & "<AS_TERM>1</AS_TERM>"                                        ''AS �����Ⱓ(������)
    	'strRst = strRst & "<AS_REPAIR_TERM>1</AS_REPAIR_TERM>"                          ''AS �ҿ�Ⱓ(�ϼ�)
    	'strRst = strRst & "<AS_RECEIVE_TYPE>10</AS_RECEIVE_TYPE>"                       ''AS ������ü 10:�Ե�Ȩ���� 20:��ü�ȳ�
    	'strRst = strRst & "<AS_DELY_TYPE>20</AS_DELY_TYPE>"                             ''AS �����ü 10:�Ե�Ȩ���� 20:���¾�ü
    	'strRst = strRst & "<AS_OUT_COM>1</AS_OUT_COM>"                                  ''AS ����ü �������SEQ //Ȯ��
    	'strRst = strRst & "<AS_RETURN_COM>1</AS_RETURN_COM>"                            ''AS ȸ����ü ȸ�������SEQ //Ȯ��
    	'strRst = strRst & "<AS_NOTE></AS_NOTE>"                                         ''AS ����
    	'strRst = strRst & "<ENTP_MAN_SEQ>1</ENTP_MAN_SEQ>"                              ''�������  ���������SEQ  //Ȯ��
    	'strRst = strRst & "<OUTPLACE_SEQ>1</OUTPLACE_SEQ>"                              ''�����  �������SEQ  //Ȯ��
    	'strRst = strRst & "<RETURN_SEQ>1</RETURN_SEQ>"                                  ''ȸ�����  ȸ�������SEQ  //Ȯ��
    	strRst = strRst & "<BUY_PRICE>"&getLTiMallSuplyPrice()&"</BUY_PRICE>"           ''���ް�(���԰�)
    	strRst = strRst & "<SALE_PRICE>"&FSellCash&"</SALE_PRICE>"                      ''�ǸŰ�
    strRst = strRst & "<TAX_YN>"&CHKIIF(FVatInclude="N","2","1")&"</TAX_YN>"        ''1:���� 2:�鼼
    	strRst = strRst & "<NOREST_ALLOW_MONTH>"&getNOREST_ALLOW_MONTH()&"</NOREST_ALLOW_MONTH>"                   ''�����ڰ�����
    	strRst = strRst & "<APPLY_DATE>"&replace(Left(now(),10),"-","")&"</APPLY_DATE>" '' ������������ �����̸� sysdate�� �ú��ʷ� ����ǰ�, ���� ���ĸ� sysdate�� 00:00:00�� ����ǰ� ��
'2012-11-07 ��������' ������������
''		strRst = strRst & getLotteSafeToReg
'    	strRst = strRst & "<SAFETY_TEST_GB></SAFETY_TEST_GB>"                           ''��������
'    	strRst = strRst & "<SAFETY_TEST_CENTER></SAFETY_TEST_CENTER>"                   ''�������
'    	strRst = strRst & "<SAFETY_TEST_NO></SAFETY_TEST_NO>"                           ''������ȣ
'    	strRst = strRst & "<SAFETY_MODEL_NAME></SAFETY_MODEL_NAME>"                     ''�𵨸�
'    	strRst = strRst & "<SAFETY_TEST_DATE></SAFETY_TEST_DATE>"                       ''������
'2012-11-07 ��������' ������������ ��
    	IF (application("Svr_Info")="Dev") THEN  ''�׼��� ���� �ȵȴٰ�..
    	    strRst = strRst & "<MODEL_NAME><![CDATA["&FItemName&"]]></MODEL_NAME>"
    	else
    	    strRst = strRst & "<MODEL_NAME><![CDATA["&FItemName&"]]></MODEL_NAME>"          ''��ǰ����� �𵨸�
        End IF
        IF (Not application("Svr_Info")="Dev") THEN  ''�׼��� ���� �ȵȴٰ�..
        	strRst = strRst & "<HEADCOPY>"&CHEADCOPY&"</HEADCOPY>"                 ''��ǰ����� ���ī��
        	strRst = strRst & "<GIFT_GOODS></GIFT_GOODS>"                                   ''��ǰ����� ����ǰ ���� text�Է�
        	strRst = strRst & "<RETURN_CONDITION></RETURN_CONDITION>"                       ''��ǰ����� ��ǰ���� ���� text�Է�
        	strRst = strRst & "<CARE_NOTE><![CDATA["&ForderComment&"]]></CARE_NOTE>"                                     ''��ǰ����� ���ǻ��� ���� text�Է�
        ENd IF
    	IF (application("Svr_Info")="Dev") THEN  ''�׼��� ���� �ȵȴٰ�..
    	    ''strRst = strRst & "<DETAIL_HTML><![CDATA[�����..]]></DETAIL_HTML>"
    	    strRst = strRst & "<DETAIL_HTML></DETAIL_HTML>"
    	ELSE
    	    strRst = strRst & "<DETAIL_HTML><![CDATA["&getLotteItemContParamToReg&"]]></DETAIL_HTML>"  ''html�Է°���, stringŸ������ �Է��ϸ� ���� ��ȯ�ؼ� �ۼ���
        END IF
    	strRst = strRst & "<MC_NOTE></MC_NOTE>"                                         ''����ϱ����**
    	IF (Not application("Svr_Info")="Dev") THEN  ''�׼��� ���� �ȵȴٰ�..
        	strRst = strRst & getLotteCateParamToReg
        	''if InStr(FbasicImage,"-")>0 then
            ''	strRst = strRst & getLotteAddImageParamToReg                          '''�̹��� ���� �����ɸ� //�̺κ� �������.
            ''end if
        end if
        IF (Not application("Svr_Info")="Dev") THEN  ''�׼��� ���� �ȵȴٰ�..
    	    ''strRst = strRst & getLotteOptionParamToReg(false,false)
    	    strRst = strRst & getLotteOptionParamToEdit
        End IF

'2012-11-07 ��������' ��ǰǰ����������
		strRst = strRst & getLotteimallItemInfoCdToReg(iORIGIN_CODE)
'    	strRst = strRst & "<LISART_INFO>"
'    	strRst = strRst & "<LISART_CODE></LISART_CODE>"
'    	strRst = strRst & "<LISART_CSTN_CODE></LISART_CSTN_CODE>"
'    	strRst = strRst & "<LISART_CSTN_DG1_CNTT></LISART_CSTN_DG1_CNTT>"
'    	strRst = strRst & "<LISART_CSTN_DG2_CNTT></LISART_CSTN_DG2_CNTT>"
'    	strRst = strRst & "<LISART_CSTN_DG3_CNTT></LISART_CSTN_DG3_CNTT>"
'    	strRst = strRst & "<LISART_CSTN_DG4_CNTT></LISART_CSTN_DG4_CNTT>"
'    	strRst = strRst & "<LISART_CSTN_DG5_CNTT></LISART_CSTN_DG5_CNTT>"
'    	strRst = strRst & "</LISART_INFO>"
'2012-11-07 ��������' ��ǰǰ���������� ��
    	strRst = strRst & "</ModifyGoodsLineItem>"
    	strRst = strRst & "</ModifyGoods>"
        strRst = strRst & "</MessageBody>"
        strRst = strRst & "</ModifyGoods_V01>"

		'��� ��ȯ
		getLTiMallItemModXML = strRst

	end Function

    '// ��ǰ ���� �Ķ���� ����
    public Function getLTiMallItemModDTXML
        dim strRst

        strRst = "<?xml version=""1.0"" encoding=""utf-8"" ?>" '''
		strRst = strRst & "<ModifyGoodsDt_V01>"
        strRst = strRst & "<MessageHeader>"
    	strRst = strRst & "<SENDER>TENBYTEN</SENDER>"
    	strRst = strRst & "<RECEIVER>LotteH</RECEIVER>"
    	strRst = strRst & "<DATETIME>"&replace(Left(FormatDateTime(now,0),10),"-","")&" "&Left(FormatDateTime(now,4),5)&Right(FormatDateTime(now,0),3)&"</DATETIME>"
    	strRst = strRst & "<DOCUMENTID>MODIFYGOODSDT</DOCUMENTID>"
    	strRst = strRst & "<ERROROCCUR>N</ERROROCCUR>"
    	strRst = strRst & "<ERRORMESSAGE>N</ERRORMESSAGE>"
        strRst = strRst & "</MessageHeader>"
		strRst = strRst & "<MessageBody>"
    	strRst = strRst & "<ModifyGoodsDt>"
    	strRst = strRst & "<ENTP_CODE>"&ENTP_CODE&"</ENTP_CODE>"   ' in incLotteiMallFunction
    	if (FItemID=210499) or (FItemID=724724) or (FItemID=692489) then
    	    strRst = strRst & "<ENTP_GOODS_CODE>"&999&FItemID&"</ENTP_GOODS_CODE>"
        else
        	strRst = strRst & "<ENTP_GOODS_CODE>"&FItemID&"</ENTP_GOODS_CODE>"
        end if
    	''strRst = strRst & getLotteOptionParamToReg(false,false)
    	strRst = strRst & getLotteOptionParamToEdit()
    	strRst = strRst & "</ModifyGoodsDt>"
        strRst = strRst & "</MessageBody>"
        strRst = strRst & "</ModifyGoodsDt_V01>"
        '��� ��ȯ
		getLTiMallItemModDTXML = strRst
	end Function

	'// ��ǰ ����-ǰ�� �Ķ���� ����
    public Function getLTiMallItemSOLDOUTDTXML
        dim strRst

        strRst = "<?xml version=""1.0"" encoding=""utf-8"" ?>" '''
		strRst = strRst & "<ModifyGoodsDt_V01>"
        strRst = strRst & "<MessageHeader>"
    	strRst = strRst & "<SENDER>TENBYTEN</SENDER>"
    	strRst = strRst & "<RECEIVER>LotteH</RECEIVER>"
    	strRst = strRst & "<DATETIME>"&replace(Left(FormatDateTime(now,0),10),"-","")&" "&Left(FormatDateTime(now,4),5)&Right(FormatDateTime(now,0),3)&"</DATETIME>"
    	strRst = strRst & "<DOCUMENTID>MODIFYGOODSDT</DOCUMENTID>"
    	strRst = strRst & "<ERROROCCUR>N</ERROROCCUR>"
    	strRst = strRst & "<ERRORMESSAGE>N</ERRORMESSAGE>"
        strRst = strRst & "</MessageHeader>"
		strRst = strRst & "<MessageBody>"
    	strRst = strRst & "<ModifyGoodsDt>"
    	strRst = strRst & "<ENTP_CODE>"&ENTP_CODE&"</ENTP_CODE>"   ' in incLotteiMallFunction
    	if (FItemID=210499) or (FItemID=724724) or (FItemID=692489) then
    	    strRst = strRst & "<ENTP_GOODS_CODE>"&999&FItemID&"</ENTP_GOODS_CODE>"
        else
        	strRst = strRst & "<ENTP_GOODS_CODE>"&FItemID&"</ENTP_GOODS_CODE>"
        end if
    	strRst = strRst & getLotteOptionParamToSoldOut ''getLotteOptionParamToReg(false,false) '' false,true
    	strRst = strRst & "</ModifyGoodsDt>"
        strRst = strRst & "</MessageBody>"
        strRst = strRst & "</ModifyGoodsDt_V01>"
        '��� ��ȯ
		getLTiMallItemSOLDOUTDTXML = strRst
	end Function

    '// ��ǰ���: MD��ǰ�� �� ���� ī�װ� �Ķ���� ����(��ǰ��Ͽ�)
	public function getLotteCateParamToReg()
		dim strSql, strRst, i
		strSql = "Select top 6 c.CateKey "
		strSql = strSql & " from db_item.dbo.tbl_LtiMall_cate_mapping as m "
		strSql = strSql & " 	join db_temp.dbo.tbl_LtiMall_Category as c "
		strSql = strSql & " 		on m.CateKey=c.CateKey "
		strSql = strSql & " where tenCateLarge='" & FtenCateLarge & "' "
		strSql = strSql & " 	and tenCateMid='" & FtenCateMid & "' "
		strSql = strSql & " 	and tenCateSmall='" & FtenCateSmall & "' "
		strSql = strSql & " 	and c.isusing='Y'"
		''strSql = strSql & " and c.cateGbn='B'"
		strSql = strSql & " order by c.cateGbn asc " ''���� ī�װ� �켱.          'B : �귣�� / D : �Ϲ�

		rsget.Open strSql,dbget,1

		if Not(rsget.EOF or rsget.BOF) then
			strRst = ""

			i=0
			Do until rsget.EOF
				strRst = strRst & "<CATEGORY_CODE"&i+1&">" & rsget("CateKey") & "</CATEGORY_CODE"&i+1&">"
				rsget.MoveNext
				i=i+1
			Loop
		end if
'        if (i=1) then
'            strRst = strRst & "<CATEGORY_CODE"&i+1&"></CATEGORY_CODE"&i+1&">"
'        end if
		rsget.Close

		getLotteCateParamToReg = strRst
	end function

	''ǰ���� ���ϵ� �ɼǱ��� // ���� �ɼǸ�� ���� �ʿ�.
	public function getLotteOptionParamToSoldOut()
	    dim strSql, strRst, i, optDc
	    dim optNm, validSellno

	    dim iGOODSDT_FDATE: iGOODSDT_FDATE=replace(Left(dateAdd("d",0,Now()),10),"-","") ''replace(Left(dateAdd("d",1,Now()),10),"-","") ''���ϵ� ����

	    strSql = "select top 100 itemid,itemoption,outmalloptname from [db_item].[dbo].tbl_OutMall_regedoption R"
        strSql = strSql & " 	where R.itemid="&FItemid
        strSql = strSql & " 	and R.mallid='"&CMALLNAME&"'"

        rsget.Open strSql,dbget,1
	     '''rw strSql
		if Not(rsget.EOF or rsget.BOF) then

			optNm = "�ɼ�"
			optDc = ""

			optNm = "<GOODSDT_D1><![CDATA["&getGOODSDT_NmFormat(optNm)&"]]></GOODSDT_D1>"
			optNm = optNm & "<GOODSDT_D2></GOODSDT_D2>"
			optNm = optNm & "<GOODSDT_D3></GOODSDT_D3>"
			Do until rsget.EOF
			    validSellno =0


				optDc = optDc & "<GOODSDT_INFO>"
                optDc = optDc & "<ENTP_DT_CODE>"&FItemID&"_"&rsget("itemoption")&"</ENTP_DT_CODE>"
                optDc = optDc & "<GOODSDT_COLOR>999</GOODSDT_COLOR>"
                optDc = optDc & "<GOODSDT_COLORNAME><![CDATA["&Replace(Replace(db2Html(rsget("outmalloptname")),":",""),",","")&"]]></GOODSDT_COLORNAME>"
                optDc = optDc & "<GOODSDT_SIZE>000</GOODSDT_SIZE>"
        	    optDc = optDc & "<GOODSDT_SIZENAME></GOODSDT_SIZENAME>"
        	    optDc = optDc & "<GOODSDT_PATTERN>000</GOODSDT_PATTERN>"
        	    optDc = optDc & "<GOODSDT_PATTERNNAME></GOODSDT_PATTERNNAME>"
            	optDc = optDc & "<GOODSDT_FDATE>"&iGOODSDT_FDATE&"</GOODSDT_FDATE>"                       ''Default: ���� (yyyymmdd)
            	optDc = optDc & "<GOODSDT_DAILY_CAPA>"&validSellno&"</GOODSDT_DAILY_CAPA>"           ''�����޷� 1�� �̻� (0���� ������ �Ͻ��ߴ�)
            	optDc = optDc & "<GOODSDT_MAX_SALE>"&validSellno&"</GOODSDT_MAX_SALE>"  ''�ǸŰ��ɷ� 1�� �̻� (0���� ������ �Ͻ��ߴ�)
            	optDc = optDc & "<GOODSDT_SAFE_STOCK>0</GOODSDT_SAFE_STOCK>"
                optDc = optDc & "</GOODSDT_INFO>"

				rsget.MoveNext
			Loop
		end if
		rsget.Close

		getLotteOptionParamToSoldOut = optNm&optDc
    end function

	''// ���� �ɼ� ���·� ���.
	public function getLotteOptionParamToOneOption()

    end function

	'// ��ǰ���� : �ɼ� �Ķ���� ����(��ǰ������)
	''//  GOODSDT_D(N) �� ���� ����� �������µ� :: ��ǰ��Ͻ� ������ ��ǰ�Ӽ� �����̿ܿ� ���ο� ��ǰ�Ӽ� ������ �Է��� �� �����ϴ�.
    '' ==> �ذ� ��� GOODSDT_D1 �Ѱ�������(�ɼǾ��°�쿡��) ��Ƽ�ɼ��ΰ�쵵 ���� �ɼ����� ����
	public function getLotteOptionParamToEdit()
	    dim strSql , i
	    dim isOptionExists, arrRows
	    dim chkMultiOpt,optNm, optDc, validSellno
	    dim itemoption,optionname
        dim optlimityn
        dim optLimit
        dim isusing
        dim optsellyn
        dim opt1name
        dim opt2name
        dim opt3name
        dim preged
        dim optNameDiff
        dim forceExpire

        dim iGOODSDT_FDATE : iGOODSDT_FDATE=replace(Left(dateAdd("d",0,Now()),10),"-","") ''replace(Left(dateAdd("d",1,Now()),10),"-","")

	    strSql = "exec db_item.dbo.sp_Ten_OutMall_optEditParamList '"&CMallName&"'," & FItemid
        rsget.CursorLocation = adUseClient
		rsget.CursorType = adOpenStatic
		rsget.LockType = adLockOptimistic
        rsget.Open strSql, dbget
        if Not(rsget.EOF or rsget.BOF) then
            arrRows = rsget.getRows
        end if
        rsget.close

        isOptionExists = isArray(arrRows)

        if (isOptionExists) then
            chkMultiOpt = false
            '// ���߿ɼ��� ��
			'#�ɼǸ� ����
			strSql = "exec [db_item].[dbo].sp_Ten_ItemOptionMultipleTypeList " & FItemid
	        rsget.CursorLocation = adUseClient
			rsget.CursorType = adOpenStatic
			rsget.LockType = adLockOptimistic
	        rsget.Open strSql, dbget

			optNm = ""
			i=1
			if Not(rsget.EOF or rsget.BOF) then
				chkMultiOpt = true
				Do until rsget.EOF
					optNm = optNm & "<GOODSDT_D"&i&"><![CDATA["&getGOODSDT_NmFormat(rsget("optionTypeName"))&"]]></GOODSDT_D"&i&">"
					i=i+1
					rsget.MoveNext
				Loop

				if i=3 then
				    optNm = optNm & "<GOODSDT_D3></GOODSDT_D3>"
				end if
			end if
			rsget.Close

            For i =0 To UBound(ArrRows,2)

                if (Not chkMultiOpt) and (i=0) then
                    if db2Html(ArrRows(2,i))<>"" then
    					optNm = ArrRows(2,i)
    				else
    					optNm = "�ɼ�"
    				end if
    				optNm = "<GOODSDT_D1><![CDATA["&getGOODSDT_NmFormat(optNm)&"]]></GOODSDT_D1>"
    				optNm = optNm & "<GOODSDT_D2></GOODSDT_D2>"
    				optNm = optNm & "<GOODSDT_D3></GOODSDT_D3>"
				end if

                validSellno=50
                itemoption = ArrRows(1,i)
                optionname = Replace(Replace(db2Html(ArrRows(3,i)),":",""),",","")

				optlimityn = ArrRows(5,i)
			    optLimit   = ArrRows(4,i)
			    isusing    = ArrRows(6,i)
			    optsellyn  = ArrRows(7,i)
			    opt1name   = Replace(Replace(db2Html(ArrRows(8,i)),":",""),",","")
			    opt2name   = Replace(Replace(db2Html(ArrRows(9,i)),":",""),",","")
			    opt3name   = Replace(Replace(db2Html(ArrRows(10,i)),":",""),",","")
			    preged     = (ArrRows(11,i)=1)
			    optNameDiff = (ArrRows(12,i)=1)
			    forceExpire = (ArrRows(13,i)=1)

			    if (Not chkMultiOpt) then
			        opt1name = optionname
			        opt2name = ""
			        opt3name = ""
			    end if

			    if (FSellyn<>"Y") or ((optlimityn="Y") and (optLimit<1)) or (isusing<>"Y") or (optsellyn<>"Y") then
			        validSellno = 0
			    end if

			    if (optlimityn="Y") then
			        validSellno = optLimit
			    end if

			    if (validSellno<CMAXLIMITSELL) then validSellno=0
			    if (optlimityn="Y") and (validSellno>0) then
			        validSellno = validSellno-CMAXLIMITSELL
			    end if
			    if (validSellno<1) then validSellno=0

                if IsSoldOut then validSellno=0
                if (preged and optNameDiff) then validSellno=0
                if (forceExpire) then validSellno=0

                if (Not preged) and (validSellno=0) then
                    ''skip
                    rw "skip itemoption="&itemoption
                else
                    rw "1itemoption="&itemoption&CHKIIF(validSellno<1," :soldout","")
    			    optDc = optDc & "<GOODSDT_INFO>"
                    optDc = optDc & "<ENTP_DT_CODE>"&FItemID&"_"&itemoption&"</ENTP_DT_CODE>"
                    optDc = optDc & "<GOODSDT_COLOR>"&CHKIIF(opt1name<>"","999","000")&"</GOODSDT_COLOR>"
                    optDc = optDc & "<GOODSDT_COLORNAME><![CDATA["&opt1name&"]]></GOODSDT_COLORNAME>"
                    optDc = optDc & "<GOODSDT_SIZE>"&CHKIIF(opt2name<>"","999","000")&"</GOODSDT_SIZE>"
            	    optDc = optDc & "<GOODSDT_SIZENAME><![CDATA["&opt2name&"]]></GOODSDT_SIZENAME>"
            	    optDc = optDc & "<GOODSDT_PATTERN>"&CHKIIF(opt3name<>"","999","000")&"</GOODSDT_PATTERN>"
            	    optDc = optDc & "<GOODSDT_PATTERNNAME><![CDATA["&opt3name&"]]></GOODSDT_PATTERNNAME>"
                	optDc = optDc & "<GOODSDT_FDATE>"&iGOODSDT_FDATE&"</GOODSDT_FDATE>"                       ''Default: ���� (yyyymmdd)
                	optDc = optDc & "<GOODSDT_DAILY_CAPA>"&CHKIIF(validSellno<1,0,10)&"</GOODSDT_DAILY_CAPA>"           ''�����޷� 1�� �̻� (0���� ������ �Ͻ��ߴ�)
                	optDc = optDc & "<GOODSDT_MAX_SALE>"&validSellno&"</GOODSDT_MAX_SALE>"  ''�ǸŰ��ɷ� 1�� �̻� (0���� ������ �Ͻ��ߴ�)
                	optDc = optDc & "<GOODSDT_SAFE_STOCK>0</GOODSDT_SAFE_STOCK>"
                    optDc = optDc & "</GOODSDT_INFO>"
                end if
            Next
        end if


        if (Not isOptionExists) then
            validSellno = 50
            if (FLimitYN="Y") THEN
                validSellno = (FLimitno-FLimitSold)
            END IF
            if (validSellno>50) then validSellno=50

            if (validSellno<CMAXLIMITSELL) then validSellno=0
            if (validSellno<1) then validSellno=0

            if (validSellno>0) then
                validSellno = validSellno-CMAXLIMITSELL
            end if

            if IsSoldOut then validSellno=0

            optNm = "<GOODSDT_D1></GOODSDT_D1><GOODSDT_D2></GOODSDT_D2><GOODSDT_D3></GOODSDT_D3>" ''��ǰ�� ��� <GOODSDT_D1></GOODSDT_D1>�� ���� ������ �ȵȴٸ�..
            optDc = "<GOODSDT_INFO>"
            optDc = optDc & "<ENTP_DT_CODE>"&FItemID&"_0000</ENTP_DT_CODE>"
            optDc = optDc & "<GOODSDT_COLOR>000</GOODSDT_COLOR>"
            optDc = optDc & "<GOODSDT_COLORNAME></GOODSDT_COLORNAME>"
            optDc = optDc & "<GOODSDT_SIZE>000</GOODSDT_SIZE>"
    	    optDc = optDc & "<GOODSDT_SIZENAME></GOODSDT_SIZENAME>"
    	    optDc = optDc & "<GOODSDT_PATTERN>000</GOODSDT_PATTERN>"
    	    optDc = optDc & "<GOODSDT_PATTERNNAME></GOODSDT_PATTERNNAME>"
        	optDc = optDc & "<GOODSDT_FDATE>"&iGOODSDT_FDATE&"</GOODSDT_FDATE>"                       ''Default: ���� (yyyymmdd)
        	optDc = optDc & "<GOODSDT_DAILY_CAPA>"&CHKIIF(validSellno<1,0,5)&"</GOODSDT_DAILY_CAPA>"    '''
        	optDc = optDc & "<GOODSDT_MAX_SALE>"&validSellno&"</GOODSDT_MAX_SALE>"
        	optDc = optDc & "<GOODSDT_SAFE_STOCK>0</GOODSDT_SAFE_STOCK>"
            optDc = optDc & "</GOODSDT_INFO>"
        end if
        rw "isOptionExists="&isOptionExists
        'rw optNm&optDc
        getLotteOptionParamToEdit = optNm&optDc
    end function

	'// ��ǰ���: �ɼ� �Ķ���� ����(��ǰ��Ͽ�) ===>  ��ϵ� ��ǰ������ ���� �ʿ�.
	public function getLotteOptionParamToReg(byVal isReg,byval isNotUsingOptionInclude)
		dim strSql, strRst, i, optYn, optNm, optDc, chkMultiOpt
		dim validSellno, optlimityn, optLimit, isusing, optsellyn , opt1name , opt2name, opt3name
		dim iGOODSDT_FDATE
		dim isItemSoldout, preged, isdtEsists
		isdtEsists = false

		if (isReg) then
		    iGOODSDT_FDATE=replace(Left(dateAdd("d",0,Now()),10),"-","") ''replace(Left(dateAdd("d",1,Now()),10),"-","") ''���ϵ� ����
		else
		    iGOODSDT_FDATE=replace(Left(dateAdd("d",0,Now()),10),"-","") ''replace(Left(dateAdd("d",1,Now()),10),"-","")
	    end if

		chkMultiOpt = false
		optYn = "N"

        if (isNotUsingOptionInclude) then FoptionCnt=1

        if (not isReg) and (FoptionCnt=0) then  ''�ɼ��� ���� ��� �������� �����ϸ�, �ɼ� ����0���� �����Ƿ� �����
            strSql = "select count(*) as CNT from [db_item].[dbo].tbl_OutMall_regedoption where itemid="&FItemid&" and mallid='"&CMALLNAME&"'"
            rsget.Open strSql,dbget,1
            if Not(rsget.EOF or rsget.BOF) then
                FoptionCnt = rsget("CNT")
            end if
            rsget.close
        end if
    ''rw "FoptionCnt="&FoptionCnt
		if (FoptionCnt>0) then
			'// ���߿ɼ��� ��
			'#�ɼǸ� ����
			strSql = "exec [db_item].[dbo].sp_Ten_ItemOptionMultipleTypeList " & FItemid
	        rsget.CursorLocation = adUseClient
			rsget.CursorType = adOpenStatic
			rsget.LockType = adLockOptimistic
	        rsget.Open strSql, dbget

			optNm = ""
			i=1
			if Not(rsget.EOF or rsget.BOF) then
				chkMultiOpt = true
				optYn = "Y"
				Do until rsget.EOF
					optNm = optNm & "<GOODSDT_D"&i&"><![CDATA["&getGOODSDT_NmFormat(rsget("optionTypeName"))&"]]></GOODSDT_D"&i&">"
					i=i+1
					rsget.MoveNext
				Loop
			end if
			rsget.Close

			'#�ɼǳ��� ����
			if chkMultiOpt then

				strSql = " select o.itemoption, o.optionTypeName, o.optionname, (o.optlimitno-o.optlimitsold) as optLimit, o.optlimityn, o.isUsing, o.optsellyn"
                strSql = strSql & " ,IsNULL((select optionKindName from db_item.dbo.tbl_item_option_Multiple p"
                strSql = strSql & "     where p.itemid="&FItemid&" and p.TypeSeq=1 and p.kindSeq=substring(o.itemoption,2,1)"
                strSql = strSql & "     ),'') as opt1name"
                strSql = strSql & " ,IsNULL((select optionKindName"
                strSql = strSql & "     from db_item.dbo.tbl_item_option_Multiple p"
                strSql = strSql & "     where p.itemid="&FItemid&" and p.TypeSeq=2 and p.kindSeq=substring(o.itemoption,3,1)"
                strSql = strSql & "     ),'') as opt2name"
                strSql = strSql & " ,IsNULL((select optionKindName"
                strSql = strSql & "     from db_item.dbo.tbl_item_option_Multiple p"
                strSql = strSql & "     where p.itemid="&FItemid&" and p.TypeSeq=3 and p.kindSeq=substring(o.itemoption,4,1)"
                strSql = strSql & "     ),'') as opt3name"
                strSql = strSql & " ,(CASe WHEN R.itemid is NULL THEN 0 ELSE 1 END ) as preged"
                strSql = strSql & " from [db_item].[dbo].tbl_item_option o"
                strSql = strSql & " 	left join [db_item].[dbo].tbl_OutMall_regedoption R"
                strSql = strSql & " 	on o.itemid=R.itemid"
                strSql = strSql & " 	and o.itemoption=R.itemoption"
                strSql = strSql & " 	and R.mallid='"&CMALLNAME&"'"
                strSql = strSql & " where o.itemid="&FItemid
                strSql = strSql & " 	and o.optaddprice=0 "                     '''�߰��ݾ� �Ұ�.
                if (isNotUsingOptionInclude) then
                    strSql = strSql & " 	and ((R.itemid is Not NULL)) "
                else
				    strSql = strSql & " 	and ((o.isUsing='Y') or (R.itemid is Not NULL)) "  '''"and optsellyn='Y' "
			    end if
			''rw strSql
				rsget.Open strSql,dbget,1

				optDc = ""

				if Not(rsget.EOF or rsget.BOF) then
					Do until rsget.EOF
					    validSellno=50
						optlimityn = rsget("optlimityn")
					    optLimit   = rsget("optLimit")
					    isusing    = rsget("isUsing")
					    optsellyn  = rsget("optsellyn")
					    opt1name   = Replace(Replace(db2Html(rsget("opt1name")),":",""),",","")
					    opt2name   = Replace(Replace(db2Html(rsget("opt2name")),":",""),",","")
					    opt3name   = Replace(Replace(db2Html(rsget("opt3name")),":",""),",","")
					    preged     = (rsget("preged")=1)

					    if (FSellyn<>"Y") or ((optlimityn="Y") and (optLimit<1)) or (isusing<>"Y") or (optsellyn<>"Y") then
					        validSellno = 0
					    end if

					    if (optlimityn="Y") then
					        validSellno = optLimit
					    end if

					    if (validSellno<CMAXLIMITSELL) then validSellno=0
					    if (optlimityn="Y") and (validSellno>0) then
					        validSellno = validSellno-CMAXLIMITSELL
					    end if
					    if (validSellno<1) then validSellno=0

                        if IsSoldOut then validSellno=0

                        ''if (FRectMode="MDT") and (Not isReg) and (preged=0) and validSellno=0 then validSellno=1 '' ��ǰ��� 0�̻��̾�� �űԴ�ǰ�߰� ����
					    '' ��ǰ��Ͻ� ������ ��ǰ�Ӽ� �����̿ܿ� ���ο� ��ǰ�Ӽ� ������ �Է��� �� �����ϴ�.



					    if ((isReg) and (validSellno>0)) or (Not isReg) then
					        if (Not isReg) and (Not preged) and (validSellno<1) then
					            ''SKIP :: �������鼭, ������ ��Ͼȵ� ��ǰ�� ���� 0�ΰ�� ����.
					            'rw rsget("optionname")&validSellno
					        '''elseif  (Not isReg) and (Not preged) then

					        else
        						optDc = optDc & "<GOODSDT_INFO>"
                                optDc = optDc & "<ENTP_DT_CODE>"&FItemID&"_"&rsget("itemoption")&"</ENTP_DT_CODE>"
                                optDc = optDc & "<GOODSDT_COLOR>"&CHKIIF(opt1name<>"","999","000")&"</GOODSDT_COLOR>"
                                optDc = optDc & "<GOODSDT_COLORNAME><![CDATA["&opt1name&"]]></GOODSDT_COLORNAME>"
                                optDc = optDc & "<GOODSDT_SIZE>"&CHKIIF(opt2name<>"","999","000")&"</GOODSDT_SIZE>"
                        	    optDc = optDc & "<GOODSDT_SIZENAME><![CDATA["&opt2name&"]]></GOODSDT_SIZENAME>"
                        	    optDc = optDc & "<GOODSDT_PATTERN>"&CHKIIF(opt3name<>"","999","000")&"</GOODSDT_PATTERN>"
                        	    optDc = optDc & "<GOODSDT_PATTERNNAME><![CDATA["&opt3name&"]]></GOODSDT_PATTERNNAME>"
                            	optDc = optDc & "<GOODSDT_FDATE>"&iGOODSDT_FDATE&"</GOODSDT_FDATE>"                       ''Default: ���� (yyyymmdd)
                            	optDc = optDc & "<GOODSDT_DAILY_CAPA>"&CHKIIF(validSellno<1,0,10)&"</GOODSDT_DAILY_CAPA>"           ''�����޷� 1�� �̻� (0���� ������ �Ͻ��ߴ�)
                            	optDc = optDc & "<GOODSDT_MAX_SALE>"&validSellno&"</GOODSDT_MAX_SALE>"  ''�ǸŰ��ɷ� 1�� �̻� (0���� ������ �Ͻ��ߴ�)
                            	optDc = optDc & "<GOODSDT_SAFE_STOCK>0</GOODSDT_SAFE_STOCK>"
                                optDc = optDc & "</GOODSDT_INFO>"

                                isdtEsists = true
                            end if
                        end if
						rsget.MoveNext
					Loop
				end if
				rsget.Close
			end if


			'// ���Ͽɼ��� ��
			if Not(chkMultiOpt) then
			    optNm = "" : optDc = ""         ''�ʱ�ȭ.

				strSql = "Select o.itemoption, (CASE WHEN convert(varchar(18),o.optionTypeName)<>o.optionTypeName THEN '�ɼǼ���' ELSE o.optionTypeName END) as optionTypeName, o.optionname, (o.optlimitno-o.optlimitsold) as optLimit, o.optlimityn, o.isUsing, o.optsellyn "
				strSql = strSql & " ,(CASE WHEN  R.itemoption is NULL THEN 0 ELSE 1 END) as preged"
				strSql = strSql & " From [db_item].[dbo].tbl_item_option o"
				strSql = strSql & "     left join [db_item].[dbo].tbl_OutMall_regedoption R"
				strSql = strSql & "     on o.itemid=R.itemid"
				strSql = strSql & "     and o.itemoption=R.itemoption"
				strSql = strSql & "     and R.mallid='"&CMALLNAME&"'"
				strSql = strSql & " where o.itemid=" & FItemid
				strSql = strSql & " 	and o.optaddprice=0 "                     '''�߰��ݾ� �Ұ�.
				if (isNotUsingOptionInclude) then
				    strSql = strSql & " 	and ((R.itemid is Not NULL)) "
				else
    				strSql = strSql & " 	and ((o.isUsing='Y') or (R.itemid is Not NULL)) "  '''"and optsellyn='Y' "
    				'''strSql = strSql & " 	and (optlimityn='N' or (optlimityn='Y' and optlimitno-optlimitsold>="&CMAXLIMITSELL&")) "
    		    end if

				rsget.Open strSql,dbget,1
	     '''rw strSql
				if Not(rsget.EOF or rsget.BOF) then
				    ''rw "���Ͽɼ�"

					optYn = "Y"
					if db2Html(rsget("optionTypeName"))<>"" then
						optNm = rsget("optionTypeName")
					else
						optNm = "�ɼ�"
					end if
					optNm = "<GOODSDT_D1><![CDATA["&getGOODSDT_NmFormat(optNm)&"]]></GOODSDT_D1>"
					optNm = optNm & "<GOODSDT_D2></GOODSDT_D2>"
					optNm = optNm & "<GOODSDT_D3></GOODSDT_D3>"
					Do until rsget.EOF
					    validSellno=50
					    optlimityn = rsget("optlimityn")
					    optLimit   = rsget("optLimit")
					    isusing    = rsget("isUsing")
					    optsellyn  = rsget("optsellyn")
					    preged     = (rsget("preged")=1)
					    if (FSellyn<>"Y") or ((optlimityn="Y") and (optLimit<1)) or (isusing<>"Y") or (optsellyn<>"Y") then
					        validSellno = 0
					    end if

					    if (optlimityn="Y") then
					        validSellno = optLimit
					    end if


					    if (validSellno<CMAXLIMITSELL) then validSellno=0

					    if (optlimityn="Y") and (validSellno>0) then
					        validSellno = validSellno-CMAXLIMITSELL
					    end if

					    if (validSellno<1) then validSellno=0

					    if IsSoldOut then validSellno=0

					    ''if (FRectMode="MDT") and (Not isReg) and (preged=0) and validSellno=0 then validSellno=1 '' ��ǰ��� 0�̻��̾�� �űԴ�ǰ�߰� ����
					    '' ��ǰ��Ͻ� ������ ��ǰ�Ӽ� �����̿ܿ� ���ο� ��ǰ�Ӽ� ������ �Է��� �� �����ϴ�.

					    if ((isReg) and (validSellno>0)) or (Not isReg) then
					        if (Not isReg) and (Not preged) and (validSellno<1) then
					            ''SKIP :: �������鼭, ������ ��Ͼȵ� ��ǰ�� ���� 0�ΰ�� ����.
					            'rw rsget("optionname")&validSellno
					        else
        						optDc = optDc & "<GOODSDT_INFO>"
                                optDc = optDc & "<ENTP_DT_CODE>"&FItemID&"_"&rsget("itemoption")&"</ENTP_DT_CODE>"
                                optDc = optDc & "<GOODSDT_COLOR>999</GOODSDT_COLOR>"
                                optDc = optDc & "<GOODSDT_COLORNAME><![CDATA["&Replace(Replace(db2Html(rsget("optionname")),":",""),",","")&"]]></GOODSDT_COLORNAME>"
                                optDc = optDc & "<GOODSDT_SIZE>000</GOODSDT_SIZE>"
                        	    optDc = optDc & "<GOODSDT_SIZENAME></GOODSDT_SIZENAME>"
                        	    optDc = optDc & "<GOODSDT_PATTERN>000</GOODSDT_PATTERN>"
                        	    optDc = optDc & "<GOODSDT_PATTERNNAME></GOODSDT_PATTERNNAME>"
                            	optDc = optDc & "<GOODSDT_FDATE>"&iGOODSDT_FDATE&"</GOODSDT_FDATE>"                       ''Default: ���� (yyyymmdd)
                            	optDc = optDc & "<GOODSDT_DAILY_CAPA>"&CHKIIF(validSellno<1,0,10)&"</GOODSDT_DAILY_CAPA>"           ''�����޷� 1�� �̻� (0���� ������ �Ͻ��ߴ�)
                            	optDc = optDc & "<GOODSDT_MAX_SALE>"&validSellno&"</GOODSDT_MAX_SALE>"  ''�ǸŰ��ɷ� 1�� �̻� (0���� ������ �Ͻ��ߴ�)
                            	optDc = optDc & "<GOODSDT_SAFE_STOCK>0</GOODSDT_SAFE_STOCK>"
                                optDc = optDc & "</GOODSDT_INFO>"

                                isdtEsists = true
                            end if
                        end if
						rsget.MoveNext
					Loop
				end if
				rsget.Close
			end if
		end if

        if (Not isdtEsists) and (optYn="Y") then optYn="N"

        IF (optYn<>"Y") THEN
            validSellno = 50
            if (FLimitYN="Y") THEN
                validSellno = (FLimitno-FLimitSold)
            END IF
            if (validSellno>50) then validSellno=50

            if (validSellno<CMAXLIMITSELL) then validSellno=0
            if (validSellno<1) then validSellno=0

            if (validSellno>0) then
                validSellno = validSellno-CMAXLIMITSELL
            end if



            if IsSoldOut then validSellno=0

            optNm = "<GOODSDT_D1></GOODSDT_D1><GOODSDT_D2></GOODSDT_D2><GOODSDT_D3></GOODSDT_D3>" ''��ǰ�� ��� <GOODSDT_D1></GOODSDT_D1>�� ���� ������ �ȵȴٸ�..
            optDc = "<GOODSDT_INFO>"
            optDc = optDc & "<ENTP_DT_CODE>"&FItemID&"_0000</ENTP_DT_CODE>"
            optDc = optDc & "<GOODSDT_COLOR>000</GOODSDT_COLOR>"
            optDc = optDc & "<GOODSDT_COLORNAME></GOODSDT_COLORNAME>"
            optDc = optDc & "<GOODSDT_SIZE>000</GOODSDT_SIZE>"
    	    optDc = optDc & "<GOODSDT_SIZENAME></GOODSDT_SIZENAME>"
    	    optDc = optDc & "<GOODSDT_PATTERN>000</GOODSDT_PATTERN>"
    	    optDc = optDc & "<GOODSDT_PATTERNNAME></GOODSDT_PATTERNNAME>"
        	optDc = optDc & "<GOODSDT_FDATE>"&iGOODSDT_FDATE&"</GOODSDT_FDATE>"                       ''Default: ���� (yyyymmdd)
        	optDc = optDc & "<GOODSDT_DAILY_CAPA>"&CHKIIF(validSellno<1,0,5)&"</GOODSDT_DAILY_CAPA>"    '''
        	optDc = optDc & "<GOODSDT_MAX_SALE>"&validSellno&"</GOODSDT_MAX_SALE>"
        	optDc = optDc & "<GOODSDT_SAFE_STOCK>0</GOODSDT_SAFE_STOCK>"
            optDc = optDc & "</GOODSDT_INFO>"
        END IF

		getLotteOptionParamToReg = optNm&optDc
	end function

    ''Ȯ�ο��..�������� // ���� �ȵǴµ�..
	public function getLotteSafeToReg()
	    'getLotteSafeToReg = ""
	    'Exit function

		Dim strRst, strSQL
		Dim igub, igov, inum
		strSQL = ""
		strSQL = strSQL & "SELECT top 1 * FROM "
		strSQL = strSQL & "db_item.dbo.tbl_item_contents " & vbcrlf
		strSQL = strSQL & "where safetyyn='Y' and itemid = '"&Fitemid&"' " & vbcrlf
		rsget.Open strSQL, dbget

			If Not(rsget.EOF or rsget.BOF) Then
				Select Case rsget("safetyDiv")
					Case "10"			'������������(KC��ũ)
						igub = "11"
						igov = "11"
					Case "20"			'�����ǰ ��������
						igub = "21"
						igov = "21"
					Case "30"			'KPS �������� ǥ��
						igub = "11"
						igov = "11"
					Case "40"			'KPS �������� Ȯ�� ǥ��
						igub = "11"
						igov = "11"
					Case "50"			'KPS ��� ��ȣ���� ǥ��
						igub = "11"
						igov = "11"
				End Select
				inum = rsget("safetyNum")

				strRst = ""
		    	strRst = strRst & "<SAFETY_TEST_GB>"&igub&"</SAFETY_TEST_GB>"                           ''��������
		    	strRst = strRst & "<SAFETY_TEST_CENTER>"&igov&"</SAFETY_TEST_CENTER>"                   ''�������
		    	strRst = strRst & "<SAFETY_TEST_NO><![CDATA["&inum&"]]></SAFETY_TEST_NO>"                           ''������ȣ
		    	strRst = strRst & "<SAFETY_MODEL_NAME>"&getItemNameFormat()&"</SAFETY_MODEL_NAME>"      ''�𵨸�
		    	strRst = strRst & "<SAFETY_TEST_DATE></SAFETY_TEST_DATE>"                       	''������
			Else
				strRst = ""
		    	strRst = strRst & "<SAFETY_TEST_GB></SAFETY_TEST_GB>"                           ''��������
		    	strRst = strRst & "<SAFETY_TEST_CENTER></SAFETY_TEST_CENTER>"                   ''�������
		    	strRst = strRst & "<SAFETY_TEST_NO></SAFETY_TEST_NO>"                           ''������ȣ
		    	strRst = strRst & "<SAFETY_MODEL_NAME></SAFETY_MODEL_NAME>"                     ''�𵨸�
		    	strRst = strRst & "<SAFETY_TEST_DATE></SAFETY_TEST_DATE>"                       ''������
			End If
		rsget.Close
		getLotteSafeToReg = strRst
	rw strRst
	end function

	'��ǰǰ���ڵ�
	Public Function getLotteimallItemInfoCdToReg(iORIGIN_CODE)
		Dim strRst, strSQL
		Dim mallinfoDiv,mallinfoCd,infoContent

		'���ϸ��� ��ó�� �̴� ����
		Dim YM, ConvertYM, SD
		strSQL = ""
		strSQL = strSQL & " SELECT top 1 F.infocontent, IC.safetyDiv " & vbcrlf
		strSQL = strSQL & " FROM db_item.dbo.tbl_OutMall_infoCodeMap M " & vbcrlf
		strSQL = strSQL & " INNER JOIN db_item.dbo.tbl_item_contents IC ON '1'+IC.infoDiv=M.mallinfoDiv  " & vbcrlf
		strSQL = strSQL & " LEFT JOIN db_item.dbo.tbl_item_infoCont F ON M.infocd=F.infocd AND F.itemid='"&Fitemid&"' " & vbcrlf
		strSQL = strSQL & " where IC.itemid='"&Fitemid&"' and M.mallinfocd = '10011' " & vbcrlf
		rsget.Open strSql,dbget,1
'rw strSQL
'response.end
		If Not(rsget.EOF or rsget.BOF) then
			YM = rsget("infocontent")
			SD = rsget("safetyDiv")
		Else
			YM = "X"
			SD = "X"
		End If
		rsget.Close

		If YM <> "X" Then
		    YM = replace(YM,".","")
		    YM = replace(YM,"/","")
		    YM = replace(YM,"-","")
		    YM = replace(YM," ","")
		    YM = TRIM(YM)

			If isNumeric(Ym) Then
				ConvertYM = Clng(YM)
			Else
				ConvertYM = "X"
			End If
		Else
			ConvertYM = YM
		End If

		strSQL = ""
		strSQL = strSQL & " SELECT TOP 100 M.* , " & vbcrlf
		strSQL = strSQL & " CASE " & vbcrlf

		If SD = "10" Then
			'��ó���� ���� ���� ���
			If ConvertYM = "X" Then
				strSQL = strSQL & " 	 WHEN (M.infoCd='00000') AND (IC.safetyyn= 'Y') AND (LEFT(IC.safetyDiv,3)='KCC') THEN IC.safetyNum " & vbcrlf
			'��ó���� ���� �ִ� ���
			Else
				'��ó���� 2012�� 7�� ������ ���
				If ConvertYM < 201207 Then
					strSQL = strSQL & " 	 WHEN (M.infoCd='00000') AND (IC.safetyyn= 'Y') AND (LEFT(IC.safetyDiv,3)='KCC') THEN '�ش����' " & vbcrlf	 '(�����ڵ尡 KCC�����̰�), (10x10���� ���������ڵ忩�ΰ� Y, ������ KC(10), 201207��)�� ��
				'��ó���� 2012�� 7�� ������ ���
				ElseIf ConvertYM >= 201207 Then
					strSQL = strSQL & " 	 WHEN (M.infoCd='00000') AND (IC.safetyyn= 'Y') AND (LEFT(IC.safetyDiv,3)='KCC') THEN IC.safetyNum " & vbcrlf '(�����ڵ尡 KCC�����̰�), (10x10���� ���������ڵ忩�ΰ� Y, ������ KC(10), 201207��)�� ��
				End If
			End If
		End If

		strSQL = strSQL & " 	 WHEN (M.infoCd='00000') AND (isNULL(IC.safetyyn,'') = 'Y') AND (M.mallinfoCd= '10205') THEN IC.safetyNum " & vbcrlf		'10206�� KC����
		strSQL = strSQL & " 	 WHEN (M.infoCd='00000') AND (isNULL(IC.safetyyn,'') <> 'Y') AND (M.mallinfoCd= '10205') THEN '�ش����'  " & vbcrlf		'10206�� KC����
		strSQL = strSQL & " 	 WHEN (M.infoCd='00000') AND (isNULL(IC.safetyyn,'') = 'Y') AND (M.mallinfoCd= '10206') THEN 'KC �������� ��'  " & vbcrlf		'10206�� KC����
		strSQL = strSQL & " 	 WHEN (M.infoCd='00000') AND (isNULL(IC.safetyyn,'') <> 'Y') AND (M.mallinfoCd= '10206') THEN '�ش����'  " & vbcrlf		'10206�� KC����
		strSQL = strSQL & " 	 WHEN (M.infoCd='00000') AND (IC.safetyyn= 'N') THEN '�ش����'  " & vbcrlf		'(�����ڵ尡 KCC�����̰�), (10x10���� ���������ڵ忩�ΰ� N)�� ��
		strSQL = strSQL & " 	 WHEN M.infoCd='00001' THEN '�ش����' " & vbcrlf
		strSQL = strSQL & " 	 WHEN M.infoCd='00002' THEN '�������� ����' " & vbcrlf
		strSQL = strSQL & " 	 WHEN F.chkDiv='Y' AND (M.infoCd='19008') THEN '������' " & vbcrlf				'�ͱݼ��� ������
		strSQL = strSQL & " 	 WHEN F.chkDiv='N' AND (M.infoCd='19008') THEN '�������� ����' " & vbcrlf
		strSQL = strSQL & " 	 WHEN F.chkDiv='Y' AND (M.infoCd='18008') THEN '��ɼ� �ɻ� ��' " & vbcrlf		'ȭ��ǰ�� ��ɼ� ȭ��ǰ ����
		strSQL = strSQL & " 	 WHEN F.chkDiv='N' AND (M.infoCd='18008') THEN '�ش����' " & vbcrlf
		strSQL = strSQL & " 	 WHEN F.chkDiv='Y' AND (M.infoCd='17008') THEN '��ǰ�������� ���� ���ԽŰ�����' " & vbcrlf		'��ǰ�������� ���� ���ԽŰ� ����	20130215���� �߰�
		strSQL = strSQL & " 	 WHEN F.chkDiv='N' AND (M.infoCd='17008') THEN '�ش����' " & vbcrlf
		strSQL = strSQL & " 	 WHEN c.infotype='M' THEN replace(F.infocontent,'.','') " & vbcrlf
		strSQL = strSQL & " 	 WHEN c.infotype='C' AND F.chkDiv='N' THEN '�ش����' " & vbcrlf
		strSQL = strSQL & " 	 WHEN c.infotype='P' THEN replace(F.infocontent,'1644-6030','1644-6035') " & vbcrlf
		strSQL = strSQL & " 	 WHEN c.infoCd='02004' and F.infocontent='' then '�ش����' " & vbcrlf
		strSQL = strSQL & " 	 ELSE F.infocontent " & vbcrlf
		strSQL = strSQL & " END AS infoContent " & vbcrlf
		strSQL = strSQL & " FROM db_item.dbo.tbl_OutMall_infoCodeMap M " & vbcrlf
		strSQL = strSQL & " INNER JOIN db_item.dbo.tbl_item_contents IC ON '1'+IC.infoDiv=M.mallinfoDiv " & vbcrlf
		strSQL = strSQL & " LEFT JOIN db_item.dbo.tbl_item_infoCode c ON M.infocd=c.infocd " & vbcrlf
		strSQL = strSQL & " LEFT JOIN db_item.dbo.tbl_item_infoCont F ON M.infocd=F.infocd AND F.itemid='"&Fitemid&"' " & vbcrlf
		strSQL = strSQL & " WHERE M.mallid = 'lotteimall' AND IC.itemid='"&Fitemid&"' " & vbcrlf
		strSQL = strSQL & " ORDER BY M.infocd ASC"
		rsget.Open strSQL,dbget,1
'rw strSQL
'response.end
		If Not(rsget.EOF or rsget.BOF) then
			Do until rsget.EOF
			    mallinfoDiv = rsget("mallinfoDiv")
			    mallinfoCd  = rsget("mallinfoCd")
			    infoContent = rsget("infoContent")

			    if IsNULL(infoContent) then
			        infoContent=""
			    end if
				'' ��Ÿ 35 ������ Ȯ�� (���Կ���..)
			    If (mallinfoDiv="135") and (mallinfoCd="10041") and (iORIGIN_CODE="0082") then
			        infoContent="�ش����"
			    End If

'                If (mallinfoCd="10206") and (infoContent<>"�ش����") then
'                   infoContent="KC �������� ��"
'                end if

                If (mallinfoCd="10205") and (infoContent="") then
                    infoContent="�ش����"
                end if

                If (mallinfoCd="10011") and (ConvertYM="X") then
                    infoContent="�ش����"
                end if

                if (mallinfoCd="10073") or (mallinfoCd="10011") then
                    infoContent = replace(infoContent,"��","")
                    infoContent = replace(infoContent,"��","")
                    infoContent = replace(infoContent,".","")
                end if
          ''rw mallinfoCd&"|"&infoContent

				strRst = strRst & "<LISART_INFO>" & vbcrlf
				strRst = strRst & "<LISART_CODE>"&mallinfoDiv&"</LISART_CODE>" & vbcrlf
		    	strRst = strRst & "<LISART_CSTN_CODE>"&mallinfoCd&"</LISART_CSTN_CODE>" & vbcrlf
		    	strRst = strRst & "<LISART_CSTN_DG1_CNTT><![CDATA["&infoContent&"]]></LISART_CSTN_DG1_CNTT>" & vbcrlf
		    	If (mallinfoCd="10011") Then  ''���ϸ��� ��ó�� / ����ǰ�ΰ�� N
		    	    If (ConvertYM<>"X") Then
        		    	strRst = strRst & "<LISART_CSTN_DG2_CNTT>Y</LISART_CSTN_DG2_CNTT>" & vbcrlf
        		    Else
        		        strRst = strRst & "<LISART_CSTN_DG2_CNTT>N</LISART_CSTN_DG2_CNTT>" & vbcrlf
        		    End If
		    	ElseIf (mallinfoCd="10201") Then ''��ǰ�������� ���� �ⱸ����ǰ ���� / �ش���� �ΰ�� N
		    	    If (infoContent<>"�ش����") Then
        		    	strRst = strRst & "<LISART_CSTN_DG2_CNTT>Y</LISART_CSTN_DG2_CNTT>" & vbcrlf
        		    Else
        		        strRst = strRst & "<LISART_CSTN_DG2_CNTT>N</LISART_CSTN_DG2_CNTT>" & vbcrlf
        		    End If
		    	ElseIf (mallinfoCd="10040") Then ''���Խ�ǰ���� / �ش���� �ΰ�� N
		    	    If (infoContent<>"�ش����") Then
        		    	strRst = strRst & "<LISART_CSTN_DG2_CNTT>Y</LISART_CSTN_DG2_CNTT>" & vbcrlf
        		    Else
        		        strRst = strRst & "<LISART_CSTN_DG2_CNTT>N</LISART_CSTN_DG2_CNTT>" & vbcrlf
        		    End If
		    	ElseIf (mallinfoCd="10041") Then ''���Կ��� / �ش���� �ΰ�� N
		    	    If (infoContent<>"�ش����") Then
        		    	strRst = strRst & "<LISART_CSTN_DG2_CNTT>Y</LISART_CSTN_DG2_CNTT>" & vbcrlf
        		    Else
        		        strRst = strRst & "<LISART_CSTN_DG2_CNTT>N</LISART_CSTN_DG2_CNTT>" & vbcrlf
        		    End If
		    	ElseIf (mallinfoCd="10007") Then ''��ɼ� ȭ��ǰ ���� / �ش���� �ΰ�� N	2013-01-02 ������ �߰�//������ '''' WHEN F.chkDiv='Y' AND (M.infoCd='18008') THEN '��ɼ� �ɻ� �� ��κ� �߰�
		    	    If (infoContent<>"�ش����") Then
        		    	strRst = strRst & "<LISART_CSTN_DG2_CNTT>Y</LISART_CSTN_DG2_CNTT>" & vbcrlf
        		    Else
        		        strRst = strRst & "<LISART_CSTN_DG2_CNTT>N</LISART_CSTN_DG2_CNTT>" & vbcrlf
        		    End If
		    	ElseIf (mallinfoCd="10002") and (mallinfoDiv = "119") Then			'�ͱݼ�/����/�ð���̰� ������ / �������� ������ ��� Y
		    	    If (infoContent<>"�������� ����") Then
        		    	strRst = strRst & "<LISART_CSTN_DG2_CNTT>Y</LISART_CSTN_DG2_CNTT>" & vbcrlf
        		    Else
        		        strRst = strRst & "<LISART_CSTN_DG2_CNTT>N</LISART_CSTN_DG2_CNTT>" & vbcrlf
        		    End If
		    	ElseIf (mallinfoCd="10019") and (mallinfoDiv = "119") Then			'�ͱݼ�/����/�ð���̰� ������ / �������� ������ ��� Y
		    	    If (infoContent<>"������") Then
        		    	strRst = strRst & "<LISART_CSTN_DG2_CNTT>N</LISART_CSTN_DG2_CNTT>" & vbcrlf
        		    Else
        		        strRst = strRst & "<LISART_CSTN_DG2_CNTT>Y</LISART_CSTN_DG2_CNTT>" & vbcrlf
        		    End If
				ElseIf (mallinfoCd="10048") Then ''���缺�� ǥ�� ��� ���� / �ش���� �ΰ�� N
					If (infoContent<>"�ش����") Then
						strRst = strRst & "<LISART_CSTN_DG2_CNTT>Y</LISART_CSTN_DG2_CNTT>" & vbcrlf
        		    Else
        		        strRst = strRst & "<LISART_CSTN_DG2_CNTT>N</LISART_CSTN_DG2_CNTT>" & vbcrlf
        		    End If
				ElseIf (mallinfoCd="10054") Then ''������ ������ ��ǰ ���� / �ش���� �ΰ�� N
					If (infoContent<>"�ش����") Then
						strRst = strRst & "<LISART_CSTN_DG2_CNTT>Y</LISART_CSTN_DG2_CNTT>" & vbcrlf
        		    Else
        		        strRst = strRst & "<LISART_CSTN_DG2_CNTT>N</LISART_CSTN_DG2_CNTT>" & vbcrlf
        		    End If
				ElseIf (mallinfoCd="10110") Then ''Ư���뵵��ǰ(�����ƽ�, ü��������ǰ) ���� / �ش���� �ΰ�� N
					If (infoContent<>"�ش����") Then
						strRst = strRst & "<LISART_CSTN_DG2_CNTT>Y</LISART_CSTN_DG2_CNTT>" & vbcrlf
        		    Else
        		        strRst = strRst & "<LISART_CSTN_DG2_CNTT>N</LISART_CSTN_DG2_CNTT>" & vbcrlf
        		    End If
        		ElseIf (mallinfoCd="10006") Then '' ���α�
        		    If (infoContent<>"�ش����") Then
						strRst = strRst & "<LISART_CSTN_DG2_CNTT>Y</LISART_CSTN_DG2_CNTT>" & vbcrlf
        		    Else
        		        strRst = strRst & "<LISART_CSTN_DG2_CNTT>N</LISART_CSTN_DG2_CNTT>" & vbcrlf
        		    End If
        		ElseIf (mallinfoCd="10206") Then 'KC��������'
        		    If (infoContent<>"�ش����") Then
        		        strRst = strRst & "<LISART_CSTN_DG2_CNTT>Y</LISART_CSTN_DG2_CNTT>" & vbcrlf
        		    else
        		        strRst = strRst & "<LISART_CSTN_DG2_CNTT>N</LISART_CSTN_DG2_CNTT>" & vbcrlf
        		    end if
				Else
					strRst = strRst & "<LISART_CSTN_DG2_CNTT>N</LISART_CSTN_DG2_CNTT>" & vbcrlf
    		    ENd IF

		    	strRst = strRst & "<LISART_CSTN_DG3_CNTT>N</LISART_CSTN_DG3_CNTT>" & vbcrlf
		    	strRst = strRst & "<LISART_CSTN_DG4_CNTT>N</LISART_CSTN_DG4_CNTT>" & vbcrlf
		    	strRst = strRst & "<LISART_CSTN_DG5_CNTT>N</LISART_CSTN_DG5_CNTT>" & vbcrlf
		    	strRst = strRst & "</LISART_INFO>" & vbcrlf
				rsget.MoveNext
			Loop
		End If
		rsget.Close
''rw strRst
		getLotteimallItemInfoCdToReg = strRst
	End Function

	'// ��ǰ���: ��ǰ���� �Ķ���� ����(��ǰ��Ͽ�)
	public function getLotteItemContParamToReg()
		dim strRst, strSQL

		strRst = ("<div align=""center"">")

		'#�⺻ ��ǰ����
		Select Case FUsingHTML
			Case "Y"
				strRst = strRst & (Fitemcontent & "<br>")
			Case "H"
				strRst = strRst & (nl2br(Fitemcontent) & "<br>")
			Case Else
				strRst = strRst & (nl2br(ReplaceBracket(Fitemcontent)) & "<br>")
		End Select

		'# �߰� ��ǰ �����̹��� ����
		strSQL = "exec [db_item].[dbo].sp_Ten_CategoryPrd_AddImage @vItemid =" & Fitemid
		rsget.CursorLocation = adUseClient
		rsget.CursorType=adOpenStatic
		rsget.Locktype=adLockReadOnly
		rsget.Open strSQL, dbget

		if Not(rsget.EOF or rsget.BOF) then
			Do Until rsget.EOF
				if rsget("imgType")="1" then
					strRst = strRst & ("<img src=""http://webimage.10x10.co.kr/item/contentsimage/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_400") & """ border=""0""><br>")
				end if
				rsget.MoveNext
			Loop
		end if

		rsget.Close

		'#�⺻ ��ǰ �����̹���
		if ImageExists(FmainImage) then strRst = strRst & ("<img src=""" & FmainImage & """ border=""0""><br>")
		if ImageExists(FmainImage2) then strRst = strRst & ("<img src=""" & FmainImage2 & """ border=""0""><br>")

		'#��� ���ǻ���
		strRst = strRst & ("<br><img src=""http://fiximage.10x10.co.kr/web2008/etc/cs_info_LTimall.jpg"">")

		strRst = strRst & ("</div>")

		getLotteItemContParamToReg = strRst

		strSQL = ""
		strSQL = strSQL & " SELECT itemid, mallid, linkgbn, textVal " & VBCRLF
		strSQL = strSQL & " FROM db_item.dbo.tbl_OutMall_etcLink " & VBCRLF
		strSQL = strSQL & " where mallid='"&CMALLNAME&"' and linkgbn = 'contents' and itemid = '"&Fitemid&"' " & VBCRLF
		rsget.Open strSQL, dbget
		if Not(rsget.EOF or rsget.BOF) then
			strRst = rsget("textVal")
			strRst = "<div align=""center"">" & strRst & "<br><img src=""http://fiximage.10x10.co.kr/web2008/etc/cs_info.jpg""></div>"
			getLotteItemContParamToReg = strRst
		End If
		rsget.Close
	end function

	'// ��ǰ���: ��ǰ�߰��̹��� �Ķ���� ����(��ǰ��Ͽ�)
	public function getLotteAddImageParamToReg()
		dim strRst, strSQL, i

        strRst = ""

        IF application("Svr_Info")="Dev" THEN
            FbasicImage = "http://61.252.133.2/images/B000151064.jpg"
        ENd IF

        strRst = strRst & "<IMG_L>"&FbasicImage&"</IMG_L>"

		'# �߰� ��ǰ �����̹��� ����
		strSQL = "exec [db_item].[dbo].sp_Ten_CategoryPrd_AddImage @vItemid =" & Fitemid
		rsget.CursorLocation = adUseClient
		rsget.CursorType=adOpenStatic
		rsget.Locktype=adLockReadOnly
		rsget.Open strSQL, dbget

		if Not(rsget.EOF or rsget.BOF) then
			for i=1 to rsget.RecordCount
				if rsget("imgType")="0" then
					strRst = strRst & "<IMG_L"& i & ">" & ("http://webimage.10x10.co.kr/image/add" & rsget("gubun") & "/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_400"))&"</IMG_L"&i&">"
				end if
				rsget.MoveNext
				if i>=5 then Exit For
			next
		end if

		rsget.Close

		getLotteAddImageParamToReg = strRst
	end function

	'// �ٹ����� ��ǰ�ɼ� �˻�
	public function checkTenItemOptionValid()
		dim strSql, chkRst, chkMultiOpt
		dim cntType, cntOpt
		chkRst = true
		chkMultiOpt = false

		if FoptionCnt>0 then
			'// ���߿ɼ�Ȯ��
			strSql = "exec [db_item].[dbo].sp_Ten_ItemOptionMultipleTypeList " & FItemid
	        rsget.CursorLocation = adUseClient
			rsget.CursorType = adOpenStatic
			rsget.LockType = adLockOptimistic
	        rsget.Open strSql, dbget

			if Not(rsget.EOF or rsget.BOF) then
				chkMultiOpt = true
				cntType = rsget.RecordCount
			end if
			rsget.Close

			if chkMultiOpt then
				'// ���߿ɼ� �϶�
				strSql = "Select optionname "
				strSql = strSql & " From [db_item].[dbo].tbl_item_option "
				strSql = strSql & " where itemid=" & FItemid
				strSql = strSql & " 	and isUsing='Y' and optsellyn='Y' "
				strSql = strSql & " 	and optaddprice=0 "
				strSql = strSql & " 	and (optlimityn='N' or (optlimityn='Y' and optlimitno-optlimitsold>="&CMAXLIMITSELL&")) "
				rsget.Open strSql,dbget,1

				if Not(rsget.EOF or rsget.BOF) then
					Do until rsget.EOF
						cntOpt = ubound(split(db2Html(rsget("optionname")),","))+1
						if cntType<>cntOpt then
							chkRst = false
						end if
						rsget.MoveNext
					Loop
				else
					chkRst = false
				end if
				rsget.Close
			Else
				'// ���Ͽɼ��� ��
				strSql = "Select optionTypeName, optionname "
				strSql = strSql & " From [db_item].[dbo].tbl_item_option "
				strSql = strSql & " where itemid=" & FItemid
				strSql = strSql & " 	and isUsing='Y' and optsellyn='Y' "
				strSql = strSql & " 	and optaddprice=0 "
				strSql = strSql & " 	and (optlimityn='N' or (optlimityn='Y' and optlimitno-optlimitsold>="&CMAXLIMITSELL&")) "
				rsget.Open strSql,dbget,1

				if (rsget.EOF or rsget.BOF) then
					chkRst = false
				end if
				rsget.Close
			end if
		end if

		'//��� ��ȯ
		checkTenItemOptionValid = chkRst

	end Function

	'// �Ե����� �Ǹſ��� ��ȯ
	public function getLTiMallSellYn()
		'�ǸŻ��� (10:�Ǹ�����, 20:ǰ��)
		if FsellYn="Y" and FisUsing="Y" then
			if FLimitYn="N" or (FLimitYn="Y" and FLimitNo-FLimitSold>=CMAXLIMITSELL) then
				getLTiMallSellYn = "Y"
			else
				getLTiMallSellYn = "N"
			end if
		else
			getLTiMallSellYn = "N"
		end if
	end Function

	'// �Ե����� ��ϻ��� ��ȯ
	public function getLotteItemStatCd()
	    getLotteItemStatCd = getLtiMallStatName
'		Select Case FLTiMallStatCd
'			Case "1"
'				getLotteItemStatCd = "�ӽõ��"
'			Case "2"
'				getLotteItemStatCd = "���ο�û"
'			Case "3"
'				getLotteItemStatCd = "���οϷ�"
'			Case "4"
'				getLotteItemStatCd = "�ݷ�"
'			Case "50"
'				getLotteItemStatCd = "���κҰ�"
'			Case "51"
'				getLotteItemStatCd = "����ο�û"
'			Case "52"
'				getLotteItemStatCd = "������û"
'		End Select
	end Function

	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub
end class

class CLotteiMall
	public FItemList()

	public FResultCount
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount

	public FRectMdCode
	public FRectDspNo
	public FRectIsMapping

	public FRectSDiv
	public FRectKeyword
	public FRectGrpCode

	public FRectCDL
	public FRectCDM
	public FRectCDS

    public FRectMode

	public FRectItemID
	public FRectItemName
	public FRectMakerid
	public FRectLotteNotReg
	public FRectMatchCate
	public FRectMatchCateNotCheck
	public FRectSellYn
	public FRectLimitYn
	public FRectLTiMallGoodNo
	public FRectMinusMigin
	public FRectonlyValidMargin
	public FRectIsSoldOut
	public FRectExpensive10x10
	public FRectLotteYes10x10No
	public FRectLotteNo10x10Yes
	public FRectOnreginotmapping
	public FRectNotJehyu
	public FRectEventid
	public FRectdiffPrc
	public FRectdisptpcd
    public FRectCateUsingYn

    ''���ļ���
    public FRectOrdType
    public FRectoptAddprcExists
    public FRectoptAddPrcRegTypeNone
    public FRectoptAddprcExistsExcept
    public FRectoptExists
    public FRectregedOptNull

    public FRectFailCntExists
    public FRectFailCntOverExcept
    public FRectExtSellYn
    public FRectInfoDiv


	Private Sub Class_Initialize()
		redim  FItemList(0)
		FCurrPage =1
		FPageSize = 30
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()
	End Sub

	'--------------------------------------------------------------------------------

    '// ���MD���
    public Sub getLotte_MDList
		dim sqlStr,i
		sqlStr = " select count(MDCode) as cnt, CEILING(CAST(Count(MDCode) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr + "From db_temp.dbo.tbl_lotte_MDInfo "

		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'������������ ��ü ���������� Ŭ �� �Լ�����
		if Cint(FCurrPage)>Cint(FTotalPage) then
			FResultCount = 0
			exit sub
		end if

		sqlStr = " select  top " + CStr(FPageSize*FCurrPage) + " * "
		sqlStr = sqlStr + " from db_temp.dbo.tbl_lotte_MDInfo "
		sqlStr = sqlStr + " order by MDCode asc"

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CLotteiMallItem
				FItemList(i).FMDCode		= rsget("MDCode")
				FItemList(i).FMDName		= db2html(rsget("MDName"))
				FItemList(i).FSellFeeType	= rsget("SellFeeType")
				FItemList(i).FNormalSellFee	= rsget("NormalSellFee")
				FItemList(i).FEventSellFee	= rsget("EventSellFee")
				FItemList(i).FisUsing		= rsget("isUsing")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end Sub


    '// ���MD��ǰ�� ���
    public Sub getLotte_MDGrpList
		dim sqlStr, i

		sqlStr = " select count(groupCode) as cnt, CEILING(CAST(Count(groupCode) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr + " From db_temp.dbo.tbl_lotte_MDCateGrp "
		sqlStr = sqlStr + " Where MDCode='" & FRectMdCode & "'"

		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'������������ ��ü ���������� Ŭ �� �Լ�����
		if Cint(FCurrPage)>Cint(FTotalPage) then
			FResultCount = 0
			exit sub
		end if

		sqlStr = " select  top " + CStr(FPageSize*FCurrPage) + " * "
		sqlStr = sqlStr + " from db_temp.dbo.tbl_lotte_MDCateGrp "
		sqlStr = sqlStr + " Where MDCode='" & FRectMdCode & "'"
		sqlStr = sqlStr + " order by groupCode asc"

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CLotteiMallItem

				FItemList(i).FgroupCode			= rsget("groupCode")
				FItemList(i).FSuperGroupName	= db2html(rsget("SuperGroupName"))
				FItemList(i).FGroupName			= db2html(rsget("GroupName"))
				FItemList(i).FisUsing			= rsget("isUsing")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end Sub

	'--------------------------------------------------------------------------------
	'// �ٹ�����-�Ե�iMall ī�װ�
	public Sub getTenLTiMallCateList
		dim sqlStr, addSql, i

		if FRectCDL<>"" then
			addSql = addSql & " and s.code_large='" & FRectCDL & "'"
		end if
		if FRectCDM<>"" then
			addSql = addSql & " and s.code_mid='" & FRectCDM & "'"
		end if
		if FRectCDS<>"" then
			addSql = addSql & " and s.code_small='" & FRectCDS & "'"
		end if
		if FRectDspNo<>"" then
			addSql = addSql & " and T.CateKey='" & FRectDspNo & "'"
		end if

		if FRectIsMapping="Y" then
			addSql = addSql & " and T.CateKey is Not null "
		elseif FRectIsMapping="N" then
			addSql = addSql & " and T.CateKey is null "
		end if

		if FRectKeyword<>"" then
			Select Case FRectSDiv
				Case "LCD"	'�Ե����� �����ڵ� �˻�
					addSql = addSql & " and T.CateKey='" & FRectKeyword & "'"
				Case "CNM"	'ī�װ���(�ٹ����� �Һз���)
					addSql = addSql & " and s.code_nm like '%" & FRectKeyword & "%'"
			End Select
		end if

		sqlStr = " select count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr + " from db_item.dbo.tbl_cate_small as s "
		sqlStr = sqlStr + " 	left join ("
		sqlStr = sqlStr + " 		select cm.CateKey, cm.tenCateLarge,cm.tenCateMid, cm.tenCateSmall, lc.CateGbn, lc.L_Code"
		sqlStr = sqlStr + " 	    From db_item.dbo.tbl_LTiMall_cate_mapping as cm "  ''����ī�� ����
		sqlStr = sqlStr + " 	    Join db_temp.dbo.tbl_LTiMall_Category as lc "
		sqlStr = sqlStr + " 		on lc.CateKey=cm.CateKey "

		if FRectdisptpcd<>"" then
            sqlStr = sqlStr & " and lc.CateGbn='"&FRectdisptpcd&"'"
        end if
        sqlStr = sqlStr + " 		) T on T.tenCateLarge=s.code_large "
		sqlStr = sqlStr + " 		and T.tenCateMid=s.code_mid "
		sqlStr = sqlStr + " 		and T.tenCateSmall=s.code_small "
		sqlStr = sqlStr + " Where 1=1 " & addSql

		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'������������ ��ü ���������� Ŭ �� �Լ�����
		if Cint(FCurrPage)>Cint(FTotalPage) then
			FResultCount = 0
			exit sub
		end if

		sqlStr = " select  top " + CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr + " 	s.code_large,s.code_mid,s.code_small "
		sqlStr = sqlStr + " 	,(Select code_nm from db_item.dbo.tbl_cate_large Where code_large=s.code_large) as large_nm "
		sqlStr = sqlStr + " 	,(Select code_nm from db_item.dbo.tbl_cate_mid Where code_large=s.code_large and code_mid=s.code_mid) as mid_nm "
		sqlStr = sqlStr + " 	,code_nm as small_nm "
		sqlStr = sqlStr + " 	,T.CateKey as DispNo, T.L_Code"
		sqlStr = sqlStr + " 	,T.D_Name as DispNm, T.L_Name as DispLrgNm, T.M_Name as DispMidNm, T.S_Name as DispSmlNm, T.D_Name as DispThnNm, T.cateGbn as disptpcd"
		sqlStr = sqlStr + " 	,T.IsUsing as CateIsUsing"
		sqlStr = sqlStr + " from db_item.dbo.tbl_cate_small as s "
		sqlStr = sqlStr + " 	left join ("
		sqlStr = sqlStr + " 		select cm.CateKey, cm.tenCateLarge,cm.tenCateMid, cm.tenCateSmall, lc.CateGbn, lc.L_Code"
		sqlStr = sqlStr + " 	    ,lc.D_Name,lc.L_Name,lc.M_Name,lc.S_Name, lc.isusing"
		sqlStr = sqlStr + " 	    From db_item.dbo.tbl_LTiMall_cate_mapping as cm "  ''����ī�� ����
		sqlStr = sqlStr + " 	    Join db_temp.dbo.tbl_LTiMall_Category as lc "
		sqlStr = sqlStr + " 		on lc.CateKey=cm.CateKey "

		if FRectdisptpcd<>"" then
            sqlStr = sqlStr & " and lc.CateGbn='"&FRectdisptpcd&"'"
        end if
        sqlStr = sqlStr + " 		) T on T.tenCateLarge=s.code_large "
		sqlStr = sqlStr + " 		and T.tenCateMid=s.code_mid "
		sqlStr = sqlStr + " 		and T.tenCateSmall=s.code_small "
		sqlStr = sqlStr + " Where 1=1 " & addSql
		sqlStr = sqlStr + " order by s.code_large,s.code_mid,s.code_small, T.CateGbn asc "

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CLotteiMallItem

				FItemList(i).FtenCDLName	= db2html(rsget("large_nm"))
				FItemList(i).FtenCDMName	= db2html(rsget("mid_nm"))
				FItemList(i).FtenCDSName	= db2html(rsget("small_nm"))

'                FItemList(i).FitemGbnKey    = db2html(rsget("itemGbnKey"))
'                FItemList(i).FitemGbnNm     = db2html(rsget("itemGbnNm"))
				FItemList(i).FDispNo		= rsget("DispNo")
				FItemList(i).FDispNm		= db2html(rsget("DispNm"))

				FItemList(i).FtenCateLarge	= rsget("code_large")
				FItemList(i).FtenCateMid	= rsget("code_mid")
				FItemList(i).FtenCateSmall	= rsget("code_small")

				FItemList(i).FgroupCode		= rsget("L_Code")
				FItemList(i).FDispLrgNm		= db2html(rsget("DispLrgNm"))
				FItemList(i).FDispMidNm		= db2html(rsget("DispMidNm"))
				FItemList(i).FDispSmlNm		= db2html(rsget("DispSmlNm"))
				FItemList(i).FDispThnNm		= db2html(rsget("DispThnNm"))

'                FItemList(i).FGbnLrgNm      = db2html(rsget("GbnLrgNm"))
'                FItemList(i).FGbnMidNm      = db2html(rsget("GbnMidNm"))
'                FItemList(i).FGbnSmlNm      = db2html(rsget("GbnSmlNm"))
'                FItemList(i).FGbnThnNm		= db2html(rsget("itemGbnNm"))

                FItemList(i).Fdisptpcd      = rsget("disptpcd")
                FItemList(i).FCateIsUsing   = rsget("CateIsUsing")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end Sub

	'// �Ե�iMall ī�װ�
	public Sub getLTiMallCategoryList
		dim sqlStr, addSql, i

        ''��ǰ�з��� �˻� ���� - ���� ���� ó��.
        addSql = addSql & " and c.cateGbn<>'M'"
        addSql = addSql + " and ((c.L_Code<>'50000000') or ((c.L_Code='50000000') and (c.M_Code='201300115948')))"  ''��Ÿ ������ �˻�����. :: �ٹ����� ������ CateGbn:B L:10500000 M:201200078827

		if FRectDspNo<>"" then
			addSql = addSql & " and c.cateKey=" & FRectDspNo
		end if

        if FRectdisptpcd<>"" then
            addSql = addSql & " and c.CateGbn='"&FRectdisptpcd&"'"
        end if

        if FRectCateUsingYn<>"" then
            addSql = addSql & " and c.isusing='"&FRectCateUsingYn&"'"
        end if

		if FRectKeyword<>"" then
			Select Case FRectSDiv
				Case "LCD"	'�Ե����� �����ڵ� �˻�
					addSql = addSql & " and c.cateKey='" & FRectKeyword & "'"
'				Case "TCD"	'�ٹ����� ī�װ��ڵ� �˻�(���߼� �����ڵ� 9�ڸ�)
'					addSql = addSql & " and m.tenCateLarge&m.tenCateMid&m.tenCateSmall='" & FRectKeyword & "'"
				Case "CNM"	'ī�װ���(�Ե� ���з���)
					addSql = addSql & " and (c.D_Name like '%" & FRectKeyword & "%'"
					addSql = addSql & " or c.S_Name like '%" & FRectKeyword & "%'"
					addSql = addSql & " or c.M_Name like '%" & FRectKeyword & "%'"
					addSql = addSql & " or c.L_Name like '%" & FRectKeyword & "%'"
					addSql = addSql & " )"
			End Select
		end if

		sqlStr = " select count(c.cateKey) as cnt, CEILING(CAST(Count(c.cateKey) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr + " from db_temp.dbo.tbl_LTiMall_Category as c "
		sqlStr = sqlStr + " Where 1=1 " & addSql

		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'������������ ��ü ���������� Ŭ �� �Լ�����
		if Cint(FCurrPage)>Cint(FTotalPage) then
			FResultCount = 0
			exit sub
		end if

		sqlStr = " select distinct top " + CStr(FPageSize*FCurrPage) + " c.* "
		sqlStr = sqlStr + " from db_temp.dbo.tbl_LTiMall_Category as c "
		sqlStr = sqlStr + " Where 1=1 " & addSql
		sqlStr = sqlStr + " order by c.cateKey" ''"c.L_Code, c.M_Code, c.S_Code, c.D_Code, c.cateKey"

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CLotteiMallItem

				FItemList(i).FDispNo		= rsget("cateKey")
				FItemList(i).FDispNm		= db2html(rsget("D_Name"))
				FItemList(i).FDispLrgNm		= db2html(rsget("L_Name"))
				FItemList(i).FDispMidNm		= db2html(rsget("M_Name"))
				FItemList(i).FDispSmlNm		= db2html(rsget("S_Name"))
				FItemList(i).FDispThnNm		= db2html(rsget("D_Name"))
                FItemList(i).Fdisptpcd      = rsget("cateGbn")
                FItemList(i).FgroupCode     = rsget("L_Code")
'				FItemList(i).FtenCateLarge	= rsget("tenCateLarge")
'				FItemList(i).FtenCateMid	= rsget("tenCateMid")
'				FItemList(i).FtenCateSmall	= rsget("tenCateSmall")
'				FItemList(i).FtenCateName	= db2html(rsget("code_nm"))

				FItemList(i).FisUsing		= rsget("isUsing")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end Sub


	'--------------------------------------------------------------------------------

	'// �Ե����� �귣��
	public Sub getLotteBrandList
		dim sqlStr, addSql, i

		if FRectMakerid<>"" then
			addSql = addSql & " and b.TenMakerid='" & FRectMakerid & "'"
		end if

		if FRectKeyword<>"" then
			Select Case FRectSDiv
				Case "LCD"	'�Ե����� �귣���ڵ� �˻�
					addSql = addSql & " and b.lotteBrandCD='" & FRectKeyword & "'"
				Case "TCD"	'�ٹ����� �귣��ID �˻�
					addSql = addSql & " and b.TenMakerid='" & FRectKeyword & "'"
				Case "BNM"	'�귣���(�ٹ����ٸ�)
					addSql = addSql & " and c.socname_kor like '%" & FRectKeyword & "%'"
			End Select
		end if

		sqlStr = " select count(b.TenMakerid) as cnt, CEILING(CAST(Count(b.TenMakerid) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr + " from db_user.dbo.tbl_user_c as c "
		sqlStr = sqlStr + " 	Join db_item.dbo.tbl_lotte_brand_mapping as b "
		sqlStr = sqlStr + " 		on c.userid=b.TenMakerid "
		sqlStr = sqlStr + " Where 1=1 " & addSql

		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'������������ ��ü ���������� Ŭ �� �Լ�����
		if Cint(FCurrPage)>Cint(FTotalPage) then
			FResultCount = 0
			exit sub
		end if

		sqlStr = " select  top " + CStr(FPageSize*FCurrPage) + " b.*, c.socname_kor "
		sqlStr = sqlStr + " from db_user.dbo.tbl_user_c as c "
		sqlStr = sqlStr + " 	Join db_item.dbo.tbl_lotte_brand_mapping as b "
		sqlStr = sqlStr + " 		on c.userid=b.TenMakerid "
		sqlStr = sqlStr + " Where 1=1 " & addSql
		sqlStr = sqlStr + " order by b.regdate desc "

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CLotteiMallItem

				FItemList(i).FlotteBrandCd		= rsget("lotteBrandCd")
				FItemList(i).FlotteBrandName	= db2html(rsget("lotteBrandName"))
				FItemList(i).FTenMakerid		= rsget("TenMakerid")
				FItemList(i).FTenBrandName		= db2html(rsget("socname_kor"))
				FItemList(i).FisUsing			= rsget("isUsing")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end Sub


	'--------------------------------------------------------------------------------

	'// �Ե�iMall ��ǰ ��� // ������ ������ �޶�� ��..
	public Sub getLTiMallRegedItemList
		dim sqlStr, addSql, i

		if FRectMakerid<>"" then
			addSql = addSql & " and i.makerid='" & FRectMakerid & "'"
		end if

		if FRectItemID<>"" then
			addSql = addSql & " and i.itemid in (" & FRectItemID & ")"
		end if

		if FRectItemName<>"" then
			addSql = addSql & " and i.itemname like '%" & FRectItemName & "%'"
		end if

		if FRectLTiMallGoodNo<>"" then
			addSql = addSql & " and m.LTiMallGoodNo='" & FRectLTiMallGoodNo & "'"
		end if

		if FRectCDL<>"" then
			addSql = addSql & " and i.cate_large='" & FRectCDL & "'"
		end if
		if FRectCDM<>"" then
			addSql = addSql & " and i.cate_mid='" & FRectCDM & "'"
		end if
		if FRectCDS<>"" then
			addSql = addSql & " and i.cate_small='" & FRectCDS & "'"
		end if

        if (FRectExtSellYn<>"") then
		    addSql = addSql + " and m.LTiMallSellYn='" & FRectExtSellYn & "'"
		end if

		Select Case FRectLotteNotReg
			Case "M"	'�̵��
			    addSql = addSql & " and m.itemid is NULL "
				''addSql = addSql & " and isNULL(m.LTiMallStatCd,0)=0 " ''((m.itemid is NULL))  " ''and m.LTiMallTmpGoodNo is Null  //or (m.LTiMallStatCd=0))
			Case "Q"	''��Ͻ���
				addSql = addSql & " and m.LTiMallStatCd=-1"
			Case "J"	'��Ͽ����̻�
				addSql = addSql & " and m.LTiMallStatCd>=0"
			Case "W"	'��Ͽ���
				addSql = addSql & " and m.LTiMallStatCd=0"
		    Case "V"	'��Ͽ����׵�ϰ���
				addSql = addSql & " and m.LTiMallStatCd=0"
		    Case "A"	'���۽õ�
				addSql = addSql & " and m.LTiMallStatCd=1"
			Case "F"	'��ϿϷ�(�ӽ�)
			    addSql = addSql & " and m.LTiMallStatCd=3"
				'addSql = addSql & " and m.LTiMallTmpGoodNo is Not Null"
				'addSql = addSql & " and m.LTiMallGoodNo is Null"
			Case "D"	'��ϿϷ�(����)
			    addSql = addSql & " and m.LTiMallStatCd=7"
				addSql = addSql & " and m.LTiMallGoodNo is Not Null"
				'addSql = addSql & " and m.LTiMallTmpGoodNo is Not Null"
			Case "R"	'�������
			    addSql = addSql & " and m.LTiMallStatCd=7"
		        addSql = addSql & " and m.LTiMallGoodNo is Not NULL"
		        addSql = addSql & " and m.LTiMallLastUpdate < i.lastupdate"
		End Select

		Select Case FRectMatchCate
			Case "Y"	'��Ī�Ϸ�
				addSql = addSql & " and c.mapCnt is Not Null"
			Case "N"	'�̸�Ī
				addSql = addSql & " and c.mapCnt is Null"
		End Select

		Select Case FRectSellYn
			Case "Y"	'�Ǹ�
				addSql = addSql & " and i.sellYn='Y'"
			Case "N"	'ǰ��
				addSql = addSql & " and i.sellYn in ('S','N')"
		End Select

		if FRectLimitYn<>"" then
			addSql = addSql & " and i.limitYn='" & FRectLimitYn & "'"
		end if

		if (FRectMinusMigin<>"") then
		   addSql = addSql & " and i.sellcash<>0"
		   addSql = addSql & " and ((i.sellcash-i.buycash)/i.sellcash)*100<"&CMAXMARGIN & VbCrlf
		   addSql = addSql & " and m.LTiMallSellYn= 'Y' " '''  ���� �߰�.
		else
		   IF (FRectonlyValidMargin<>"") then
		        addSql = addSql & " and i.sellcash<>0"
		        addSql = addSql & " and ((i.sellcash-i.buycash)/i.sellcash)*100>="&CMAXMARGIN & VbCrlf
		   END IF
		   ''addSql = addSql & " and m.LTiMallSellYn<> 'X' " '''  ���� �߰�.
		end if

		if FRectExpensive10x10 <> "" then
		   addSql = addSql & " and m.LTiMallPrice is Not Null and i.sellcash > m.LTiMallPrice "
		end if

        if FRectdiffPrc <> "" then
		   addSql = addSql & " and m.LTiMallPrice is Not Null and i.sellcash <> m.LTiMallPrice "
		end if

		if FRectLotteYes10x10No <> "" then
		   ''addSql = addSql & " and m.LTiMallPrice is Not Null and i.sellcash > m.LTiMallPrice "
		   addSql = addSql & " and m.LTiMallPrice is Not Null and (m.LTiMallSellYn= 'Y' and i.sellyn <> 'Y')"
		end if

		if FRectLotteNo10x10Yes <> "" then
		   addSql = addSql & " and m.LTiMallPrice is Not Null and (m.LTiMallSellYn= 'N' and i.sellyn='Y' and (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>="&CMAXLIMITSELL&")))"
		end if


		if FRectOnreginotmapping <> "" then
		    addSql = addSql & " and m.LTiMallTmpGoodNo is Not Null and IsNULL(c.mapCnt,0)>0" '''c.mapCnt is Null
		end if

		if FRectEventid<>"" then
			addSql = addSql & " and i.itemid in (Select itemid From [db_event].[dbo].tbl_eventitem Where evt_code='" & FRectEventid & "')" + VbCrlf
		end if

		if (FRectLotteNotReg<>"M" and FRectLotteNotReg<>"Q" and FRectLotteNotReg<>"V") then ''
		else
            if FRectLotteYes10x10No = "" then
        		'//���޸� �ǸŸ� ���
        		addSql = addSql & " and i.isExtUsing='Y'"
        		'//���ҹ�� ��ǰ ����
        		addSql = addSql & " and i.deliverytype<>'7'"
        		'//���ǹ�� 10000�� �̻�
        		IF (CUPJODLVVALID) then
                    addSql = addSql + " and ((i.deliveryType<>'9') or ((i.deliveryType='9') and (i.sellcash>=10000)))"
                ELSE
                     addSql = addSql + " and (i.deliveryType<>'9')"
                ENd IF
            end if
        end if

        ''�ɼ��߰��ݾ� �����ǰ.
		if (FRectoptAddprcExists<>"") and (FRectLotteNotReg<>"M") then
		    addSql = addSql & " and m.optAddPrcCnt>0"
'		    addSql = addSql & " and i.itemid in ("
'		    addSql = addSql & "     select distinct ii.itemid "
'		    addSql = addSql & "     from db_item.dbo.tbl_item ii "
'		    addSql = addSql & "     Join db_item.dbo.tbl_item_option o "
'		    addSql = addSql & "     on ii.itemid=o.itemid and o.optaddprice>0 and o.isusing='Y'"
'		    addSql = addSql & " )"
		end if

		if (FRectoptAddPrcRegTypeNone<>"") then          ''�ɼ��߰��ݾ׻�ǰ �̼��� ��ǰ.
		    addSql = addSql & " and m.optAddPrcCnt>0"
		    addSql = addSql & " and m.optAddPrcRegType=0"
		end if

		''�ɼ��߰��ݾ� �����ǰ ����
		if (FRectoptAddprcExistsExcept<>"") then
		    addSql = addSql & " and i.itemid Not in ("
		    addSql = addSql & "     select distinct ii.itemid "
		    addSql = addSql & "     from db_item.dbo.tbl_item ii "
		    addSql = addSql & "     Join db_item.dbo.tbl_item_option o "
		    addSql = addSql & "     on ii.itemid=o.itemid and o.optaddprice>0 and o.isusing='Y'"
		    addSql = addSql & " )"
		end if

		if (FRectoptExists<>"") then
            addSql = addSql & " and i.optioncnt>0"
        end if

        if (FRectregedOptNull<>"") then
            addSql = addSql & " and isNULL(m.regedOptCnt,0)=0"
        end if

        if (FRectFailCntExists<>"") then
            addSql = addSql & " and m.accFailCNT>0"
        end if

        if (FRectFailCntOverExcept<>"") then
            addSql = addSql & " and m.accFailCNT<"&FRectFailCntOverExcept
        end if

        if (FRectInfoDiv<>"") then
		    if (FRectInfoDiv="YY") then
		        addSql = addSql & " and isNULL(ct.infodiv,'')<>''"
		    elseif (FRectInfoDiv="NN") then
		        addSql = addSql & " and isNULL(ct.infodiv,'')=''"
		    else
    		    addSql = addSql & " and ct.infodiv='"&FRectInfoDiv&"'"
    		end if
		end if

        ''����
        ''addSql = addSql + " and (select count(*) from db_item.dbo.tbl_item_option o where o.itemid=i.itemid and o.optaddprice>0)<1" ''�ɼ��߰��ݾ׾��°Ÿ�.

		sqlStr = " select count(i.itemid) as cnt, CEILING(CAST(Count(i.itemid) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr + " from db_item.dbo.tbl_item as i "
		IF (FRectLotteNotReg<>"M") and (FRectLotteNotReg<>"") then
		    sqlStr = sqlStr + " 	join db_item.dbo.tbl_LTiMall_regItem as m "
		ELSE
		    sqlStr = sqlStr + " 	left join db_item.dbo.tbl_LTiMall_regItem as m "
	    END IF
		sqlStr = sqlStr + " 		on i.itemid=m.itemid "
		sqlStr = sqlStr + " 	left Join db_item.dbo.tbl_OutMall_CateMap_Summary as c "
		sqlStr = sqlStr + " 		on c.mallid='"&CMALLNAME&"' and c.tenCateLarge=i.cate_large and c.tenCateMid=i.cate_mid and c.tenCateSmall=i.cate_small "
		sqlStr = sqlStr + " 	left join db_user.dbo.tbl_user_c uc"
		sqlStr = sqlStr + " 	on i.makerid=uc.userid"
		sqlStr = sqlStr + " 	join db_item.dbo.tbl_item_contents as ct "
		sqlStr = sqlStr + " 	on i.itemid=ct.itemid"
		sqlStr = sqlStr + " where 1=1"

		''IF (FRectLotteNotReg="D" or FRectLotteNotReg="R") then  ''�̹� ��ǰ�� ��ϵ� CASE  FRectLotteNotReg
		''IF (FRectLotteNotReg="") or (FRectLotteYes10x10No <> "") THEN
		if (FRectLotteNotReg<>"M" and FRectLotteNotReg<>"Q" and FRectLotteNotReg<>"V") then

		ELSE
    		sqlStr = sqlStr + "     and i.isusing='Y' "
    		sqlStr = sqlStr + "     and i.deliverfixday not in ('C','X') "
    		sqlStr = sqlStr + "     and i.basicimage is not null "
    		sqlStr = sqlStr + "     and i.itemdiv<50 "  '''and i.itemdiv<>'08'
    		sqlStr = sqlStr + "     and i.cate_large<>'' "
    		sqlStr = sqlStr + "     and i.cate_large<>'999' "
    		sqlStr = sqlStr + "		and i.makerid not in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "	'������� �귣��
    		sqlStr = sqlStr + "		and i.itemid not in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "		'������� ��ǰ
    		sqlStr = sqlStr + "     and i.sellcash>=1000 "
    		sqlStr = sqlStr + "     and i.itemdiv<>'06'" ''�ֹ����� ��ǰ ���� 2013/01/15
    		sqlStr = sqlStr + "		and uc.isExtUsing='Y'"  ''20130304 �귣�� ���޻�뿩�� Y��.
    	END IF
		sqlStr = sqlStr & addSql

''rw "=FRectLotteNotReg="&FRectLotteNotReg
''rw sqlStr
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'������������ ��ü ���������� Ŭ �� �Լ�����
		if Cint(FCurrPage)>Cint(FTotalPage) then
			FResultCount = 0
			exit sub
		end if

		sqlStr = " select  top " + CStr(FPageSize*FCurrPage) + " i.itemid,i.itemname,i.smallImage "
		sqlStr = sqlStr + "		, i.makerid, i.regdate, i.lastUpdate, i.orgPrice, i.sellcash, i.buycash"
		sqlStr = sqlStr + "		, i.sellYn, i.sailyn, i.LimitYn, i.LimitNo, i.LimitSold, i.deliverytype, i.optionCnt"
		sqlStr = sqlStr + "		, m.LTiMallRegdate, m.LTiMallLastUpdate, m.LTiMallGoodNo, m.LTiMallTmpGoodNo, m.LTiMallPrice, m.LTiMallSellYn, m.regUserid, IsNULL(m.LTiMallStatCd,-9) as LTiMallStatCd "
		sqlStr = sqlStr + "		, c.mapCnt, m.regedOptCnt, m.rctSellCNT, m.accFailCNT, m.lastErrStr "
		sqlStr = sqlStr + "     ,uc.defaultdeliverytype, uc.defaultfreeBeasongLimit"
		sqlStr = sqlStr + "		, Ct.infoDiv, m.optAddPrcCnt, m.optAddPrcRegType"
		sqlStr = sqlStr + " from db_item.dbo.tbl_item as i "
		IF (FRectLotteNotReg<>"M") and (FRectLotteNotReg<>"") then
		    sqlStr = sqlStr + " 	Join db_item.dbo.tbl_LTiMall_regItem as m "
		ELSE
		    sqlStr = sqlStr + " 	left join db_item.dbo.tbl_LTiMall_regItem as m "
	    END IF
		sqlStr = sqlStr + " 		on i.itemid=m.itemid "
		sqlStr = sqlStr + " 	left Join db_item.dbo.tbl_OutMall_CateMap_Summary as c "
		sqlStr = sqlStr + " 		on c.mallid='"&CMALLNAME&"' and c.tenCateLarge=i.cate_large and c.tenCateMid=i.cate_mid and c.tenCateSmall=i.cate_small "
		sqlStr = sqlStr + " 	left join db_user.dbo.tbl_user_c uc"
		sqlStr = sqlStr + " 	on i.makerid=uc.userid"
		sqlStr = sqlStr + " 	join db_item.dbo.tbl_item_contents as ct "
		sqlStr = sqlStr + " 	on i.itemid=ct.itemid"
		sqlStr = sqlStr + " where 1=1"

		''if (FRectLotteNotReg="D" or FRectLotteNotReg="R") then
		''IF (FRectLotteNotReg="") or (FRectLotteYes10x10No <> "") THEN
		if (FRectLotteNotReg<>"M" and FRectLotteNotReg<>"Q" and FRectLotteNotReg<>"V") then
		ELSE
    		sqlStr = sqlStr + "     and i.isusing='Y' "
    		sqlStr = sqlStr + "     and i.deliverfixday not in ('C','X') "
    		sqlStr = sqlStr + "     and i.basicimage is not null "
    		sqlStr = sqlStr + "     and i.itemdiv<50 "  ''and i.itemdiv<>'08'
    		sqlStr = sqlStr + "     and i.cate_large<>'' "
    		sqlStr = sqlStr + "     and i.cate_large<>'999' "
    		sqlStr = sqlStr + "		and i.makerid not in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "	'������� �귣��
    		sqlStr = sqlStr + "		and i.itemid not in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "		'������� ��ǰ
    		sqlStr = sqlStr + "     and i.sellcash>=1000 "
    		sqlStr = sqlStr + "     and i.itemdiv<>'06'" ''�ֹ����� ��ǰ ���� 2013/01/15
    		sqlStr = sqlStr + "		and uc.isExtUsing='Y'"  ''20130304 �귣�� ���޻�뿩�� Y��.
    	END If
		sqlStr = sqlStr & addSql
		'sqlStr = sqlStr + " order by i.itemid desc "
		IF (FRectLotteNotReg="F") then
		    sqlStr = sqlStr + " order by m.LtiMallLastupdate "
		ELSEIF (FRectOrdType="B") then
		    sqlStr = sqlStr + " order by i.itemscore desc, i.itemid desc "
		ELSEIF (FRectOrdType="BM") then
		    sqlStr = sqlStr + " order by m.rctSellCNT desc,i.itemscore desc, m.itemid desc"
		else
		    sqlStr = sqlStr + " order by i.itemid desc " '' m.regdate desc
	    end if
''rw sqlStr
''response.end

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CLotteiMallItem

				FItemList(i).Fitemid			= rsget("itemid")
				FItemList(i).Fitemname			= db2html(rsget("itemname"))
				FItemList(i).FsmallImage		= rsget("smallImage")
				FItemList(i).Fmakerid			= rsget("makerid")
				FItemList(i).Fregdate			= rsget("regdate")
				FItemList(i).FlastUpdate		= rsget("lastUpdate")
				FItemList(i).ForgPrice			= rsget("orgPrice")
				FItemList(i).FSellCash			= rsget("sellcash")
				FItemList(i).FBuyCash			= rsget("buycash")
				FItemList(i).FsellYn			= rsget("sellYn")
				FItemList(i).FsaleYn			= rsget("sailyn")
				FItemList(i).FLimitYn			= rsget("LimitYn")
				FItemList(i).FLimitNo			= rsget("LimitNo")
				FItemList(i).FLimitSold			= rsget("LimitSold")

				FItemList(i).FLTiMallRegdate		= rsget("LTiMallRegdate")
				FItemList(i).FLTiMallLastUpdate	= rsget("LTiMallLastUpdate")
				FItemList(i).FLTiMallGoodNo		= rsget("LTiMallGoodNo")
				FItemList(i).FLTiMallTmpGoodNo	= rsget("LTiMallTmpGoodNo")
				FItemList(i).FLTiMallPrice		= rsget("LTiMallPrice")
				FItemList(i).FLTiMallSellYn		= rsget("LTiMallSellYn")
				FItemList(i).FregUserid			= rsget("regUserid")
				FItemList(i).FLTiMallStatCd		= rsget("LTiMallStatCd")
				FItemList(i).FCateMapCnt		= rsget("mapCnt")
                FItemList(i).Fdeliverytype      = rsget("deliverytype")
                FItemList(i).Fdefaultdeliverytype = rsget("defaultdeliverytype")
                FItemList(i).FdefaultfreeBeasongLimit = rsget("defaultfreeBeasongLimit")

				if Not(FItemList(i).FsmallImage="" or isNull(FItemList(i).FsmallImage)) then
					FItemList(i).FsmallImage = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("smallImage")
				else
					FItemList(i).FsmallImage = "http://fiximage.10x10.co.kr/images/spacer.gif"
				end if
                FItemList(i).FoptionCnt         = rsget("optionCnt")
                FItemList(i).FregedOptCnt       = rsget("regedOptCnt")
                FItemList(i).FrctSellCNT        = rsget("rctSellCNT")
                FItemList(i).FaccFailCNT      = rsget("accFailCNT")
                FItemList(i).FlastErrStr      = rsget("lastErrStr")
                FItemList(i).FinfoDiv           = rsget("infoDiv")

                FItemList(i).FoptAddPrcCnt      = rsget("optAddPrcCnt")
                FItemList(i).FoptAddPrcRegType  = rsget("optAddPrcRegType")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end Sub

    ''' ��ϵ��� ���ƾ� �� ��ǰ..
    public Sub getLotteReqExpireItemList
		dim sqlStr, addSql, i

		sqlStr = " select count(i.itemid) as cnt, CEILING(CAST(Count(i.itemid) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr + " from db_item.dbo.tbl_item as i "
		sqlStr = sqlStr + " 	join db_item.dbo.tbl_LtiMall_regItem as m "
		sqlStr = sqlStr + " 		on i.itemid=m.itemid "
		sqlStr = sqlStr + " 		and m.LTiMallGoodNo is Not Null"
		sqlStr = sqlStr + " 		and m.LTiMallSellYn= 'Y' "                ''' �Ե� �Ǹ����ΰŸ�.
		sqlStr = sqlStr + " 	join db_user.dbo.tbl_user_c c"
		sqlStr = sqlStr + " 	    on i.makerid=c.userid"
		sqlStr = sqlStr + " 	Join db_item.dbo.tbl_item_contents ct"
		sqlStr = sqlStr + " 	on i.itemid=ct.itemid"
		sqlStr = sqlStr + " where (i.isusing<>'Y' or i.isExtUsing<>'Y' "
		sqlStr = sqlStr + "     or i.deliverytype in ('7') "
		'//���ǹ�� 10000�� �̻�
		IF (CUPJODLVVALID) then
		    sqlStr = sqlStr + "     or ((i.deliveryType=9) and (i.sellcash<10000) )" ''
		ELSE
            sqlStr = sqlStr + "     or ((i.deliveryType=9) and (i.sellcash<isNULL(c.defaultFreebeasongLimit,0)) )" ''
        END IF
		sqlStr = sqlStr + "     or i.deliverfixday in ('C','X') "
		sqlStr = sqlStr + "     or i.itemdiv='06'" ''�ֹ����� ��ǰ ���� 2013/01/15
		sqlStr = sqlStr + "     or i.itemdiv>=50 or i.itemdiv='08' or i.cate_large='999' or i.cate_large=''"
		sqlStr = sqlStr + "		or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "	'������� �귣��
		sqlStr = sqlStr + "		or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "		'������� ��ǰ
		sqlStr = sqlStr + "		or c.isExtUsing='N'"
		sqlStr = sqlStr + "		or isNULL(ct.infodiv,'') in ('','18','20','21','22')"  ''ȭ��ǰ, ��ǰ�� ����
        sqlStr = sqlStr + " )"

        ''//���� ���ܻ�ǰ
        sqlStr = sqlStr & " and i.itemid not in ("
        sqlStr = sqlStr & "     select itemid from db_temp.dbo.tbl_jaehyumall_not_edit_itemid"
        sqlStr = sqlStr & "     where stDt<getdate()"
        sqlStr = sqlStr & "     and edDt>getdate()"
        sqlStr = sqlStr & "     and mallgubun='"&CMALLNAME&"'"
        sqlStr = sqlStr & " )"

        if FRectMakerid<>"" then
			sqlStr = sqlStr & " and i.makerid='" & FRectMakerid & "'"
		end if

		if FRectItemID<>"" then
			sqlStr = sqlStr & " and i.itemid in (" & FRectItemID & ")"
		end if

		''2013/05/29 �߰�
		if (FRectInfoDiv<>"") then
		    if (FRectInfoDiv="YY") then
		        sqlStr = sqlStr & " and isNULL(ct.infodiv,'')<>''"
		    elseif (FRectInfoDiv="NN") then
		        sqlStr = sqlStr & " and isNULL(ct.infodiv,'')=''"
		    else
    		    sqlStr = sqlStr & " and ct.infodiv='"&FRectInfoDiv&"'"
    		end if
		end if
''rw sqlStr
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'������������ ��ü ���������� Ŭ �� �Լ�����
		if Cint(FCurrPage)>Cint(FTotalPage) then
			FResultCount = 0
			exit sub
		end if

		sqlStr = " select  top " + CStr(FPageSize*FCurrPage) + " i.* "
		sqlStr = sqlStr + "		, m.LTiMallRegdate, m.LTiMallLastUpdate, m.LTiMallGoodNo, m.LTiMallTmpGoodNo, m.LTiMallPrice, m.LTiMallSellYn, m.regUserid, m.LTiMallStatCd "
		sqlStr = sqlStr + "		, 1 as mapCnt "
		sqlStr = sqlStr + "     ,c.defaultdeliverytype, c.defaultfreeBeasongLimit"
		sqlStr = sqlStr + "     ,ct.infoDiv, m.optAddPrcCnt, m.optAddPrcRegType"
		sqlStr = sqlStr + " from db_item.dbo.tbl_item as i "
		sqlStr = sqlStr + " 	join db_item.dbo.tbl_LtiMall_regItem as m "
		sqlStr = sqlStr + " 		on i.itemid=m.itemid "
		sqlStr = sqlStr + " 		and m.LTiMallGoodNo is Not Null"
		sqlStr = sqlStr + " 		and m.LTiMallSellYn= 'Y' "                ''' �Ե� �Ǹ����ΰŸ�.
		sqlStr = sqlStr + " 	join db_user.dbo.tbl_user_c c"
		sqlStr = sqlStr + " 	    on i.makerid=c.userid"
		sqlStr = sqlStr + " 	Join db_item.dbo.tbl_item_contents ct"
		sqlStr = sqlStr + " 	on i.itemid=ct.itemid"
		sqlStr = sqlStr + " where (i.isusing<>'Y' or i.isExtUsing<>'Y' "
		sqlStr = sqlStr + "     or i.deliverytype in ('7') "
		'//���ǹ�� 10000�� �̻�
		IF (CUPJODLVVALID) then
		    sqlStr = sqlStr + "     or ((i.deliveryType=9) and (i.sellcash<10000) )" ''
		ELSE
            sqlStr = sqlStr + "     or ((i.deliveryType=9) and (i.sellcash<isNULL(c.defaultFreebeasongLimit,0)) )" ''
        ENd IF
		sqlStr = sqlStr + "     or i.deliverfixday in ('C','X') "
		sqlStr = sqlStr + "     or i.itemdiv='06'" ''�ֹ����� ��ǰ ���� 2013/01/15
		sqlStr = sqlStr + "     or i.itemdiv>=50 or i.itemdiv='08' or i.cate_large='999' or i.cate_large=''"
		sqlStr = sqlStr + "		or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "	'������� �귣��
		sqlStr = sqlStr + "		or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "		'������� ��ǰ
		sqlStr = sqlStr + "		or c.isExtUsing='N'"
		sqlStr = sqlStr + "		or isNULL(ct.infodiv,'') in ('','18','20','21','22')"
        sqlStr = sqlStr + " )"

        ''//���� ���ܻ�ǰ //���� ������ �ҵ�.
        sqlStr = sqlStr & " and i.itemid not in ("
        sqlStr = sqlStr & "     select itemid from db_temp.dbo.tbl_jaehyumall_not_edit_itemid"
        sqlStr = sqlStr & "     where stDt<getdate()"
        sqlStr = sqlStr & "     and edDt>getdate()"
        sqlStr = sqlStr & "     and mallgubun='"&CMALLNAME&"'"
        sqlStr = sqlStr & " )"

        if FRectMakerid<>"" then
			sqlStr = sqlStr & " and i.makerid='" & FRectMakerid & "'"
		end if

		if FRectItemID<>"" then
			sqlStr = sqlStr & " and i.itemid in (" & FRectItemID & ")"
		end if

		''2013/05/29 �߰�
		if (FRectInfoDiv<>"") then
		    if (FRectInfoDiv="YY") then
		        sqlStr = sqlStr & " and isNULL(ct.infodiv,'')<>''"
		    elseif (FRectInfoDiv="NN") then
		        sqlStr = sqlStr & " and isNULL(ct.infodiv,'')=''"
		    else
    		    sqlStr = sqlStr & " and ct.infodiv='"&FRectInfoDiv&"'"
    		end if
		end if

		sqlStr = sqlStr + " order by m.regdate desc, i.itemid desc "
''rw sqlStr
''response.end

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CLotteiMallItem

				FItemList(i).Fitemid			= rsget("itemid")
				FItemList(i).Fitemname			= db2html(rsget("itemname"))
				FItemList(i).FsmallImage		= rsget("smallImage")
				FItemList(i).Fmakerid			= rsget("makerid")
				FItemList(i).Fregdate			= rsget("regdate")
				FItemList(i).FlastUpdate		= rsget("lastUpdate")
				FItemList(i).ForgPrice			= rsget("orgPrice")
				FItemList(i).FSellCash			= rsget("sellcash")
				FItemList(i).FBuyCash			= rsget("buycash")
				FItemList(i).FsellYn			= rsget("sellYn")
				FItemList(i).FsaleYn			= rsget("sailyn")
				FItemList(i).FLimitYn			= rsget("LimitYn")
				FItemList(i).FLimitNo			= rsget("LimitNo")
				FItemList(i).FLimitSold			= rsget("LimitSold")

				FItemList(i).FLTiMallRegdate		= rsget("LTiMallRegdate")
				FItemList(i).FLTiMallLastUpdate	= rsget("LTiMallLastUpdate")
				FItemList(i).FLTiMallGoodNo		= rsget("LTiMallGoodNo")
				FItemList(i).FLTiMallTmpGoodNo	= rsget("LTiMallTmpGoodNo")
				FItemList(i).FLTiMallPrice		= rsget("LTiMallPrice")
				FItemList(i).FLTiMallSellYn		= rsget("LTiMallSellYn")
				FItemList(i).FregUserid			= rsget("regUserid")
				FItemList(i).FLTiMallStatCd		= rsget("LTiMallStatCd")
				FItemList(i).FCateMapCnt		= rsget("mapCnt")
                FItemList(i).Fdeliverytype      = rsget("deliverytype")
                FItemList(i).Fdefaultdeliverytype = rsget("defaultdeliverytype")
                FItemList(i).FdefaultfreeBeasongLimit = rsget("defaultfreeBeasongLimit")

				if Not(FItemList(i).FsmallImage="" or isNull(FItemList(i).FsmallImage)) then
					FItemList(i).FsmallImage = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("smallImage")
				else
					FItemList(i).FsmallImage = "http://fiximage.10x10.co.kr/images/spacer.gif"
				end if

                FItemList(i).FinfoDiv = rsget("infoDiv")
                FItemList(i).FoptAddPrcCnt      = rsget("optAddPrcCnt")
                FItemList(i).FoptAddPrcRegType  = rsget("optAddPrcRegType")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end Sub

	'--------------------------------------------------------------------------------
	'// �̵�� ��ǰ ���(��Ͽ�)
	public Sub getLTiMallNotRegItemList
		dim strSql, addSql, i

		if FRectItemID<>"" then
			addSql = addSql & " and i.itemid in (" & FRectItemID & ")"
			''' �ɼ� �߰��ݾ� �ִ°�� ��� �Ұ�. //�ɼ� ��ü ǰ���� ��� ��� �Ұ�.
			addSql = addSql & " and i.itemid not in ("
			addSql = addSql & " select itemid from ("
            addSql = addSql & "     select itemid"
            addSql = addSql & " 	,count(*) as optCNT"
            addSql = addSql & " 	,sum(CASE WHEN optAddPrice>0 then 1 ELSE 0 END) as optAddCNT"
            addSql = addSql & " 	,sum(CASE WHEN (optsellyn='N') or (optlimityn='Y' and (optlimitno-optlimitsold<1)) then 1 ELSE 0 END) as optNotSellCnt"
            addSql = addSql & " 	from db_item.dbo.tbl_item_option"
            addSql = addSql & " 	where itemid in (" & FRectItemID & ")"
            addSql = addSql & " 	and isusing='Y'"
            addSql = addSql & " 	group by itemid"
            addSql = addSql & " ) T"
            addSql = addSql & " where optAddCNT>0"
            addSql = addSql & " or (optCnt-optNotSellCnt<1)"
            addSql = addSql & " )"

            ''' 2013/05/29 Ư��ǰ�� ��� �Ұ� (ȭ��ǰ, ��ǰ��)
            addSql = addSql & " and isNULL(c.infodiv,'') not in ('','18','20','21','22')"
		end if

		strSql = "Select top " & FPageSize & " i.* "
		strSql = strSql & "		, c.keywords, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent "
		strSql = strSql & "		, '"&CitemGbnKey&"' as itemGbnKey"
		strSql = strSql & "		, isNULL(R.LtiMallStatCD,-9) as LtiMallStatCD"
		strSql = strSql & " From db_item.dbo.tbl_item as i "
		strSql = strSql & " 	join db_item.dbo.tbl_item_contents as c "
		strSql = strSql & " 		on i.itemid=c.itemid "
		strSql = strSql & " 	Join (Select tenCateLarge, tenCateMid, tenCateSmall, count(*) as mapCnt From db_item.dbo.tbl_LTiMall_cate_mapping Group by tenCateLarge, tenCateMid, tenCateSmall ) as cm "
		strSql = strSql & " 		on cm.tenCateLarge=i.cate_large and cm.tenCateMid=i.cate_mid and cm.tenCateSmall=i.cate_small "
		''strSql = strSql & " 	Join db_item.dbo.tbl_LTiMall_cateGbn_mapping G"
		''strSql = strSql & " 		on G.tenCateLarge=i.cate_large and G.tenCateMid=i.cate_mid and G.tenCateSmall=i.cate_small "
		strSql = strSql & " 	left join db_item.dbo.tbl_LtiMall_regItem R"
		strSql = strSql & " 	on i.itemid=R.itemid"
		strSql = strSql & " Where i.isusing='Y' "
		strSql = strSql & "     and i.isExtUsing='Y' "
		strSql = strSql & "     and i.deliverytype not in ('7')"
		IF (CUPJODLVVALID) then
		    strSql = strSql & " 	and ((i.deliveryType<>9) or ((i.deliveryType=9) and (i.sellcash>=10000)))"
		ELSE
		    strSql = strSql & " 	and (i.deliveryType<>9)"
	    END IF
		strSql = strSql & "     and i.sellyn='Y' "
		strSql = strSql & "     and i.deliverfixday not in ('C','X') "																				'�ö��/ȭ����� ��ǰ ����
		strSql = strSql & "     and i.basicimage is not null "
		strSql = strSql & "     and i.itemdiv<50 and i.itemdiv<>'08' "
		strSql = strSql & "     and i.cate_large<>'' "
		strSql = strSql & "     and i.cate_large<>'999' "
		strSql = strSql & "     and i.sellcash>0 "
		strSql = strSql & "     and ((i.LimitYn='N') or ((i.LimitYn='Y') and (i.LimitNo-i.LimitSold>="&CMAXLIMITSELL&")) )" ''���� ǰ�� �� ��� ����.
		''strSql = strSql & "     and i.sellcash=i.orgprice"              '''��а� ���� ���ϴ°͸�.. // ���ݼ��� ��� ����..?
		''strSql = strSql & " 	and (i.orgprice<>0 and ((i.orgprice-i.orgSuplyCash)/i.orgprice)*100>=" & CMAXMARGIN & ")"							'������ ��ǰ ����
		strSql = strSql & " 	and (i.sellcash<>0 and ((i.sellcash-i.buycash)/i.sellcash)*100>=" & CMAXMARGIN & ")"
		strSql = strSql & "		and i.makerid not in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "	'������� �귣��
		strSql = strSql & "		and i.itemid not in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "		'������� ��ǰ
		strSql = strSql & "		and i.itemid not in (Select itemid From db_item.dbo.tbl_LtiMall_regItem where LtiMallStatCD>3) "	''LtiMallStatCD>=3 ��ϿϷ��̻��� ��Ͼȵ�.										'�Ե���ϻ�ǰ ����
		''strSql = strSql & "		and cm.mapCnt is Not Null "	& addSql
		strSql = strSql & "		"	& addSql																				'ī�װ� ��Ī ��ǰ��
''rw strSql
		rsget.Open strSql,dbget,1

		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CLotteiMallItem
				FItemList(i).Fitemid			= rsget("itemid")

				FItemList(i).FtenCateLarge		= rsget("cate_large")
				FItemList(i).FtenCateMid		= rsget("cate_mid")
				FItemList(i).FtenCateSmall		= rsget("cate_small")
				FItemList(i).Fitemname			= db2html(rsget("itemname"))
				FItemList(i).FitemDiv			= rsget("itemdiv")
				FItemList(i).FsmallImage		= rsget("smallImage")
				FItemList(i).Fmakerid			= rsget("makerid")
				FItemList(i).Fregdate			= rsget("regdate")
				FItemList(i).FlastUpdate		= rsget("lastUpdate")
				FItemList(i).ForgPrice			= rsget("orgPrice")
				FItemList(i).ForgSuplyCash		= rsget("orgSuplyCash")
				FItemList(i).FSellCash			= rsget("sellcash")
				FItemList(i).FBuyCash			= rsget("buycash")
				FItemList(i).FsellYn			= rsget("sellYn")
				FItemList(i).FsaleYn			= rsget("sailyn")
				FItemList(i).FisUsing			= rsget("isusing")
				FItemList(i).FLimitYn			= rsget("LimitYn")
				FItemList(i).FLimitNo			= rsget("LimitNo")
				FItemList(i).FLimitSold			= rsget("LimitSold")
				FItemList(i).Fkeywords			= rsget("keywords")
				FItemList(i).Fvatinclude        = rsget("vatinclude")
				FItemList(i).ForderComment		= db2html(rsget("ordercomment"))
				FItemList(i).FoptionCnt			= rsget("optionCnt")
				FItemList(i).FbasicImage		= "http://webimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("basicImage")
				FItemList(i).FmainImage			= "http://webimage.10x10.co.kr/image/main/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("mainimage")
				FItemList(i).FmainImage2		= "http://webimage.10x10.co.kr/image/main2/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("mainimage2")
				FItemList(i).Fsourcearea		= db2html(rsget("sourcearea"))
				FItemList(i).Fmakername			= db2html(rsget("makername"))
				FItemList(i).FUsingHTML			= rsget("usingHTML")
				FItemList(i).Fitemcontent		= db2html(rsget("itemcontent"))

                FItemList(i).FitemGbnKey        = rsget("itemGbnKey")
                FItemList(i).FLtiMallStatCD     = rsget("LtiMallStatCD")

                FItemList(i).FRectMode			= FRectMode
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end Sub


	'--------------------------------------------------------------------------------
	'// �Ե�iMall ��ǰ ���(������)
	public Sub getLTiMallEditedItemList
		dim strSql, addSql, i

		if FRectItemID<>"" then
			'���û�ǰ�� �ִٸ�
			addSql = " and i.itemid in (" & FRectItemID & ")"
		elseif FRectNotJehyu="Y" then
			'���޸� ��ǰ�� �ƴѰ�
			addSql = " and i.isExtUsing='N' "
		else
			'������ ��ǰ��
			''addSql = " and m.LTiMallLastUpdate<i.lastupdate"
		end if

        ''//���� ���ܻ�ǰ
        addSql = addSql & " and i.itemid not in ("
        addSql = addSql & "     select itemid from db_temp.dbo.tbl_jaehyumall_not_edit_itemid"
        addSql = addSql & "     where stDt<getdate()"
        addSql = addSql & "     and edDt>getdate()"
        addSql = addSql & "     and mallgubun='"&CMALLNAME&"'"
        addSql = addSql & " )"

		strSql = "Select top " & FPageSize & " i.* "
		strSql = strSql & "		, c.keywords, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent "
		strSql = strSql & "		, m.LTiMallGoodNo, m.LTiMallTmpGoodNo, m.LTiMallSellYn "
		strSql = strSql & " From db_item.dbo.tbl_item as i "
		strSql = strSql & " 	join db_item.dbo.tbl_item_contents as c "
		strSql = strSql & " 		on i.itemid=c.itemid "
		strSql = strSql & " 	join db_item.dbo.tbl_LtiMall_regItem as m "
		strSql = strSql & " 		on i.itemid=m.itemid "
		if (FRectMatchCateNotCheck<>"on") then
    		strSql = strSql & " 	Join (Select tenCateLarge, tenCateMid, tenCateSmall, count(*) as mapCnt From db_item.dbo.tbl_LTiMall_cate_mapping Group by tenCateLarge, tenCateMid, tenCateSmall ) as cm "
    		strSql = strSql & " 		on cm.tenCateLarge=i.cate_large and cm.tenCateMid=i.cate_mid and cm.tenCateSmall=i.cate_small "
    	end if
		''strSql = strSql & " 	Join db_item.dbo.tbl_LTiMall_cateGbn_mapping G"
		''strSql = strSql & " 		on G.tenCateLarge=i.cate_large and G.tenCateMid=i.cate_mid and G.tenCateSmall=i.cate_small "
		strSql = strSql & " Where 1=1 " & addSql
		strSql = strSql & " and m.LTiMallGoodNo is Not Null "									'#��� ��ǰ��

''rw strSql
		rsget.Open strSql,dbget,1

		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CLotteiMallItem
				FItemList(i).Fitemid			= rsget("itemid")

				FItemList(i).FtenCateLarge		= rsget("cate_large")
				FItemList(i).FtenCateMid		= rsget("cate_mid")
				FItemList(i).FtenCateSmall		= rsget("cate_small")
				FItemList(i).Fitemname			= db2html(rsget("itemname"))
				FItemList(i).FitemDiv			= rsget("itemdiv")
				FItemList(i).FsmallImage		= rsget("smallImage")
				FItemList(i).Fmakerid			= rsget("makerid")
				FItemList(i).Fregdate			= rsget("regdate")
				FItemList(i).FlastUpdate		= rsget("lastUpdate")
				FItemList(i).ForgPrice			= rsget("orgPrice")
				FItemList(i).ForgSuplyCash		= rsget("orgSuplyCash")
				FItemList(i).FSellCash			= rsget("sellcash")
				FItemList(i).FBuyCash			= rsget("buycash")
				FItemList(i).FsellYn			= rsget("sellYn")
				FItemList(i).FsaleYn			= rsget("sailyn")
				FItemList(i).FisUsing			= rsget("isusing")
				FItemList(i).FLimitYn			= rsget("LimitYn")
				FItemList(i).FLimitNo			= rsget("LimitNo")
				FItemList(i).FLimitSold			= rsget("LimitSold")
				FItemList(i).Fkeywords			= rsget("keywords")
				FItemList(i).ForderComment		= db2html(rsget("ordercomment"))
				FItemList(i).FoptionCnt			= rsget("optionCnt")
				FItemList(i).FbasicImage		= "http://webimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("basicImage")
				FItemList(i).FmainImage			= "http://webimage.10x10.co.kr/image/main/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("mainimage")
				FItemList(i).FmainImage2		= "http://webimage.10x10.co.kr/image/main2/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("mainimage2")
				FItemList(i).Fsourcearea		= db2html(rsget("sourcearea"))
				FItemList(i).Fmakername			= db2html(rsget("makername"))
				FItemList(i).FUsingHTML			= rsget("usingHTML")
				FItemList(i).Fitemcontent		= db2html(rsget("itemcontent"))
				FItemList(i).FLTiMallGoodNo		= rsget("LTiMallGoodNo")
				FItemList(i).FLTiMallTmpGoodNo	= rsget("LTiMallTmpGoodNo")
				FItemList(i).FLTiMallSellYn		= rsget("LTiMallSellYn")
                FItemList(i).Fvatinclude        = rsget("vatinclude")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end Sub

	'--------------------------------------------------------------------------------
	public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function
end class

'// MD��ǰ�� ���û��� ���
Function printLotteCateGrpSelectBox(fnm,selcd)
	dim strSql, rstStr
	rstStr = "<Select name='" & fnm & "' class='select'>"
	rstStr = rstStr & "<option value=''>��ü</option>"

	strSql = "Select * From db_temp.dbo.tbl_lotte_MDCateGrp Where isUsing='Y'"
	rsget.Open strSql,dbget,1
	if Not(rsget.EOF or rsget.BOF) then
		Do Until rsget.EOF
			if cStr(rsget("groupCode"))=cStr(selcd) then
				rstStr = rstStr & "<option value='" & rsget("groupCode") & "' selected>" & rsget("groupName")& "</option>"
			else
				rstStr = rstStr & "<option value='" & rsget("groupCode") & "'>" & rsget("groupName")& "</option>"
			end if
			rsget.MoveNext
		Loop
	end if
	rsget.Close

	rstStr = rstStr & "</select>"

	printLotteCateGrpSelectBox = rstStr
end Function

'// ��ǰ�̹��� ���翩�� �˻�
function ImageExists(byval iimg)
	if (IsNull(iimg)) or (trim(iimg)="") or (Right(trim(iimg),1)="\") or (Right(trim(iimg),1)="/") then
		ImageExists = false
	else
		ImageExists = true
	end if
end function

Function GetRaiseValue(value)
    If Fix(value) < value Then
    GetRaiseValue = Fix(value) + 1
    Else
    GetRaiseValue = Fix(value)
    End If
End Function
%>