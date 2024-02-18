<%
CONST CMAXMARGIN = 15			'' MaxMagin��.. '(�Ե����� 11%)
CONST CMAXLIMITSELL = 5        '' �� ���� �̻��̾�� �Ǹ���. // �ɼ������� ��������.
CONST CMALLNAME = "lotteCom"

class CLotteItem
	Public FRegId
	Public FLastupdateId

	public FLastUpdate
	public FisUsing

	'���MD
	public FMDCode
	public FMDName
	public FSellFeeType
	public FNormalSellFee
	public FEventSellFee

	'MD��ǰ��
	public FgroupCode
	public FSuperGroupName
	public FGroupName

	'�Ե����� ī�װ�
	public FDispNo
	public FDispNm
	public FDispLrgNm
	public FDispMidNm
	public FDispSmlNm
	public FDispThnNm
	public FtenCateLarge
	public FtenCateMid
	public FtenCateSmall
	public FtenCDLName
	public FtenCDMName
	public FtenCDSName
	public FtenCateName
    public Fdisptpcd
    public FCateisusing

	'�Ե����� �귣��
	public FlotteBrandCd
	public FlotteBrandName
	public FTenMakerid
	public FTenBrandName

	'�Ե����� ��ǰ���
	public FLotteRegdate
	public FLotteLastUpdate
	public FLotteGoodNo				'�ǻ�ǰ��ȣ
	public FLotteTmpGoodNo			'�ӽû�ǰ��ȣ
	public FLottePrice
	public FLotteSellYn
	public FregUserid
	public FLotteDispCnt
	public FCateMapCnt
	public FLotteStatCd				'��ǰ��ϻ���
    public FrctSellCNT              '6�����Ǹŷ�
    public FregedOptCnt             '�Ե���Ͽɼ�
    public FaccFailCNT              '��ϼ��� ���� Ƚ��
    public FlastErrStr              '��������

    '''public Fcorp_dlvp_sn             '��ǰ�ּ����ڵ�

	'�ٹ����� ��ǰ���
	public Fidx
	public FNewitemname
	public FItemnameChange
	Public FItemoption
	Public Foptsellyn
	Public Foptlimityn
	Public Foptlimitno
	Public Foptlimitsold

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
    public FrequireMakeDay
    public Fdefaultdeliverytype
    public FdefaultfreeBeasongLimit
	Public FOptaddprice
	Public FOptionname
	Public FRegedOptionname
	Public FRegedItemname

    ''ǰ������ �� ������������.
    public FinfoDiv
    public Fsafetyyn
    public FsafetyDiv
    public FsafetyNum

    public FoptAddPrcCnt
    public FoptAddPrcRegType
    public FLastcateChgDate

    public FmaySoldOut ''���޸� �����Ե�
	Public FSpecialPrice
	Public FStartDate
	Public FEndDate

	Public Function getRealItemname
		If FitemnameChange = "" Then
			getRealItemname = FNewitemname
		Else
			getRealItemname = FItemnameChange
		End If
	End Function

    public function getLotteReturnDLVCode
        dim sqlStr
        dim retReturnCode

        getLotteReturnDLVCode = "72125"    ''�⺻ ��ǰ�� �ڵ�(����)
    Exit function '' ������ ������ 2013/02/25
        if Not (Fdeliverytype="2" or Fdeliverytype="7" or Fdeliverytype="9") then Exit function  ''��ü����̸� ����


'        ''�ӽ� ����
'        if (Fcorp_dlvp_sn<>"72125") and (Fcorp_dlvp_sn<>"113045") and (Fcorp_dlvp_sn<>"113044") and (Fcorp_dlvp_sn<>"114747") then
'            Fcorp_dlvp_sn = "72125"
'        end if

        '' ��ǰ�ּ����� ������ �˻� ������� ����
        sqlStr = " select R.returnCode from db_item.dbo.tbl_OutMall_BrandReturnCode R"
        sqlStr = sqlStr & "  	Join db_temp.dbo.tbl_jaehyumall_returnInfo T"
        sqlStr = sqlStr & "	on R.makerid='"&Fmakerid&"'"
        sqlStr = sqlStr & "	and R.returnCode=T.returnCode"
        sqlStr = sqlStr & "	Join db_partner.dbo.tbl_partner p"
        sqlStr = sqlStr & "	on p.id=R.makerid"
        sqlStr = sqlStr & " where replace(T.returnAddress,' ','')=replace(replace(p.return_zipCode,'-','') +  p.return_address + p.return_address2,' ','')"

        rsget.Open sqlStr,dbget,1
		if Not(rsget.EOF or rsget.BOF) then
		    retReturnCode = rsget("returnCode")
        end if
        rsget.close

        if isNULL(retReturnCode) then Exit function
        if (retReturnCode="") then Exit function

        getLotteReturnDLVCode = CStr(retReturnCode)
    end function

    public function getDisptpcdName
        if (Fdisptpcd="10") then
            getDisptpcdName = "�Ϲ�"
        elseif (Fdisptpcd="11") then
            getDisptpcdName = "�귣��"
        elseif (Fdisptpcd="12") then
            getDisptpcdName = "<font color='blue'>����</font>"
        elseif (Fdisptpcd="99") then
            getDisptpcdName = "<font color='red'>�ű�</font>"
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

    public function getItemNameFormat()
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

    public function getLimitEa()
        dim ret
        ret = FLimitNo-FLimitSold

        if (ret<1) then ret=0
        getLimitEa = ret
    end function

    public function getLimitLotteEa()
        dim ret
        ret = FLimitNo-FLimitSold-5

        if (ret<1) then ret=0
        getLimitLotteEa = ret
    end function

	Function getLimitHtmlStr()
	    If IsNULL(FLimityn) Then Exit Function
	    If (FLimityn = "Y") Then
	        getLimitHtmlStr = "<font color=blue>����:"&getLimitEa&"</font>"
	    End if
	End Function

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
	public Function getLotteItemRegParameter(isEdit)
		dim strRst
		strRst = "subscriptionId=" & lotteAuthNo																						'�Ե����� ������ȣ	(*)
		if (isEdit) then
		   strRst = strRst & "&goods_req_no="&FLotteTmpGoodNo
		end if
		strRst = strRst & "&brnd_no=" & tenBrandCd																						'�귣���ڵ�			(*)
		strRst = strRst & "&goods_nm=" & Server.URLEncode(Trim(getItemNameFormat))																	'��ǰ��				(*)
		strRst = strRst & "&sch_kwd_1_nm=" & Server.URLEncode(getItemKeywordArray(0))													'Ű����1
		strRst = strRst & "&sch_kwd_2_nm=" & Server.URLEncode(getItemKeywordArray(1))													'Ű����2
		strRst = strRst & "&sch_kwd_3_nm=" & Server.URLEncode(getItemKeywordArray(2))													'Ű����3
		strRst = strRst & "&pmct_fix_cd=2"																					 			'������������		(*:����������)
		strRst = strRst & "&pur_shp_cd=2" 																								'��������			(*:�Ǹźи���)
		strRst = strRst & "&sale_shp_cd=10" 																							'�Ǹ�����			(*:����)
		strRst = strRst & "&sale_prc=" & cLng(GetRaiseValue(FSellCash/10)*10)																				'�ǸŰ�(���ǸŰ�)	(*) �Һ��ڰ��� ������ ����..?
		strRst = strRst & "&mrgn_rt=12" 																								'������				(*:11%) ==> 2013/01/01 12%
		strRst = strRst & "&pur_prc=" & cLng(FSellCash*0.88)																			'���ް�				(*)
		strRst = strRst & "&tdf_sct_cd=1" 																								'���鼼�ڵ�			(*:����)
		strRst = strRst & "&card_dsct_yn=Y"																					 			'�Ե�ī�����ο���	(*:���)
		if getLotteSellYn="Y" then										 																'�ǸŻ���			(*:10:�Ǹ�,20:ǰ��)
			strRst = strRst & "&sale_stat_cd=10"
		else
			strRst = strRst & "&sale_stat_cd=20"
		end if

		IF (FLimitYn="Y") then
		    strRst = strRst & "&inv_mgmt_yn=Y"
		    if FoptionCnt=0 then
		        strRst = strRst & "&inv_qty="&getLimitLotteEa()
		    end if
		ELSE
    		strRst = strRst & "&inv_mgmt_yn=N" 																								'����������		(*:��������)
    	END IF

		if FitemDiv="06" then
			strRst = strRst & "&add_choc_tp_cd_20=" & Server.URLEncode("�ֹ����ۻ�ǰ") 													'�����Է����ɼ�
		end if
		strRst = strRst & "&dlv_proc_tp_cd=1" 																							'�������			(*:����)
		strRst = strRst & "&box_pkg_yn=Y" 																								'���Box����		(*:���尡��)
		strRst = strRst & "&fut_msg_yn=N" 																								'�������忩��		(*:�Ұ�)
		strRst = strRst & "&shop_fwd_msg_yn=N"	 																						'��������			(*:������)
		strRst = strRst & "&dlv_mean_cd=10" 																							'��ۼ���			(*:�ù�)

    	strRst = strRst & getLotteGoodDLVDtParams
		strRst = strRst & "&dlvp_stn_grp_cd=01" 																						'��۰�������		(*:����)
		strRst = strRst & "&byr_age_lmt_cd=0" 																							'�����ڳ�������		(*:��ü)
		strRst = strRst & "&exch_rtgs_sct_cd=21" 																						'��ȯ��ǰ����		(*:�д㱳ȯ)
		strRst = strRst & "&dlv_polc_no=" & tenDlvCd																					'�����å��ȣ		(*:��ü)
		strRst = strRst & "&corp_dlvp_sn=" & getLotteReturnDLVCode 																			 			'��ǰ������ڵ�		(*:��������)
		strRst = strRst & "&dcom_asgn_rtgs_hdc_use_yn=Y"                                                                 ''��ǰ��ȯ �����ù� ��뿩��  dcom_asgn_rtgs_hdc_use_yn
		strRst = strRst & "&orpl_nm=" & Server.URLEncode(chkIIF(trim(Fsourcearea)="" or isNull(Fsourcearea),"��ǰ���� ����",Fsourcearea))	'������				(*)
		strRst = strRst & "&mfcp_nm=" & Server.URLEncode(chkIIF(trim(Fmakername)="" or isNull(Fmakername),"��ǰ���� ����",Fmakername))		'������				(*)
		strRst = strRst & "&img_url=" & Server.URLEncode(FbasicImage)																	'��ǥ�̹���URL		(*)
		strRst = strRst & "&attd_mtr_cont=" & Server.URLEncode(ForderComment)															'�ֹ��� ���ǻ���(-)
		'strRst = strRst & "&md_ntc_2_fcont=" & server.URLEncode("<img src=""http://fiximage.10x10.co.kr/web2008/etc/cs_info.jpg"">")	'MD����
		strRst = strRst & "&brnd_intro_cont=" & server.URLEncode("Design Your Life! ���ο� �ϻ��� ����� ������Ȱ�귣�� �ٹ�����")		'�귣�� ����
		strRst = strRst & "&corp_goods_no=" & Fitemid																					'��ü��ǰ��ȣ(���ٻ�ǰ�ڵ�)
		strRst = strRst & getLotteCateParamToReg()																						'MD��ǰ�� �� �ش� ����ī�װ�
		strRst = strRst & getLotteOptionParamToReg()																					'��ǰ�ɼ� ���� �� �ɼǳ���
		strRst = strRst & getLotteItemContParamToReg()																					'��ǰ�󼼼���		(*)
		strRst = strRst & getLotteAddImageParamToReg()																					'��ǰ �߰� �̹���
        strRst = strRst & getLotteItemInfoCdToReg()   ''����
		'��� ��ȯ
		getLotteItemRegParameter = strRst
	end Function

    public Function getLotteItemEditParameter2()
        dim strRst
        strRst = getLotteItemRegParameter(true)

        getLotteItemEditParameter2 = strRst
    end function

	'// ��ǰ���� �Ķ���� ����
	public Function getLotteItemEditParameter()
		dim strRst
		strRst = "subscriptionId=" & lotteAuthNo																						'�Ե����� ������ȣ	(*)
		strRst = strRst & "&goods_no=" & FLotteGoodNo																					'�Ե����� ��ǰ��ȣ	(*)
		strRst = strRst & "&brnd_no=" & tenBrandCd																						'�귣���ڵ�			(*)
		strRst = strRst & "&goods_nm=" & Server.URLEncode(Trim(getItemNameFormat))																	'��ǰ��				(*)
		''if (FItemid="443553") then strRst = strRst & "."
		strRst = strRst & "&sch_kwd_1_nm=" & Server.URLEncode(getItemKeywordArray(0))													'Ű����1
		strRst = strRst & "&sch_kwd_2_nm=" & Server.URLEncode(getItemKeywordArray(1))													'Ű����2
		strRst = strRst & "&sch_kwd_3_nm=" & Server.URLEncode(getItemKeywordArray(2))													'Ű����3
		strRst = strRst & "&pur_shp_cd=2" 																								'��������			(*:�Ǹźи���)

		''getLotteOptionParamToEdit �� ����
'		if getLotteSellYn="Y" then										 																'�ǸŻ���			(*:10:�Ǹ�,20:ǰ��)
'			strRst = strRst & "&sale_stat_cd=10"
'		else
'			strRst = strRst & "&sale_stat_cd=20"
'		end if

		''getLotteOptionParamToEdit �� ����
'		IF (FLimitYn="Y") then
'		    strRst = strRst & "&inv_mgmt_yn=Y"
'		    if FoptionCnt=0 then
'		        strRst = strRst & "&inv_qty="&getLimitLotteEa()
'		    end if
'		ELSE
'    		strRst = strRst & "&inv_mgmt_yn=N" 																								'����������		(*:��������)
'    	END IF

		if FitemDiv="06" then
		    ''���� �ʿ� �̰� ���� ����..? // �̰��� ��� �ֹ����� �ɼ��� ��������� �ʴµ�.
			'strRst = strRst & "&add_choc_tp_cd_20=" & Server.URLEncode("�ֹ����ۻ�ǰ") 													'�����Է����ɼ�
	        ''strRst = strRst & "&add_choc_tp_cd_20="
		end if
		strRst = strRst & "&dlv_proc_tp_cd=1" 																							'�������			(*:����)
		strRst = strRst & "&box_pkg_yn=Y" 																								'���Box����		(*:���尡��)
		strRst = strRst & "&fut_msg_yn=N" 																								'�������忩��		(*:�Ұ�)
		strRst = strRst & "&shop_fwd_msg_yn=N"	 																						'��������			(*:������)
		strRst = strRst & "&dlv_mean_cd=10" 																							'��ۼ���			(*:�ù�)

    	strRst = strRst & getLotteGoodDLVDtParams
		strRst = strRst & "&dlvp_stn_grp_cd=01" 																						'��۰�������		(*:����)
		strRst = strRst & "&byr_age_lmt_cd=0" 																							'�����ڳ�������		(*:��ü)
		strRst = strRst & "&exch_rtgs_sct_cd=21" 																						'��ȯ��ǰ����		(*:�д㱳ȯ)
		strRst = strRst & "&dlv_polc_no=" & tenDlvCd																					'�����å��ȣ		(*:��ü)
		strRst = strRst & "&corp_dlvp_sn=" & getLotteReturnDLVCode 																			 			'��ǰ������ڵ�		(*:��������)
		strRst = strRst & "&dcom_asgn_rtgs_hdc_use_yn=Y"                                                                                               ''��ǰ��ȯ �����ù� ��뿩�� 2013/02/26
		strRst = strRst & "&orpl_nm=" & Server.URLEncode(chkIIF(trim(Fsourcearea)="" or isNull(Fsourcearea),"��ǰ���� ����",Fsourcearea))	'������				(*)
		strRst = strRst & "&mfcp_nm=" & Server.URLEncode(chkIIF(trim(Fmakername)="" or isNull(Fmakername),"��ǰ���� ����",Fmakername))		'������				(*)
		strRst = strRst & "&img_url=" & Server.URLEncode(FbasicImage)																	'��ǥ�̹���URL		(*)
		'strRst = strRst & "&md_ntc_2_fcont=" & server.URLEncode("<img src=""http://fiximage.10x10.co.kr/web2008/etc/cs_info.jpg"">")	'MD����
		strRst = strRst & "&brnd_intro_cont=" & server.URLEncode("Design Your Life! ���ο� �ϻ��� ����� ������Ȱ�귣�� �ٹ�����")		'�귣�� ����
		if (Fitemid="409536") then
		    strRst = strRst & "&attd_mtr_cont="
		else
    		strRst = strRst & "&attd_mtr_cont=" & Server.URLEncode(ForderComment)															'�ֹ��� ���ǻ���
    	end if
		strRst = strRst & "&corp_goods_no=" & Fitemid																					'��ü��ǰ��ȣ(���ٻ�ǰ�ڵ�)
		strRst = strRst & getLotteItemContParamToReg()																					'��ǰ�󼼼���		(*)
		strRst = strRst & getLotteAddImageParamToReg()																					'��ǰ �߰� �̹���
        strRst = strRst & getLotteOptionParamToEdit()
        strRst = strRst & getLotteItemInfoCdToReg()																						'��ǰǰ������/2012-11-02������ ����
        ''strRst = strRst & getLotteCateParamToReg() ''''''''20120831

		'��� ��ȯ
		getLotteItemEditParameter = strRst
	end Function

    '// ���� ���� �Ķ���� ����
    public Function getLotteItemPriceEditParameter()
        ''http://openapidev.lotte.com/openapi/updateGoodsNmOpenApi.lotte?subscriptionId=[����Ű]&strGoodsNo=XXXXXX&strReqSalePrc=[��ǰ����]&strChgCausCont=[�������]
        dim strRst
        strRst = "subscriptionId=" & lotteAuthNo
        strRst = strRst & "&strGoodsNo=" & FLotteGoodNo
        strRst = strRst & "&strReqSalePrc=" & GetRaiseValue(FSellCash/10)*10

        ''strRst = strRst & "&mrgn_rt=12"
        ''strRst = strRst & "&pur_prc=" & cLng(FSellCash*0.88)
        ''strRst = strRst & "&strChgCausCont=" & Server.URLEncode("���ݺ���")

        getLotteItemPriceEditParameter = strRst
        ''rw  strRst
    end function

    ''//��ǰ�� ���� �Ķ���� ����
    public Function getLotteItemNameEditParameter()
        dim strRst
        strRst = "subscriptionId=" & lotteAuthNo
        strRst = strRst & "&strGoodsNo=" & FLotteGoodNo
        strRst = strRst & "&strGoodsNm=" & Server.URLEncode(Trim(getItemNameFormat))
        strRst = strRst & "&strMblGoodsNm=" & Server.URLEncode(Trim(getItemNameFormat))
        strRst = strRst & "&strChgCausCont=" & Server.URLEncode("api ��ǰ�� ����")
        getLotteItemNameEditParameter = strRst
    end function

    public function getLotteGoodDLVDtParams()
        dim strRst
        strRst = ""
        if (FtenCateLarge="055") or (FtenCateLarge="040") then ''����/�к긯 15�Ϸ�
		    strRst = strRst & "&dlv_goods_sct_cd=03"
		    strRst = strRst & "&dlv_dday=15"
		elseif (FtenCateLarge="080") or (FtenCateLarge="100") then  ''���/���̺� 5��
		    strRst = strRst & "&dlv_goods_sct_cd=03" 																						'��ۻ�ǰ����		(*:�ֹ�����03)
		    strRst = strRst & "&dlv_dday=5"
		elseif ((FtenCateLarge="045") and (FtenCateMid="001" or FtenCateMid="004")) then  ''����/��Ȱ> ��/�̺Ҽ��� or �ֹ���� 10�� - ���ƾ���û 2013/01/22
		    strRst = strRst & "&dlv_goods_sct_cd=03"
		    strRst = strRst & "&dlv_dday=10"
		elseif ((FtenCateLarge="025") and (FtenCateMid="107")) then  ''������ > ��Ÿ ����Ʈ��� ���̽�  10�� - ���ƾ���û 2013/01/22
		    strRst = strRst & "&dlv_goods_sct_cd=03"
		    strRst = strRst & "&dlv_dday=10"
	    elseif ((FtenCateLarge="050") and (FtenCateMid="777")) then   ''Ȩ/���� > �ſ�   - ���񾾿�û 2013/03/08
	        strRst = strRst & "&dlv_goods_sct_cd=03"
		    strRst = strRst & "&dlv_dday=10"
		elseif ((FtenCateLarge="045") and (FtenCateMid="002") and (FtenCateSmall="001")) then    ''HOME > ����/��Ȱ > ����/������ǰ > ������ 			�ֹ�����15�� 045&cdm=002&cds=001
		    strRst = strRst & "&dlv_goods_sct_cd=03"
		    strRst = strRst & "&dlv_dday=15"
		elseif ((FtenCateLarge="045") and (FtenCateMid="002") and (FtenCateSmall="002")) then    ''HOME > ����/��Ȱ > ����/������ǰ > ƴ��������			�ֹ�����10��
            strRst = strRst & "&dlv_goods_sct_cd=03"
		    strRst = strRst & "&dlv_dday=10"
        elseif ((FtenCateLarge="045") and (FtenCateMid="002") and (FtenCateSmall="005")) then    ''HOME > ����/��Ȱ > ����/������ǰ > �������� 			�ֹ�����10��
            strRst = strRst & "&dlv_goods_sct_cd=03"
		    strRst = strRst & "&dlv_dday=10"
		elseif ((FtenCateLarge="045") and (FtenCateMid="006") and (FtenCateSmall="001")) then    ''HOME > ����/��Ȱ > ���ڼ��� > ���ڽ� 				�ֹ�����10��
            strRst = strRst & "&dlv_goods_sct_cd=03"
		    strRst = strRst & "&dlv_dday=10"
        elseif ((FtenCateLarge="045") and (FtenCateMid="006") and (FtenCateSmall="007")) then    ''HOME > ����/��Ȱ > ���ڼ��� > �������ڽ� 			               �ֹ�����10��
            strRst = strRst & "&dlv_goods_sct_cd=03"
		    strRst = strRst & "&dlv_dday=10"
        elseif ((FtenCateLarge="050") and (FtenCateMid="060") and (FtenCateSmall="070")) then    ''HOME > Ȩ/���� > ��ǰ�ڽ�/�ٱ��� > �������ڽ�			�ֹ�����10�� cdl=050&cdm=060&cds=070
            strRst = strRst & "&dlv_goods_sct_cd=03"
		    strRst = strRst & "&dlv_dday=10"
        elseif ((FtenCateLarge="110") and (FtenCateMid="090") and (FtenCateSmall="040")) then    ''HOME > ����ä�� > DIY > �����θ���� 				�ֹ�����10�� 110&cdm=090&cds=040
            strRst = strRst & "&dlv_goods_sct_cd=03"
		    strRst = strRst & "&dlv_dday=10"
        elseif ((FtenCateLarge="045") and (FtenCateMid="010")) then   ''����/��Ȱ > �����μ���  - ���񾾿�û 2013/03/08
	        strRst = strRst & "&dlv_goods_sct_cd=03"
		    strRst = strRst & "&dlv_dday=10"
'		elseif (FtenCateLarge="025")  then  ''������ 10�� - ���񾾿�û 2013/01/17
'		    strRst = strRst & "&dlv_goods_sct_cd=03" 																						'��ۻ�ǰ����		(*:�ֹ�����03)
'		    strRst = strRst & "&dlv_dday=10"
		elseif ((FitemDiv="06") or (FitemDiv="16")) then    ''�ֹ�(��)���ۻ�ǰ
		    strRst = strRst & "&dlv_goods_sct_cd=03"
		    if (FrequireMakeDay>7) then
		        strRst = strRst & "&dlv_dday="&CStr(FrequireMakeDay)
		    elseif (FrequireMakeDay<1) then
		        strRst = strRst & "&dlv_dday=7"
		    else
		        strRst = strRst & "&dlv_dday="&(FrequireMakeDay+1)
		    end if
		else
		    strRst = strRst & "&dlv_goods_sct_cd=01" 																						'��ۻ�ǰ����		(*:�Ϲݻ�ǰ)
    		strRst = strRst & "&dlv_dday=3" 																								'��۱���			(*:3���̳�)
    	end if
    	getLotteGoodDLVDtParams = strRst
    end function

	'// ��ǰ���: MD��ǰ�� �� ���� ī�װ� �Ķ���� ����(��ǰ��Ͽ�)
	public function getLotteCateParamToReg()
		dim strSql, strRst, i, ogrpCode
		strSql = "Select top 6 c.groupCode, m.dispNo, c.disptpcd "
		strSql = strSql & " from db_item.dbo.tbl_lotte_cate_mapping as m "
		strSql = strSql & " 	join db_temp.dbo.tbl_lotte_Category as c "
		strSql = strSql & " 		on m.DispNO=c.DispNO "
		strSql = strSql & " where tenCateLarge='" & FtenCateLarge & "' "
		strSql = strSql & " 	and tenCateMid='" & FtenCateMid & "' "
		strSql = strSql & " 	and tenCateSmall='" & FtenCateSmall & "' "
	''strSql = strSql & " 	and c.disptpcd<>'99'"
	    strSql = strSql & " 	and c.isusing='Y'"
		strSql = strSql & " order by (CASE WHEN c.disptpcd='12' THEN 'ZZ' ELSE c.disptpcd END) desc"           ''''//�������� �⺻ ī�װ���..
		rsget.Open strSql,dbget,1

		if Not(rsget.EOF or rsget.BOF) then
		    ogrpCode = rsget("groupCode")
			strRst = "&md_gsgr_no=" & ogrpCode				            'md��ǰ�� �ڵ� (md��ǰ�� �ڵ� ���� ī�װ��� ��ϰ���)
            ''if (rsget("groupCode")="1598") then
            ''    strRst = "&md_gsgr_no=" & "1596"
            ''end if

			i=0
			Do until rsget.EOF
				''if i=0 then
				''	strRst = strRst & "&disp_no=" & rsget("dispNo")		'�⺻ ����ī�װ�
				''else
				''	strRst = strRst & "&disp_no_b=" & rsget("dispNo") 	'�߰� ����ī�װ�
				''end if

				if (rsget("disptpcd")="12") then                        ''������ ī�װ��� �⺻���� �϶��.. /2012/06/14
				    strRst = strRst & "&disp_no=" & rsget("dispNo")		'�⺻ ����ī�װ�
				else
				    IF (ogrpCode=rsget("groupCode")) then
    				    strRst = strRst & "&disp_no_b=" & rsget("dispNo") 	'�߰� ����ī�װ�
    				End IF
			    end if
				rsget.MoveNext
				i=i+1
			Loop
		end if

		rsget.Close

		getLotteCateParamToReg = strRst
	end function

    public function getLotteOptionParamToEditNew()
        dim ret : ret = ""
        dim i
        dim strSql, arrRows, iErrStr
        dim isOptionExists
        dim mayOptionCnt : mayOptionCnt = 0
        dim item_sale_stat_cd,outmalloptcode, optLimit
        dim item_noStr, item_sale_stat_cdStr, inv_qtyStr
        dim optValidExists : optValidExists = false
        dim preMaxOutmalloptcode : preMaxOutmalloptcode=-1

        strSql = "exec db_item.dbo.sp_Ten_OutMall_optEditParamList_Lotte '"&CMallName&"'," & FItemid
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
            mayOptionCnt = UBound(ArrRows,2)
            mayOptionCnt = mayOptionCnt + 1
        end if

        ''if (FoptionCnt<>mayOptionCnt) then
        if (FregedOptCnt<>mayOptionCnt) then
            ''�����ȸ.
            rw "FregedOptCnt="&FregedOptCnt&".."&"mayOptionCnt="&mayOptionCnt

            CALL LotteOneItemCheckStock(Fitemid,iErrStr)
        end if

        ret = ""
        if (Not isOptionExists) then
            IF (FLimitYn="Y") then
    		    ret = ret & "&inv_mgmt_yn=Y"
    		    ret = ret & "&item_no=0"
    		    ret = ret & "&item_sale_stat_cd=10"
    		    ret = ret & "&inv_qty="&getLimitLotteEa()                                                                                   '������ �� ��� ��ǰ�� �־��..
    		ELSE
        		ret = ret & "&inv_mgmt_yn=N" 																								'����������		(*:��������)
        	END IF
        else
            if FLimitYn="Y" then
			    ret = ret&"&inv_mgmt_yn=Y"
			else
			    ret = ret&"&inv_mgmt_yn=N"
		    end if

            For i =0 To UBound(ArrRows,2)
                item_sale_stat_cd = "10"    ''10:�Ǹ�����,20:ǰ��,30:�Ǹ�����
			    outmalloptcode = ArrRows(2,i)
			    if IsNULL(outmalloptcode) then
			        outmalloptcode=preMaxOutmalloptcode+1
			    else
			        if (preMaxOutmalloptcode>outmalloptcode) then
			            preMaxOutmalloptcode=preMaxOutmalloptcode
			        else
			            preMaxOutmalloptcode=outmalloptcode
			        end if
			    end if

			    if FLimitYn="Y" then
			        optLimit = ArrRows(4,i)-5
			    else
			        optLimit = "50"
			    end if
			    if (optLimit<1) then optLimit=0
			    if (ArrRows(6,i)="N") or (ArrRows(7,i)="N") then item_sale_stat_cd="20"
			    if (FLimitYn="Y") and (optLimit<1) then item_sale_stat_cd="20"

			    if ((ArrRows(11,i)="1") and (ArrRows(12,i)="1")) or (ArrRows(13,i)="1") then
			        optLimit=0
			        item_sale_stat_cd="20"
			    end if

			    item_noStr = item_noStr & "&item_no="&outmalloptcode
			    item_sale_stat_cdStr = item_sale_stat_cdStr & "&item_sale_stat_cd="&item_sale_stat_cd
			    inv_qtyStr = inv_qtyStr & "&inv_qty="&optLimit

			    if (item_sale_stat_cd = "10") then optValidExists=TRUE
            next

            ret = ret&item_noStr&item_sale_stat_cdStr&inv_qtyStr
        end if

	    ''rw ret
	    if (Not isOptionExists) then   ''�ɼ��� ������.
    		if getLotteSellYn="Y" then										 																'�ǸŻ���			(*:10:�Ǹ�,20:ǰ��)
    			ret = ret & "&sale_stat_cd=10"
    		else
    		    FSellyn="N"
    			ret = ret & "&sale_stat_cd=20"
    		end if
    	else
    	    if (optValidExists) and (getLotteSellYn="Y") then  ''�Ǹ��� �̰� �ɼ� �ǸŰ����̸�.
    	        ret = ret & "&sale_stat_cd=10"
    	    else
    	        rw "None Exists Valid Option"
    	        FSellyn="N"
    	        ret = ret & "&sale_stat_cd=20"
    	    end if
        end if

        getLotteOptionParamToEditNew = ret

    end function

    '// ��ǰ����: �ɼ� �Ķ���� ����(��ǰ������) :: �� �ɼ��� ���� �Ǿ� �־�� ��..
    ''�ɼǸ� ������ �ȵ�.. ���� API �ִ��� ����.
    public function getLotteOptionParamToEdit()
        if (TRUE) or (FItemID="138371") or (FItemID="295139") or (FItemID="830724") or (FItemID="830728") or (FItemID="830816") or (FItemID="830795") then
            getLotteOptionParamToEdit=getLotteOptionParamToEditNew
            exit function
        end if

        dim ret : ret = ""
        dim strSql
        dim optYn, item_noStr, item_sale_stat_cdStr, inv_qtyStr, optLimit, item_sale_stat_cd, outmalloptcode
        dim optValidExists : optValidExists = FALSE
        ''getLotteOptionParamToEdit = ret
        ''Exit function

        ''�ɼ��� ������ �Ե� ����� �ɼ� ������ ������..
        if (FoptionCnt>0) then
            ''�����ȸ.
        end if

'        if (FItemid=379468) then
'            ret = ret&"&inv_mgmt_yn=Y"
'            ret = ret&"&item_no=0&item_no=1&item_no=2"
'            ret = ret&"&item_sale_stat_cd=20&item_sale_stat_cd=20&item_sale_stat_cd=10"
'            ret = ret&"&inv_qty=0&inv_qty=0&inv_qty=2"
'        end if
'
'        if (FItemid=330185) then
'            ret = ret&"&inv_mgmt_yn=Y"
'            ret = ret&"&item_no=0&item_no=1&item_no=2&item_no=3&item_no=4&item_no=5&item_no=6&item_no=7&item_no=8&item_no=9"
'            ret = ret&"&item_sale_stat_cd=10&item_sale_stat_cd=10&item_sale_stat_cd=10&item_sale_stat_cd=10&item_sale_stat_cd=10&item_sale_stat_cd=10&item_sale_stat_cd=10&item_sale_stat_cd=10&item_sale_stat_cd=10&item_sale_stat_cd=10"
'            ret = ret&"&inv_qty=10&inv_qty=10&inv_qty=10&inv_qty=10&inv_qty=5&inv_qty=5&inv_qty=10&inv_qty=10&inv_qty=10&inv_qty=10"
'        end if

        ''553443

        ''&inv_mgmt_yn=Y&item_no=0&item_no=1&item_no=2&item_no=3&item_no=4&item_no=5&item_no=6&item_no=7&item_no=8&item_no=9&item_no=10&item_no=11&item_no=12&item_no=13&item_sale_stat_cd=10&item_sale_stat_cd=10&item_sale_stat_cd=10&item_sale_stat_cd=10&item_sale_stat_cd=10&item_sale_stat_cd=10&item_sale_stat_cd=10&item_sale_stat_cd=10&item_sale_stat_cd=10&item_sale_stat_cd=20&item_sale_stat_cd=10&item_sale_stat_cd=10&item_sale_stat_cd=10&item_sale_stat_cd=10&inv_qtyStr=5&inv_qtyStr=6&inv_qtyStr=9&inv_qtyStr=7&inv_qtyStr=11&inv_qtyStr=5&inv_qtyStr=8&inv_qtyStr=3&inv_qtyStr=27&inv_qtyStr=0&inv_qtyStr=13&inv_qtyStr=11&inv_qtyStr=4&inv_qtyStr=10


        strSql = "Select o.itemoption, (CASE WHEN convert(varchar(18),o.optionTypeName)<>o.optionTypeName THEN '�ɼǼ���' ELSE o.optionTypeName END) as optionTypeName, o.optionname, (o.optlimitno-o.optlimitsold) as optLimit, o.optlimityn, o.isUsing, o.optsellyn "
		strSql = strSql & " ,R.outmalloptcode"
		strSql = strSql & " ,(CASE WHEN  R.itemoption is NULL THEN 0 ELSE 1 END) as preged"
		strSql = strSql & " From [db_item].[dbo].tbl_item_option o"
		strSql = strSql & "     left join [db_item].[dbo].tbl_OutMall_regedoption R"
		strSql = strSql & "     on o.itemid=R.itemid"
		strSql = strSql & "     and o.itemoption=R.itemoption"
		strSql = strSql & "     and R.itemoption<>''"
		strSql = strSql & "     and R.mallid='"&CMALLNAME&"'"
		strSql = strSql & " where o.itemid=" & FItemid
		strSql = strSql & " 	and o.optaddprice=0 "                     '''�߰��ݾ� �Ұ�.
		strSql = strSql & " 	and ((o.isUsing='Y') or (R.itemid is Not NULL)) "  '
		strSql = strSql & " 	and R.outmalloptcode is Not NULL"         '''��� �ɼǸ� ���� ����. :: ���ʿ� ��� �ɼ��� ����ؾ��ҵ�..
		strSql = strSql & " and isNULL(R.outmallsellyn,'')<>'X'"

		'''strSql = strSql & " 	and (optlimityn='N' or (optlimityn='Y' and optlimitno-optlimitsold>="&CMAXLIMITSELL&")) "

		rsget.Open strSql,dbget,1

		if Not(rsget.EOF or rsget.BOF) then
			optYn = "Y"
			if FLimitYn="Y" then
			    ret = ret&"&inv_mgmt_yn=Y"
			else
			    ret = ret&"&inv_mgmt_yn=N"
		    end if

			Do until rsget.EOF
			    item_sale_stat_cd = "10"    ''10:�Ǹ�����,20:ǰ��,30:�Ǹ�����
			    outmalloptcode = rsget("outmalloptcode")
			    optLimit = rsget("optLimit")-5

			    if (optLimit<1) then optLimit=0
			    if (rsget("isUsing")="N") or (rsget("optsellyn")="N") then item_sale_stat_cd="20"
			    if (FLimitYn="Y") and (optLimit<1) then item_sale_stat_cd="20"

			    item_noStr = item_noStr & "&item_no="&outmalloptcode
			    item_sale_stat_cdStr = item_sale_stat_cdStr & "&item_sale_stat_cd="&item_sale_stat_cd
			    inv_qtyStr = inv_qtyStr & "&inv_qty="&optLimit

			    if (item_sale_stat_cd = "10") then optValidExists=TRUE
			    rsget.MoveNext
			Loop
		end if
		rsget.Close


		if optYn <> "Y" then
    		IF (FLimitYn="Y") then
    		    ret = ""
    		    ret = ret & "&inv_mgmt_yn=Y"
    		    ret = ret & "&item_no=0"
    		    ret = ret & "&item_sale_stat_cd=10"
    		    ret = ret & "&inv_qty="&getLimitLotteEa()
    		ELSE
        		ret = ret & "&inv_mgmt_yn=N" 																								'����������		(*:��������)
        	END IF

		else
		    ret = ret&item_noStr&item_sale_stat_cdStr&inv_qtyStr
	    end if

	    ''rw ret
	    if optYn <> "Y" then   ''�ɼ��� ������.
    		if getLotteSellYn="Y" then										 																'�ǸŻ���			(*:10:�Ǹ�,20:ǰ��)
    			ret = ret & "&sale_stat_cd=10"
    		else
    			ret = ret & "&sale_stat_cd=20"
    		end if
    	else
    	    if (optValidExists) and (getLotteSellYn="Y") then  ''�Ǹ��� �̰� �ɼ� �ǸŰ����̸�.
    	        ret = ret & "&sale_stat_cd=10"
    	    else
    	        ret = ret & "&sale_stat_cd=20"
    	    end if
        end if

        getLotteOptionParamToEdit = ret
    end function

	'// ��ǰ���: �ɼ� �Ķ���� ����(��ǰ��Ͽ�)
	public function getLotteOptionParamToReg()
		dim strSql, strRst, i, optYn, optNm, optDc, chkMultiOpt, optLimit
		chkMultiOpt = false
		optYn = "N"

		if FoptionCnt>0 then
			'// ���߿ɼ��� ��
			'#�ɼǸ� ����
			strSql = "exec [db_item].[dbo].sp_Ten_ItemOptionMultipleTypeList " & FItemid
	        rsget.CursorLocation = adUseClient
			rsget.CursorType = adOpenStatic
			rsget.LockType = adLockOptimistic
	        rsget.Open strSql, dbget

			optNm = ""
			if Not(rsget.EOF or rsget.BOF) then
				chkMultiOpt = true
				optYn = "Y"
				Do until rsget.EOF
					optNm = optNm & Replace(db2Html(rsget("optionTypeName")),":","")
					rsget.MoveNext
					if Not(rsget.EOF) then optNm = optNm & ":"
				Loop
			end if
			rsget.Close

			'#�ɼǳ��� ����
			if chkMultiOpt then
				strSql = "Select optionname, (optlimitno-optlimitsold) as optLimit "
				strSql = strSql & " From [db_item].[dbo].tbl_item_option "
				strSql = strSql & " where itemid=" & FItemid
				strSql = strSql & " 	and isUsing='Y' and optsellyn='Y' "
				strSql = strSql & " 	and optaddprice=0 "
				'''strSql = strSql & " 	and (optlimityn='N' or (optlimityn='Y' and optlimitno-optlimitsold>="&CMAXLIMITSELL&")) " ''�ϴ� �Է�

				rsget.Open strSql,dbget,1

				optDc = ""
				if Not(rsget.EOF or rsget.BOF) then
					Do until rsget.EOF
					    optLimit = rsget("optLimit")
					    optLimit = optLimit-5
					    if (optLimit<1) then optLimit=0
						optDc = optDc & Replace(Replace(db2Html(rsget("optionname")),":",""),"'","") & "," & optLimit
						rsget.MoveNext
						if Not(rsget.EOF) then optDc = optDc & ":"
					Loop
				end if
				rsget.Close
			end if


			'// ���Ͽɼ��� ��
			if Not(chkMultiOpt) then
				strSql = "Select optionTypeName, optionname, (optlimitno-optlimitsold) as optLimit "
				strSql = strSql & " From [db_item].[dbo].tbl_item_option "
				strSql = strSql & " where itemid=" & FItemid
				strSql = strSql & " 	and isUsing='Y' and optsellyn='Y' "
				strSql = strSql & " 	and optaddprice=0 "
				''strSql = strSql & " 	and (optlimityn='N' or (optlimityn='Y' and optlimitno-optlimitsold>="&CMAXLIMITSELL&")) "
				rsget.Open strSql,dbget,1

				if Not(rsget.EOF or rsget.BOF) then
					optYn = "Y"
					if db2Html(rsget("optionTypeName"))<>"" then
						optNm = Replace(db2Html(rsget("optionTypeName")),":","")
					else
						optNm = "�ɼ�"
					end if
					Do until rsget.EOF
					    optLimit = rsget("optLimit")
					    optLimit = optLimit-5
					    if (optLimit<1) then optLimit=0
						optDc = optDc & Replace(Replace(Replace(db2Html(rsget("optionname")),":",""),",",""),"'","") & "," & optLimit
						rsget.MoveNext
						if Not(rsget.EOF) then optDc = optDc & ":"
					Loop
				end if
				rsget.Close
			end if
		end if

		strRst = strRst & "&item_mgmt_yn=" & optYn						'��ǰ��������(�ɼ�)
		strRst = strRst & "&opt_nm=" & server.URLEncode(optNm)			'�ɼǸ�
		strRst = strRst & "&item_list=" & server.URLEncode(optDc)		'�ɼǻ�

		getLotteOptionParamToReg = strRst
	end function

	'// ��ǰ���: ��ǰ���� �Ķ���� ����(��ǰ��Ͽ�)
	public function getLotteItemContParamToReg()
		dim strRst, strSQL

		strRst = Server.URLEncode("<div align=""center"">")

		'#�⺻ ��ǰ����
		Select Case FUsingHTML
			Case "Y"
				strRst = strRst & Server.URLEncode(oLotteitem.FItemList(i).Fitemcontent & "<br>")
			Case "H"
				strRst = strRst & Server.URLEncode(nl2br(oLotteitem.FItemList(i).Fitemcontent) & "<br>")
			Case Else
				strRst = strRst & Server.URLEncode(nl2br(ReplaceBracket(oLotteitem.FItemList(i).Fitemcontent)) & "<br>")
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
					strRst = strRst & Server.URLEncode("<img src=""http://webimage.10x10.co.kr/item/contentsimage/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_400") & """ border=""0""><br>")
				end if
				rsget.MoveNext
			Loop
		end if

		rsget.Close

		'#�⺻ ��ǰ �����̹���
		if ImageExists(FmainImage) then strRst = strRst & Server.URLEncode("<img src=""" & FmainImage & """ border=""0""><br>")
		if ImageExists(FmainImage2) then strRst = strRst & Server.URLEncode("<img src=""" & FmainImage2 & """ border=""0""><br>")

		'#��� ���ǻ���
		strRst = strRst & Server.URLEncode("<br><img src=""http://fiximage.10x10.co.kr/web2008/etc/cs_info.jpg"">")

		strRst = strRst & Server.URLEncode("</div>")

		getLotteItemContParamToReg = "&dtl_info_fcont=" & strRst

		''660877 db_item.dbo.tbl_OutMall_etcLink ���� �� ���� �����ϸ� ����
		strSQL = ""
		strSQL = strSQL & " SELECT itemid, mallid, linkgbn, textVal " & VBCRLF
		strSQL = strSQL & " FROM db_item.dbo.tbl_OutMall_etcLink " & VBCRLF
		strSQL = strSQL & " where mallid='"&CMALLNAME&"' and linkgbn = 'contents' and itemid = '"&Fitemid&"' " & VBCRLF
		rsget.Open strSQL, dbget
		if Not(rsget.EOF or rsget.BOF) then
			strRst = Server.URLEncode(""&rsget("textVal")&"")
			strRst = Server.URLEncode("<div align=""center"">") & strRst & Server.URLEncode("<br><img src=""http://fiximage.10x10.co.kr/web2008/etc/cs_info.jpg""></div>")
			getLotteItemContParamToReg = "&dtl_info_fcont=" & strRst
		End If
		rsget.Close
		''660877 db_item.dbo.tbl_OutMall_etcLink ���� �� ���� �����ϸ� ��

'		if (FItemID="502049") or (FItemID="660877") then
'		    ''getLotteItemContParamToReg = "&dtl_info_fcont="&Server.URLEncode("<div></div>")     ''
'		    getLotteItemContParamToReg = ""                                           '' �������Ϸ��� �Ķ���� ���� �Ǵº�
'		end if
	end function

	'// ��ǰ���: ��ǰ�߰��̹��� �Ķ���� ����(��ǰ��Ͽ�)
	public function getLotteAddImageParamToReg()
		dim strRst, strSQL, i

		'# �߰� ��ǰ �����̹��� ����
		strSQL = "exec [db_item].[dbo].sp_Ten_CategoryPrd_AddImage @vItemid =" & Fitemid
		rsget.CursorLocation = adUseClient
		rsget.CursorType=adOpenStatic
		rsget.Locktype=adLockReadOnly
		rsget.Open strSQL, dbget

		if Not(rsget.EOF or rsget.BOF) then
			for i=1 to rsget.RecordCount
				if rsget("imgType")="0" then
					strRst = strRst & "&img_url" & i & "=" & Server.URLEncode("http://webimage.10x10.co.kr/image/add" & rsget("gubun") & "/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_400"))
				end if
				rsget.MoveNext
				if i>=5 then Exit For
			next
		end if

		rsget.Close

		getLotteAddImageParamToReg = strRst
	end function

    ''���꿩��
    public Function isMadeInKorea(iVal)
        isMadeInKorea = False
        if isNULL(iVal) then Exit Function
        iVal = Trim(iVal)

        if (iVal="�ѱ�") or (iVal="���ѹα�") or (iVal="KOREA") or (iVal="KOREA �ѱ�") or (iVal="KOREA ������") then
            isMadeInKorea = True
        end if

        if (iVal="����") or (iVal="��������") or (iVal="�ѱ�OEM") or (iVal="��������(�ѱ�)") or (iVal="����") then
            isMadeInKorea = True
        end if

        if (iVal="�ѱ� / ������Ʈ") then
            isMadeInKorea = True
        end if
    end Function

	'2012/11/02 ������ ���� ��ǰǰ������ �Ķ��Ÿ
	Public Function getLotteItemInfoCdToReg()
		Dim strRst, strSQL
		Dim anjunInfo, mallinfoDiv, mallinfoCdAll,mallinfoCd, infoCDVal, psourceArea
        Dim bufTxt : bufTxt=""

        ''������������
		If (Fsafetyyn="Y" and FsafetyDiv<>0) Then
			anjunInfo = anjunInfo & "&sft_cert_tgt_yn=Y"
			If (FsafetyDiv=10) Then
				anjunInfo = anjunInfo & "&kps_1_no="&Server.URLEncode(FsafetyNum)
			Elseif (FsafetyDiv=20) Then
				anjunInfo = anjunInfo & "&kps_1_no="
				anjunInfo = anjunInfo & "&kps_2_no="&Server.URLEncode(FsafetyNum)
			Elseif (FsafetyDiv=30) Then
				anjunInfo = anjunInfo & "&kps_1_no="
				anjunInfo = anjunInfo & "&kps_3_no="&Server.URLEncode(FsafetyNum)
			Elseif (FsafetyDiv=40) Then
				anjunInfo = anjunInfo & "&kps_1_no="
				anjunInfo = anjunInfo & "&kps_4_no="&Server.URLEncode(FsafetyNum)
			Elseif (FsafetyDiv=50) Then
				anjunInfo = anjunInfo & "&kps_1_no="
				anjunInfo = anjunInfo & "&kps_5_no="&Server.URLEncode(FsafetyNum)
			Else
				anjunInfo = ""
			End if
		End If

		strSql = ""
		strSql = strSql & " SELECT top 100 M.* " & vbcrlf
		strSql = strSql & " ,isNULL(CASE WHEN M.infocd='00000' then 'N' " '''db_item.dbo.[fn_LotteCom_SaftyFormat](IC.safetyyn,IC.safetyDiv,IC.safetyNum) " & vbcrlf
		strSql = strSql & " 	  WHEN M.infocd='99999' then M.infoETC"& vbcrlf
		strSql = strSql & " 	  WHEN M.infocd='00005' then '�ش����'"& vbcrlf
		strSql = strSql & " 	  WHEN M.infocd='00006' then '��ǰ �� ����'"& vbcrlf
		strSql = strSql & " 	  WHEN c.infotype='C' and F.chkDiv='N' THEN '�ش����' " & vbcrlf
		strSql = strSql & " 	  WHEN c.infotype='P' THEN replace(c.infoDesc,'1644-6030','1644-6035') " & vbcrlf
		'2014-07-14 16:07 ������ �ϴ� �߰�. ���Ƹ� ��û "ǰ����������" �տ� �ؽ�Ʈ ���� �߰�
		strSql = strSql & " 	  WHEN (c.infoItemName= 'ǰ����������') and (isnull(ET.itemid, '') <> '') THEN 'ǰ���������� ���ù� �� �Һ��� �����ذ� ���ؿ� ����, ' + F.infocontent + isNULL(F2.infocontent,'') " & vbcrlf
		strSql = strSql & "  ELSE F.infocontent + isNULL(F2.infocontent,'') " & vbcrlf
		strSql = strSql & "  END,'') as infoCDVal " & vbcrlf
		strSql = strSql & "  , L.shortVal, ET.itemid as ETText " & vbcrlf
		strSql = strSql & " FROM db_item.dbo.tbl_OutMall_infoCodeMap M " & vbcrlf
		strSql = strSql & " Join db_item.dbo.tbl_item_contents IC " & vbcrlf
		strSql = strSql & " on IC.infoDiv=M.mallinfoDiv " & vbcrlf
		strSql = strSql & " left Join db_item.dbo.tbl_item_infoCode c " & vbcrlf
		strSql = strSql & " on M.infocd=c.infocd " & vbcrlf
		strSql = strSql & " Left Join db_item.dbo.tbl_item_infoCont F " & vbcrlf
		strSql = strSql & " on M.infocd=F.infocd and F.itemid=" & FItemid & vbcrlf
		strSql = strSql & " left join db_item.dbo.tbl_item_infoCont F2 " & vbcrlf
		strSql = strSql & " on M.infocdAdd=F2.infocd and F2.itemid=" & FItemid & vbcrlf
		strSql = strSql & " left join db_item.dbo.tbl_OutMall_etcLink as L " & vbcrlf
		strSql = strSql & " on ((L.mallid = M.mallid) OR (isnull(L.mallid,'') = '')) and L.itemid =" & FItemid & vbcrlf
		strSql = strSql & " and L.linkgbn='infoDiv21Lotte'" & vbcrlf
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_Outmall_etcTextItem as ET on IC.itemid = ET.itemid and ET.mallid = '"&CMALLNAME&"' " & vbcrlf
		strSql = strSql & " where M.mallid = '"&CMALLNAME&"' AND IC.itemid=" & FItemid
		strSql = strSql & " order by M.mallinfoCd"
'response.write strSql
'response.end
		rsget.Open strSql,dbget,1
		Dim mat_name, mat_percent, mat_place, material

		psourceArea = ""
		If Not(rsget.EOF or rsget.BOF) then
			mallinfoDiv = "&ec_goods_artc_cd="&Server.URLEncode(rsget("mallinfoDiv"))
			Do until rsget.EOF
			    mallinfoCd = rsget("mallinfoCd")
			    infoCDVal  = rsget("infoCDVal")

			    IF mallinfoCd="2105" Then
			    	If isNull(rsget("shortVal")) = FALSE Then
						material = Split(rsget("shortVal"),"!!^^")
						mat_name	= material(0)
						mat_percent	= material(1)
						mat_place	= material(2)

				        bufTxt = "&mmtr_nm="&Server.URLEncode(""&mat_name&"")
				        bufTxt = bufTxt&"&cmps_rt="&Server.URLEncode(""&mat_percent&"")
				        bufTxt = bufTxt&"&mmtr_orpl_nm="&Server.URLEncode(""&mat_place&"")
			    	End If
			    Else
			        if (mallinfoCd="3503") then
			            '' ǰ��(35 ��Ÿ�ΰ��) �������� ����
			            psourceArea = rsget("infoCDVal")
			        end if

			        if (mallinfoCd="3504") then  ''�����ΰ�� �ش�������� �־�޶�� ��. �����ΰ�츸 ǥ�� // 2014-07-14 14:30 3504�� �� "�귣�����,�ش����" ������ ó����û
			            if (isMadeInKorea(psourceArea)) then
			                infoCDVal = Fmakername & ",�ش����"
    			        end if
			        end if
        		    mallinfoCdAll = mallinfoCdAll & "&"&mallinfoCd&"=" &Server.URLEncode(infoCDVal)
    			End If
				rsget.MoveNext
			Loop
		End If
		rsget.Close
		strRst = anjunInfo & mallinfoDiv & mallinfoCdAll & bufTxt
		getLotteItemInfoCdToReg = strRst
	End Function

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
				''strSql = strSql & " 	and (optlimityn='N' or (optlimityn='Y' and optlimitno-optlimitsold>="&CMAXLIMITSELL&")) "
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
				''strSql = strSql & " 	and (optlimityn='N' or (optlimityn='Y' and optlimitno-optlimitsold>="&CMAXLIMITSELL&")) "
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
	public function getLotteSellYn()
		'�ǸŻ��� (10:�Ǹ�����, 20:ǰ��)
		if FsellYn="Y" and FisUsing="Y" then
			if FLimitYn="N" or (FLimitYn="Y" and FLimitNo-FLimitSold>=CMAXLIMITSELL) then
				getLotteSellYn = "Y"
			else
				getLotteSellYn = "N"
			end if
		else
			getLotteSellYn = "N"
		end if
	end Function

	'// �Ե����� ��ϻ��� ��ȯ
	public function getLotteItemStatCd()
		Select Case FLotteStatCd
		    Case "00"
				getLotteItemStatCd = "��Ͽ���"
			Case "10"
				getLotteItemStatCd = "���۽õ�"         ''��Žõ�( �� �ӽõ��)
			Case "20"
				getLotteItemStatCd = "���ο�û"         ''1�����
			Case "30"
				getLotteItemStatCd = "���οϷ�"
			Case "40"
				getLotteItemStatCd = "�ݷ�"
			Case "50"
				getLotteItemStatCd = "���κҰ�"
			Case "51"
				getLotteItemStatCd = "����ο�û"
			Case "52"
				getLotteItemStatCd = "������û"
			Case "-9"
				getLotteItemStatCd = "�̵��"
		End Select
	end Function

	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub
end class

class CLotte
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
	public FRectExtNotReg
	public FRectIsReged
	public FRectNotinmakerid
	Public FRectNotinitemid
	Public FRectExcTrans
	public FRectPriceOption

	public FRectSailYn
	public FRectIsMadeHand
	public FRectIsOption
	public FRectInfoDiv

	public FRectItemID
	public FRectItemName
	public FRectMakerid
	public FRectLotteNotReg
	public FRectMatchCate
	public FRectMatchCateNotCheck
	public FRectSellYn
	public FRectLimitYn
	public FRectLotteGoodNo
	public FRectLotteTmpGoodNo
	public FRectMinusMigin
	public FRectonlyValidMargin
	Public FRectStartMargin
	Public FRectEndMargin
	public FRectIsSoldOut
	public FRectExpensive10x10
	public FRectLotteYes10x10No
	public FRectReqEdit
	Public FRectDeliverytype
	Public FRectMwdiv
	Public FRectIsextusing

	public FRectLotteNo10x10Yes
	public FRectOnreginotmapping
	public FRectNotJehyu
	public FRectEventid
	public FRectdiffPrc
	public FRectdisptpcd
    public FRectoptAddprcExists
    public FRectoptAddPrcRegTypeNone
    public FRectoptAddprcExistsExcept
    public FRectoptExists
    public FRectregedOptNull
    public FRectOrdType
    public FRectFailCntExists
    public FRectFailCntOverExcept
    public FRectLimitOver
    public FRectExtSellYn
    public FRectInfoDivYn
    public FRectOnlyNotUsingCheck
	Public FRectIsSpecialPrice

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
				set FItemList(i) = new CLotteItem
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
				set FItemList(i) = new CLotteItem

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
	'// �ٹ�����-�Ե����� ī�װ� :: ������ ī�װ��� ���� �Ǿ�� ��..
	public Sub getTenLotteCateList
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
			addSql = addSql & " and cm.dispNo='" & FRectDspNo & "'"
		end if

		if FRectIsMapping="Y" then
			addSql = addSql & " and cm.DispNo is Not null "
		elseif FRectIsMapping="N" then
			addSql = addSql & " and cm.DispNo is null "
		end if

		if FRectKeyword<>"" then
			Select Case FRectSDiv
				Case "LCD"	'�Ե����� �����ڵ� �˻�
					addSql = addSql & " and cm.DispNo='" & FRectKeyword & "'"
				Case "CNM"	'ī�װ���(�ٹ����� �Һз���)
					addSql = addSql & " and s.code_nm like '%" & FRectKeyword & "%'"
			End Select
		end if

		sqlStr = " select count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr + " from db_item.dbo.tbl_cate_small as s "
		sqlStr = sqlStr + " 	left join db_item.dbo.tbl_lotte_cate_mapping as cm "
		sqlStr = sqlStr + " 		on cm.tenCateLarge=s.code_large "
		sqlStr = sqlStr + " 			and cm.tenCateMid=s.code_mid "
		sqlStr = sqlStr + " 			and cm.tenCateSmall=s.code_small "
		sqlStr = sqlStr + " 	left Join db_temp.dbo.tbl_lotte_Category as lc "
		sqlStr = sqlStr + " 		on lc.DispNo=cm.DispNo "
		if FRectdisptpcd<>"" then
            sqlStr = sqlStr & " and lc.disptpcd='" & FRectdisptpcd &"'"
        end if
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
		sqlStr = sqlStr + " 	,cm.DispNo, lc.DispNm, lc.DispLrgNm, lc.DispMidNm, lc.DispSmlNm, lc.DispThnNm, lc.groupCode, lc.disptpcd "
		sqlStr = sqlStr + " 	,lc.isusing"
		sqlStr = sqlStr + " from db_item.dbo.tbl_cate_small as s "
		sqlStr = sqlStr + " 	left join db_item.dbo.tbl_lotte_cate_mapping as cm "
		sqlStr = sqlStr + " 		on cm.tenCateLarge=s.code_large "
		sqlStr = sqlStr + " 			and cm.tenCateMid=s.code_mid "
		sqlStr = sqlStr + " 			and cm.tenCateSmall=s.code_small "
		sqlStr = sqlStr + " 	left Join db_temp.dbo.tbl_lotte_Category as lc "
		sqlStr = sqlStr + " 		on lc.DispNo=cm.DispNo "
		if FRectdisptpcd<>"" then
            sqlStr = sqlStr & " and lc.disptpcd='" & FRectdisptpcd &"'"
        end if
		sqlStr = sqlStr + " Where 1=1 " & addSql
		sqlStr = sqlStr + " order by s.code_large,s.code_mid,s.code_small"

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CLotteItem

				FItemList(i).FtenCDLName	= db2html(rsget("large_nm"))
				FItemList(i).FtenCDMName	= db2html(rsget("mid_nm"))
				FItemList(i).FtenCDSName	= db2html(rsget("small_nm"))

				FItemList(i).FDispNo		= rsget("DispNo")
				FItemList(i).FDispNm		= db2html(rsget("DispNm"))

				FItemList(i).FtenCateLarge	= rsget("code_large")
				FItemList(i).FtenCateMid	= rsget("code_mid")
				FItemList(i).FtenCateSmall	= rsget("code_small")

				FItemList(i).FgroupCode		= rsget("groupCode")
				FItemList(i).FDispLrgNm		= db2html(rsget("DispLrgNm"))
				FItemList(i).FDispMidNm		= db2html(rsget("DispMidNm"))
				FItemList(i).FDispSmlNm		= db2html(rsget("DispSmlNm"))
				FItemList(i).FDispThnNm		= db2html(rsget("DispThnNm"))

                FItemList(i).Fdisptpcd      = rsget("disptpcd")
                FItemList(i).FCateisusing   = rsget("isusing")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end Sub

	'// �Ե����� ī�װ�
	public Sub getLotteCategoryList
		dim sqlStr, addSql, i

		if FRectDspNo<>"" then
			addSql = addSql & " and c.dispNo=" & FRectDspNo
		end if

'		if FRectIsMapping="Y" then
'			addSql = addSql & " and m.DispNo is Not null "
'		elseif FRectIsMapping="N" then
'			addSql = addSql & " and m.DispNo is null "
'		end if

		if FRectGrpCode<>"" then
			addSql = addSql & " and c.groupCode=" & FRectGrpCode
		end if

        if FRectdisptpcd<>"" then
            addSql = addSql & " and c.disptpcd='" & FRectdisptpcd &"'"
        end if

		if FRectKeyword<>"" then
			Select Case FRectSDiv
				Case "LCD"	'�Ե����� �����ڵ� �˻�
					addSql = addSql & " and c.DispNo='" & FRectKeyword & "'"
'				Case "TCD"	'�ٹ����� ī�װ��ڵ� �˻�(���߼� �����ڵ� 9�ڸ�)
'					addSql = addSql & " and m.tenCateLarge&m.tenCateMid&m.tenCateSmall='" & FRectKeyword & "'"
				Case "CNM"	'ī�װ���(�Ե� ���з���)
					''addSql = addSql & " and c.dispNm like '%" & FRectKeyword & "%'"
					addSql = addSql & " and ((c.dispNm like '%" & FRectKeyword & "%') or (c.dispsmlNm like '%" & FRectKeyword & "%'))"
			End Select
		end if

		sqlStr = " select count(c.DispNo) as cnt, CEILING(CAST(Count(c.DispNo) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr + " from db_temp.dbo.tbl_lotte_Category as c "
'		sqlStr = sqlStr + " 	Left Join db_item.dbo.tbl_lotte_cate_mapping as m "
'		sqlStr = sqlStr + " 		on c.DispNo=m.DispNo "
		sqlStr = sqlStr + " Where c.DispMidNm<>'�ٺ����' " & addSql

		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'������������ ��ü ���������� Ŭ �� �Լ�����
		if Cint(FCurrPage)>Cint(FTotalPage) then
			FResultCount = 0
			exit sub
		end if

'		sqlStr = " select  top " + CStr(FPageSize*FCurrPage) + " c.*, m.tenCateLarge, m.tenCateMid, m.tenCateSmall, s.code_nm "
		sqlStr = " select distinct top " + CStr(FPageSize*FCurrPage) + " c.* "
		sqlStr = sqlStr + " from db_temp.dbo.tbl_lotte_Category as c "
'		sqlStr = sqlStr + " 	Left Join db_item.dbo.tbl_lotte_cate_mapping as m "
'		sqlStr = sqlStr + " 		on c.DispNo=m.DispNo "
'		sqlStr = sqlStr + " 	Left Join db_item.dbo.tbl_cate_small as s "
'		sqlStr = sqlStr + " 		on s.code_large=m.tenCateLarge "
'		sqlStr = sqlStr + " 		and s.code_mid=m.tenCateMid "
'		sqlStr = sqlStr + " 		and s.code_small=m.tenCateSmall "
		sqlStr = sqlStr + " Where c.DispMidNm<>'�ٺ����' " & addSql
		sqlStr = sqlStr + " order by c.DispLrgNm, c.DispMidNm, c.DispSmlNm, c.DispThnNm, c.DispNo"

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CLotteItem

				FItemList(i).FDispNo		= rsget("DispNo")
				FItemList(i).FDispNm		= db2html(rsget("DispNm"))
				FItemList(i).FDispLrgNm		= db2html(rsget("DispLrgNm"))
				FItemList(i).FDispMidNm		= db2html(rsget("DispMidNm"))
				FItemList(i).FDispSmlNm		= db2html(rsget("DispSmlNm"))
				FItemList(i).FDispThnNm		= db2html(rsget("DispThnNm"))
                FItemList(i).Fdisptpcd      = rsget("disptpcd")

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
				set FItemList(i) = new CLotteItem

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
	'�ɼ��߰��ݾ� ��ǰ ����Ʈ
	public Sub getLotteAddOptionRegedItemList
		Dim sqlStr, addSql, i
		'�귣��˻�
		If FRectMakerid <> "" Then
			addSql = addSql & " and i.makerid='" & FRectMakerid & "'"
		End If

		'��ǰ�ڵ� �˻�
        If (FRectItemid <> "") then
            If Right(Trim(FRectItemid) ,1) = "," Then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            Else
				FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" + FRectItemid + ")"
            End If
        End If

		'��ǰ�� �˻�
		If FRectItemName <> "" Then
			addSql = addSql & " and i.itemname like '%" & FRectItemName & "%'"
		End if

		'�Ե����� ��ǰ��ȣ �˻�
        If (FRectLotteGoodNo <> "") then
            If Right(Trim(FRectLotteGoodNo) ,1) = "," Then
            	FRectLotteGoodNo = Replace(FRectLotteGoodNo,",,",",")
            	addSql = addSql & " and J.LotteGoodNo in ('" & replace(Left(FRectLotteGoodNo, Len(FRectLotteGoodNo)-1),",","','") & "')"
            Else
				FRectLotteGoodNo = Replace(FRectLotteGoodNo,",,",",")
            	addSql = addSql & " and J.LotteGoodNo in ('" & replace(FRectLotteGoodNo,",","','") & "')"
            End If
        End If

		'�Ե����� ������ ��ǰ��ȣ �˻�
        If (FRectLotteTmpGoodNo <> "") then
            If Right(Trim(FRectLotteTmpGoodNo) ,1) = "," Then
            	FRectLotteTmpGoodNo = Replace(FRectLotteTmpGoodNo,",,",",")
            	addSql = addSql & " and J.LotteTmpGoodNo in ('" & replace(Left(FRectLotteTmpGoodNo, Len(FRectLotteTmpGoodNo)-1),",","','") & "')"
            Else
				FRectLotteTmpGoodNo = Replace(FRectLotteTmpGoodNo,",,",",")
            	addSql = addSql & " and J.LotteTmpGoodNo in ('" & replace(FRectLotteTmpGoodNo,",","','") & "')"
            End If
        End If

		'ī�װ� �˻�
		If FRectCDL <> "" Then
			addSql = addSql & " and i.cate_large='" & FRectCDL & "'"
		End if
		If FRectCDM <> "" Then
			addSql = addSql & " and i.cate_mid='" & FRectCDM & "'"
		End if
		If FRectCDS <> "" Then
			addSql = addSql & " and i.cate_small='" & FRectCDS & "'"
		End If

		'��Ͽ��� �˻�
		Select Case FRectExtNotReg
			Case "Q"	''��Ͻ���
				addSql = addSql & " and J.LotteStatCd = '-1'"
			Case "W"	'��Ͽ���
				addSql = addSql & " and J.LotteStatCd = '00'"
			    addSql = addSql & " and J.LotteGoodNo is Null "
			Case "C"	'�ݷ�
			    addSql = addSql & " and J.LotteStatCd = '40'"
			    addSql = addSql & " and J.LotteGoodNo is Null "
			Case "F"	'��ϿϷ�(�ӽ�)
			    addSql = addSql & " and J.LotteTmpGoodNo is Not Null "
			    addSql = addSql & " and J.LotteStatCd <> '40'"
			    addSql = addSql & " and J.LotteGoodNo is Null "
			Case "D"	'��ϿϷ�(����)
			    addSql = addSql & " and J.LotteTmpGoodNo is Not Null"
				addSql = addSql & " and J.LotteGoodNo is Not Null"
			Case "R"	'�������		'�����ٸ����� ���
				addSql = addSql & " and J.LotteLastUpdate < i.lastupdate"
				addSql = addSql & " and isnull(J.LotteGoodNo, '') <> '' "
		End Select

		'�̵�� ������ư Ŭ�� ��
		Select Case FRectIsReged
			Case "N"	'��Ͽ����̻�
			    addSql = addSql & " and J.midx is NULL  and (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>5)) "
		End Select

		'�Ǹſ��� �˻�
		Select Case FRectSellYn
			Case "Y"	addSql = addSql & " and i.sellYn='Y'"			'�Ǹ�
			Case "N"	addSql = addSql & " and i.sellYn in ('S','N')"	'ǰ��
		End Select

		'�ٹ����� �������� �˻�
		If FRectLimitYn <> "" Then
			addSql = addSql & " and i.limitYn = '" & FRectLimitYn & "'"
		End If

		'�ٹ����� ���Ͽ��� �˻�
		If FRectSailYn <> "" Then
			addSql = addSql & " and i.sailYn = '" & FRectSailYn & "'"
		End If

		'������ �� ���� CMAXMARGIN �̻� �˻�
		If (FRectonlyValidMargin <> "") Then
			IF (FRectonlyValidMargin = "Y") Then
				addSql = addSql & " and Round(((i.sellcash-i.buycash)/(CASE WHEN i.sellcash=0 THEN 1 ELSE i.sellcash END))*100,0) >= " & CMAXMARGIN & VbCrlf
			Else
				addSql = addSql & " and Round(((i.sellcash-i.buycash)/(CASE WHEN i.sellcash=0 THEN 1 ELSE i.sellcash END))*100,0) < " & CMAXMARGIN & VbCrlf
			End If
		End If

		If (FRectStartMargin <> "") OR (FRectEndMargin <> "") Then
			If (FRectStartMargin <> "") And (FRectEndMargin <> "") Then
				addSql = addSql & " and ("
				addSql = addSql & " 	convert(int, ((i.sellcash-i.buycash)/(CASE WHEN i.sellcash=0 THEN 1 ELSE i.sellcash END))*100)>="&FRectStartMargin & VbCrlf
				addSql = addSql & " 	and convert(int, ((i.sellcash-i.buycash)/(CASE WHEN i.sellcash=0 THEN 1 ELSE i.sellcash END))*100)<="&FRectEndMargin & VbCrlf
				addSql = addSql & " ) "
			ElseIf (FRectStartMargin <> "") And (FRectEndMargin = "") Then
				addSql = addSql & " and convert(int, ((i.sellcash-i.buycash)/(CASE WHEN i.sellcash=0 THEN 1 ELSE i.sellcash END))*100)>="&FRectStartMargin & VbCrlf
			ElseIf (FRectStartMargin = "") And (FRectEndMargin <> "") Then
				addSql = addSql & " and convert(int, ((i.sellcash-i.buycash)/(CASE WHEN i.sellcash=0 THEN 1 ELSE i.sellcash END))*100)<="&FRectEndMargin & VbCrlf
			End If
		End If

		'�ֹ����� ���� �˻�
		If FRectisMadeHand <> "" Then
			If (FRectisMadeHand = "Y") Then
				addSql = addSql & " and i.itemdiv in ('06', '16')" & VbCrlf
			Else
				addSql = addSql & " and i.itemdiv not in ('06', '16')" & VbCrlf
			End If
		End if

		'�ɼ� ���� �˻�
		If FRectIsOption <> "" Then
			If FRectIsOption = "optAll" Then			'�ɼ���ü
				addSql = addSql & " and i.optioncnt > 0"
			ElseIf FRectIsOption = "optaddpricey" Then	'�߰��ݾ�Y
				addSql = addSql & " and i.optioncnt > 0"
 				addSql = addSql & " and J.optAddPrcCnt > 0"
			ElseIf FRectIsOption = "optaddpricen" Then	'�߰��ݾ�N
				addSql = addSql & " and i.optioncnt > 0"
				addSql = addSql & " and isNULL(J.optAddPrcCnt,0)=0"
			ElseIf FRectIsOption = "optN" Then			'��ǰ
				addSql = addSql & " and i.optioncnt = 0"
			ElseIf FRectIsOption = "optAddPrcRegType" Then
				addSql = addSql & " and J.optAddPrcCnt > 0"
				addSql = addSql & " and J.optAddPrcRegType = 0"
			End If
		End If

		'�ٹ����� ǰ������ �˻�
		If (FRectInfoDiv <> "") then
			If (FRectInfoDiv = "YY") Then
				addSql = addSql & " and isNULL(ct.infodiv,'')<>''"
			ElseIf (FRectInfoDiv = "NN") Then
				addSql = addSql & " and isNULL(ct.infodiv,'')=''"
			Else
				addSql = addSql & " and ct.infodiv = '"&FRectInfoDiv&"'"
			End If
		End If

		'�ٹ����� ������� �귣�� ���� �˻�
		If (FRectNotinmakerid <> "") then
			If (FRectNotinmakerid = "Y") Then
				addSql = addSql & " and exists(SELECT top 1 n.makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid n with (nolock) WHERE n.makerid=i.makerid and n.mallgubun = 'lotteCom') "
			ElseIf (FRectNotinmakerid = "N") Then
				addSql = addSql & " and not exists(SELECT top 1 n.makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid n with (nolock) WHERE n.makerid=i.makerid and n.mallgubun = 'lotteCom') "
			End If
		End If

		'�ٹ����� ������� ��ǰ ���� �˻�
		If (FRectNotinitemid <> "") then
			If (FRectNotinitemid = "Y") Then
				addSql = addSql & " and exists(SELECT top 1 n.itemid FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid n with (nolock) WHERE n.itemid=i.itemid and n.mallgubun = 'lotteCom') "
			ElseIf (FRectNotinitemid = "N") Then
				addSql = addSql & " and not exists(SELECT top 1 n.itemid FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid n with (nolock) WHERE n.itemid=i.itemid and n.mallgubun = 'lotteCom') "
			End If
		End If

		'���޸� �������� ��ǰ �˻�
		If (FRectExcTrans <> "") then
			If (FRectExcTrans = "Y") Then
				addSql = addSql & " and 'Y' = (CASE WHEN i.isusing='N' "
				addSql = addSql & " or i.makerid in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='lotteCom') "
				addSql = addSql & " or i.itemid in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='lotteCom') "
				addSql = addSql & " or i.isExtUsing='N' "
				addSql = addSql & " or uc.isExtUsing='N' "
				addSql = addSql & " or i.deliveryType = 7 "
				addSql = addSql & " or ((i.deliveryType = 9) and (i.sellcash < 10000)) "
				addSql = addSql & " or i.itemdiv = '21' "
				addSql = addSql & " or i.deliverfixday in ('C','X') "
				addSql = addSql & " or i.itemdiv >= 50 "
				addSql = addSql & " or i.itemdiv = '08' "
				addSql = addSql & " or i.itemdiv = '09' "
				addSql = addSql & " or i.cate_large = '999' "
				addSql = addSql & " or i.cate_large='' "
				addSql = addSql & " or not (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>5)) "
				addSql = addSql & " or not ( "
				addSql = addSql & " 	i.optioncnt = 0 "
				addSql = addSql & " 	or "
				addSql = addSql & " 	exists(SELECT top 1 o.itemid FROM [db_item].[dbo].tbl_item_option o WHERE o.isUsing='Y' and o.optsellyn='Y' and o.itemid=i.itemid and (o.optlimityn <> 'Y' or (o.optlimitno-o.optlimitsold)>5)) "
				addSql = addSql & " ) "
				addSql = addSql & " THEN 'Y' ELSE 'N' END) "
			ElseIf (FRectExcTrans = "F") Then
				addSql = addSql & " and i.makerid not in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='lotteCom') "
				addSql = addSql & " and i.itemid not in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='lotteCom') "
				addSql = addSql & " and i.isusing='Y' "
				addSql = addSql & " and i.isExtUsing='Y' "											'// �ܺθ�����ǰ
				addSql = addSql & " and uc.isExtUsing='Y' "
				addSql = addSql & " and i.deliveryType <> 7 "										'// ��ü����
				addSql = addSql & " and i.itemdiv <> '21' "											'// ����ǰ
				addSql = addSql & " and i.deliverfixday not in ('C','X') "							'// �ɹ��, ȭ�����
				addSql = addSql & " and not ((i.deliveryType = 9) and (i.sellcash < 10000)) "		'// �ǸŰ�(���ΰ�) 1���� �̸�
				addSql = addSql & " and i.itemdiv <> '08' "											'// Ƽ��(����) ��ǰ
				addSql = addSql & " and i.itemdiv <> '09' "											'// Present��ǰ
				addSql = addSql & " and i.itemdiv < 50 "
				addSql = addSql & " and (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>5)) "
				addSql = addSql & " and ( "
				addSql = addSql & " 	i.optioncnt = 0 "
				addSql = addSql & " 	or "
				addSql = addSql & " 	exists(SELECT top 1 o.itemid FROM [db_item].[dbo].tbl_item_option o WHERE o.isUsing='Y' and o.optsellyn='Y' and o.itemid=i.itemid and (o.optlimityn <> 'Y' or (o.optlimitno-o.optlimitsold)>5)) "
				addSql = addSql & " ) "
				addSql = addSql & " and 'Y' = (CASE WHEN i.cate_large = '999' "
				addSql = addSql & " or i.cate_large='' "
				addSql = addSql & " or J.accFailCnt > 0 "
				addSql = addSql & " THEN 'Y' ELSE 'N' END) "
			ElseIf (FRectExcTrans = "N") Then
				addSql = addSql & " and not exists(SELECT top 1 n.makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid n with (nolock) WHERE n.makerid=i.makerid and n.mallgubun = 'lotteCom') "
				addSql = addSql & " and not exists(SELECT top 1 n.itemid FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid n with (nolock) WHERE n.itemid=i.itemid and n.mallgubun = 'lotteCom') "
				addSql = addSql & " and i.isusing='Y' "
				addSql = addSql & " and i.isExtUsing='Y' "											'// �ܺθ�����ǰ
				addSql = addSql & " and uc.isExtUsing='Y' "
				addSql = addSql & " and i.deliveryType <> 7 "										'// ��ü����
				addSql = addSql & " and i.itemdiv <> '21' "											'// ����ǰ
				addSql = addSql & " and i.deliverfixday not in ('C','X') "							'// �ɹ��, ȭ�����
				addSql = addSql & " and not ((i.deliveryType = 9) and (i.sellcash < 10000)) "		'// �ǸŰ�(���ΰ�) 1���� �̸�
				addSql = addSql & " and i.itemdiv <> '08' "											'// Ƽ��(����) ��ǰ
				addSql = addSql & " and i.itemdiv <> '09' "											'// Present��ǰ
				addSql = addSql & " and i.cate_large <> '999' "										'// ī�װ� ������
				addSql = addSql & " and i.cate_large <> '' "										'// ī�װ� ������
				addSql = addSql & " and i.itemdiv < 50 "
				addSql = addSql & " and (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>5)) "
				addSql = addSql & " and ( "
				addSql = addSql & " 	i.optioncnt = 0 "
				addSql = addSql & " 	or "
				addSql = addSql & " 	exists(SELECT top 1 o.itemid FROM [db_item].[dbo].tbl_item_option o WHERE o.isUsing='Y' and o.optsellyn='Y' and o.itemid=i.itemid and (o.optlimityn <> 'Y' or (o.optlimitno-o.optlimitsold)>5)) "
				addSql = addSql & " ) "
				addSql = addSql & " and ( "
				addSql = addSql & " 	i.optioncnt = 0 "
				addSql = addSql & " 	or "
				addSql = addSql & " 	not exists(SELECT top 1 o.itemid FROM [db_item].[dbo].tbl_item_option o WHERE o.isUsing='Y' and o.itemid=i.itemid and o.optAddPrice > 0) "
				addSql = addSql & " ) "
				''addSql = addSql & " and  G.optAddPrcCnt = 0 "
			End If
		End If

		'�ɼ��߰��ݾ�New
		If (FRectPriceOption <> "") then
			If (FRectPriceOption = "Y") Then
				addSql = addSql & " and i.itemid in (SELECT itemid FROM db_item.[dbo].[tbl_const_OptAddPrice_Exists]) "
			ElseIf (FRectPriceOption = "N") Then
				addSql = addSql & " and i.itemid not in (SELECT itemid FROM db_item.[dbo].[tbl_const_OptAddPrice_Exists]) "
			End If
		End If

		'�Ե����� �Ǹſ���
		If (FRectExtSellYn<>"") then
			If (FRectExtSellYn = "YN") Then
				addSql = addSql & " and J.LotteSellYn <> 'X'"
			Else
				addSql = addSql & " and J.LotteSellYn='" & FRectExtSellYn & "'"
			End if
		End If

		'��ϼ���������ǰ
		Select Case FRectFailCntExists
			Case "Y"	'����1ȸ�̻�
				addSql = addSql & " and J.accFailCNT>0"
			Case "N"	'����0ȸ
				addSql = addSql & " and J.accFailCNT=0"
		End Select

		'�Ե����� ī�װ� ��Ī ����
		Select Case FRectMatchCate
			Case "Y"	'��Ī�Ϸ�
				addSql = addSql & " and isnull(c.mapCnt, '') <> ''"
			Case "N"	'�̸�Ī
				addSql = addSql & " and isnull(c.mapCnt, '') = ''"
		End Select

        '�Ե����� < 10x10 ����
		If (FRectexpensive10x10 <> "") Then
			addSql = addSql & " and J.LottePrice is Not Null and J.LottePrice < i.sellcash"
		End If

		'���ݻ�����ü����
		If FRectdiffPrc <> "" Then
			addSql = addSql & " and J.LottePrice is Not Null and i.sellcash <> J.LottePrice "
		End If

		'�Ե������Ǹ�,  10x10 ǰ��
		If (FRectLotteYes10x10No <> "") Then
			addSql = addSql & " and i.sellyn<>'Y'"
			addSql = addSql & " and J.LotteSellYn='Y'"
		End If

		'�Ե�����ǰ��&�ٹ������ǸŰ���(�Ǹ���,����>=10) ��ǰ����
		If FRectLotteNo10x10Yes <> "" Then
			addSql = addSql & " and (J.LotteSellYn= 'N' and i.sellyn='Y' and (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>"&CMAXLIMITSELL&")))"
		End If

		'���������ǰ����(����������Ʈ�� ����)
		If FRectReqEdit <> "" Then
			addSql = addSql & " and J.LotteLastUpdate < i.lastupdate "
		End If

		'�����ٸ����� ��� ����Ƚ�� ����
		If (FRectFailCntOverExcept <> "") Then
			addSql = addSql & " and J.accFailCNT < "&FRectFailCntOverExcept
		End If

		'�����ٸ����� ��� ��Ʈ������Ʈ ���� ����
		If (FRectOrdType = "LU") Then
		    addSql = addSql & " and isnull(J.lastStatCheckDate,'') = '' "
		    addSql = addSql & " and Left(i.lastupdate, 10) <> Left(J.LotteLastUpdate, 10) "
		End If

		'��۱���
		If (FRectDeliverytype <> "") Then
			addSql = addSql & " and i.deliverytype='" & FRectDeliverytype & "'"
		End If

		'�ŷ�����
		If FRectMWDiv = "MW" Then
			addSql = addSql & " and (i.mwdiv='M' or i.mwdiv='W')"
		ElseIf FRectMWDiv<>"" Then
			addSql = addSql & " and i.mwdiv='"& FRectMWDiv & "'"
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(i.itemid) as cnt, CEILING(CAST(Count(i.itemid) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i "
		sqlStr = sqlStr & " JOIN db_etcmall.[dbo].[tbl_Outmall_option_Manager] as M on M.itemid = i.itemid and M.mallid = '"&CMALLNAME&"' "
		sqlStr = sqlStr & "	LEFT JOIN db_item.dbo.tbl_item_option as o on i.itemid = o.itemid and M.itemoption = o.itemoption and M.mallid = '"&CMALLNAME&"' "
		If (FRectIsReged = "N") OR (FRectIsReged = "A") Then		'//�̵���� �ƴϸ� JOIN
			sqlStr = sqlStr & "	LEFT JOIN db_etcmall.dbo.tbl_lotteAddoption_regitem as J on J.midx = M.idx "
		Else
			sqlStr = sqlStr & "	JOIN db_etcmall.dbo.tbl_lotteAddoption_regitem as J on J.midx = M.idx "
		End If
		sqlStr = sqlStr & "	LEFT Join db_item.dbo.tbl_OutMall_CateMap_Summary as c on c.mallid = 'lotteCom' and c.tenCateLarge = i.cate_large and c.tenCateMid = i.cate_mid and c.tenCateSmall = i.cate_small "
		sqlStr = sqlStr & " LEFT join db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_contents as ct on i.itemid = ct.itemid"
		sqlStr = sqlStr & " WHERE 1 = 1"
		If (FRectIsReged <> "N" and FRectExtNotReg <> "Q")  Then		'// �̵�ϵ� �ƴϰ� ��Ͻ��е� �ƴϸ� ���� ����
			If FRectIsReged = "Q" Then
				sqlStr = sqlStr & " and J.LotteGoodNo is Not Null "
				sqlStr = sqlStr & " and (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>5)) "
				sqlStr = sqlStr & " and 'N' = (CASE WHEN i.isusing='N'  "
				sqlStr = sqlStr & " or i.isExtUsing='N' "
				sqlStr = sqlStr & " or uc.isExtUsing='N' "
				sqlStr = sqlStr & " or ((i.deliveryType = 9) and (i.sellcash < 10000)) "
				sqlStr = sqlStr & " or i.sellyn<>'Y' "
				sqlStr = sqlStr & " or i.deliverfixday in ('C','X') "
				sqlStr = sqlStr & " or i.itemdiv >= 50 or i.itemdiv = '08' or i.cate_large = '999' or i.cate_large='' "
				sqlStr = sqlStr & " or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "
				sqlStr = sqlStr & " or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "
				sqlStr = sqlStr & " THEN 'Y' ELSE 'N' END) "
			End If
		Else
    		sqlStr = sqlStr & " and i.isusing='Y' "
    		sqlStr = sqlStr & " and i.deliverfixday not in ('C','X') "
    		sqlStr = sqlStr & " and i.basicimage is not null "
    		sqlStr = sqlStr & " and i.itemdiv<50 "  ''and i.itemdiv<>'08'
    		sqlStr = sqlStr & " and i.itemdiv not in ('08','09')"
    		sqlStr = sqlStr & " and i.cate_large<>'' "
		    sqlStr = sqlStr & " and ((i.cate_large <> '999') or ((i.cate_large='999')and(i.makerid='ftroupe'))) " & VBCRLF
    		sqlStr = sqlStr & "	and i.makerid not in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "	'������� �귣��
    		sqlStr = sqlStr & "	and i.itemid not in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "		'������� ��ǰ
			If FRectExtNotReg <> "" Then
				sqlStr = sqlStr & " and i.sellcash>=1000 "  & VBCRLF
				'sqlStr = sqlStr & " and i.itemdiv<>'06'" & VBCRLF				'�ֹ�����
			End If
    		sqlStr = sqlStr & "	and uc.isExtUsing='Y'"  ''20130304 �귣�� ���޻�뿩�� Y��.
			sqlStr = sqlStr & " and i.isExtUsing='Y'"														'//���޸� �ǸŸ� ���
			sqlStr = sqlStr & " and i.deliverytype not in ('7')"											'//���ҹ�� ��ǰ ����
		End If
		sqlStr = sqlStr & addSql
rw sqlStr
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'������������ ��ü ���������� Ŭ �� �Լ�����
		if Cint(FCurrPage)>Cint(FTotalPage) then
			FResultCount = 0
			exit sub
		end if

		sqlStr = ""
		sqlStr = sqlStr & " SELECT top " & CStr(FPageSize*FCurrPage) & " M.idx, isnull(M.itemnameChange, '') as itemnameChange, isnull(M.newitemname, '') as newitemname, i.itemid, i.itemname, i.smallImage "
		sqlStr = sqlStr & "	, i.makerid, i.regdate, i.lastUpdate, i.orgPrice, i.sellcash, i.buycash"
		sqlStr = sqlStr & "	, i.sellYn, i.sailyn, i.LimitYn, i.LimitNo, i.LimitSold, i.deliverytype, i.optionCnt"
		sqlStr = sqlStr & "	, J.LotteRegdate, J.LotteLastUpdate, J.LotteGoodNo, J.LotteTmpGoodNo, J.LottePrice, J.LotteSellYn, J.regUserid, IsNULL(J.LotteStatCd,-9) as LotteStatCd "
		sqlStr = sqlStr & "	, c.mapCnt, J.rctSellCNT, J.accFailCNT, J.lastErrStr "
		sqlStr = sqlStr & " ,uc.defaultdeliverytype, uc.defaultfreeBeasongLimit"
		sqlStr = sqlStr & "	, Ct.infoDiv, i.itemdiv"
		sqlStr = sqlStr & "	, o.itemoption , o.optaddprice, o.optionname, o.optlimitno, o.optlimitsold, o.optsellyn "
		sqlStr = sqlStr & "	, M.optionname as regedOptionname, M.itemname as regedItemname "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i "
		sqlStr = sqlStr & " JOIN db_etcmall.[dbo].[tbl_Outmall_option_Manager] as M on M.itemid = i.itemid and M.mallid = '"&CMALLNAME&"' "
		sqlStr = sqlStr & "	LEFT JOIN db_item.dbo.tbl_item_option as o on i.itemid = o.itemid and M.itemoption = o.itemoption and M.mallid = '"&CMALLNAME&"' "
		If (FRectIsReged = "N") OR (FRectIsReged = "A") Then		'//�̵���� �ƴϸ� JOIN
			sqlStr = sqlStr & "	LEFT JOIN db_etcmall.dbo.tbl_lotteAddoption_regitem as J on J.midx = M.idx "
		Else
			sqlStr = sqlStr & "	JOIN db_etcmall.dbo.tbl_lotteAddoption_regitem as J on J.midx = M.idx "
		End If
		sqlStr = sqlStr & " LEFT JOIN db_item.dbo.tbl_OutMall_CateMap_Summary as c on c.mallid = 'lotteCom' and c.tenCateLarge = i.cate_large and c.tenCateMid = i.cate_mid and c.tenCateSmall = i.cate_small "
		sqlStr = sqlStr & "	LEFT JOIN db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_contents as ct on i.itemid = ct.itemid"
		sqlStr = sqlStr & " where 1 = 1"
		If (FRectIsReged <> "N" and FRectExtNotReg <> "Q")  Then		'// �̵�ϵ� �ƴϰ� ��Ͻ��е� �ƴϸ� ���� ����
			If FRectIsReged = "Q" Then
				sqlStr = sqlStr & " and J.LotteGoodNo is Not Null "
				sqlStr = sqlStr & " and (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>5)) "
				sqlStr = sqlStr & " and 'N' = (CASE WHEN i.isusing='N'  "
				sqlStr = sqlStr & " or i.isExtUsing='N' "
				sqlStr = sqlStr & " or uc.isExtUsing='N' "
				sqlStr = sqlStr & " or ((i.deliveryType = 9) and (i.sellcash < 10000)) "
				sqlStr = sqlStr & " or i.sellyn<>'Y' "
				sqlStr = sqlStr & " or i.deliverfixday in ('C','X') "
				sqlStr = sqlStr & " or i.itemdiv >= 50 or i.itemdiv = '08' or i.cate_large = '999' or i.cate_large='' "
				sqlStr = sqlStr & " or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "
				sqlStr = sqlStr & " or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "
				sqlStr = sqlStr & " THEN 'Y' ELSE 'N' END) "
			End If
		Else
    		sqlStr = sqlStr & " and i.isusing='Y' "
    		sqlStr = sqlStr & " and i.deliverfixday not in ('C','X') "
    		sqlStr = sqlStr & " and i.basicimage is not null "
    		sqlStr = sqlStr & " and i.itemdiv<50 "  ''and i.itemdiv<>'08'
    		sqlStr = sqlStr & " and i.itemdiv not in ('08','09')"
    		sqlStr = sqlStr & " and i.cate_large<>'' "
		    sqlStr = sqlStr & " and ((i.cate_large <> '999') or ((i.cate_large='999')and(i.makerid='ftroupe'))) " & VBCRLF
    		sqlStr = sqlStr & "	and i.makerid not in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "	'������� �귣��
    		sqlStr = sqlStr & "	and i.itemid not in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "		'������� ��ǰ
			If FRectExtNotReg <> "" Then
				sqlStr = sqlStr & " and i.sellcash>=1000 "  & VBCRLF
				'sqlStr = sqlStr & " and i.itemdiv<>'06'" & VBCRLF				'�ֹ�����
			End If
    		sqlStr = sqlStr & "	and uc.isExtUsing='Y'"  ''20130304 �귣�� ���޻�뿩�� Y��.
			sqlStr = sqlStr & " and i.isExtUsing='Y'"														'//���޸� �ǸŸ� ���
			sqlStr = sqlStr & " and i.deliverytype not in ('7')"											'//���ҹ�� ��ǰ ����
    	End If
		sqlStr = sqlStr & addSql
		If (FRectOrdType = "LS") AND (FRectLotteNotReg = "F") Then
			sqlStr = sqlStr & " ORDER BY J.lastStatCheckDate, J.LotteLastupdate"
		ElseIf (FRectLotteNotReg = "F") Then
		    sqlStr = sqlStr & " ORDER BY J.LotteLastupdate "
		ElseIf (FRectOrdType = "B") Then
		    sqlStr = sqlStr & " ORDER BY i.itemscore DESC, i.itemid DESC "
		ElseIf (FRectOrdType = "BM") Then
		    sqlStr = sqlStr & " ORDER BY J.rctSellCNT DESC, i.itemscore DESC, i.itemid DESC"
		Else
		    sqlStr = sqlStr & " ORDER BY i.itemid DESC" '' m.regdate desc
	    End If
		rsget.pagesize = FPageSize
'rw sqlStr
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CLotteItem
					FItemList(i).Fidx				= rsget("idx")
					FItemList(i).FNewitemname		= rsget("newitemname")
					FItemList(i).FItemnameChange	= rsget("itemnameChange")
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

					FItemList(i).FLotteRegdate		= rsget("LotteRegdate")
					FItemList(i).FLotteLastUpdate	= rsget("LotteLastUpdate")
					FItemList(i).FLotteGoodNo		= rsget("LotteGoodNo")
					FItemList(i).FLotteTmpGoodNo	= rsget("LotteTmpGoodNo")
					FItemList(i).FLottePrice		= rsget("LottePrice")
					FItemList(i).FLotteSellYn		= rsget("LotteSellYn")
					FItemList(i).FregUserid			= rsget("regUserid")
					FItemList(i).FLotteStatCd		= rsget("lotteStatCd")
					FItemList(i).FCateMapCnt		= rsget("mapCnt")
					FItemList(i).Fdefaultdeliverytype = rsget("defaultdeliverytype")
	                FItemList(i).FdefaultfreeBeasongLimit = rsget("defaultfreeBeasongLimit")
	                FItemList(i).Fdeliverytype      = rsget("deliverytype")

					if Not(FItemList(i).FsmallImage="" or isNull(FItemList(i).FsmallImage)) then
						FItemList(i).FsmallImage = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("smallImage")
					else
						FItemList(i).FsmallImage = "http://fiximage.10x10.co.kr/images/spacer.gif"
					end if
					FItemList(i).FoptionCnt       = rsget("optionCnt")
	                FItemList(i).FrctSellCNT      = rsget("rctSellCNT")
	                FItemList(i).FaccFailCNT      = rsget("accFailCNT")
	                FItemList(i).FlastErrStr      = rsget("lastErrStr")
	                FItemList(i).FinfoDiv      		= rsget("infoDiv")
					FItemList(i).FItemoption		= rsget("itemoption")
					FItemList(i).FOptaddprice		= rsget("optaddprice")
					FItemList(i).FOptionname		= rsget("optionname")
					FItemList(i).FOptlimitno		= rsget("optlimitno")
					FItemList(i).FOptlimitsold		= rsget("optlimitsold")
					FItemList(i).FOptsellyn			= rsget("optsellyn")
	                FItemList(i).FRegedOptionname	= rsget("regedOptionname")
	                FItemList(i).FRegedItemname		= rsget("regedItemname")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end Sub

	'// �Ե����� ��ǰ ��� // ������ ������ �޶�� ��..
	public Sub getLotteRegedItemList
		Dim sqlStr, addSql, i

		if (FRectItemName <> "") then
			sqlStr = " select top 1000 B.itemid into #TMPSearchItem"
			sqlStr = sqlStr + " from [DBAPPWISH].db_AppWish.dbo.tbl_item_SearchBase B"
			if (FRectMakerid <> "") then
    			sqlStr = sqlStr + " Join [DBAPPWISH].[db_AppWish].dbo.tbl_item ai"
            	sqlStr = sqlStr + " on B.itemid=ai.itemid"
            	sqlStr = sqlStr + " and ai.makerid='"&FRectMakerid&"'"
	        end if
	        sqlStr = sqlStr + " where contains(B.searchKey,'""" + CStr(FRectItemName) + """') "
            sqlStr = sqlStr + " order by B.itemid desc "
            dbget.Execute sqlStr
		end if

		'��ǰ�� �˻�
		If FRectItemName <> "" Then
			''addSql = addSql & " and i.itemname like '%" & FRectItemName & "%'"
			''addSql = addSql & " and i.itemname = '" & FRectItemName & "'"
			addSql = addSql & " and i.itemid in (select itemid from #TMPSearchItem )"
		End if

		'�귣��˻�
		If FRectMakerid <> "" Then
			addSql = addSql & " and i.makerid='" & FRectMakerid & "'"
		End If

		'��ǰ�ڵ� �˻�
        If (FRectItemid <> "") then
            If Right(Trim(FRectItemid) ,1) = "," Then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            Else
				FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" + FRectItemid + ")"
            End If
        End If


		'�Ե����̸� ��ǰ��ȣ �˻�
        If (FRectLotteGoodNo <> "") then
            If Right(Trim(FRectLotteGoodNo) ,1) = "," Then
            	FRectLotteGoodNo = Replace(FRectLotteGoodNo,",,",",")
            	addSql = addSql & " and J.LotteGoodNo in ('" & replace(Left(FRectLotteGoodNo, Len(FRectLotteGoodNo)-1),",","','") & "')"
            Else
				FRectLotteGoodNo = Replace(FRectLotteGoodNo,",,",",")
            	addSql = addSql & " and J.LotteGoodNo in ('" & replace(FRectLotteGoodNo,",","','") & "')"
            End If
        End If

		'�Ե����� ������ ��ǰ��ȣ �˻�
        If (FRectLotteTmpGoodNo <> "") then
            If Right(Trim(FRectLotteTmpGoodNo) ,1) = "," Then
            	FRectLotteTmpGoodNo = Replace(FRectLotteTmpGoodNo,",,",",")
            	addSql = addSql & " and J.LotteTmpGoodNo in ('" & replace(Left(FRectLotteTmpGoodNo, Len(FRectLotteTmpGoodNo)-1),",","','") & "')"
            Else
				FRectLotteTmpGoodNo = Replace(FRectLotteTmpGoodNo,",,",",")
            	addSql = addSql & " and J.LotteTmpGoodNo in ('" & replace(FRectLotteTmpGoodNo,",","','") & "')"
            End If
        End If

		'ī�װ� �˻�
		If FRectCDL <> "" Then
			addSql = addSql & " and i.cate_large='" & FRectCDL & "'"
		End if
		If FRectCDM <> "" Then
			addSql = addSql & " and i.cate_mid='" & FRectCDM & "'"
		End if
		If FRectCDS <> "" Then
			addSql = addSql & " and i.cate_small='" & FRectCDS & "'"
		End If

		'��Ͽ��� �˻�
		Select Case FRectExtNotReg
			Case "Q"	''��Ͻ���
				addSql = addSql & " and J.LotteStatCd = '-1'"
			Case "W"	'��Ͽ���
				addSql = addSql & " and J.LotteStatCd = '00'"
			    addSql = addSql & " and J.LotteGoodNo is Null "
			Case "F"	'��ϿϷ�(�ӽ�)
			    addSql = addSql & " and J.LotteTmpGoodNo is Not Null "
				addSql = addSql & " and J.LotteStatCd <> '40'"
			    addSql = addSql & " and J.LotteGoodNo is Null "
			Case "C"	'�ݷ�
			    addSql = addSql & " and J.LotteStatCd = '40'"
			    addSql = addSql & " and J.LotteGoodNo is Null "
			Case "D"	'��ϿϷ�(����)
			    addSql = addSql & " and J.LotteTmpGoodNo is Not Null"
				addSql = addSql & " and J.LotteGoodNo is Not Null"
			Case "R"	'�������		'�����ٸ����� ���
				addSql = addSql & " and J.LotteLastUpdate < i.lastupdate"
				addSql = addSql & " and isnull(J.LotteGoodNo, '') <> '' "
		End Select

		'�̵�� ������ư Ŭ�� ��
		Select Case FRectIsReged
			Case "N"	'��Ͽ����̻�
			    addSql = addSql & " and J.itemid is NULL  and (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>5)) "
		End Select

		'�Ǹſ��� �˻�
		Select Case FRectSellYn
			Case "Y"	addSql = addSql & " and i.sellYn='Y'"			'�Ǹ�
			Case "N"	addSql = addSql & " and i.sellYn in ('S','N')"	'ǰ��
		End Select

		'�ٹ����� �������� �˻�
		If FRectLimitYn <> "" Then
			addSql = addSql & " and i.limitYn = '" & FRectLimitYn & "'"
		End If

		'�ٹ����� ���Ͽ��� �˻�
		If FRectSailYn <> "" Then
			addSql = addSql & " and i.sailYn = '" & FRectSailYn & "'"
		End If

		'������ �� ���� CMAXMARGIN �̻� �˻�
		If (FRectonlyValidMargin <> "") Then
			IF (FRectonlyValidMargin = "Y") Then
				'addSql = addSql & " and i.sellcash<>0"
				'addSql = addSql & " and i.sellcash - i.buycash > 0 "
				addSql = addSql & " and convert(int, ((i.sellcash-i.buycash)/(CASE WHEN i.sellcash=0 THEN 1 ELSE i.sellcash END))*100)>="&CMAXMARGIN & VbCrlf
			Else
				'addSql = addSql & " and i.sellcash<>0"
				'addSql = addSql & " and i.sellcash - i.buycash > 0 "
				addSql = addSql & " and convert(int, ((i.sellcash-i.buycash)/(CASE WHEN i.sellcash=0 THEN 1 ELSE i.sellcash END))*100)<"&CMAXMARGIN & VbCrlf
			End If
		End If

		If (FRectStartMargin <> "") OR (FRectEndMargin <> "") Then
			If (FRectStartMargin <> "") And (FRectEndMargin <> "") Then
				addSql = addSql & " and ("
				addSql = addSql & " 	convert(int, ((i.sellcash-i.buycash)/(CASE WHEN i.sellcash=0 THEN 1 ELSE i.sellcash END))*100)>="&FRectStartMargin & VbCrlf
				addSql = addSql & " 	and convert(int, ((i.sellcash-i.buycash)/(CASE WHEN i.sellcash=0 THEN 1 ELSE i.sellcash END))*100)<="&FRectEndMargin & VbCrlf
				addSql = addSql & " ) "
			ElseIf (FRectStartMargin <> "") And (FRectEndMargin = "") Then
				addSql = addSql & " and convert(int, ((i.sellcash-i.buycash)/(CASE WHEN i.sellcash=0 THEN 1 ELSE i.sellcash END))*100)>="&FRectStartMargin & VbCrlf
			ElseIf (FRectStartMargin = "") And (FRectEndMargin <> "") Then
				addSql = addSql & " and convert(int, ((i.sellcash-i.buycash)/(CASE WHEN i.sellcash=0 THEN 1 ELSE i.sellcash END))*100)<="&FRectEndMargin & VbCrlf
			End If
		End If

		'�ֹ����� ���� �˻�
		If FRectisMadeHand <> "" Then
			If (FRectisMadeHand = "Y") Then
				addSql = addSql & " and i.itemdiv in ('06', '16')" & VbCrlf
			Else
				addSql = addSql & " and i.itemdiv not in ('06', '16')" & VbCrlf
			End If
		End if

		'�ɼ� ���� �˻�
		If FRectIsOption <> "" Then
			If FRectIsOption = "optAll" Then			'�ɼ���ü
				addSql = addSql & " and i.optioncnt > 0"
			ElseIf FRectIsOption = "optaddpricey" Then	'�߰��ݾ�Y
				addSql = addSql & " and i.optioncnt > 0"
 				addSql = addSql & " and J.optAddPrcCnt > 0"
			ElseIf FRectIsOption = "optaddpricen" Then	'�߰��ݾ�N
				addSql = addSql & " and i.optioncnt > 0"
				addSql = addSql & " and isNULL(J.optAddPrcCnt,0)=0"
			ElseIf FRectIsOption = "optN" Then			'��ǰ
				addSql = addSql & " and i.optioncnt = 0"
			ElseIf FRectIsOption = "optAddPrcRegType" Then
				addSql = addSql & " and J.optAddPrcCnt > 0"
				addSql = addSql & " and J.optAddPrcRegType = 0"
			End If
		End If

		'�ٹ����� ǰ������ �˻�
		If (FRectInfoDiv <> "") then
			If (FRectInfoDiv = "YY") Then
				addSql = addSql & " and isNULL(ct.infodiv,'')<>''"
			ElseIf (FRectInfoDiv = "NN") Then
				addSql = addSql & " and isNULL(ct.infodiv,'')=''"
			Else
				addSql = addSql & " and ct.infodiv = '"&FRectInfoDiv&"'"
			End If
		End If

		'�ٹ����� ������� �귣�� ���� �˻�
		If (FRectNotinmakerid <> "") then
			If (FRectNotinmakerid = "Y") Then
				addSql = addSql & " and exists(SELECT top 1 n.makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid n with (nolock) WHERE n.makerid=i.makerid and n.mallgubun = 'lotteCom') "
			ElseIf (FRectNotinmakerid = "N") Then
				addSql = addSql & " and not exists(SELECT top 1 n.makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid n with (nolock) WHERE n.makerid=i.makerid and n.mallgubun = 'lotteCom') "
			End If
		End If

		'�ٹ����� ������� ��ǰ ���� �˻�
		If (FRectNotinitemid <> "") then
			If (FRectNotinitemid = "Y") Then
				addSql = addSql & " and exists(SELECT top 1 n.itemid FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid n with (nolock) WHERE n.itemid=i.itemid and n.mallgubun = 'lotteCom') "
			ElseIf (FRectNotinitemid = "N") Then
				addSql = addSql & " and not exists(SELECT top 1 n.itemid FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid n with (nolock) WHERE n.itemid=i.itemid and n.mallgubun = 'lotteCom') "
			End If
		End If

		'���޸� �������� ��ǰ �˻�
		If (FRectExcTrans <> "") then
			If (FRectExcTrans = "Y") Then
				addSql = addSql & " and 'Y' = (CASE WHEN i.isusing='N' "
				addSql = addSql & " or i.makerid in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='lotteCom') "
				addSql = addSql & " or i.itemid in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='lotteCom') "
				addSql = addSql & " or i.isExtUsing='N' "
				addSql = addSql & " or uc.isExtUsing='N' "
				addSql = addSql & " or i.deliveryType = 7 "
				addSql = addSql & " or ((i.deliveryType = 9) and (i.sellcash < 10000)) "
				addSql = addSql & " or i.itemdiv = '21' "
				addSql = addSql & " or i.deliverfixday in ('C','X') "
				addSql = addSql & " or i.itemdiv >= 50 "
				addSql = addSql & " or i.itemdiv = '08' "
				addSql = addSql & " or i.itemdiv = '09' "
				addSql = addSql & " or i.cate_large = '999' "
				addSql = addSql & " or i.cate_large='' "
				addSql = addSql & " or not (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>5)) "
				addSql = addSql & " or not ( "
				addSql = addSql & " 	i.optioncnt = 0 "
				addSql = addSql & " 	or "
				addSql = addSql & " 	exists(SELECT top 1 o.itemid FROM [db_item].[dbo].tbl_item_option o WHERE o.isUsing='Y' and o.optsellyn='Y' and o.itemid=i.itemid and (o.optlimityn <> 'Y' or (o.optlimitno-o.optlimitsold)>5)) "
				addSql = addSql & " ) "
				addSql = addSql & " THEN 'Y' ELSE 'N' END) "
			ElseIf (FRectExcTrans = "F") Then
				addSql = addSql & " and i.makerid not in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='lotteCom') "
				addSql = addSql & " and i.itemid not in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='lotteCom') "
				addSql = addSql & " and i.isusing='Y' "
				addSql = addSql & " and i.isExtUsing='Y' "											'// �ܺθ�����ǰ
				addSql = addSql & " and uc.isExtUsing='Y' "
				addSql = addSql & " and i.deliveryType <> 7 "										'// ��ü����
				addSql = addSql & " and i.itemdiv <> '21' "											'// ����ǰ
				addSql = addSql & " and i.deliverfixday not in ('C','X') "							'// �ɹ��, ȭ�����
				addSql = addSql & " and not ((i.deliveryType = 9) and (i.sellcash < 10000)) "		'// �ǸŰ�(���ΰ�) 1���� �̸�
				addSql = addSql & " and i.itemdiv <> '08' "											'// Ƽ��(����) ��ǰ
				addSql = addSql & " and i.itemdiv <> '09' "											'// Present��ǰ
				addSql = addSql & " and i.itemdiv < 50 "
				addSql = addSql & " and (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>5)) "
				addSql = addSql & " and ( "
				addSql = addSql & " 	i.optioncnt = 0 "
				addSql = addSql & " 	or "
				addSql = addSql & " 	exists(SELECT top 1 o.itemid FROM [db_item].[dbo].tbl_item_option o WHERE o.isUsing='Y' and o.optsellyn='Y' and o.itemid=i.itemid and (o.optlimityn <> 'Y' or (o.optlimitno-o.optlimitsold)>5)) "
				addSql = addSql & " ) "
				addSql = addSql & " and 'Y' = (CASE WHEN i.cate_large = '999' "
				addSql = addSql & " or i.cate_large='' "
				addSql = addSql & " or J.accFailCnt > 0 "
				addSql = addSql & " THEN 'Y' ELSE 'N' END) "
			ElseIf (FRectExcTrans = "N") Then
				addSql = addSql & " and not exists(SELECT top 1 n.makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid n with (nolock) WHERE n.makerid=i.makerid and n.mallgubun = 'lotteCom') "
				addSql = addSql & " and not exists(SELECT top 1 n.itemid FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid n with (nolock) WHERE n.itemid=i.itemid and n.mallgubun = 'lotteCom') "
				addSql = addSql & " and i.isusing='Y' "
				addSql = addSql & " and i.isExtUsing='Y' "											'// �ܺθ�����ǰ
				addSql = addSql & " and uc.isExtUsing='Y' "
				addSql = addSql & " and i.deliveryType <> 7 "										'// ��ü����
				addSql = addSql & " and i.itemdiv <> '21' "											'// ����ǰ
				addSql = addSql & " and i.deliverfixday not in ('C','X') "							'// �ɹ��, ȭ�����
				addSql = addSql & " and not ((i.deliveryType = 9) and (i.sellcash < 10000)) "		'// �ǸŰ�(���ΰ�) 1���� �̸�
				addSql = addSql & " and i.itemdiv <> '08' "											'// Ƽ��(����) ��ǰ
				addSql = addSql & " and i.itemdiv <> '09' "											'// Present��ǰ
				addSql = addSql & " and i.cate_large <> '999' "										'// ī�װ� ������
				addSql = addSql & " and i.cate_large <> '' "										'// ī�װ� ������
				addSql = addSql & " and i.itemdiv < 50 "
				addSql = addSql & " and (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>5)) "
				addSql = addSql & " and ( "
				addSql = addSql & " 	i.optioncnt = 0 "
				addSql = addSql & " 	or "
				addSql = addSql & " 	exists(SELECT top 1 o.itemid FROM [db_item].[dbo].tbl_item_option o WHERE o.isUsing='Y' and o.optsellyn='Y' and o.itemid=i.itemid and (o.optlimityn <> 'Y' or (o.optlimitno-o.optlimitsold)>5)) "
				addSql = addSql & " ) "
				addSql = addSql & " and ( "
				addSql = addSql & " 	i.optioncnt = 0 "
				addSql = addSql & " 	or "
				addSql = addSql & " 	not exists(SELECT top 1 o.itemid FROM [db_item].[dbo].tbl_item_option o WHERE o.isUsing='Y' and o.itemid=i.itemid and o.optAddPrice > 0) "
				addSql = addSql & " ) "
				''addSql = addSql & " and  G.optAddPrcCnt = 0 "
				addSql = addSql & " and not (i.optionCnt > 0 and IsNull(J.regedOptCnt,0) = 0) "
			End If
		End If

        'Ư�� ��ǰ ����
        If (FRectIsSpecialPrice <> "") then
            If (FRectIsSpecialPrice = "Y") Then
				addSql = addSql & " and (GETDATE() > mi.startDate and GETDATE() <= mi.endDate) "
            End If
        End If

		'�ɼ��߰��ݾ�New
		If (FRectPriceOption <> "") then
			If (FRectPriceOption = "Y") Then
				addSql = addSql & " and i.itemid in (SELECT itemid FROM db_item.[dbo].[tbl_const_OptAddPrice_Exists]) "
			ElseIf (FRectPriceOption = "N") Then
				addSql = addSql & " and i.itemid not in (SELECT itemid FROM db_item.[dbo].[tbl_const_OptAddPrice_Exists]) "
			End If
		End If

		'�Ե����̸� �Ǹſ���
		If (FRectExtSellYn<>"") then
			If (FRectExtSellYn = "YN") Then
				addSql = addSql & " and J.LotteSellYn <> 'X'"
			Else
				addSql = addSql & " and J.LotteSellYn='" & FRectExtSellYn & "'"
			End if
		End If

		'��ϼ���������ǰ
		Select Case FRectFailCntExists
			Case "Y"	'����1ȸ�̻�
				addSql = addSql & " and J.accFailCNT>0"
			Case "N"	'����0ȸ
				addSql = addSql & " and J.accFailCNT=0"
		End Select

		'�Ե����̸� ī�װ� ��Ī ����
		Select Case FRectMatchCate
			Case "Y"	'��Ī�Ϸ�
				addSql = addSql & " and isnull(c.mapCnt, '') <> ''"
			Case "N"	'�̸�Ī
				addSql = addSql & " and isnull(c.mapCnt, '') = ''"
		End Select

        '�Ե����̸� < 10x10 ����
		If (FRectexpensive10x10 <> "") Then
			addSql = addSql & " and J.LottePrice is Not Null and J.LottePrice < i.sellcash"
		End If

		'���ݻ�����ü����
		If FRectdiffPrc <> "" Then
			addSql = addSql & " and J.LottePrice is Not Null and i.sellcash <> J.LottePrice "
		End If

		'�Ե����̸��Ǹ�,  10x10 ǰ��
		If (FRectLotteYes10x10No <> "") Then
			addSql = addSql & " and i.sellyn<>'Y'"
			addSql = addSql & " and J.LotteSellYn='Y'"
		End If

		'�Ե����̸�ǰ��&�ٹ������ǸŰ���(�Ǹ���,����>=10) ��ǰ����
		If FRectLotteNo10x10Yes <> "" Then
			addSql = addSql & " and (J.LotteSellYn= 'N' and i.sellyn='Y' and (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>"&CMAXLIMITSELL&")))"
		End If

		'���������ǰ����(����������Ʈ�� ����)
		If FRectReqEdit <> "" Then
			addSql = addSql & " and J.LotteLastUpdate < i.lastupdate "
		End If

		'�����ٸ����� ��� ����Ƚ�� ����
		If (FRectFailCntOverExcept <> "") Then
			addSql = addSql & " and J.accFailCNT < "&FRectFailCntOverExcept
		End If

		'�����ٸ����� ��� ��Ʈ������Ʈ ���� ����
		If (FRectOrdType = "LU") Then
		    addSql = addSql & " and isnull(J.lastStatCheckDate,'') = '' "
		    addSql = addSql & " and Left(i.lastupdate, 10) <> Left(J.LotteLastUpdate, 10) "
		End If

		'��۱���
		If (FRectDeliverytype <> "") Then
			addSql = addSql & " and i.deliverytype='" & FRectDeliverytype & "'"
		End If

		'�ŷ�����
		If FRectMWDiv = "MW" Then
			addSql = addSql & " and (i.mwdiv='M' or i.mwdiv='W')"
		ElseIf FRectMWDiv<>"" Then
			addSql = addSql & " and i.mwdiv='"& FRectMWDiv & "'"
		End If

		'���� ��� ����
		If (FRectIsextusing <> "") Then
			addSql = addSql & " and i.isextusing='" & FRectIsextusing & "'"
		End If

		sqlStr = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED; "
		sqlStr = sqlStr & " SELECT count(i.itemid) as cnt, CEILING(CAST(Count(i.itemid) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i "
		If (FRectIsReged = "N") OR (FRectIsReged = "A") Then		'//�̵���� �ƴϸ� JOIN
			sqlStr = sqlStr & "	LEFT JOIN db_item.dbo.tbl_lotte_regItem as J on J.itemid = i.itemid "
		Else
			sqlStr = sqlStr & "	JOIN db_item.dbo.tbl_lotte_regItem as J on J.itemid = i.itemid "
		End If
		sqlStr = sqlStr & "	LEFT Join db_item.dbo.tbl_OutMall_CateMap_Summary as c on c.mallid = 'lotteCom' and c.tenCateLarge = i.cate_large and c.tenCateMid = i.cate_mid and c.tenCateSmall = i.cate_small "
		sqlStr = sqlStr & " LEFT join db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_contents as ct on i.itemid = ct.itemid"
		sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_outmall_mustPriceItem as mi with (nolock) on mi.itemid = i.itemid and mi.mallgubun = '"& CMALLNAME &"' "
		sqlStr = sqlStr & " WHERE 1 = 1"
		If (FRectIsReged <> "N" and FRectExtNotReg <> "Q")  Then		'// �̵�ϵ� �ƴϰ� ��Ͻ��е� �ƴϸ� ���� ����
			If FRectIsReged = "Q" Then
				sqlStr = sqlStr & " and J.LotteGoodNo is Not Null "
				sqlStr = sqlStr & " and (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>5)) "
				sqlStr = sqlStr & " and 'N' = (CASE WHEN i.isusing='N'  "
				sqlStr = sqlStr & " or i.isExtUsing='N' "
				sqlStr = sqlStr & " or uc.isExtUsing='N' "
				sqlStr = sqlStr & " or ((i.deliveryType = 9) and (i.sellcash < 10000)) "
				sqlStr = sqlStr & " or i.sellyn<>'Y' "
				sqlStr = sqlStr & " or i.deliverfixday in ('C','X') "
				sqlStr = sqlStr & " or i.itemdiv >= 50 or i.itemdiv = '08' or i.cate_large = '999' or i.cate_large='' "
				sqlStr = sqlStr & " or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "
				sqlStr = sqlStr & " or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "
				sqlStr = sqlStr & " THEN 'Y' ELSE 'N' END) "
			End If
		Else
    		'sqlStr = sqlStr & " and i.isusing='Y' "
    		'sqlStr = sqlStr & " and i.deliverfixday not in ('C','X') "
    		'sqlStr = sqlStr & " and i.basicimage is not null "
    		'sqlStr = sqlStr & " and i.itemdiv<50 "  ''and i.itemdiv<>'08'
    		'sqlStr = sqlStr & " and i.itemdiv not in ('08','09')"
    		'sqlStr = sqlStr & " and i.cate_large<>'' "
		    'sqlStr = sqlStr & " and ((i.cate_large <> '999') or ((i.cate_large='999')and(i.makerid='ftroupe'))) " & VBCRLF
    		'sqlStr = sqlStr & "	and i.makerid not in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "	'������� �귣��
    		'sqlStr = sqlStr & "	and i.itemid not in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "		'������� ��ǰ
    		'sqlStr = sqlStr & "	and uc.isExtUsing='Y'"  ''20130304 �귣�� ���޻�뿩�� Y��.
			'sqlStr = sqlStr & " and i.isExtUsing='Y'"														'//���޸� �ǸŸ� ���
			'sqlStr = sqlStr & " and i.deliverytype not in ('7')"											'//���ҹ�� ��ǰ ����

			If FRectExtNotReg <> "" Then
				sqlStr = sqlStr & " and i.sellcash>=1000 "  & VBCRLF
				'sqlStr = sqlStr & " and i.itemdiv<>'06'" & VBCRLF				'�ֹ�����
			End If

			sqlStr = sqlStr & " and NOT EXISTS (select 1 from db_etcmall.dbo.tbl_outmall_const_Except_item ex where ex.itemid=i.itemid)"
			sqlStr = sqlStr & " and NOT EXISTS (select 1 from db_etcmall.dbo.tbl_outmall_const_Except_item_lottecom ex where ex.itemid=i.itemid)"


		End If
		sqlStr = sqlStr & addSql
		''response.write sqlStr
		''response.end
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'������������ ��ü ���������� Ŭ �� �Լ�����
		if Cint(FCurrPage)>Cint(FTotalPage) then
			FResultCount = 0
			exit sub
		end if

		sqlStr = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED; "
		sqlStr = sqlStr & " SELECT top " + CStr(FPageSize*FCurrPage) + " i.itemid, i.itemname, i.smallImage "
		sqlStr = sqlStr & "	, i.makerid, i.regdate, i.lastUpdate, i.orgPrice, i.orgSuplyCash, i.sellcash, i.buycash"
		sqlStr = sqlStr & "	, i.sellYn, i.sailyn, i.LimitYn, i.LimitNo, i.LimitSold, i.deliverytype, i.optionCnt"
		sqlStr = sqlStr & "	, J.LotteRegdate, J.LotteLastUpdate, J.LotteGoodNo, J.LotteTmpGoodNo, J.LottePrice, J.LotteSellYn, J.regUserid, IsNULL(J.LotteStatCd,-9) as LotteStatCd "
		sqlStr = sqlStr & "	, c.mapCnt, J.regedOptCnt, J.rctSellCNT, J.accFailCNT, J.lastErrStr "
		sqlStr = sqlStr & " ,uc.defaultdeliverytype, uc.defaultfreeBeasongLimit"
		sqlStr = sqlStr & "	, Ct.infoDiv, J.optAddPrcCnt, J.optAddPrcRegType, i.itemdiv, isnull(J.lastcateChgDate, '') as lastcateChgDate, mi.mustPrice as specialPrice, mi.startDate, mi.endDate "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i "
		If (FRectIsReged = "N") OR (FRectIsReged = "A") Then		'//�̵���� �ƴϸ� JOIN
			sqlStr = sqlStr & "	LEFT JOIN db_item.dbo.tbl_lotte_regItem as J on J.itemid = i.itemid "
		Else
			sqlStr = sqlStr & "	JOIN db_item.dbo.tbl_lotte_regItem as J on J.itemid = i.itemid "
		End If
		sqlStr = sqlStr & " LEFT JOIN db_item.dbo.tbl_OutMall_CateMap_Summary as c on c.mallid = 'lotteCom' and c.tenCateLarge = i.cate_large and c.tenCateMid = i.cate_mid and c.tenCateSmall = i.cate_small "
		sqlStr = sqlStr & "	LEFT JOIN db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_contents as ct on i.itemid = ct.itemid"
		sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_outmall_mustPriceItem as mi with (nolock) on mi.itemid = i.itemid and mi.mallgubun = '"& CMALLNAME &"' "
		sqlStr = sqlStr & " where 1 = 1"
		If (FRectIsReged <> "N" and FRectExtNotReg <> "Q")  Then		'// �̵�ϵ� �ƴϰ� ��Ͻ��е� �ƴϸ� ���� ����
			If FRectIsReged = "Q" Then
				sqlStr = sqlStr & " and J.LotteGoodNo is Not Null "
				sqlStr = sqlStr & " and (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>5)) "
				sqlStr = sqlStr & " and 'N' = (CASE WHEN i.isusing='N'  "
				sqlStr = sqlStr & " or i.isExtUsing='N' "
				sqlStr = sqlStr & " or uc.isExtUsing='N' "
				sqlStr = sqlStr & " or ((i.deliveryType = 9) and (i.sellcash < 10000)) "
				sqlStr = sqlStr & " or i.sellyn<>'Y' "
				sqlStr = sqlStr & " or i.deliverfixday in ('C','X') "
				sqlStr = sqlStr & " or i.itemdiv >= 50 or i.itemdiv = '08' or i.cate_large = '999' or i.cate_large='' "
				sqlStr = sqlStr & " or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "
				sqlStr = sqlStr & " or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "
				sqlStr = sqlStr & " THEN 'Y' ELSE 'N' END) "
			End If
		Else
    		'sqlStr = sqlStr & " and i.isusing='Y' "
    		'sqlStr = sqlStr & " and i.deliverfixday not in ('C','X') "
    		'sqlStr = sqlStr & " and i.basicimage is not null "
    		'sqlStr = sqlStr & " and i.itemdiv<50 "  ''and i.itemdiv<>'08'
    		'sqlStr = sqlStr & " and i.itemdiv not in ('08','09')"
    		'sqlStr = sqlStr & " and i.cate_large<>'' "
		    'sqlStr = sqlStr & " and ((i.cate_large <> '999') or ((i.cate_large='999')and(i.makerid='ftroupe'))) " & VBCRLF
    		'sqlStr = sqlStr & "	and i.makerid not in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "	'������� �귣��
    		'sqlStr = sqlStr & "	and i.itemid not in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "		'������� ��ǰ
    		'sqlStr = sqlStr & "	and uc.isExtUsing='Y'"  ''20130304 �귣�� ���޻�뿩�� Y��.
			'sqlStr = sqlStr & " and i.isExtUsing='Y'"														'//���޸� �ǸŸ� ���
			'sqlStr = sqlStr & " and i.deliverytype not in ('7')"											'//���ҹ�� ��ǰ ����

			If FRectExtNotReg <> "" Then
				sqlStr = sqlStr & " and i.sellcash>=1000 "  & VBCRLF
				'sqlStr = sqlStr & " and i.itemdiv<>'06'" & VBCRLF				'�ֹ�����
			End If


			sqlStr = sqlStr & " and NOT EXISTS (select 1 from db_etcmall.dbo.tbl_outmall_const_Except_item ex where ex.itemid=i.itemid)"
			sqlStr = sqlStr & " and NOT EXISTS (select 1 from db_etcmall.dbo.tbl_outmall_const_Except_item_lottecom ex where ex.itemid=i.itemid)"
    	End If
		sqlStr = sqlStr & addSql
		If (FRectOrdType = "LS") AND (FRectLotteNotReg = "F") Then
			sqlStr = sqlStr & " ORDER BY J.lastStatCheckDate, J.LotteLastupdate"
		ElseIf (FRectLotteNotReg = "F") Then
		    sqlStr = sqlStr & " ORDER BY J.LotteLastupdate "
		ElseIf (FRectOrdType = "B") Then
		    sqlStr = sqlStr & " ORDER BY i.itemscore DESC, i.itemid DESC "
		ElseIf (FRectOrdType = "BM") Then
		    sqlStr = sqlStr & " ORDER BY J.rctSellCNT DESC, i.itemscore DESC, J.itemid DESC"
		Else
		    sqlStr = sqlStr & " ORDER BY i.itemid DESC" '' m.regdate desc
	    End If
		rsget.pagesize = FPageSize
'rw sqlStr
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CLotteItem

				FItemList(i).Fitemid			= rsget("itemid")
				FItemList(i).Fitemname			= db2html(rsget("itemname"))
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
				FItemList(i).FLimitYn			= rsget("LimitYn")
				FItemList(i).FLimitNo			= rsget("LimitNo")
				FItemList(i).FLimitSold			= rsget("LimitSold")

				FItemList(i).FLotteRegdate		= rsget("LotteRegdate")
				FItemList(i).FLotteLastUpdate	= rsget("LotteLastUpdate")
				FItemList(i).FLotteGoodNo		= rsget("LotteGoodNo")
				FItemList(i).FLotteTmpGoodNo	= rsget("LotteTmpGoodNo")
				FItemList(i).FLottePrice		= rsget("LottePrice")
				FItemList(i).FLotteSellYn		= rsget("LotteSellYn")
				FItemList(i).FregUserid			= rsget("regUserid")
				FItemList(i).FLotteStatCd		= rsget("lotteStatCd")
				FItemList(i).FCateMapCnt		= rsget("mapCnt")
				FItemList(i).Fdefaultdeliverytype = rsget("defaultdeliverytype")
                FItemList(i).FdefaultfreeBeasongLimit = rsget("defaultfreeBeasongLimit")
                FItemList(i).Fdeliverytype      = rsget("deliverytype")

				if Not(FItemList(i).FsmallImage="" or isNull(FItemList(i).FsmallImage)) then
					FItemList(i).FsmallImage = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("smallImage")
				else
					FItemList(i).FsmallImage = "http://fiximage.10x10.co.kr/images/spacer.gif"
				end if
				FItemList(i).FoptionCnt       = rsget("optionCnt")
                FItemList(i).FrctSellCNT      = rsget("rctSellCNT")
                FItemList(i).FregedOptCnt     = rsget("regedOptCnt")
                FItemList(i).FaccFailCNT      = rsget("accFailCNT")
                FItemList(i).FlastErrStr      = rsget("lastErrStr")

                FItemList(i).FinfoDiv      		= rsget("infoDiv")
                FItemList(i).FoptAddPrcCnt      = rsget("optAddPrcCnt")
                FItemList(i).FoptAddPrcRegType  = rsget("optAddPrcRegType")
                FItemList(i).FLastcateChgDate  = rsget("lastcateChgDate")
				FItemList(i).FSpecialPrice		= rsget("specialPrice")
				FItemList(i).FStartDate	      	= rsget("startDate")
				FItemList(i).FEndDate			= rsget("endDate")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close

		if (FRectItemName <> "") then
            sqlStr = " drop table #TMPSearchItem"
			dbget.Execute sqlStr
        end if
	end Sub

    ''' ��ϵ��� ���ƾ� �� ��ǰ..
    public Sub getLotteReqExpireItemList
		dim sqlStr, addSql, i

		sqlStr = " select count(i.itemid) as cnt, CEILING(CAST(Count(i.itemid) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr + " from db_item.dbo.tbl_item as i "
		sqlStr = sqlStr + " 	join db_item.dbo.tbl_lotte_regItem as m "
		sqlStr = sqlStr + " 		on i.itemid=m.itemid "
		sqlStr = sqlStr + " 		and m.LotteGoodNo is Not Null"
		sqlStr = sqlStr + " 	join db_item.dbo.tbl_item_contents as ct "
		sqlStr = sqlStr + " 	on i.itemid=ct.itemid"
		sqlStr = sqlStr + " 	left join db_user.dbo.tbl_user_c uc"
		sqlStr = sqlStr + " 	on i.makerid=uc.userid"
		sqlStr = sqlStr + " where 1=1"
		sqlStr = sqlStr + "     and (i.isusing<>'Y' or i.isExtUsing<>'Y' "
		sqlStr = sqlStr + "     or i.deliverytype in ('7') "
        sqlStr = sqlStr + "     or ((i.deliveryType=9) and (i.sellcash<10000))"
		sqlStr = sqlStr + "     or i.deliverfixday in ('C','X') "
		sqlStr = sqlStr + "     or i.itemdiv>=50 or i.itemdiv='08' or i.cate_large='999' or i.cate_large=''"
		sqlStr = sqlStr + "		or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "	'������� �귣��
		sqlStr = sqlStr + "		or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "		'������� ��ǰ
		sqlStr = sqlStr & "		or ((i.LimitYn='Y') and (i.LimitNo-i.LimitSold<"&CMAXLIMITSELL&")) "
		sqlStr = sqlStr + "		or uc.isExtUsing='N'"
        sqlStr = sqlStr + "	)"
        sqlStr = sqlStr & " and i.itemid not in ("
        sqlStr = sqlStr & "     select itemid from db_item.dbo.tbl_OutMall_etcLink"
        sqlStr = sqlStr & "     where stDt<getdate()"
        sqlStr = sqlStr & "     and edDt>getdate()"
        sqlStr = sqlStr & "     and mallid='"&CMALLNAME&"'"
        sqlStr = sqlStr & "     and linkgbn='donotEdit'"
        sqlStr = sqlStr & " )"

        if FRectMakerid<>"" then
			sqlStr = sqlStr & " and i.makerid='" & FRectMakerid & "'"
		end if

		if FRectItemID<>"" then
			sqlStr = sqlStr & " and i.itemid in (" & FRectItemID & ")"
		end if

		if (FRectExtSellYn<>"") then
		    if (FRectExtSellYn="YN") then
		        sqlStr = sqlStr + " and m.lotteSellYn<>'X'"
		    else
		        sqlStr = sqlStr + " and m.lotteSellYn='" & FRectExtSellYn & "'"
		    end if
		end if

		Select Case FRectSellYn
			Case "Y"	'�Ǹ�
				sqlStr = sqlStr & " and i.sellYn='Y'"
			Case "N"	'ǰ��
				sqlStr = sqlStr & " and i.sellYn in ('S','N')"
		End Select

		if (FRectInfoDivYn<>"") then
		    if FRectInfoDivYn="Y" then
		        sqlStr = sqlStr + " and isNULL(ct.infoDiv,'')<>''"
		    elseif FRectInfoDivYn="N" then
    		    sqlStr = sqlStr + " and isNULL(ct.infoDiv,'')=''"
    		else
    		    sqlStr = sqlStr + " and ct.infoDiv='"&FRectInfoDivYn&"'"
    		end if
		end if

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
		sqlStr = sqlStr + "		, m.LotteRegdate, m.LotteLastUpdate, m.LotteGoodNo, m.LotteTmpGoodNo, m.LottePrice, m.LotteSellYn, m.regUserid, m.lotteStatCd "
		sqlStr = sqlStr + "		, 1 as mapCnt, m.rctSellCNT "
		sqlStr = sqlStr + "		, ct.infoDiv"
		sqlStr = sqlStr + "     ,uc.defaultdeliverytype, uc.defaultfreeBeasongLimit"
		sqlStr = sqlStr + " from db_item.dbo.tbl_item as i "
		sqlStr = sqlStr + " 	join db_item.dbo.tbl_lotte_regItem as m "
		sqlStr = sqlStr + " 		on i.itemid=m.itemid "
		sqlStr = sqlStr + " 		and m.LotteGoodNo is Not Null"
		sqlStr = sqlStr + " 	join db_item.dbo.tbl_item_contents as ct "
		sqlStr = sqlStr + " 	on i.itemid=ct.itemid"
		sqlStr = sqlStr + " 	left join db_user.dbo.tbl_user_c uc"
		sqlStr = sqlStr + " 	on i.makerid=uc.userid"
		sqlStr = sqlStr + " where 1=1"
		sqlStr = sqlStr + "     and (i.isusing<>'Y' or i.isExtUsing<>'Y' "
		sqlStr = sqlStr + "     or i.deliverytype in ('7') "
        sqlStr = sqlStr + "     or ((i.deliveryType=9) and (i.sellcash<10000))"
		sqlStr = sqlStr + "     or i.deliverfixday in ('C','X') "
		sqlStr = sqlStr + "     or i.itemdiv>=50 or i.itemdiv='08' or i.cate_large='999' or i.cate_large=''"
		sqlStr = sqlStr + "		or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "	'������� �귣��
		sqlStr = sqlStr + "		or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "		'������� ��ǰ
		sqlStr = sqlStr & "		or ((i.LimitYn='Y') and (i.LimitNo-i.LimitSold<"&CMAXLIMITSELL&")) "
		sqlStr = sqlStr + "		or uc.isExtUsing='N'"
        sqlStr = sqlStr + "	)"
        sqlStr = sqlStr & " and i.itemid not in ("
        sqlStr = sqlStr & "     select itemid from db_item.dbo.tbl_OutMall_etcLink"
        sqlStr = sqlStr & "     where stDt<getdate()"
        sqlStr = sqlStr & "     and edDt>getdate()"
        sqlStr = sqlStr & "     and mallid='"&CMALLNAME&"'"
        sqlStr = sqlStr & "     and linkgbn='donotEdit'"
        sqlStr = sqlStr & " )"

        if FRectMakerid<>"" then
			sqlStr = sqlStr & " and i.makerid='" & FRectMakerid & "'"
		end if

		if FRectItemID<>"" then
			sqlStr = sqlStr & " and i.itemid in (" & FRectItemID & ")"
		end if

		if (FRectExtSellYn<>"") then
		    if (FRectExtSellYn="YN") then
		        sqlStr = sqlStr + " and m.lotteSellYn<>'X'"
		    else
		        sqlStr = sqlStr + " and m.lotteSellYn='" & FRectExtSellYn & "'"
		    end if
		end if

		Select Case FRectSellYn
			Case "Y"	'�Ǹ�
				sqlStr = sqlStr & " and i.sellYn='Y'"
			Case "N"	'ǰ��
				sqlStr = sqlStr & " and i.sellYn in ('S','N')"
		End Select

		if (FRectInfoDivYn<>"") then
		    if FRectInfoDivYn="Y" then
		        sqlStr = sqlStr + " and isNULL(ct.infoDiv,'')<>''"
		    elseif FRectInfoDivYn="N" then
    		    sqlStr = sqlStr + " and isNULL(ct.infoDiv,'')=''"
    		else
    		    sqlStr = sqlStr + " and ct.infoDiv='"&FRectInfoDivYn&"'"
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
				set FItemList(i) = new CLotteItem

				FItemList(i).Fitemid			= rsget("itemid")
				FItemList(i).Fitemname			= db2html(rsget("itemname"))
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
				FItemList(i).FLimitYn			= rsget("LimitYn")
				FItemList(i).FLimitNo			= rsget("LimitNo")
				FItemList(i).FLimitSold			= rsget("LimitSold")

				FItemList(i).FLotteRegdate		= rsget("LotteRegdate")
				FItemList(i).FLotteLastUpdate	= rsget("LotteLastUpdate")
				FItemList(i).FLotteGoodNo		= rsget("LotteGoodNo")
				FItemList(i).FLotteTmpGoodNo	= rsget("LotteTmpGoodNo")
				FItemList(i).FLottePrice		= rsget("LottePrice")
				FItemList(i).FLotteSellYn		= rsget("LotteSellYn")
				FItemList(i).FregUserid			= rsget("regUserid")
				FItemList(i).FLotteStatCd		= rsget("lotteStatCd")
				FItemList(i).FCateMapCnt		= rsget("mapCnt")
                FItemList(i).Fdeliverytype      = rsget("deliverytype")
                FItemList(i).Fdefaultdeliverytype = rsget("defaultdeliverytype")
                FItemList(i).FdefaultfreeBeasongLimit = rsget("defaultfreeBeasongLimit")

				if Not(FItemList(i).FsmallImage="" or isNull(FItemList(i).FsmallImage)) then
					FItemList(i).FsmallImage = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("smallImage")
				else
					FItemList(i).FsmallImage = "http://fiximage.10x10.co.kr/images/spacer.gif"
				end if
                FItemList(i).FrctSellCNT      = rsget("rctSellCNT")
                FItemList(i).FinfoDiv       = rsget("infoDiv")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end Sub

	'--------------------------------------------------------------------------------
	'// �̵�� ��ǰ ���(��Ͽ�)
	public Sub getLotteNotRegItemList
		dim sqlStr, addSql, i

		if FRectItemID<>"" then
			addSql = addSql & " and i.itemid in (" & FRectItemID & ")"

			''' �ɼ� �߰��ݾ� �ִ°�� ��� �Ұ�. //�ɼ� ��ü ǰ���� ��� ��� �Ұ�. 20120723�߰�
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
		end if

		strSql = "Select top " & FPageSize & " i.* "
		strSql = strSql & "		, c.keywords, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent, isNULL(c.requireMakeDay,0) as requireMakeDay "
		''strSql = strSql & "		, B.returnCode"
		strSql = strSql & "		, C.infoDiv,isNULL(C.safetyyn,'N') as safetyyn,isNULL(C.safetyDiv,0) as safetyDiv,C.safetyNum " '' ǰ������ �� ������������. �߰� 20121102
		strSql = strSql & " From db_item.dbo.tbl_item as i "
		strSql = strSql & " 	join db_item.dbo.tbl_item_contents as c "
		strSql = strSql & " 		on i.itemid=c.itemid "
		strSql = strSql & " 	left Join (Select tenCateLarge, tenCateMid, tenCateSmall, count(*) as mapCnt From db_item.dbo.tbl_lotte_cate_mapping Group by tenCateLarge, tenCateMid, tenCateSmall ) as cm "
		strSql = strSql & " 		on cm.tenCateLarge=i.cate_large and cm.tenCateMid=i.cate_mid and cm.tenCateSmall=i.cate_small "
		''strSql = strSql & " 	left Join db_item.dbo.tbl_OutMall_BrandReturnCode B"
		''strSql = strSql & " 	    on i.makerid=B.makerid and B.mallid='"&CMALLNAME&"'"
		strSql = strSql & " Where i.isusing='Y' "
		strSql = strSql & "     and i.isExtUsing='Y' "
		strSql = strSql & "     and i.deliverytype not in ('7')"
		strSql = strSql & " 	and ((i.deliveryType<>9) or ((i.deliveryType=9) and (i.sellcash>=10000)))"
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
		strSql = strSql & "		and i.itemid not in (Select itemid From db_item.dbo.tbl_lotte_regItem  where lottestatCD not in ('00','10') ) "		    '�Ե���ϻ�ǰ ����  ,'10' -- ����
		strSql = strSql & "		and cm.mapCnt is Not Null "	& addSql																				'ī�װ� ��Ī ��ǰ��
''rw strSql
		rsget.Open strSql,dbget,1

		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CLotteItem
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
                ''FItemList(i).Fcorp_dlvp_sn      = rsget("returnCode")
                FItemList(i).Fdeliverytype      = rsget("deliverytype")

                FItemList(i).FrequireMakeDay    = rsget("requireMakeDay")

                FItemList(i).FinfoDiv       = rsget("infoDiv")
                FItemList(i).Fsafetyyn      = rsget("safetyyn")
                FItemList(i).FsafetyDiv     = rsget("safetyDiv")
                FItemList(i).FsafetyNum     = rsget("safetyNum")


				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end Sub


	'--------------------------------------------------------------------------------
	'// �Ե����� ��ǰ ���(������)
	public Sub getLotteEditedItemList
		dim sqlStr, addSql, i

		if FRectItemID<>"" then
			'���û�ǰ�� �ִٸ�
			addSql = " and i.itemid in (" & FRectItemID & ")"
		elseif FRectNotJehyu="Y" then
			'���޸� ��ǰ�� �ƴѰ�
			addSql = " and i.isExtUsing='N' "
		else
			'������ ��ǰ��
			addSql = " and m.LotteLastUpdate<i.lastupdate"
		end if

        ''//���� ���ܻ�ǰ
        addSql = addSql & " and i.itemid not in ("
        addSql = addSql & "     select itemid from db_item.dbo.tbl_OutMall_etcLink"
        addSql = addSql & "     where stDt<getdate()"
        addSql = addSql & "     and edDt>getdate()"
        addSql = addSql & "     and mallid='"&CMALLNAME&"'"
        addSql = addSql & "     and linkgbn='donotEdit'"
        addSql = addSql & " )"

		strSql = "Select top " & FPageSize & " i.* "
		strSql = strSql & "		, c.keywords, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent, isNULL(c.requireMakeDay,0) as requireMakeDay "
		strSql = strSql & "		, m.LotteGoodNo, m.LotteTmpGoodNo, m.LotteSellYn, isNULL(m.regedOptCnt,0) as regedOptCnt "
		strSql = strSql & "		, m.accFailCNT, m.lastErrStr "
		''strSql = strSql & "		, B.returnCode"
		strSql = strSql & "		, C.infoDiv,isNULL(C.safetyyn,'N') as safetyyn,isNULL(C.safetyDiv,0) as safetyDiv,C.safetyNum " '' ǰ������ �� ������������. �߰� 20121102

        strSql = strSql & "		,(CASE WHEN i.isusing='N' "
		strSql = strSql & "		or i.isExtUsing='N'"
		strSql = strSql & "		or uc.isExtUsing='N'"
		strSql = strSql & "		or ((i.deliveryType=9) and (i.sellcash<10000))"
		strSql = strSql & "		or i.sellyn<>'Y'"
		strSql = strSql & "		or i.deliverfixday in ('C','X')"
		strSql = strSql & "		or i.itemdiv>=50 or i.itemdiv='08' or i.cate_large='999' or i.cate_large=''"
		strSql = strSql & "		or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "		or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "		THEN 'Y' ELSE 'N' END) as maySoldOut"

		strSql = strSql & " From db_item.dbo.tbl_item as i "
		strSql = strSql & " 	join db_item.dbo.tbl_item_contents as c "
		strSql = strSql & " 		on i.itemid=c.itemid "
		strSql = strSql & " 	join db_item.dbo.tbl_lotte_regItem as m "
		strSql = strSql & " 		on i.itemid=m.itemid "
		strSql = strSql & " 	left Join (Select tenCateLarge, tenCateMid, tenCateSmall, count(*) as mapCnt From db_item.dbo.tbl_lotte_cate_mapping Group by tenCateLarge, tenCateMid, tenCateSmall ) as cm "
		strSql = strSql & " 		on cm.tenCateLarge=i.cate_large and cm.tenCateMid=i.cate_mid and cm.tenCateSmall=i.cate_small "
		strSql = strSql & " 	left join db_user.dbo.tbl_user_c uc"
		strSql = strSql & " 		on i.makerid=uc.userid"

		''strSql = strSql & " 	left Join db_item.dbo.tbl_OutMall_BrandReturnCode B"
		''strSql = strSql & " 	    on i.makerid=B.makerid and B.mallid='"&CMALLNAME&"'"
		strSql = strSql & " Where 1=1"
		if (FRectMatchCateNotCheck<>"on") then
		    strSql = strSql & " and cm.mapCnt is Not Null "
	    end if
		strSql = strSql & addSql
		''strSql = strSql & " and m.LotteGoodNo is Not Null "									'#��� ��ǰ��
		strSql = strSql & " and isNULL(m.LotteTmpGoodNo,m.LotteGoodNo) is Not Null "									'#��� ��ǰ��
''rw strSql
		rsget.Open strSql,dbget,1

		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CLotteItem
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
				FItemList(i).FLotteGoodNo		= rsget("LotteGoodNo")
				FItemList(i).FLotteTmpGoodNo	= rsget("LotteTmpGoodNo")
				FItemList(i).FLotteSellYn		= rsget("LotteSellYn")

                FItemList(i).FoptionCnt         = rsget("optionCnt")
                FItemList(i).FregedOptCnt       = rsget("regedOptCnt")
                FItemList(i).FaccFailCNT        = rsget("accFailCNT")
                FItemList(i).FlastErrStr        = rsget("lastErrStr")
                ''FItemList(i).Fcorp_dlvp_sn      = rsget("returnCode")
                FItemList(i).Fdeliverytype      = rsget("deliverytype")
                FItemList(i).FrequireMakeDay    = rsget("requireMakeDay")

                FItemList(i).FinfoDiv       = rsget("infoDiv")
                FItemList(i).Fsafetyyn      = rsget("safetyyn")
                FItemList(i).FsafetyDiv     = rsget("safetyDiv")
                FItemList(i).FsafetyNum     = rsget("safetyNum")

                FItemList(i).FmaySoldOut    = rsget("maySoldOut")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end Sub

	Public Sub OnlyLotteNotUpdateMakeridList
		Dim sqlStr, i, addsql

		If FRectMakerId <> "" Then
			addSql = " and makerid in (" & FRectMakerId & ")"
		End If

		sqlStr = " select count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " From db_temp.dbo.tbl_Lotte_not_in_makerid_By_KimJinYoung "
		sqlStr = sqlStr & " Where 1=1 " & addsql
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'������������ ��ü ���������� Ŭ �� �Լ�����
		If Cint(FCurrPage)>Cint(FTotalPage) Then
			FResultCount = 0
			Exit sub
		End If

		sqlStr = " select  top " + CStr(FPageSize*FCurrPage) + " * "
		sqlStr = sqlStr + " from db_temp.dbo.tbl_Lotte_not_in_makerid_By_KimJinYoung "
		sqlStr = sqlStr & " Where 1=1 " & addsql
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CLotteItem
					FItemList(i).FMakerid		= rsget("makerid")
					FItemList(i).FRegdate		= rsget("regdate")
					FItemList(i).FRegId			= rsget("regId")
					FItemList(i).FLastupdate	= rsget("lastupdate")
					FItemList(i).FLastupdateId	= rsget("lastupdateId")
					FItemList(i).FIsUsing		= rsget("isUsing")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	End Sub

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

function getLotteItemIdByTenItemID(iitemid)
    dim sqlStr, retVal
    sqlStr = " select isNULL(lotteGoodNo,lotteTmpGoodNo) as lotteGoodNo "&VbCRLF
    sqlStr = sqlStr & " from db_item.dbo.tbl_lotte_regItem"&VbCRLF
    sqlStr = sqlStr & " where itemid="&iitemid&VbCRLF

    rsget.Open sqlStr,dbget,1
	if Not(rsget.EOF or rsget.BOF) then
	    retVal = rsget("lotteGoodNo")
	end if
	rsget.Close

	if IsNULL(retVal) then retVal=""
	getLotteItemIdByTenItemID = retVal
end function

''//��ǰ�� ���� �Ķ���� ����
Function fnGetLotteItemNameEditParameter(iLotteGoodNo,iItemName)
    dim strRst
    strRst = "subscriptionId=" & lotteAuthNo
    strRst = strRst & "&strGoodsNo=" & iLotteGoodNo
    strRst = strRst & "&strGoodsNm=" & Server.URLEncode(Trim(iItemName))
    strRst = strRst & "&strMblGoodsNm=" & Server.URLEncode(Trim(iItemName))
    strRst = strRst & "&strChgCausCont=" & Server.URLEncode("api ��ǰ�� ����")
    fnGetLotteItemNameEditParameter = strRst
end function

%>
